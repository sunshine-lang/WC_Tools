import hashlib
import re
from io import BytesIO
from collections import Counter, defaultdict
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from PIL import Image as PillowImage
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

from .key_classifier import CATEGORIES, EMOTIONS, PENDING, KeyClassifier
from .dialogue_context import DIALOGUE_CONTEXT_POLICY, DialogueContextAnalyzer


MATERIAL_DIR = Path(__file__).resolve().parents[2]
DEFAULT_REFERENCE = MATERIAL_DIR / 'WC_UI文案翻译-1016.xlsx'
DEFAULT_GUIDE = MATERIAL_DIR / '柠檬微趣_WC指导材料.xlsx'
CONTEXT_POLICY = 'wc-ui-compact-v2'
STATUSES = ('已关联', '缺资料', '待确认')
DIALOGUE_KEY = re.compile(r'mdnew_(.+)_mdnew_decoration_(intro|\d+)_(\d+)', re.IGNORECASE)


def text(value):
    return '' if value is None else str(value)


def dialogue_identity(key, speaker, note):
    """兼容客户把完整对白 Key 放在 note、说话人放在 key 的表格。"""
    key, speaker, note = (text(value).strip() for value in (key, speaker, note))
    if not DIALOGUE_KEY.fullmatch(key) and DIALOGUE_KEY.fullmatch(note):
        if not key or key.casefold() == speaker.casefold() or (
                not speaker and re.fullmatch(r"[A-Za-z][A-Za-z .'-]{0,60}", key)):
            return note, speaker or key
    return key, speaker


def validate_material_name(name):
    normalized = text(name).replace('\\', '/').rsplit('/', 1)[-1]
    if re.search(r'done(?:\s*\(\d+\))?$', Path(normalized).stem, re.IGNORECASE):
        raise ValueError('不能使用 done 结尾的材料，请选择原始待译文件或原始指导材料')


class ContextProcessor:
    def __init__(self, reference_path=DEFAULT_REFERENCE, guide_path=DEFAULT_GUIDE, materials_dir=MATERIAL_DIR):
        self.reference_path = Path(reference_path)
        self.guide_path = Path(guide_path)
        self.materials_dir = Path(materials_dir).resolve()

    def resolve_material(self, name):
        validate_material_name(name)
        candidate = (self.materials_dir / name).resolve()
        validate_material_name(candidate.name)
        if candidate.parent != self.materials_dir or candidate.suffix.lower() != '.xlsx' or candidate.name.startswith(('.', '~$')):
            raise ValueError('参考材料只能选择 WC-UI 目录中的 XLSX 文件，不可使用工具缓存或外部目录')
        return candidate

    def material_options(self):
        result = []
        for path in sorted(self.materials_dir.glob('*.xlsx')):
            try:
                self.resolve_material(path.name)
            except ValueError:
                continue
            result.append(path.name)
        return result

    def material_status(self):
        return [{'name': path.name, 'ready': path.is_file(), 'kind': kind}
                for kind, path in [('本目录待译材料', self.reference_path), ('人设与背景', self.guide_path)]]

    @staticmethod
    def read_records(workbook, source_name, classified=False):
        records = []
        for sheet in workbook:
            headers = {text(c.value).strip().casefold(): c.column for c in sheet[1] if c.value is not None}
            en_col = headers.get('game_text_en') or headers.get('en')
            if not en_col:
                continue
            if classified and '分类' not in headers:
                raise ValueError(f'“{sheet.title}”缺少“分类”列，请先完成第一步')
            if classified and 'context' in headers:
                raise ValueError('文件已有 Context 列，请上传第一步的分类结果')

            def value(row, *names):
                col = next((headers[name] for name in names if name in headers), None)
                return text(sheet.cell(row, col).value) if col else ''

            width = max(c.column for row in sheet for c in row if c.value is not None)
            for row in range(2, sheet.max_row + 1):
                english = value(row, 'game_text_en', 'en')
                if not english and not (classified and value(row, '分类')):
                    continue
                key = value(row, 'key').strip()
                speaker = value(row, 'speaker').strip()
                emotion = value(row, 'animation').strip()
                chinese = value(row, 'game_text_cn', 'cn')
                note = value(row, 'note', '备注')
                if not speaker and chinese.strip().casefold() in EMOTIONS and re.fullmatch(r"[A-Za-z][A-Za-z .'-]{0,60}", key):
                    speaker, emotion, chinese = key, chinese.strip(), ''
                    # 原始混合需求表把对白 Key 放在 note 列，工作副本中则没有。
                    key = note.strip() if re.match(r'^(mdnew_|story|Fire)', note.strip()) else ''
                    if key:
                        note = ''
                key, speaker = dialogue_identity(key, speaker, note)
                if not emotion and chinese.strip().casefold() in EMOTIONS and speaker:
                    emotion = chinese.strip()
                bubble = value(row, 'bubbletype', 'bubble_type').strip()
                if re.fullmatch(r'\d+\.0+', bubble):
                    bubble = bubble.split('.')[0]
                records.append({
                    'id': f'{sheet.title}:{row}', 'sheet': sheet.title, 'row': row, 'width': width,
                    'source': f'{source_name} / {sheet.title}!{row}', 'key': key,
                    'english': english, 'speaker': speaker, 'emotion': emotion,
                    'bubble': bubble, 'chinese': chinese,
                    'note': note.strip(), 'category': value(row, '分类').strip(),
                    'formula': sheet.cell(row, en_col).data_type == 'f',
                })
        return records

    @staticmethod
    def relation(record, activity_hint=''):
        key = record['key'].casefold()
        card = re.fullmatch(r'a(.+)_set(\d+)(?:_(\d+))?_name', key)
        decor = re.fullmatch(r'mdnew_(.+)_mdnew_decoration_name_(\d+)', key)
        dialogue = DIALOGUE_KEY.fullmatch(key)
        if card:
            activity, number, item = card.groups()
            return activity, f'{activity}/set/{int(number)}', f'卡册{int(number)}', int(item or 0), ''
        if decor:
            activity, number = decor.groups()
            return activity, f'{activity}/decor/{int(number)}', f'装饰物{int(number)}', int(number), str(int(number))
        if dialogue:
            activity, number, line = dialogue.groups()
            label = '活动开场' if number == 'intro' else f'装饰物{int(number)}对白'
            return activity, f'{activity}/dialogue/{number}', label, int(line), number
        activity = ''
        if key.startswith(('cat_multiboards_', 'catmultiboards_')) or key == 'offer_cat':
            activity = 'cat'
        elif key.startswith('midautumn_multiboards_') or key == 'offer_midautumn':
            activity = 'midautumn'
        elif key.startswith('album_'):
            activity = key.split('_')[1]
        elif key.startswith('event_') and '_mdnew_' in key:
            activity = key.split('_mdnew_')[0][6:]
        elif key.startswith('chain_gift_title_chain_'):
            activity = key[len('chain_gift_title_chain_'):]
        elif key.startswith('forestconcert_'):
            activity = 'forestconcert'
        if not key:
            activity = activity_hint
        if not activity or record.get('category') == CATEGORIES[3]:
            return '', '', '', 0, ''
        module = '商店宣传' if not key else '活动UI'
        if re.match(r'^forestconcert_(road|rank)_', key):
            module = '进度奖励' if '_road_' in key else '排行榜'
        return activity, f'{activity}/ui/{module}', module, record['row'], ''

    def prepare_relations(self, records):
        hints = defaultdict(set)
        for record in records:
            activity = self.relation(record)[0]
            if activity:
                hints[record['sheet']].add(activity)
        for record in records:
            choices = hints[record['sheet']]
            hint = next(iter(choices)) if len(choices) == 1 else ''
            activity, group, label, order, number = self.relation(record, hint)
            record.update(activity=activity, group=group, label=label, order=order, number=number)

    @staticmethod
    def load_roles(path, source_name=None):
        if not path.is_file():
            return {}
        workbook = load_workbook(path, data_only=True)
        roles = {}
        if '人设表' in workbook:
            sheet = workbook['人设表']
            for row in range(3, sheet.max_row + 1):
                names = text(sheet.cell(row, 4).value).splitlines()
                if not names:
                    continue
                description = text(sheet.cell(row, 7).value)
                # 仅引用开头两句人物事实，避免混入后续剧透和其他语种列。
                facts = ''.join(re.split(r'(?<=[。！？])', description)[:2]).strip()
                age = text(sheet.cell(row, 5).value).strip()
                role = {'facts': (f'年龄：{age}；' if age and age != '/' else '') + facts,
                        'source': f'{source_name or path.name} / 人设表!G{row}'}
                for name in names:
                    if name.strip():
                        roles[name.strip().casefold()] = role
        workbook.close()
        return roles

    @staticmethod
    def find_reference(record, by_key, by_english):
        if record['key']:
            candidates = by_key.get(record['key'], [])
            if candidates:
                matching = [r for r in candidates if r['english'].strip() == record['english'].strip()]
                if not matching:
                    return None, '待确认', '相同 Key 的参考英文与本条不一致，请核对版本'
                candidates = matching
            else:
                return None, '缺资料', '参考文件未找到此 Key'
        else:
            candidates = by_english.get(record['english'].strip(), [])
        if record['speaker']:
            candidates = [r for r in candidates if r['speaker'] == record['speaker']
                          and (not record['emotion'] or r['emotion'] == record['emotion'])]
        if len(candidates) != 1:
            return None, '待确认' if candidates else '缺资料', (
                f'找到 {len(candidates)} 条参考候选，无法唯一定位' if candidates else '未找到源文及角色一致的参考记录')
        return candidates[0], '已关联', ''

    @staticmethod
    def extract_images(workbook, references, asset_dir, source_name):
        asset_dir.mkdir(parents=True, exist_ok=True)
        images, warnings = [], []
        for sheet in workbook:
            sheet_records = [r for r in references if r['sheet'] == sheet.title]
            activities = {r['activity'] for r in sheet_records if r['activity']}
            for index, image in enumerate(sheet._images):
                anchor = getattr(image.anchor, '_from', None)
                end = getattr(image.anchor, 'to', None)
                if anchor is None:
                    warnings.append(f'{sheet.title} 图片{index + 1}没有可识别锚点')
                    continue
                start_row = anchor.row + 1
                end_row = max(start_row, end.row if end else start_row)
                covered = [r for r in sheet_records if start_row <= r['row'] <= end_row]
                card_groups = {r['group'] for r in covered if '/set/' in r['group']}
                group = next(iter(card_groups)) if len(card_groups) == 1 else ''
                activity = next(iter(activities)) if len(activities) == 1 else ''
                if not activity:
                    covered_activities = {r['activity'] for r in covered if r['activity']}
                    activity = next(iter(covered_activities)) if len(covered_activities) == 1 else ''
                data = image._data()
                with PillowImage.open(BytesIO(data)) as picture:
                    extension = picture.format.lower()
                if extension not in ('png', 'jpeg', 'jpg', 'gif', 'bmp'):
                    warnings.append(f'{sheet.title} 图片{index + 1}格式暂不支持')
                    continue
                name = hashlib.sha256(data).hexdigest()[:16] + '.' + extension
                (asset_dir / name).write_bytes(data)
                caption = f'{source_name} / {sheet.title} 图片{index + 1}（锚点第{start_row}行）'
                if group:
                    scope = '套组参考：图片覆盖行的 Key 属于同一卡册，未逐张裁切'
                elif activity:
                    scope = '活动级参考，未定位到单个物件'
                else:
                    scope = '图片归属待确认'
                    warnings.append(caption + '：归属待确认')
                images.append({'name': name, 'path': f'images/{name}', 'caption': caption,
                               'scope': scope, 'activity': activity, 'group': group})
        return images, warnings

    @staticmethod
    def ui_unit(record):
        key = record['key'].casefold().replace('catmultiboards_', 'cat_multiboards_')
        if '_book_' in key:
            return key.split('_book_', 1)[0] + '/book'
        if '_guide_' in key:
            return key.split('_guide_', 1)[0] + '/intro'
        if key.endswith('_pre_desc'):
            return key.removesuffix('_pre_desc') + '/intro'
        if key.endswith('_title') and (key.removesuffix('_title').endswith('_multiboards') or key.startswith('album_')):
            return key.removesuffix('_title') + '/intro'
        rank = re.fullmatch(r'(.+_rank)_(title|score_need|start_desc)', key)
        if rank:
            return rank.group(1) + '/entry'
        return re.sub(r'_(?:title|desc\d*)$', '', key)

    def related_records(self, record, groups, activities):
        peers = groups.get(record['group'], [])
        activity_rows = activities.get(record['activity'], [])
        overview = [r for r in activity_rows if '/ui/' in r['group']
                    and r['key'].casefold().endswith(('_pre_desc', '_prediction_desc'))]
        if '/set/' in record['group'] or '/dialogue/' in record['group']:
            selected = list(peers)
        elif '/decor/' in record['group']:
            dialogue = groups.get(f"{record['activity']}/dialogue/{record['number']}", [])
            selected = [*peers, *dialogue]
            if not dialogue and overview:
                intro = overview[0]
                title = intro['key'].removesuffix('_pre_desc').removesuffix('_desc') + '_title'
                selected.extend(r for r in activity_rows if r['key'] == title)
                selected.append(intro)
        else:
            selected = [r for r in activity_rows if '/ui/' in r['group'] and self.ui_unit(r) == self.ui_unit(record)]
            if len(selected) < 2:
                titles = [r for r in peers if r['key'].casefold().endswith('_title')
                          and not any(part in r['key'].casefold() for part in ('_complete_', '_success_', '_fail_', '_book_'))]
                anchors = overview if record['key'].casefold().startswith(('offer_', 'chain_gift_')) else titles + overview
                anchor = next((r for r in anchors if r['id'] != record['id'] and r['english'] != record['english']), None)
                if anchor:
                    selected.append(anchor)
        selected = selected or [record]
        if '/dialogue/' in record['group']:
            selected.extend(groups.get(f"{record['activity']}/decor/{record['number']}", []))
        unique = {r['id']: r for r in selected}
        return list(unique.values())

    @staticmethod
    def line(record):
        label = record['key'] or record['speaker'] or '无 Key'
        details = [record['speaker'], record['emotion']]
        if record['bubble']:
            details.append('思考' if record['bubble'] == '2' else '正常对话' if record['bubble'] == '1' else '气泡' + record['bubble'])
        details = [item for item in details if item]
        return label + (' (' + ' / '.join(details) + ')' if details else '') + ': ' + record['english']

    def build_context(self, record, reference, status, warning, groups, activities, roles, images, enhanced=False):
        effective = dict(reference or record)
        effective['english'] = record['english']
        warnings = [warning] if warning else []
        if not reference:
            self.prepare_relations([effective])
        if record['category'] == PENDING or record['category'] not in CATEGORIES:
            status, warnings = '待确认', ['分类尚未确认，请先核对第一步结果']
        if record['formula'] or not record['english'].strip():
            status, warnings = '待确认', ['英文源文为空或为公式，请先核对源文']
        if not effective.get('group'):
            message = ('缺少完整对白 Key，不能按说话人或行号推断分段' if record['speaker']
                       else '无 Key 或活动归属未明确，无法可靠分组')
            status = '待确认'
            warnings = warnings or [message]
        else:
            group = effective['group']
            expected = CATEGORIES[1] if '/set/' in group else CATEGORIES[2] if '/decor/' in group else CATEGORIES[3] if '/dialogue/' in group else CATEGORIES[0]
            if record['category'] in CATEGORIES and record['category'] != expected:
                status, warnings = '待确认', ['分类与参考 Key 结构不一致，请先核对分类']
        effective['warning'] = '；'.join(warnings)
        if status == '待确认':
            return status, '', [], effective

        selected = self.related_records(effective, groups, activities)
        if effective['group'] == 'forestconcert/dialogue/12':
            # 只引用上传文件内已核对的前段线索，不用模型补写。
            prior = [r for r in activities.get(effective['activity'], [])
                     if r['key'] == 'mdnew_forestconcert_mdnew_decoration_5_10']
            if any('treble' in r['english'].casefold() for r in selected) and prior and 'treble clef' in prior[0]['english'].casefold():
                selected.extend(prior)
            else:
                warnings.append('未找到或无法验证第5段的高音谱号解释')
        selected = list({r['id']: r for r in selected}.values())
        selected_images = [im for im in images if im['group'] == effective['group']
                           or (not im['group'] and im['activity'] == effective['activity'])]
        role_lines = []
        if enhanced:
            speakers = list(dict.fromkeys(r['speaker'] for r in selected if r['speaker']))
            for speaker in speakers:
                role = roles.get(speaker.casefold())
                if role and role['facts']:
                    role_lines.append(speaker + ': ' + role['facts'])
                else:
                    warnings.append(f'未找到 {speaker} 的人设')
            if effective['speaker'] and not effective['bubble']:
                warnings.append('原始材料未提供气泡类型，不能推断')
            if record['category'] in (CATEGORIES[1], CATEGORIES[2]) and not selected_images:
                warnings.append('原始材料未提供可确定归属的图片')
        useful = len(selected) > 1 or role_lines or selected_images
        if not useful:
            warnings.append('未找到可用关联原文，Context 留空')
        if warnings:
            status = '缺资料'
        lines = ['[关联语境参考]', *(self.line(r) for r in selected)] if useful else []
        if role_lines:
            lines.extend(['[人物参考]', *role_lines])
        if selected_images:
            lines.append('[配图] ' + '、'.join(dict.fromkeys(im['path'] for im in selected_images)))
        effective['warning'] = '；'.join(dict.fromkeys(warnings))
        effective['related_count'] = len(selected) if useful else 0
        effective['related_records'] = selected
        return status, '\n'.join(lines), selected_images, effective

    def process(self, input_file, output_file, task_dict, task_id, reference_path=None, guide_path=None, enhance=False,
                analyze_dialogue=False, config_path=None, materials=None, scene_cache_dir=None):
        task = task_dict[task_id]
        workbook = reference_workbook = None
        try:
            task['status'] = '正在读取分类结果...'
            reference_path = Path(reference_path or self.reference_path)
            guide_path = Path(guide_path or self.guide_path)
            for path, name in [(Path(input_file), task.get('input_name', '')),
                               (reference_path, task.get('reference_name', '')),
                               (guide_path, task.get('guide_name', ''))]:
                validate_material_name(path.name)
                validate_material_name(path.resolve().name)
                validate_material_name(name)
            if enhance:
                reference_path = self.resolve_material(str(reference_path))
                guide_path = self.resolve_material(str(guide_path))
            workbook = load_workbook(input_file)
            records = self.read_records(workbook, task.get('input_name', Path(input_file).name), classified=True)
            if not records:
                raise ValueError('未找到待处理记录，须包含英文源文列和“分类”列')
            references, roles = [], {}
            if enhance:
                if not reference_path.is_file():
                    raise ValueError('所选本目录材料不存在，请重新选择或关闭可选补充')
                reference_workbook = load_workbook(reference_path, data_only=True)
                references = self.read_records(reference_workbook, task.get('reference_name', reference_path.name))
                if not references:
                    raise ValueError('原始参考未找到 game_text_en 或 en 英文源文列')
                roles = self.load_roles(guide_path, task.get('guide_name'))
            by_key, by_english = defaultdict(list), defaultdict(list)
            for r in references:
                if r['key']:
                    by_key[r['key']].append(r)
                by_english[r['english'].strip()].append(r)
            effective_records, matches = [], {}
            for record in records:
                effective = dict(record)
                status, warning = '已关联', ''
                if enhance:
                    match, status, warning = self.find_reference(record, by_key, by_english)
                    if match:
                        effective.update({name: match[name] for name in ('key', 'speaker', 'emotion', 'bubble', 'note')})
                        if not effective['chinese']:
                            effective['chinese'] = match['chinese']
                        effective['reference_source'] = match['source']
                effective_records.append(effective)
                matches[record['id']] = (effective, status, warning)
            self.prepare_relations(effective_records)
            groups, activities = defaultdict(list), defaultdict(list)
            for r in effective_records:
                status = matches[r['id']][1]
                if r['group'] and status != '待确认' and r['category'] in CATEGORIES and not r['formula']:
                    groups[r['group']].append(r)
                    activities[r['activity']].append(r)
            for peers in groups.values():
                peers.sort(key=lambda r: (r['order'], r['row']))
            output = Path(output_file)
            asset_dir = output.with_name(output.stem + '_assets') / 'images'
            images, image_warnings = [], []
            if enhance:
                self.prepare_relations(references)
                images, image_warnings = self.extract_images(reference_workbook, references, asset_dir,
                                                             task.get('reference_name', reference_path.name))
            task.update(progress=30, status='正在按活动、套组和对白段关联...')
            result = []
            context_columns = {}
            scene_results, blocked_scenes = {}, {}
            analyzer = None
            for record in records:
                sheet = workbook[record['sheet']]
                if sheet.title not in context_columns:
                    column = record['width'] + 1
                    context_columns[sheet.title] = column
                    header = sheet.cell(1, column, 'Context')
                    header.font = Font(bold=True, color='FFFFFF')
                    header.fill = PatternFill('solid', fgColor='334155')
                    sheet.column_dimensions[header.column_letter].width = 90
                effective, status, warning = matches[record['id']]
                status, context, selected, effective = self.build_context(record, effective, status, warning,
                                                                        groups, activities, roles, images, enhanced=enhance)
                group = effective.get('group', '')
                if analyze_dialogue and record['category'] == CATEGORIES[3] and group and status != '待确认':
                    if group not in scene_results and group not in blocked_scenes:
                        peers = sorted((r for r in effective_records if r['group'] == group), key=lambda r: r['order'])
                        if any(r['formula'] or not r['english'].strip() or r['category'] != CATEGORIES[3]
                               or matches[r['id']][1] == '待确认' for r in peers):
                            blocked_scenes[group] = '同段存在异常源文、分类或版本冲突，场景分析留空，请核对后重跑'
                        elif len({r['key'].casefold() for r in peers}) != len(peers) or len({r['order'] for r in peers}) != len(peers):
                            blocked_scenes[group] = '同段完整 Key 重复，无法确定对白顺序，场景分析留空'
                        else:
                            if analyzer is None:
                                analyzer = DialogueContextAnalyzer(config_path, materials=materials, cache_dir=scene_cache_dir)
                            input_warnings = []
                            if any(not r['bubble'] for r in peers):
                                input_warnings.append('本组存在未提供气泡类型的台词，交流媒介或内心独白只能依据明确原文判断')
                            orders = [r['order'] for r in peers]
                            if orders[0] > 1 or any(b - a != 1 for a, b in zip(orders, orders[1:])):
                                input_warnings.append('本批对白句序存在缺号，仅分析已提供原文，不补写缺句')
                            references = [r for r in effective.get('related_records', []) if r['group'] != group]
                            task['status'] = f'正在通读对白场景 {group}（{len(peers)} 句）...'
                            try:
                                scene_results[group] = analyzer.analyze(group, peers, references, input_warnings)
                                task['dialogue_scenes'] = list(scene_results.values())
                            except ValueError as error:
                                raise ValueError(f'对白场景 {group}：{error}') from error
                    if group in blocked_scenes:
                        status, context = '待确认', ''
                        effective['warning'] = blocked_scenes[group]
                    else:
                        scene = scene_results[group]
                        context = scene['context']
                        prior_warning = effective.get('warning', '').replace('未找到可用关联原文，Context 留空', '').strip('；')
                        effective['warning'] = '；'.join(filter(None, [prior_warning, *scene['warnings']]))
                        effective['related_count'] = len(scene['sources']) + len(scene['references'])
                        status = '缺资料' if effective['warning'] else '已关联'
                if len(context) > 32767:
                    raise ValueError(f'{sheet.title} 第{record["row"]}行 Context 超过 Excel 单元格上限，请缩小关联范围')
                cell = sheet.cell(record['row'], context_columns[sheet.title], context)
                cell.data_type = 's'
                cell.alignment = Alignment(vertical='top', wrap_text=True)
                result.append({'id': record['id'], 'sheet': record['sheet'], 'row': record['row'],
                               'key': effective['key'], 'speaker': effective['speaker'],
                               'english': record['english'], 'category': record['category'],
                               'group': effective.get('group', ''), 'status': status, 'context': context,
                               'images': selected, 'warning': effective.get('warning', ''),
                               'source': record['source'], 'reference_source': effective.get('reference_source', ''),
                               'related_count': effective.get('related_count', 0)})
                task['progress'] = 30 + int(60 * len(result) / len(records))
            KeyClassifier.save_workbook(workbook, output_file)
            bundle_file = None
            if images:
                bundle_file = str(output.with_suffix('.zip'))
                with ZipFile(bundle_file, 'w', ZIP_DEFLATED) as archive:
                    archive.write(output, task.get('dl_name', 'WC_context.xlsx'))
                    for name in {im['name'] for im in images}:
                        archive.write(asset_dir / name, f'images/{name}')
                    manifest = '\n'.join(f"{im['path']}\n{im['caption']}\n{im['scope']}\n" for im in images)
                    archive.writestr('图片来源.txt', manifest)
            counts = Counter(r['status'] for r in result)
            task.update(progress=100, state='completed', status=f'关联完成，共 {len(result)} 条；待确认 {counts["待确认"]} 条，缺资料 {counts["缺资料"]} 条',
                        result_file=output_file, records=result, counts={s: counts[s] for s in STATUSES},
                        bundle_file=bundle_file, image_warnings=image_warnings,
                        asset_paths={im['name']: str(asset_dir / im['name']) for im in images},
                        reference_name=task.get('reference_name', reference_path.name) if enhance else '未启用可选补充',
                        enhanced=enhance, context_policy=DIALOGUE_CONTEXT_POLICY if analyze_dialogue else CONTEXT_POLICY,
                        dialogue_scenes=list(scene_results.values()),
                        dialogue_cached_scenes=sum(s.get('cache_hit', False) for s in scene_results.values()),
                        dialogue_model_requests=sum(not s.get('cache_hit', False) for s in scene_results.values()))
        except Exception as error:
            task.update(status=f'关联失败：{error}', state='error')
        finally:
            if workbook:
                workbook.close()
            if reference_workbook:
                reference_workbook.close()
