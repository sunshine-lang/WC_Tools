import csv
import hashlib
import html
import re
from collections import Counter
from io import BytesIO, StringIO
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

from .context_processor import validate_material_name
from .key_classifier import EMOTIONS, KeyClassifier


TERM_POLICY = 'wc-en-de-v1'
TERM_STATUSES = ('命中术语', '未命中', '待确认')

# Only known noun forms are expanded; names and unrelated endings are not stemmed.
NOUN_PLURALS = {
    'board': 'boards', 'item': 'items', 'reward': 'rewards', 'review': 'reviews',
    'event': 'events', 'coin': 'coins', 'buddy': 'buddies', 'safe': 'safes',
    'butterfly': 'butterflies', 'key': 'keys', 'chain': 'chains', 'bird': 'birds',
    'castle': 'castles', 'fireman': 'firemen', 'donut': 'donuts', 'bowl': 'bowls',
    'case': 'cases', 'box': 'boxes', 'hive': 'hives', 'chest': 'chests',
    'pack': 'packs', 'truck': 'trucks', 'gift': 'gifts', 'gem': 'gems',
    'frog': 'frogs', 'ghost': 'ghosts', 'jacket': 'jackets', 'scarf': 'scarves',
    'babysitter': 'babysitters', 'station': 'stations', 'lane': 'lanes',
    'clue': 'clues', 'token': 'tokens', 'investigation': 'investigations',
    'bookshelf': 'bookshelves', 'passage': 'passages', 'ring': 'rings',
    'room': 'rooms', 'suitcase': 'suitcases', 'ball': 'balls', 'trap': 'traps',
    'tripwire': 'tripwires', 'stock': 'stocks', 'lodge': 'lodges',
    'claim': 'claims', 'apiary': 'apiaries', 'company': 'companies', 'date': 'dates',
}


def normalized(text):
    return str(text).replace('’', "'").replace('‘', "'")


class TermScanner:
    def __init__(self, tb_file_path, source_name=None):
        path = Path(tb_file_path)
        self.filename = source_name or path.name
        validate_material_name(path.name)
        validate_material_name(path.resolve().name)
        validate_material_name(self.filename)
        if path.suffix.lower() not in ('.xlsx', '.csv'):
            raise ValueError('术语库须为包含 EN、DE 两列的 XLSX 或 UTF-8 CSV 文件')
        contents = path.read_bytes()
        self.sha256 = hashlib.sha256(contents).hexdigest()
        self.row_count = 0
        by_english = {}

        def add_term(english, german, source, formula=False):
            if english in (None, '') and german in (None, ''):
                return
            if formula or any(not isinstance(v, str) or not v.strip() for v in (english, german)):
                raise ValueError(f'{source} 的 EN／DE 为空、非文本或为公式，请核对术语库')
            english, german = english.strip(), german.strip()
            key = ' '.join(normalized(english).casefold().split())
            entry = by_english.setdefault(key, {'english': english, 'german': [], 'sources': []})
            if german not in entry['german']:
                entry['german'].append(german)
            entry['sources'].append(f'{self.filename} / {source}')
            self.row_count += 1

        def headers(values):
            names = [str(v).strip().casefold() if v is not None else '' for v in values]
            if any(names.count(name) > 1 for name in ('en', 'de')):
                raise ValueError('术语库的 EN／DE 表头重复，请保留各一列')
            return {name: index for index, name in enumerate(names) if name}

        if path.suffix.lower() == '.xlsx':
            workbook = load_workbook(BytesIO(contents), data_only=False)
            try:
                for sheet in workbook:
                    columns = headers(c.value for c in sheet[1])
                    if not {'en', 'de'} <= columns.keys():
                        continue
                    for row in range(2, sheet.max_row + 1):
                        en, de = (sheet.cell(row, columns[name] + 1) for name in ('en', 'de'))
                        if en.font.strike or de.font.strike:
                            continue
                        add_term(en.value, de.value, f'{sheet.title} 第{row}行（{en.coordinate}:{de.coordinate}）',
                                 formula=en.data_type == 'f' or de.data_type == 'f')
            finally:
                workbook.close()
        else:
            try:
                table = csv.reader(StringIO(contents.decode('utf-8-sig')), strict=True)
                columns = headers(next(table, []))
                if not {'en', 'de'} <= columns.keys():
                    raise ValueError('术语 CSV 须在首行提供 EN、DE 表头')
                for number, row in enumerate(table, 2):
                    if not row:
                        continue
                    values = [row[columns[name]] if columns[name] < len(row) else None for name in ('en', 'de')]
                    add_term(*values, f'CSV 第{number}行')
            except UnicodeDecodeError as error:
                raise ValueError('请将术语 CSV 保存为 UTF-8 编码后上传') from error
        self.terms = list(by_english.values())
        if not self.terms:
            raise ValueError('术语库未找到有效 EN／DE 词条，请检查首行表头')

        self.conflict_count = sum(len(term['german']) > 1 for term in self.terms)
        self.patterns = []
        for index, term in enumerate(self.terms):
            words = normalized(term['english']).split()
            forms = [(' '.join(words), '原词')]
            last = words[-1].casefold()
            for singular, plural in NOUN_PLURALS.items():
                if last in (singular, plural):
                    alternate = plural if last == singular else singular
                    forms.append((' '.join([*words[:-1], alternate]), '词形'))
                    break
            for form, kind in forms:
                body = r'\s+'.join(re.escape(word) for word in form.split())
                self.patterns.append((index, kind, re.compile(r'(?<!\w)' + body + r'(?!\w)', re.IGNORECASE)))

    def status(self):
        return {'ready': True, 'filename': self.filename, 'row_count': self.row_count,
                'term_count': len(self.terms), 'conflict_count': self.conflict_count,
                'sha256': self.sha256}

    def find_matches(self, text):
        if not isinstance(text, str) or not text.strip():
            return []
        visible = re.sub(r'<!--[\s\S]*?-->', '', html.unescape(text))
        # Inline formatting joins visible text; sprite objects separate it.
        visible = re.sub(r'<sprite\b[^>]*>', ' \ufffc ', visible, flags=re.IGNORECASE)
        visible = re.sub(r'</?[A-Za-z][^>]*>', '', visible)
        visible = normalized(visible)
        candidates = []
        for index, kind, pattern in self.patterns:
            for match in pattern.finditer(visible):
                candidates.append((match.start(), match.end(), index, kind, match.group()))
        candidates.sort(key=lambda m: (-(m[1] - m[0]), m[3] != '原词', m[2], m[0]))
        occupied, selected = [], []
        for start, end, index, kind, matched in candidates:
            if any(start < right and end > left for left, right in occupied):
                continue
            occupied.append((start, end))
            selected.append((start, index, kind, matched))
        result, seen = [], set()
        for _, index, kind, matched in sorted(selected):
            if index in seen:
                continue
            seen.add(index)
            result.append({**self.terms[index], 'matched_text': matched, 'match_type': kind})
        return result

    @staticmethod
    def format_matches(matches):
        return ' | '.join(f"{m['english']} ({' / '.join(m['german'])})" for m in matches)

    def scan_text(self, text):
        return self.format_matches(self.find_matches(text))

    def process(self, trans_file_path, output_file, task_dict, task_id):
        task = task_dict[task_id]
        workbook = None
        try:
            path = Path(trans_file_path)
            validate_material_name(path.name)
            validate_material_name(path.resolve().name)
            validate_material_name(task.get('input_name', ''))
            if path.suffix.lower() != '.xlsx':
                raise ValueError('请上传第二步下载的 XLSX 文件')
            task.update(progress=5, status='正在读取英文原文和英德术语库...')
            workbook = load_workbook(path)
            records = []
            for sheet in workbook:
                headers = {str(c.value).strip().casefold(): c.column for c in sheet[1] if c.value is not None}
                en_col = headers.get('game_text_en') or headers.get('en')
                if not en_col:
                    continue
                if not {'分类', 'context'} <= headers.keys():
                    raise ValueError(f'“{sheet.title}”缺少分类或 Context 列，请上传第二步下载的结果')
                if '关联术语' in headers:
                    raise ValueError('文件已有“关联术语”列，请上传第二步结果，避免覆盖旧数据')
                width = max(c.column for row in sheet for c in row if c.value is not None)
                column = width + 1
                cell = sheet.cell(1, column, '关联术语')
                cell.font = Font(bold=True, color='FFFFFF')
                cell.fill = PatternFill('solid', fgColor='334155')
                sheet.column_dimensions[cell.column_letter].width = 65

                def value(row, name):
                    col = headers.get(name)
                    raw = sheet.cell(row, col).value if col else None
                    return '' if raw is None else str(raw)

                for row in range(2, sheet.max_row + 1):
                    if not any(sheet.cell(row, col).value is not None for col in range(1, width + 1)):
                        continue
                    source = sheet.cell(row, en_col)
                    english = value(row, 'game_text_en') if 'game_text_en' in headers else value(row, 'en')
                    warning = ''
                    matches = []
                    if source.data_type == 'f' or not isinstance(source.value, str) or not english.strip():
                        warning = '英文源文为空、非文本或为公式，请先核对源文'
                    else:
                        matches = self.find_matches(english)
                        if any(len(m['german']) > 1 for m in matches):
                            warning = '同一英文存在多个德语候选，请结合语境确认'
                    result = self.format_matches(matches)
                    if len(result) > 32767:
                        raise ValueError(f'{sheet.title} 第{row}行术语参考超过 Excel 单元格上限')
                    cell = sheet.cell(row, column, result)
                    cell.data_type = 's'
                    cell.alignment = Alignment(vertical='top', wrap_text=True)
                    key = value(row, 'key')
                    speaker = value(row, 'speaker')
                    emotion = value(row, 'game_text_cn') or value(row, 'cn')
                    if not speaker and emotion.strip().casefold() in EMOTIONS:
                        speaker, key = key, ''
                    records.append({'sheet': sheet.title, 'row': row, 'key': key, 'speaker': speaker,
                                    'english': english, 'category': value(row, '分类'),
                                    'status': '待确认' if warning else '命中术语' if matches else '未命中',
                                    'terms': result, 'matches': matches, 'warning': warning})
                task.update(progress=80, status=f'已扫描 {len(records)} 条，正在保存 Excel...')
            if not records:
                raise ValueError('未找到待处理记录，须包含英文源文、分类和 Context 列')
            KeyClassifier.save_workbook(workbook, output_file)
            counts = Counter(r['status'] for r in records)
            task.update(progress=100, state='completed', result_file=str(output_file),
                        records=records, counts={s: counts[s] for s in TERM_STATUSES},
                        term_db=self.status(), term_policy=TERM_POLICY,
                        status=f'匹配完成，共 {len(records)} 条；命中 {counts["命中术语"]} 条，未命中 {counts["未命中"]} 条，待确认 {counts["待确认"]} 条')
        except Exception as error:
            task.update(status=f'术语关联失败：{error}', state='error')
        finally:
            if workbook:
                workbook.close()
