import csv
import fnmatch
import hashlib
import json
import os
import re
import uuid
from collections import Counter
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill


CATEGORIES = ('活动 UI 与商店文案', '集卡名称', '装饰物名称', '活动对白')
PENDING = '待确认'
META_HEADERS = ('分类',)
EMOTIONS = {
    'neutral', 'happy', 'thinking', 'confused', 'investigating', 'proud', 'laughing',
    'eureka', 'annoyed', 'shocked', 'surprised', 'fear', 'angry', 'sad', 'flirting',
    'blushing', 'crying', 'worried', 'normal',
}


class KeyClassifier:
    def __init__(self, config_path):
        self.config_path = config_path

    def load_rules(self):
        with open(self.config_path, encoding='utf-8') as file:
            rules = json.load(file)
        self.validate_rules(rules)
        return rules

    @staticmethod
    def validate_rules(rules):
        if not isinstance(rules, dict) or set(rules) != set(CATEGORIES):
            raise ValueError('请保留 WC 的四种分类')
        for patterns in rules.values():
            if not isinstance(patterns, list) or len(patterns) > 50:
                raise ValueError('每类最多配置 50 条特征规则')
            if any(not isinstance(p, str) or not p.strip() or len(p) > 120 for p in patterns):
                raise ValueError('特征规则须为 1–120 个字符的非空文本')

    def save_rules(self, rules):
        self.validate_rules(rules)
        rules = {category: list(dict.fromkeys(p.strip() for p in rules[category])) for category in CATEGORIES}
        temp = f'{self.config_path}.{uuid.uuid4().hex}.tmp'
        try:
            with open(temp, 'w', encoding='utf-8') as file:
                json.dump(rules, file, ensure_ascii=False, indent=2)
            os.replace(temp, self.config_path)
        finally:
            if os.path.exists(temp):
                os.remove(temp)

    @staticmethod
    def text(value):
        return '' if value is None else str(value)

    def classify(self, key, english, emotion, speaker, rules):
        if not english.strip():
            return PENDING, '缺少英文源文'
        if speaker.strip() or (
            emotion.casefold() in EMOTIONS
            and re.fullmatch(r"[A-Za-z][A-Za-z .'-]{0,60}", key)
        ):
            return CATEGORIES[3], '识别到说话人和情绪字段' if not speaker else '识别到 Speaker 字段'

        def match(category):
            for pattern in rules[category]:
                p = pattern.casefold()
                if (fnmatch.fnmatchcase(key.casefold(), p) if any(c in p for c in '*?[') else p in key.casefold()):
                    return pattern
            return None

        hits = [(category, match(category)) for category in CATEGORIES[1:]]
        hits = [(category, pattern) for category, pattern in hits if pattern]
        if len(hits) > 1:
            return PENDING, '同时命中多类规则：' + '、'.join(category for category, _ in hits)
        if hits:
            category, pattern = hits[0]
            return category, f'Key 命中规则：{pattern}'
        pattern = match(CATEGORIES[0])
        if pattern:
            return CATEGORIES[0], f'Key 命中规则：{pattern}'
        if not key:
            return CATEGORIES[0], '无 Key 文案，暂按活动 UI 归类，请确认用途'
        return PENDING, '未命中分类规则'

    @staticmethod
    def counts(records):
        counts = Counter(record['category'] for record in records)
        return {category: counts[category] for category in (*CATEGORIES, PENDING)}

    @staticmethod
    def save_workbook(workbook, path):
        temp = f'{path}.{uuid.uuid4().hex}.tmp.xlsx'
        try:
            workbook.save(temp)
            os.replace(temp, path)
        finally:
            if os.path.exists(temp):
                os.remove(temp)

    def process(self, input_file, output_file, task_dict, task_id):
        task = task_dict[task_id]
        try:
            task['status'] = '正在读取文件...'
            rules = self.load_rules()
            if Path(input_file).suffix.lower() == '.csv':
                with open(input_file, encoding='utf-8-sig', newline='') as file:
                    data = list(csv.reader(file))
                workbook = Workbook()
                sheet = workbook.active
                for row in data:
                    sheet.append(row)
                for row in sheet:
                    for cell in row:
                        cell.data_type = 's'
            else:
                workbook = load_workbook(input_file)
            fingerprint = hashlib.sha256(Path(input_file).read_bytes()).hexdigest()[:12]
            records = []
            for sheet_index, sheet in enumerate(workbook.worksheets):
                headers = {self.text(c.value).strip().casefold(): c.column for c in sheet[1] if c.value is not None}
                en_col = headers.get('game_text_en') or headers.get('en')
                if not en_col:
                    continue
                if any(name.casefold() in headers for name in (*META_HEADERS, '记录ID', '文本类型', '分类依据', '检查提示')):
                    raise ValueError('文件已有分类结果列，请上传原始待译文件')
                width = max(c.column for row in sheet for c in row if c.value is not None)
                columns = {name: width + i for i, name in enumerate(META_HEADERS, 1)}
                for name, col in columns.items():
                    cell = sheet.cell(1, col, name)
                    cell.font = Font(bold=True, color='FFFFFF')
                    cell.fill = PatternFill('solid', fgColor='334155')
                    sheet.column_dimensions[cell.column_letter].width = 30

                def value(row, *names):
                    col = next((headers[name] for name in names if name in headers), None)
                    return self.text(sheet.cell(row, col).value).strip() if col else ''

                seen_keys = {}
                for row in range(2, sheet.max_row + 1):
                    if not any(sheet.cell(row, c).value is not None for c in range(1, width + 1)):
                        continue
                    raw_key = value(row, 'key')
                    emotion = value(row, 'animation') or value(row, 'game_text_cn', 'cn')
                    speaker = value(row, 'speaker')
                    source_cell = sheet.cell(row, en_col)
                    english = self.text(source_cell.value)
                    category, basis = self.classify(raw_key, english, emotion, speaker, rules)
                    warnings = []
                    if source_cell.data_type == 'f':
                        category, basis = PENDING, '英文单元格为公式，请核对实际源文'
                    if not english.strip():
                        warnings.append('缺少英文源文')
                    inferred_dialogue = not speaker and emotion.casefold() in EMOTIONS and bool(
                        re.fullmatch(r"[A-Za-z][A-Za-z .'-]{0,60}", raw_key)
                    )
                    if not raw_key or inferred_dialogue:
                        warnings.append('缺少文案 Key；使用内部记录 ID 定位')
                    if not raw_key:
                        warnings.append('请确认无 Key 文案的用途')
                    if raw_key and not inferred_dialogue:
                        if raw_key in seen_keys:
                            warnings.append(f'Key 与第 {seen_keys[raw_key]} 行重复')
                        seen_keys[raw_key] = row
                    limit = value(row, 'limit')
                    if not limit:
                        warnings.append('未提供长度限制')
                    record = {
                        'id': f'{fingerprint}:{sheet_index + 1}:{row}', 'sheet': sheet.title,
                        'row': row, 'key': '' if inferred_dialogue else raw_key,
                        'speaker': raw_key if inferred_dialogue else speaker,
                        'emotion': emotion if speaker or inferred_dialogue else '',
                        'english': english, 'limit': limit, 'category': category,
                        'basis': basis, 'warning': '；'.join(warnings),
                    }
                    records.append(record)
                    self.write_record(sheet, columns, record)
                task['status'] = f'已识别 {len(records)} 条记录，正在生成分类文件...'
                task['progress'] = 80
            if not records:
                raise ValueError('未找到待分类记录；表头须包含 game_text_en 或 en')
            self.save_workbook(workbook, output_file)
            task.update(progress=100, status=f'分类完成，共 {len(records)} 条',
                        state='completed', result_file=output_file, records=records,
                        counts=self.counts(records))
        except Exception as error:
            task.update(status=f'处理失败：{error}', state='error')

    @staticmethod
    def write_record(sheet, columns, record):
        for name, value in zip(META_HEADERS, (record['category'],)):
            cell = sheet.cell(record['row'], columns[name], value)
            cell.alignment = Alignment(vertical='top', wrap_text=True)

