"""Supported language pairs and workbook field mapping shared by every stage."""
from dataclasses import dataclass, field
import re


EMOTIONS = {
    'neutral', 'happy', 'thinking', 'confused', 'investigating', 'proud', 'laughing',
    'eureka', 'annoyed', 'shocked', 'surprised', 'fear', 'angry', 'sad', 'flirting',
    'blushing', 'crying', 'worried', 'normal',
}
FIELD_ALIASES = {
    'key': ('key',), 'note': ('note', '备注'), 'speaker': ('speaker', '说话人'),
    'emotion': ('animation', 'emotion', '情绪'),
    'bubble': ('bubbletype', 'bubble_type', '气泡类型'), 'limit': ('limit',),
}
HAN = r'[\u3400-\u9fff\U00020000-\U0002fa1f]'


@dataclass(frozen=True)
class LanguagePair:
    source_lang: str = 'en'
    target_lang: str = 'de'
    columns: dict = field(default_factory=dict)

    def __post_init__(self):
        src = str(self.source_lang).lower().replace('_', '-')
        tgt = str(self.target_lang).lower().replace('_', '-')
        if src in ('en', 'en-us') and tgt in ('de', 'de-de'):
            src, tgt = 'en', 'de'
        elif src in ('zh', 'zh-cn', 'zh-hans', 'cn') and tgt in ('en', 'en-us'):
            src, tgt = 'zh-CN', 'en-US'
        else:
            raise ValueError('当前支持英语→德语和简体中文→美式英语（en-US）')
        if not isinstance(self.columns, dict):
            raise ValueError('列映射必须是对象')
        cols = {}
        for key, value in self.columns.items():
            if key not in (*FIELD_ALIASES, 'source', 'target') or not isinstance(value, str) or len(value) > 120:
                raise ValueError('列映射包含未知字段或无效列名')
            if value.strip():
                cols[key] = value.strip()
        names = [v.casefold() for v in cols.values()]
        if len(names) != len(set(names)):
            raise ValueError('不同字段不能映射到同一列；中文源文与情绪须分开')
        object.__setattr__(self, 'source_lang', src)
        object.__setattr__(self, 'target_lang', tgt)
        object.__setattr__(self, 'columns', cols)

    @property
    def chinese(self):
        return self.source_lang == 'zh-CN'

    @property
    def key(self):
        return 'zh-en-US' if self.chinese else 'en-de'

    @property
    def labels(self):
        return ('中', '英') if self.chinese else ('英', '德')

    @property
    def name(self):
        return '简体中文→美式英语（en-US）' if self.chinese else '英语→德语'

    def as_dict(self):
        return {'source_lang': self.source_lang, 'target_lang': self.target_lang, 'columns': dict(self.columns)}

    def aliases(self, role):
        if self.columns.get(role):
            return (self.columns[role].casefold(),)
        if role == 'source':
            return ('game_text_cn', 'game_text_zh', 'zh', 'cn', 'source_text', '中文') if self.chinese else ('game_text_en', 'en')
        if role == 'target':
            return ('game_text_en', 'en', 'target_text', '英文') if self.chinese else ('game_text_de', 'de')
        return FIELD_ALIASES[role]

    def column(self, headers, role):
        hits = [headers[name] for name in self.aliases(role) if name in headers]
        if role in ('source', 'target') and len(hits) > 1:
            raise ValueError(f'{role} 有多个候选列，请在上传时指定列映射')
        return hits[0] if hits else None

    def headers(self, values):
        names = [str(v).strip().casefold() if v is not None else '' for v in values]
        if (any(n in names for n in self.aliases('source'))
                and len([n for n in names if n]) != len(set(n for n in names if n))):
            raise ValueError('表头重复，无法确定字段位置')
        headers = {name: i + 1 for i, name in enumerate(names) if name}
        source = self.column(headers, 'source')
        if source:
            for role in ('target', 'emotion', 'speaker', 'key', 'note', 'bubble', 'limit'):
                if self.column(headers, role) == source:
                    raise ValueError(f'源文列不能同时作为 {role} 字段')
        return headers

    def value(self, getter, role):
        return getter(*self.aliases(role))

    def emotion(self, getter):
        value = self.value(getter, 'emotion')
        # Older EN-DE customer exports stored emotions under the CN heading.
        if not value and not self.chinese and not self.columns.get('emotion'):
            legacy = getter('game_text_cn', 'cn')
            if legacy.strip().casefold() in EMOTIONS:
                value = legacy
        return value

    def term_columns(self, values):
        names = [str(v).strip().casefold() if v is not None else '' for v in values]
        src = ('zh', 'zh-cn', 'zh-hans', 'cn', '中文') if self.chinese else ('en', 'en-us')
        tgt = ('en-us', 'en', '英文') if self.chinese else ('de', 'de-de')
        selected = []
        for aliases in (src, tgt):
            hits = [i for i, n in enumerate(names) if n in aliases]
            if len(hits) > 1:
                raise ValueError('术语库包含重复或多个语言列，请保留唯一语言对')
            selected.append(hits[0] if hits else None)
        return selected

    def tm_language(self, code, source=False):
        code = code.lower().replace('_', '-')
        if self.chinese:
            return code in (('zh', 'zh-cn', 'zh-hans', 'zh-hans-cn') if source else ('en', 'en-us'))
        return code.split('-')[0] == ('en' if source else 'de')


def language_pair(value=None, columns=None):
    if isinstance(value, LanguagePair):
        return value if columns is None else LanguagePair(value.source_lang, value.target_lang, columns)
    value = value or {}
    return LanguagePair(value.get('source_lang', value.get('src', 'en')),
                        value.get('target_lang', value.get('tgt', 'de')),
                        value.get('columns', {}) if columns is None else columns)


def source_tokens(text, chinese=False):
    if not chinese:
        return re.findall(r"[^\W_]+(?:'[^\W_]+)?", text.casefold())
    # Character bigrams avoid requiring a Chinese segmenter and retain mixed Latin words.
    result = re.findall(r"[a-z0-9]+(?:'[a-z]+)?", text.casefold())
    for run in re.findall(HAN + '+', text):
        result.extend(run if len(run) == 1 else (run[i:i + 2] for i in range(len(run) - 1)))
    return result
