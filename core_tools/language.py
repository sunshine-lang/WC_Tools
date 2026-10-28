"""Supported language pairs and workbook field mapping shared by every stage."""
from dataclasses import dataclass, field
import json
import math
from pathlib import Path
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
LANGUAGE_DIR = Path(__file__).resolve().parents[1] / 'profiles/language'
PROMPTS_ROOT = Path(__file__).resolve().parents[1] / 'prompts'


def _code(value):
    return value.casefold().replace('_', '-')


def _read_profile(path):
    try:
        value = json.loads(path.read_text('utf-8'))
        if not isinstance(value, dict):
            raise ValueError()
        return value
    except (OSError, ValueError) as error:
        raise ValueError(f'无法读取语向配置：{path.name}') from error


def _columns(value):
    if not isinstance(value, dict):
        raise ValueError('列映射必须是对象')
    cols = {}
    for key, name in value.items():
        if key not in (*FIELD_ALIASES, 'source', 'target') or not isinstance(name, str) or len(name) > 120:
            raise ValueError('列映射包含未知字段或无效列名')
        if name.strip():
            cols[key] = name.strip()
    names = [name.casefold() for name in cols.values()]
    if len(names) != len(set(names)):
        raise ValueError('不同字段不能映射到同一列；源文与情绪须分开')
    return cols


def prompt_locations(profile, root=None):
    root = Path(root or PROMPTS_ROOT).resolve()
    result = []
    for key in ('prompt_source', 'prompt_output'):
        value = profile.get(key)
        if not isinstance(value, str) or not value or Path(value).is_absolute():
            raise ValueError(f'语向配置缺少有效的 {key} 相对路径')
        path = (root / value).resolve()
        if not path.is_relative_to(root):
            raise ValueError('提示词路径必须位于 prompts 目录内')
        result.append(path)
    if result[0] == result[1]:
        raise ValueError('提示词源目录和输出目录不能相同')
    if result[1].is_relative_to(result[0] / 'scenes'):
        raise ValueError('提示词输出不能覆盖 scenes 源文件')
    return tuple(result)


def language_profiles():
    """One registry for the UI, CLI and processing stages; pair files are the allowlist."""
    catalog = _read_profile(LANGUAGE_DIR / '_catalog.json')

    def strings(value, allow_empty=False):
        return (isinstance(value, list) and (allow_empty or bool(value))
                and all(isinstance(s, str) and s.strip() for s in value))

    for role in ('sources', 'targets'):
        definitions = catalog.get(role)
        if not isinstance(definitions, dict) or not definitions:
            raise ValueError(f'语种配置缺少 {role}')
        aliases = set()
        for code, spec in definitions.items():
            if not isinstance(spec, dict):
                raise ValueError(f'语种 {code} 配置无效')
            if any(not isinstance(spec.get(k), str) or not spec[k].strip()
                   for k in ('name', 'label', 'term_guidance')):
                raise ValueError(f'语种 {code} 缺少名称、标签或术语规则')
            if (any(not strings(spec.get(k)) for k in ('aliases', 'columns', 'term_columns', 'tm_codes'))
                    or not strings(spec.get('tm_prefixes'), allow_empty=True)):
                raise ValueError(f'语种 {code} 的语言代码或列别名无效')
            names = {_code(n) for n in spec['aliases']}
            if _code(code) not in names or names & aliases:
                raise ValueError(f'语种 {code} 的别名缺少本名或与其他语种冲突')
            aliases.update(names)
            if role == 'sources':
                if spec.get('strategy') not in ('english', 'han', 'words'):
                    raise ValueError(f'语种 {code} 的源语言处理策略尚未实现')
                floor = spec.get('bm25_floor')
                if (type(floor) not in (int, float) or not math.isfinite(floor) or floor < 0
                        or not isinstance(spec.get('stop_words'), str)
                        or not strings(spec.get('legacy_emotion_columns'), allow_empty=True)
                        or type(spec.get('reject_emotion_only_source')) is not bool):
                    raise ValueError(f'语种 {code} 的检索或情绪规则无效')
                try:
                    re.compile(spec['negative_pattern'])
                except (KeyError, TypeError, re.error) as error:
                    raise ValueError(f'语种 {code} 的否定表达规则无效') from error

    profiles, combinations, outputs = {}, set(), set()
    for path in sorted(LANGUAGE_DIR.glob('*.json')):
        if path.name.startswith('_'):
            continue
        profile = _read_profile(path)
        src, tgt = profile.get('source_lang'), profile.get('target_lang')
        if (not isinstance(src, str) or not isinstance(tgt, str)
                or src not in catalog['sources'] or tgt not in catalog['targets']):
            raise ValueError(f'{path.name} 引用了未登记的源语言或目标语言')
        if (src, tgt) in combinations:
            raise ValueError(f'{path.name} 与已有语向重复')
        profile['columns'] = _columns(profile.get('columns', {}))
        for key in ('termbases', 'tms'):
            if not strings(profile.get(key, []), allow_empty=True):
                raise ValueError(f'{path.name} 的 {key} 必须是路径列表')
        source, output = prompt_locations(profile)
        if output in outputs:
            raise ValueError(f'{path.name} 的 Prompt 输出目录与另一语向重复')
        if not (source / 'base.md').is_file() or not list((source / 'scenes').glob('*.md')):
            raise ValueError(f'{path.name} 的提示词源文件不完整')
        combinations.add((src, tgt))
        outputs.add(output)
        profiles[path.stem] = {**profile, 'key': path.stem,
                               'source': catalog['sources'][src], 'target': catalog['targets'][tgt],
                               'default': path.stem == catalog.get('default_pair')}
    if not profiles or not any(p['default'] for p in profiles.values()):
        raise ValueError('未登记默认语向，请检查 _catalog.json 的 default_pair')
    return profiles


def language_profile(key):
    profiles = language_profiles()
    if key not in profiles:
        raise ValueError('未登记的语向；当前支持：' + '、'.join(profiles))
    return profiles[key]


def supported_language_pairs():
    return [{'key': key, 'source_lang': p['source_lang'], 'target_lang': p['target_lang'],
             'source_name': p['source']['name'], 'target_name': p['target']['name'],
             'source_column': p.get('columns', {}).get('source') or p['source']['columns'][0],
             'target_column': p.get('columns', {}).get('target') or p['target']['columns'][0],
             'emotion_columns': [p['columns']['emotion']] if p['columns'].get('emotion') else list(FIELD_ALIASES['emotion']),
             'legacy_emotion_columns': [] if p['columns'].get('emotion') else p['source']['legacy_emotion_columns'],
             'default': p['default']}
            for key, p in language_profiles().items()]


@dataclass(frozen=True)
class LanguagePair:
    source_lang: str | None = None
    target_lang: str | None = None
    columns: dict = field(default_factory=dict)
    _profile: dict = field(init=False, repr=False, compare=False)

    def __post_init__(self):
        profiles = language_profiles()
        if self.source_lang is None and self.target_lang is None:
            selected = [p for p in profiles.values() if p['default']]
        elif isinstance(self.source_lang, str) and isinstance(self.target_lang, str):
            selected = [p for p in profiles.values()
                        if _code(self.source_lang) in {_code(a) for a in p['source']['aliases']}
                        and _code(self.target_lang) in {_code(a) for a in p['target']['aliases']}]
        else:
            selected = []
        if len(selected) != 1:
            raise ValueError('当前支持：' + '、'.join(f"{p['source']['name']}→{p['target']['name']}" for p in profiles.values()))
        profile = selected[0]
        cols = _columns({**profile['columns'], **_columns(self.columns)})
        object.__setattr__(self, 'source_lang', profile['source_lang'])
        object.__setattr__(self, 'target_lang', profile['target_lang'])
        object.__setattr__(self, 'columns', cols)
        object.__setattr__(self, '_profile', profile)

    @property
    def source_rules(self):
        return self._profile['source']

    @property
    def target_rules(self):
        return self._profile['target']

    @property
    def key(self):
        return self._profile['key']

    @property
    def labels(self):
        return self.source_rules['label'], self.target_rules['label']

    @property
    def name(self):
        return f"{self.source_rules['name']}→{self.target_rules['name']}"

    def as_dict(self):
        return {'source_lang': self.source_lang, 'target_lang': self.target_lang, 'columns': dict(self.columns)}

    def cache_identity(self):
        return {**self.as_dict(), 'key': self.key, 'source_rules': self.source_rules, 'target_rules': self.target_rules}

    def aliases(self, role):
        if self.columns.get(role):
            return (self.columns[role].casefold(),)
        if role in ('source', 'target'):
            return tuple(n.casefold() for n in self._profile[role]['columns'])
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
        legacy_columns = self.source_rules['legacy_emotion_columns']
        if not value and legacy_columns and not self.columns.get('emotion'):
            legacy = getter(*legacy_columns)
            if legacy.strip().casefold() in EMOTIONS:
                value = legacy
        return value

    def term_columns(self, values):
        names = [str(v).strip().casefold() if v is not None else '' for v in values]
        selected = []
        for role in ('source', 'target'):
            aliases = [n.casefold() for n in self._profile[role]['term_columns']]
            hits = [i for i, n in enumerate(names) if n in aliases]
            if len(hits) > 1:
                raise ValueError('术语库包含重复或多个语言列，请保留唯一语言对')
            selected.append(hits[0] if hits else None)
        return selected

    def tm_language(self, code, source=False):
        spec = self.source_rules if source else self.target_rules
        code = _code(code)
        return (code in {_code(c) for c in spec['tm_codes']}
                or code.split('-')[0] in {_code(c) for c in spec['tm_prefixes']})


def language_pair(value=None, columns=None):
    if isinstance(value, LanguagePair):
        return value if columns is None else LanguagePair(value.source_lang, value.target_lang, columns)
    value = {} if value is None else value
    if not isinstance(value, dict):
        raise ValueError('语向配置必须是对象')
    return LanguagePair(value.get('source_lang', value.get('src')),
                        value.get('target_lang', value.get('tgt')),
                        value.get('columns', {}) if columns is None else columns)


def source_tokens(text, strategy='english'):
    if strategy in ('english', 'words'):
        return re.findall(r"[^\W_]+(?:'[^\W_]+)?", text.casefold())
    if strategy != 'han':
        raise ValueError(f'未实现的源语言处理策略：{strategy}')
    # Character bigrams avoid requiring a Chinese segmenter and retain mixed Latin words.
    result = re.findall(r"[a-z0-9]+(?:'[a-z]+)?", text.casefold())
    for run in re.findall(HAN + '+', text):
        result.extend(run if len(run) == 1 else (run[i:i + 2] for i in range(len(run) - 1)))
    return result
