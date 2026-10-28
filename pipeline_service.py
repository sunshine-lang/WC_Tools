"""WC 翻译准备工作台：项目、资料库、批次与流水线的落盘服务。

这个模块提供本地翻译准备服务：
- 项目 / 资料库 / 批次都落到 workspace/pipeline/，重启不丢；
- 每个节点单独执行，节点之间用上一节点产物衔接；
- 自动节点自动往下走，需要确认的节点停在 review；
- 失败保留前序产物，可重试本步或重试并继续；
- 准备自检节点暂时停用；代码保留，PREFLIGHT_ENABLED=True 时恢复。

只做本地单用户使用，不引入数据库。
"""
from __future__ import annotations

import csv
import json
import os
import re
import shutil
import threading
import time
import uuid
from collections import Counter
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook

from core_tools.term_filter import TermFilter
from core_tools.term_scanner import TermScanner
from core_tools.language import language_pair, supported_language_pairs, EMOTIONS
from core_tools.text_metrics import display_length
from core_tools.tm_matcher import TMMatcher
from core_tools.tm_semantic import SemanticTMRetriever
from pipeline import (
    run_classify,
    run_context,
    run_normalize,
    run_term,
    run_term_filter,
    run_tm,
)


ROOT = Path(__file__).resolve().parent
DEFAULT_WORKSPACE = ROOT / 'workspace' / 'pipeline'
STORE_VERSION = 2

SLOTS = (
    {'key': 'termbase', 'name': '术语库', 'req': False, 'icon': '📗', 'accept': '.xlsx,.csv',
     'desc': '当前语向的词条对照；中文支持连续文本，英文按词边界匹配。'},
    {'key': 'tm', 'name': '记忆库', 'req': False, 'icon': '🗂', 'accept': '.tmx',
     'desc': '历史源文与译文，TMX 格式，用于找相同或相近的句子。'},
    {'key': 'cast', 'name': '人物人设', 'req': False, 'icon': '👤', 'accept': '.md,.txt',
     'desc': '角色身份、性格与表达习惯，全文作为术语筛选的参考语境。'},
    {'key': 'relationships', 'name': '人物关系', 'req': False, 'icon': '👥', 'accept': '.md,.txt',
     'desc': '人物间的关系、称呼与适用剧情阶段，单独上传和维护。'},
    {'key': 'world', 'name': '游戏背景', 'req': False, 'icon': '🌍', 'accept': '.md,.txt',
     'desc': '世界观、剧情梗概；术语筛选时作为语境参考。'},
    {'key': 'style', 'name': '风格指南', 'req': False, 'icon': '✒️', 'accept': '.md,.txt',
     'desc': '客户的语言规范与禁用写法；术语筛选时作为语境参考。'},
)
SLOT_KEYS = {slot['key'] for slot in SLOTS}

PREFLIGHT_ENABLED = False  # 暂时停用 06 准备自检；改为 True 并重启即可恢复

PREFLIGHT_STAGE = {
    'key': 'preflight', 'no': '06', 'title': '准备自检', 'sub': '交付前的最后一道闸门',
    'artifact': '06-自检报告.md', 'confirm': True,
    'desc': '九条断言逐条校验，输出报告而不改数据。阻断项失败就不允许交付，告警项只记录。',
}

STAGES = (
    {'key': 'normalize', 'no': '00', 'title': '规范化', 'sub': '先修表，再动内容',
     'artifact': '00-规范化.xlsx', 'confirm': False,
     'desc': '客户导出的工作表常带只有格式、没有内容的空列；说话人写在 key 列里，分类器认不出对白。这一步补齐形状，不改业务判断。'},
    {'key': 'classify', 'no': '01', 'title': '分类', 'sub': '决定下游用哪套提示词',
     'artifact': '01-分类.xlsx', 'confirm': False,
     'desc': '按项目规则给每一行打文本类型。它同时决定下游选哪套提示词，以及长度校验对哪些行生效。'},
    {'key': 'context', 'no': '02', 'title': '上下文', 'sub': '关联原文，通读对白情景',
     'artifact': '02-上下文.xlsx', 'confirm': True,
     'desc': 'UI、集卡和装饰物关联原文。活动对白按完整 Key 分组、句序排序，结合资料库的人物人设、人物关系与同段源文，生成「场景结构＋情境与基调」，同组共用一份；两类原文依据可展开核对。'},
    {'key': 'term', 'no': '03', 'title': '术语', 'sub': '词边界匹配，长词优先',
     'artifact': '03-术语.xlsx', 'confirm': False,
     'desc': '按当前源语言匹配术语，中文支持连续文本，英文保留词边界与词形规则。长词优先，候选交给下一步语境筛选。'},
    {'key': 'term_filter', 'no': '04', 'title': '术语筛选', 'sub': '结合语境判真假',
     'artifact': '04-术语筛选.xlsx', 'confirm': True,
     'desc': '把候选词放回句子里，让模型判断它是否真的生效。项目资料库里的背景、人物人设、人物关系与风格指南会全文作为参考语境送给模型。'},
    {'key': 'tm', 'no': '05', 'title': 'TM 参考', 'sub': '精确优先，语义召回与重排',
     'artifact': '05-TM.xlsx', 'confirm': True,
     'desc': '原文一致优先；其余通过语义向量和关键词召回，再由模型重排。默认召回 30 条、最多选 5 条历史原文，保留不同译法和差异提示。首次运行会生成并缓存记忆库向量。'},
)
if PREFLIGHT_ENABLED:
    STAGES = STAGES + (PREFLIGHT_STAGE,)
STAGE_BY_KEY = {stage['key']: stage for stage in STAGES}
ORDER = tuple(stage['key'] for stage in STAGES)
ARTIFACTS = {stage['key']: stage['artifact'] for stage in STAGES}
RUNNERS = {
    'normalize': run_normalize,
    'classify': run_classify,
    'context': run_context,
    'term': run_term,
    'term_filter': run_term_filter,
    'tm': run_tm,
}
if PREFLIGHT_ENABLED:
    from pipeline import run_preflight
    RUNNERS['preflight'] = run_preflight

CATEGORY_COLORS = {
    '活动 UI 与商店文案': '#0ea5e9',
    '集卡名称': '#3b82f6',
    '装饰物名称': '#10b981',
    '活动对白': '#8b5cf6',
    '待确认': '#f59e0b',
}
TEXT_SUFFIXES = {'.md', '.txt'}
MAX_TEXT_BYTES = 2 * 1024 * 1024


class PipelineError(Exception):
    """可以直接返回给前端的中文错误。"""


def now() -> str:
    return datetime.now().strftime('%Y-%m-%d %H:%M:%S')


def text(value) -> str:
    if value is None:
        return ''
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def snippet(value, limit=70) -> str:
    value = re.sub(r'\s+', ' ', text(value)).strip()
    return value if len(value) <= limit else value[:limit - 1] + '…'


def read_json(path: Path, default=None):
    try:
        return json.loads(Path(path).read_text('utf-8'))
    except (OSError, ValueError):
        return {} if default is None else default


def write_json(path: Path, data) -> None:
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    temp = path.with_name(f'{path.name}.{uuid.uuid4().hex}.tmp')
    temp.write_text(json.dumps(data, ensure_ascii=False, indent=2), 'utf-8')
    os.replace(temp, path)


def safe_name(name: str) -> str:
    name = text(name).replace('\\', '/').rsplit('/', 1)[-1].strip()
    name = re.sub(r'[\x00-\x1f<>:"/\\|?*]', '_', name)
    return name[:120] or 'file'


def workbook_rows(path: Path):
    """读第一张有数据的表，返回 (表头, 行字典, 全部行值)。"""
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        sheet = workbook.worksheets[0]
        iterator = sheet.iter_rows(values_only=True)
        header = [text(cell).strip() for cell in next(iterator, ())]
        body = []
        for values in iterator:
            if not any(value not in (None, '') for value in values):
                continue
            body.append({header[i]: (values[i] if i < len(values) else None)
                         for i in range(len(header)) if header[i]})
        return header, body
    finally:
        workbook.close()


def column(row: dict, *names):
    for name in names:
        for key, value in row.items():
            if key.casefold() == name.casefold():
                return value
    return ''


def display_len(value: str) -> int:
    return display_length(text(value))


class PipelineService:
    def __init__(self, root: Path = ROOT, workspace: Path | None = None):
        self.root = Path(root).resolve()
        self.workspace = Path(workspace) if workspace else DEFAULT_WORKSPACE
        self.workspace = self.workspace.resolve() if self.workspace.is_absolute() else (self.root / self.workspace).resolve()
        self.projects_file = self.workspace / 'projects.json'
        self.lock = threading.RLock()
        self.running: dict[str, dict] = {}
        self.workspace.mkdir(parents=True, exist_ok=True)
        self._ensure_store()
        self._recover()

    # ------------------------------------------------------------ 存储

    def _ensure_store(self):
        with self.lock:
            if self.projects_file.is_file():
                store = self._read_store()
                if store.get('version', 1) == 1:
                    for project in store['projects'].values():
                        library = project.setdefault('library', {})
                        if library.get('cast'):
                            library['cast']['legacy_combined'] = True
                        library.setdefault('relationships', None)
                    store['version'] = STORE_VERSION
                    self._write_store(store)
                return
            write_json(self.projects_file, self._default_store())

    def _default_store(self):
        world = self.root.parent / 'WC游戏介绍.md'
        style = self.root.parent / 'WC翻译客户要求.md'
        library = {
            'termbase': None,
            'tm': None,
            'cast': None,
            'relationships': None,
            'world': self._text_item(world) if world.is_file() else None,
            'style': self._text_item(style) if style.is_file() else None,
        }
        pair = language_pair()
        project = {
            'id': 'wc',
            'name': '听风城堡',
            'src': pair.source_lang,
            'tgt': pair.target_lang,
            'game': 'wc',
            'created_at': now(),
            'library': library,
            'active_batch': None,
        }
        return {'version': STORE_VERSION, 'active_project': 'wc', 'projects': {'wc': project}}

    def _library_item(self, path: Path, name: str, meta: str):
        return {'path': self._relative(path), 'name': name, 'meta': meta, 'updated_at': now()}

    def _text_item(self, path: Path):
        try:
            content = path.read_text('utf-8', errors='replace')
            meta = f'{len([line for line in content.splitlines() if line.strip()])} 行'
        except OSError:
            meta = ''
        return self._library_item(path, path.name, meta)

    def _relative(self, path: Path) -> str:
        path = Path(path).resolve()
        try:
            return str(path.relative_to(self.root))
        except ValueError:
            return str(path)

    def _resolve(self, value: str | None) -> Path | None:
        if not value:
            return None
        path = Path(value)
        return path if path.is_absolute() else (self.root / path)

    def _read_store(self) -> dict:
        store = read_json(self.projects_file)
        if not isinstance(store, dict) or not isinstance(store.get('projects'), dict):
            store = self._default_store()
            write_json(self.projects_file, store)
        return store

    def _write_store(self, store: dict):
        write_json(self.projects_file, store)

    def _project(self, store: dict, project_id: str | None = None) -> dict:
        if project_id:
            project = store['projects'].get(project_id)
            if not project:
                raise PipelineError('项目不存在')
            return project
        project_id = store.get('active_project')
        project = store['projects'].get(project_id)
        if not project:
            if not store['projects']:
                raise PipelineError('还没有项目')
            project = next(iter(store['projects'].values()))
            store['active_project'] = project['id']
        return project

    def _project_dir(self, project_id: str) -> Path:
        return self.workspace / 'projects' / project_id

    def _batch_dir(self, project_id: str, batch_id: str) -> Path:
        return self._project_dir(project_id) / 'batches' / batch_id

    def _batch_file(self, project_id: str, batch_id: str) -> Path:
        return self._batch_dir(project_id, batch_id) / 'state.json'

    # ------------------------------------------------------------ 恢复

    def _recover(self):
        with self.lock:
            store = self._read_store()
            changed = False
            for project in store['projects'].values():
                for batch in self._all_batches(project['id']):
                    state = read_json(self._batch_file(project['id'], batch['id']))
                    if not state:
                        continue
                    if state.get('running') or any(stage.get('state') == 'running'
                                                    for stage in state.get('stages', {}).values()):
                        state['running'] = False
                        state['active_stage'] = None
                        for stage in state.get('stages', {}).values():
                            if stage.get('state') == 'running':
                                stage['state'] = 'failed'
                                stage['message'] = '服务重启，节点中断，请重试本步'
                                stage['progress'] = 100
                        state['updated_at'] = now()
                        write_json(self._batch_file(project['id'], batch['id']), state)
                        changed = True
            if changed:
                self._write_store(store)

    # ------------------------------------------------------------ 项目

    def _project_view(self, project: dict) -> dict:
        library = []
        for slot in SLOTS:
            item = project.get('library', {}).get(slot['key']) or {}
            path = self._resolve(item.get('path'))
            ready = bool(path and path.is_file())
            library.append({**slot, 'file': item.get('name', '') if ready else '',
                            'meta': item.get('meta', '') if ready else '',
                            'notice': ('已保留原“人设 / 关系”合并文件，请分别上传人物人设和人物关系。'
                                       if ready and item.get('legacy_combined') else ''),
                            'ready': ready, 'updated_at': item.get('updated_at', '')})
        return {'id': project['id'], 'name': project['name'], 'src': project['src'],
                'tgt': project['tgt'], 'use_tm': project.get('use_tm', True), 'game': project.get('game', 'wc'),
                'library': library, 'created_at': project.get('created_at', '')}

    def _project_list(self, store: dict, active_id: str) -> list[dict]:
        return [{'id': project['id'], 'name': project['name'], 'src': project['src'],
                 'tgt': project['tgt'], 'active': project['id'] == active_id}
                for project in store['projects'].values()]

    def bootstrap(self, project_id: str | None = None, batch_id: str | None = None) -> dict:
        with self.lock:
            store = self._read_store()
            project = self._project(store, project_id)
            if project_id and store.get('active_project') != project['id']:
                store['active_project'] = project['id']
                self._write_store(store)
            batches = self._all_batches(project['id'])
            target_batch = batch_id or project.get('active_batch')
            if target_batch and not any(batch['id'] == target_batch for batch in batches):
                target_batch = None
            if not target_batch and batches:
                target_batch = batches[0]['id']
            if target_batch and project.get('active_batch') != target_batch:
                project['active_batch'] = target_batch
                self._write_store(store)
            batch_state = self._load_batch(project['id'], target_batch) if target_batch else None
            if batch_state:
                self._merge_runtime(batch_state)
            return {
                'language_pairs': supported_language_pairs(),
                'projects': self._project_list(store, project['id']),
                'project': self._project_view(project),
                'stage_defs': [{key: stage[key] for key in
                                ('key', 'no', 'title', 'sub', 'confirm', 'artifact', 'desc')}
                               for stage in STAGES],
                'batches': [self._batch_brief(project['id'], batch) for batch in batches],
                'batch': self._batch_view(project['id'], batch_state) if batch_state else None,
                'server': self._server_status(),
            }

    def create_project(self, name: str, src: str | None = None, tgt: str | None = None, use_tm: bool = True) -> dict:
        name = text(name).strip()
        if not name:
            raise PipelineError('请填写项目名称')
        if len(name) > 80:
            raise PipelineError('项目名称最多 80 个字符')
        try:
            pair = language_pair({'src': src, 'tgt': tgt})
        except ValueError as error:
            raise PipelineError(str(error)) from error
        if type(use_tm) is not bool:
            raise PipelineError('use_tm 必须是布尔值')
        src, tgt = pair.source_lang, pair.target_lang
        with self.lock:
            store = self._read_store()
            project_id = f'p_{uuid.uuid4().hex[:10]}'
            store['projects'][project_id] = {
                'id': project_id, 'name': name, 'src': src, 'tgt': tgt, 'game': 'wc', 'use_tm': use_tm,
                'created_at': now(), 'library': {key: None for key in SLOT_KEYS},
                'active_batch': None,
            }
            store['active_project'] = project_id
            self._write_store(store)
            self._project_dir(project_id).mkdir(parents=True, exist_ok=True)
            return self.bootstrap(project_id)

    def activate_project(self, project_id: str) -> dict:
        with self.lock:
            store = self._read_store()
            if project_id not in store['projects']:
                raise PipelineError('项目不存在')
            store['active_project'] = project_id
            self._write_store(store)
            return self.bootstrap(project_id)

    # ------------------------------------------------------------ 资料库

    def _project_busy(self, project_id: str) -> bool:
        return any(self._load_batch(project_id, batch_id) is not None
                   for batch_id in list(self.running))

    def upload_library(self, project_id: str, slot: str, upload) -> dict:
        if slot not in SLOT_KEYS:
            raise PipelineError('未知的资料库槽位')
        if not upload or not upload.filename:
            raise PipelineError('请选择文件')
        filename = safe_name(upload.filename)
        suffix = Path(filename).suffix.lower()
        slot_def = next(item for item in SLOTS if item['key'] == slot)
        if suffix not in {part.strip() for part in slot_def['accept'].split(',')}:
            raise PipelineError(f'该槽位只支持 {slot_def["accept"]}')
        if slot in ('cast', 'relationships', 'world', 'style') and suffix not in TEXT_SUFFIXES:
            raise PipelineError('人物人设 / 人物关系 / 背景 / 风格资料只支持 .md 或 .txt')

        with self.lock:
            store = self._read_store()
            project = self._project(store, project_id)
            if self._project_busy(project['id']):
                raise PipelineError('有节点正在执行，资料库暂不可修改')
            folder = self._project_dir(project['id']) / 'library' / slot
            folder.mkdir(parents=True, exist_ok=True)
            stored = folder / f'{datetime.now().strftime("%Y%m%d%H%M%S")}_{uuid.uuid4().hex[:6]}_{filename}'
            upload.save(stored)
            try:
                meta = self._validate_library(slot, stored, language_pair(project))
            except Exception as error:
                stored.unlink(missing_ok=True)
                raise PipelineError(f'{slot_def["name"]}校验失败：{error}') from error
            project.setdefault('library', {})[slot] = {
                'path': self._relative(stored), 'name': filename, 'meta': meta, 'updated_at': now(),
            }
            if slot == 'tm':
                project['use_tm'] = True
            self._write_store(store)
            return self.bootstrap(project['id'])

    def remove_library(self, project_id: str, slot: str) -> dict:
        if slot not in SLOT_KEYS:
            raise PipelineError('未知的资料库槽位')
        with self.lock:
            store = self._read_store()
            project = self._project(store, project_id)
            if self._project_busy(project['id']):
                raise PipelineError('有节点正在执行，资料库暂不可修改')
            item = project.get('library', {}).get(slot)
            if not item:
                return self.bootstrap(project['id'])
            path = self._resolve(item.get('path'))
            project['library'][slot] = None
            self._write_store(store)
            if path and self._is_in_project(path, project['id']):
                path.unlink(missing_ok=True)
            return self.bootstrap(project['id'])

    def _validate_library(self, slot: str, path: Path, language=None) -> str:
        if slot == 'termbase':
            status = TermScanner(str(path), language=language).status()
            return f'{status["term_count"]} 条'
        if slot == 'tm':
            status = TMMatcher(str(path), language=language).status()
            return f'{status["tu_count"]:,} 对'
        if path.stat().st_size > MAX_TEXT_BYTES:
            raise ValueError('文本资料不能超过 2MB')
        content = path.read_text('utf-8', errors='strict')
        if not content.strip():
            raise ValueError('文件是空的')
        return f'{len([line for line in content.splitlines() if line.strip()])} 行'

    def _is_in_project(self, path: Path, project_id: str) -> bool:
        try:
            path.resolve().relative_to(self._project_dir(project_id).resolve())
            return True
        except ValueError:
            return False

    # ------------------------------------------------------------ 批次

    def _all_batches(self, project_id: str) -> list[dict]:
        root = self._project_dir(project_id) / 'batches'
        if not root.is_dir():
            return []
        result = []
        for folder in root.iterdir():
            state = read_json(folder / 'state.json')
            if not state:
                continue
            result.append(state)
        result.sort(key=lambda item: item.get('created_at', ''), reverse=True)
        return result

    def _batch_brief(self, project_id: str, state: dict) -> dict:
        done = sum(1 for key in ORDER if state.get('stages', {}).get(key, {}).get('state') == 'done')
        return {'id': state['id'], 'name': state.get('name', ''), 'rows': state.get('rows', 0),
                'created_at': state.get('created_at', ''), 'updated_at': state.get('updated_at', ''),
                'running': bool(state.get('running')), 'done': done, 'total': len(ORDER)}

    def _load_batch(self, project_id: str, batch_id: str | None) -> dict | None:
        if not batch_id:
            return None
        state = read_json(self._batch_file(project_id, batch_id))
        if not state:
            return None
        state.setdefault('stages', {})
        for key in list(state['stages']):
            if key not in ORDER:
                state['stages'].pop(key)
        if state.get('active_stage') not in ORDER:
            state['active_stage'] = None
        for key in ORDER:
            state['stages'].setdefault(key, {'state': 'pending', 'progress': 0, 'status': ''})
        return state

    def _save_batch(self, project_id: str, state: dict):
        state['updated_at'] = now()
        write_json(self._batch_file(project_id, state['id']), state)

    def create_batch(self, project_id: str, upload, columns_map=None) -> dict:
        if not upload or not upload.filename:
            raise PipelineError('请选择待译文件')
        filename = safe_name(upload.filename)
        suffix = Path(filename).suffix.lower()
        if suffix not in ('.xlsx', '.csv'):
            raise PipelineError('待译文件只支持 .xlsx 或 .csv')
        with self.lock:
            store = self._read_store()
            project = self._project(store, project_id)
            batch_id = f'b_{datetime.now().strftime("%Y%m%d%H%M%S")}_{uuid.uuid4().hex[:6]}'
            folder = self._batch_dir(project['id'], batch_id)
            folder.mkdir(parents=True, exist_ok=True)
            source = folder / f'source{suffix}'
            upload.save(source)
            try:
                pair = language_pair(project, columns_map if columns_map is not None else {})
                rows, columns, _ = self._inspect_source(source, pair)
            except Exception as error:
                shutil.rmtree(folder, ignore_errors=True)
                raise PipelineError(f'文件无法读取：{error}') from error
            state = {
                'id': batch_id, 'project_id': project['id'], 'name': filename, 'rows': rows,
                'columns': columns, 'language': pair.as_dict(), 'source': source.name, 'created_at': now(), 'updated_at': now(),
                'running': False, 'active_stage': None,
                'stages': {key: {'state': 'pending', 'progress': 0, 'status': ''} for key in ORDER},
            }
            self._save_batch(project['id'], state)
            project['active_batch'] = batch_id
            self._write_store(store)
            return self.bootstrap(project['id'], batch_id)

    def activate_batch(self, project_id: str, batch_id: str) -> dict:
        with self.lock:
            store = self._read_store()
            project = self._project(store, project_id)
            if not self._load_batch(project['id'], batch_id):
                raise PipelineError('批次不存在')
            project['active_batch'] = batch_id
            self._write_store(store)
            return self.bootstrap(project['id'], batch_id)

    def _merge_runtime(self, state: dict):
        runtime = self.running.get(state.get('id'))
        if not runtime:
            return state
        task = runtime.get('task') or {}
        stage = state['stages'].get(state.get('active_stage'))
        if stage and stage.get('state') == 'running':
            stage['progress'] = task.get('progress', stage.get('progress', 0))
            stage['status'] = task.get('status', stage.get('status', ''))
            if task.get('status'):
                stage['message'] = task.get('status')
        return state

    def clone_batch(self, project_id: str, batch_id: str) -> dict:
        """基于同一源文件新建一个待处理批次；旧批次和产物保留。"""
        with self.lock:
            store = self._read_store()
            project = self._project(store, project_id)
            batch = self._load_batch(project['id'], batch_id)
            if not batch:
                raise PipelineError('批次不存在')
            if batch.get('running') or self._project_busy(project['id']):
                raise PipelineError('当前有节点正在执行，等结束再重跑')
            source = self._batch_dir(project['id'], batch_id) / batch['source']
            if not source.is_file():
                raise PipelineError('源文件不存在，无法重跑')
            new_id = f'b_{datetime.now().strftime("%Y%m%d%H%M%S")}_{uuid.uuid4().hex[:6]}'
            folder = self._batch_dir(project['id'], new_id)
            folder.mkdir(parents=True, exist_ok=True)
            shutil.copyfile(source, folder / batch['source'])
            state = {
                'id': new_id, 'project_id': project['id'], 'name': batch.get('name', ''),
                'rows': batch.get('rows', 0), 'columns': batch.get('columns', []),
                'language': batch.get('language', language_pair(project).as_dict()),
                'source': batch['source'], 'created_at': now(), 'updated_at': now(),
                'running': False, 'active_stage': None, 'parent_batch': batch_id,
                'stages': {key: {'state': 'pending', 'progress': 0, 'status': ''} for key in ORDER},
            }
            self._save_batch(project['id'], state)
            project['active_batch'] = new_id
            self._write_store(store)
            return self.bootstrap(project['id'], new_id)

    def batch_state(self, project_id: str, batch_id: str) -> dict:
        with self.lock:
            state = self._load_batch(project_id, batch_id)
            if not state:
                raise PipelineError('批次不存在')
            self._merge_runtime(state)
            return self._batch_view(project_id, state)

    def _batch_view(self, project_id: str, state: dict) -> dict:
        stages = []
        for meta in STAGES:
            info = state['stages'].get(meta['key'], {})
            status = info.get('state', 'pending')
            artifact = self._batch_dir(project_id, state['id']) / meta['artifact']
            item = {
                **meta,
                'desc': info.get('desc', meta['desc']),
                'state': status,
                'skipped': bool(info.get('skipped')),
                'seconds': info.get('seconds', 0),
                'message': info.get('message', ''),
                'stats': info.get('stats', []),
                'preview': info.get('preview', {}),
                'checks': info.get('checks', []),
                'review': info.get('review', []),
                'progress': info.get('progress', 100 if status in ('done', 'review', 'failed') else 0),
                'status': info.get('status', ''),
                'download_url': (f'/api/pipeline/projects/{project_id}/batches/{state["id"]}/artifact/{meta["key"]}'
                                 if artifact.is_file() else ''),
            }
            stages.append(item)
        done = sum(1 for item in stages if item['state'] == 'done')
        final_path = self._batch_dir(project_id, state['id']) / ARTIFACTS['tm']
        final_name = f'{Path(state.get("name") or "batch").stem}_ready.xlsx'
        return {
            'id': state['id'], 'project_id': project_id, 'name': state.get('name', ''),
            'rows': state.get('rows', 0), 'columns': state.get('columns', []),
            'created_at': state.get('created_at', ''), 'updated_at': state.get('updated_at', ''),
            'running': bool(state.get('running')), 'active_stage': state.get('active_stage'),
            'stages': stages, 'done': done, 'total': len(ORDER),
            'final_download_url': (f'/api/pipeline/projects/{project_id}/batches/{state["id"]}/final'
                                   if final_path.is_file() else ''),
            'final_name': final_name if final_path.is_file() else '',
            'language': state.get('language', {}),
        }

    def _inspect_source(self, path: Path, language=None):
        pair = language_pair(language)
        book = None
        try:
            if path.suffix.lower() == '.csv':
                with path.open(encoding='utf-8-sig', newline='') as file:
                    sheets = [('csv', list(csv.reader(file)))]
            else:
                book = load_workbook(path, read_only=True, data_only=False)
                sheets = [(sheet.title, list(sheet.iter_rows(values_only=True))) for sheet in book]
            count, first, source_values = 0, None, []
            for name, rows in sheets:
                if not rows:
                    continue
                headers = pair.headers(rows[0])
                source = pair.column(headers, 'source')
                if not source:
                    continue
                for role, mapped in pair.columns.items():
                    if role != 'target' and mapped.casefold() not in headers:
                        raise ValueError(f'{name} 缺少指定的 {role} 列：{mapped}')
                if first is None:
                    first = ([text(c).strip() for c in rows[0]], name)
                values = [text(row[source - 1]).strip() for row in rows[1:] if len(row) >= source and text(row[source - 1]).strip()]
                count += len(values)
                source_values.extend(values)
            if first is None or not count:
                raise ValueError('未找到有效源文，请核对当前语向及源文列映射：' + ' / '.join(pair.aliases('source')))
            if pair.source_rules['reject_emotion_only_source'] and all(v.casefold() in EMOTIONS for v in source_values):
                raise ValueError(f"选择的{pair.source_rules['name']}源文列实际为情绪标记，请指定台词列")
            return count, first[0], first[1]
        finally:
            if book:
                book.close()

    # ------------------------------------------------------------ 执行

    def _key_rules_path(self):
        profile = read_json(self.root / 'profiles/game/wc.json', {})
        return self.root / profile.get('key_rules', 'profiles/game/wc_key_rules.json')

    def _server_status(self):
        try:
            status = TermFilter().status()
        except Exception as error:  # noqa: BLE001
            status = {'ready': False, 'model': '', 'error': str(error)}
        try:
            tm_status = SemanticTMRetriever(self.root / 'config/config.json').status()
        except (ValueError, OSError):
            tm_status = {'ready': False}
        key_rules = self._key_rules_path()
        return {'llm_ready': bool(status.get('ready')), 'model': status.get('model', ''),
                'tm_ready': bool(tm_status.get('ready')),
                'key_rules_ready': key_rules.is_file(),
                'error': status.get('error', '')}

    def _project_term(self, project: dict):
        item = (project.get('library') or {}).get('termbase') or {}
        if not item:
            return None
        path = self._resolve(item.get('path'))
        if not path or not path.is_file():
            raise PipelineError('已登记的术语库文件不存在，请重新上传或移除该记录')
        return path

    def _project_tm(self, project: dict):
        if not project.get('use_tm', True):
            return None
        item = (project.get('library') or {}).get('tm') or {}
        if not item:
            return None
        path = self._resolve(item.get('path'))
        if not path or not path.is_file():
            raise PipelineError('已登记的记忆库文件不存在，请重新上传或移除该记录')
        return path

    def _materials(self, project: dict) -> dict:
        result = {}
        for key in ('world', 'cast', 'relationships', 'style'):
            item = (project.get('library') or {}).get(key) or {}
            path = self._resolve(item.get('path'))
            if path and path.is_file():
                try:
                    result[key] = path.read_text('utf-8', errors='replace')
                except OSError:
                    continue
        return result

    def _ctx(self, project: dict, batch: dict) -> dict:
        term = self._project_term(project)
        tm = self._project_tm(project)
        profile = read_json(self.root / 'profiles/game/wc.json', {})
        key_rules = self._key_rules_path()
        if not key_rules.is_file():
            raise PipelineError('缺少分类规则 profiles/game/wc_key_rules.json')
        library = project.get('library') or {}
        reference = self._resolve((library.get('world') or {}).get('path'))
        guide = self._resolve((library.get('style') or {}).get('path'))
        return {
            'game': {
                'speakers': profile.get('speakers', []),
                'key_rules': str(key_rules),
                'context': {'reference': str(reference) if reference and reference.is_file() else None,
                            'guide': str(guide) if guide and guide.is_file() else None},
            },
            'language': {**language_pair(batch.get('language') or project).as_dict(),
                         'termbases': [str(term)] if term else [], 'tms': [str(tm)] if tm else [],
                         'use_tm': project.get('use_tm', True)},
            'materials': self._materials(project),
            'config_path': self.root / 'config/config.json',
            'tm_cache_dir': self.workspace / 'tm_cache' / language_pair(project).key,
            'batch_dir': str(self._batch_dir(project['id'], batch['id'])),
            'batch_name': batch.get('name', ''),
        }

    def _input_for(self, project_id: str, batch: dict, key: str) -> Path:
        folder = self._batch_dir(project_id, batch['id'])
        if key == 'normalize':
            return folder / batch['source']
        index = ORDER.index(key)
        if index == 0:
            return folder / batch['source']
        previous = folder / ARTIFACTS[ORDER[index - 1]]
        if not previous.is_file():
            raise PipelineError(f'前置节点产物不存在：{ARTIFACTS[ORDER[index - 1]]}')
        return previous

    def start(self, project_id: str, batch_id: str, stage: str | None = None,
              mode: str = 'run', then_continue: bool = False) -> dict:
        with self.lock:
            store = self._read_store()
            project = self._project(store, project_id)
            batch = self._load_batch(project['id'], batch_id)
            if not batch:
                raise PipelineError('批次不存在')
            if batch.get('running') or batch_id in self.running:
                raise PipelineError('这个批次正在执行，请等当前节点结束')
            if mode == 'run' and any(stage.get('state') == 'review'
                                     for stage in batch['stages'].values()):
                raise PipelineError('还有节点等待确认，请先确认')
            # 已登记但丢失的文件仍须报错；未上传的库允许跳过。
            self._project_term(project)
            self._project_tm(project)
            key_rules = self._key_rules_path()
            if not key_rules.is_file():
                raise PipelineError('缺少分类规则 profiles/game/wc_key_rules.json')

            if mode == 'rerun':
                if not stage:
                    raise PipelineError('缺少要重跑的节点')
                if stage not in STAGE_BY_KEY:
                    raise PipelineError('未知节点')
                if stage == 'context':
                    cache = self._batch_dir(project['id'], batch_id) / 'context_scenes'
                    if cache.exists():
                        shutil.rmtree(cache)
                self._invalidate_after(batch, stage)
                start_key = stage
                chain = False
                auto_confirm = False
            elif mode == 'retry':
                if not stage or stage not in STAGE_BY_KEY:
                    raise PipelineError('缺少要重试的节点')
                if batch['stages'][stage].get('state') != 'failed':
                    raise PipelineError('只有失败的节点可以重试')
                self._invalidate_after(batch, stage)
                start_key = stage
                chain = bool(then_continue)
                auto_confirm = bool(then_continue)
            else:
                if stage:
                    if stage not in STAGE_BY_KEY:
                        raise PipelineError('未知节点')
                    index = ORDER.index(stage)
                    if any(batch['stages'][ORDER[i]].get('state') != 'done' for i in range(index)):
                        raise PipelineError('前置节点还没有完成')
                    if batch['stages'][stage].get('state') not in ('pending', 'failed'):
                        raise PipelineError('这个节点不需要重新开始')
                    start_key = stage
                else:
                    review = next((key for key in ORDER if batch['stages'][key].get('state') == 'review'), None)
                    if review:
                        raise PipelineError('还有节点等待确认，请先确认')
                    start_key = next((key for key in ORDER if batch['stages'][key].get('state') == 'pending'), None)
                    if not start_key:
                        raise PipelineError('没有待执行的节点')
                chain = True
                auto_confirm = False

            batch['running'] = True
            batch['active_stage'] = start_key
            for key in ORDER:
                if batch['stages'][key].get('state') == 'running':
                    batch['stages'][key].update(state='pending', progress=0, status='')
            self._save_batch(project['id'], batch)
            runtime = {'task': {}, 'started_at': time.time()}
            self.running[batch_id] = runtime
            thread = threading.Thread(target=self._worker,
                                      args=(project['id'], batch_id, start_key, chain, auto_confirm),
                                      daemon=True)
            runtime['thread'] = thread
            thread.start()
            return self._batch_view(project['id'], batch)

    def confirm(self, project_id: str, batch_id: str) -> dict:
        with self.lock:
            store = self._read_store()
            project = self._project(store, project_id)
            batch = self._load_batch(project['id'], batch_id)
            if not batch:
                raise PipelineError('批次不存在')
            if batch.get('running') or batch_id in self.running:
                raise PipelineError('当前节点还在执行，请稍等')
            review = next((key for key in ORDER if batch['stages'][key].get('state') == 'review'), None)
            if not review:
                raise PipelineError('没有等待确认的节点')
            batch['stages'][review].update(state='done', confirmed_at=now(), progress=100)
            next_key = next((key for key in ORDER if batch['stages'][key].get('state') == 'pending'), None)
            if next_key:
                batch['running'] = True
                batch['active_stage'] = next_key
                self._save_batch(project['id'], batch)
                runtime = {'task': {}, 'started_at': time.time()}
                self.running[batch_id] = runtime
                thread = threading.Thread(target=self._worker,
                                          args=(project['id'], batch_id, next_key, True, False),
                                          daemon=True)
                runtime['thread'] = thread
                thread.start()
            else:
                batch['running'] = False
                batch['active_stage'] = None
                self._save_batch(project['id'], batch)
            return self._batch_view(project['id'], batch)

    def reset(self, project_id: str, batch_id: str) -> dict:
        with self.lock:
            store = self._read_store()
            project = self._project(store, project_id)
            batch = self._load_batch(project['id'], batch_id)
            if not batch:
                raise PipelineError('批次不存在')
            if batch.get('running') or batch_id in self.running:
                raise PipelineError('正在执行，不能重置')
            folder = self._batch_dir(project['id'], batch_id)
            for artifact in ARTIFACTS.values():
                (folder / artifact).unlink(missing_ok=True)
            (folder / '06-自检报告.md').unlink(missing_ok=True)
            batch['stages'] = {key: {'state': 'pending', 'progress': 0, 'status': ''} for key in ORDER}
            batch['running'] = False
            batch['active_stage'] = None
            self._save_batch(project['id'], batch)
            return self._batch_view(project['id'], batch)

    def _invalidate_after(self, batch: dict, stage: str):
        index = ORDER.index(stage)
        folder = self._batch_dir(batch['project_id'], batch['id'])
        for key in ORDER[index + 1:]:
            batch['stages'][key] = {'state': 'pending', 'progress': 0, 'status': ''}
            (folder / ARTIFACTS[key]).unlink(missing_ok=True)
        batch['stages'][stage] = {'state': 'pending', 'progress': 0, 'status': ''}
        if stage in ('context', 'tm'):
            # Model reruns must not expose the previous output after a failure.
            (folder / ARTIFACTS[stage]).unlink(missing_ok=True)

    def _worker(self, project_id: str, batch_id: str, start_key: str,
                chain: bool, auto_confirm: bool):
        try:
            keys = ORDER[ORDER.index(start_key):]
            first = True
            for key in keys:
                with self.lock:
                    batch = self._load_batch(project_id, batch_id)
                    if not batch:
                        return
                    stage = batch['stages'][key]
                    if stage.get('state') not in ('pending', 'failed'):
                        break
                    stage.update(state='running', progress=0, status='正在准备…', message='')
                    batch['running'] = True
                    batch['active_stage'] = key
                    self._save_batch(project_id, batch)
                    runtime = self.running.get(batch_id)
                    task = runtime['task'] if runtime else {}
                    task.clear()
                try:
                    self._run_one(project_id, batch_id, key, task,
                                  auto_confirm=auto_confirm if first else False)
                    first = False
                except Exception as error:  # noqa: BLE001 — 单节点失败要落盘，不把线程打崩
                    with self.lock:
                        batch = self._load_batch(project_id, batch_id) or batch
                        batch['stages'][key].update(state='failed', message=str(error),
                                                    progress=100, status=str(error))
                        batch['running'] = False
                        batch['active_stage'] = key
                        self._save_batch(project_id, batch)
                    return
                with self.lock:
                    batch = self._load_batch(project_id, batch_id)
                    state = batch['stages'][key].get('state')
                if state == 'failed' or state == 'review':
                    break
                if not chain:
                    break
            with self.lock:
                batch = self._load_batch(project_id, batch_id)
                if batch:
                    batch['running'] = False
                    if not any(stage.get('state') == 'review' for stage in batch['stages'].values()):
                        batch['active_stage'] = None
                    self._save_batch(project_id, batch)
        finally:
            self.running.pop(batch_id, None)

    def _run_one(self, project_id: str, batch_id: str, key: str, task: dict,
                 auto_confirm: bool = False):
        with self.lock:
            store = self._read_store()
            project = self._project(store, project_id)
            batch = self._load_batch(project_id, batch_id)
            if not batch:
                raise PipelineError('批次不存在')
            ctx = self._ctx(project, batch)
            source = self._input_for(project_id, batch, key)
            output = self._batch_dir(project_id, batch_id) / ARTIFACTS[key]
            stage = batch['stages'][key]
            stage.update(state='running', progress=0, status='正在准备…', message='')
            self._save_batch(project_id, batch)
        started = time.time()
        RUNNERS[key](ctx, source, output, task)
        seconds = round(time.time() - started, 1)
        with self.lock:
            batch = self._load_batch(project_id, batch_id)
            stage = batch['stages'][key]
            state = task.get('state', 'completed')
            message = task.get('status', '')
            progress = task.get('progress', 100)
            if state == 'error':
                stage.update(state='failed', message=message or '节点执行失败', progress=100,
                             status=message, seconds=seconds)
                # 失败时也保留已经生成的预览/自检结果，便于排查。
                try:
                    stage.update(self._summarize(key, project_id, batch, key, task, source, output, seconds))
                except Exception:
                    pass
                batch['running'] = False
                batch['active_stage'] = key
                self._save_batch(project_id, batch)
                return
            summary = self._summarize(key, project_id, batch, key, task, source, output, seconds)
            confirm = STAGE_BY_KEY[key]['confirm'] and not auto_confirm and not task.get('skipped')
            stage.update(state='review' if confirm else 'done', progress=100,
                         skipped=bool(task.get('skipped')),
                         message=message, status=message, seconds=seconds,
                         confirmed_at=None if confirm else now())
            stage.update(summary)
            batch['running'] = True
            batch['active_stage'] = key
            self._save_batch(project_id, batch)

    # ------------------------------------------------------------ 预览

    def _summarize(self, key: str, project_id: str, batch: dict, stage_key: str,
                   task: dict, source: Path, output: Path, seconds: float) -> dict:
        if key == 'preflight':
            review = task.get('review', [])
            return {'stats': task.get('stats', []), 'preview': task.get('preview', {}),
                    'checks': review, 'review': review}
        try:
            header, rows = workbook_rows(output)
        except Exception:
            header, rows = [], []
        total = len(rows)
        pair = language_pair(batch.get('language') or self._project(self._read_store(), project_id))
        if task.get('skipped'):
            reason = task['skip_reason']
            return {'stats': [['已跳过', total, ' 行']], 'desc': reason,
                    'preview': {'type': 'code', 'cap': '未提供资料或未启用，已跳过', 'text': reason},
                    'checks': [reason], 'review': []}
        if key == 'normalize':
            try:
                source_rows = self._inspect_source(source, pair)[0]
            except Exception:
                source_rows = 0
            stats = [['输入行数', source_rows, ''], ['删除空列', task.get('trimmed_columns', 0), ''],
                     ['补出 speaker', task.get('speaker_rows', 0), ''], ['输出列', len(header), '']]
            sample = []
            for row in rows[:4]:
                sample.append([text(column(row, 'limit')), text(column(row, 'key')),
                               snippet(column(row, *pair.aliases('source')), 80), text(column(row, 'speaker'))])
            preview = {'type': 'table', 'cap': '产物前 4 行（新增 speaker 列）',
                       'head': ['limit', 'key', pair.aliases('source')[0], 'speaker'], 'rows': sample}
            return {'stats': stats, 'preview': preview}
        if key == 'classify':
            counts = task.get('counts', {})
            total_rows = sum(counts.values()) or total
            stats = [['总行数', total_rows, ''], ['活动对白', counts.get('活动对白', 0), ''],
                     ['集卡名称', counts.get('集卡名称', 0), ''], ['待确认', counts.get('待确认', 0), '']]
            bars = [[name, counts.get(name, 0), CATEGORY_COLORS.get(name, '#64748b')]
                    for name in ('集卡名称', '活动对白', '活动 UI 与商店文案', '装饰物名称')]
            return {'stats': stats, 'preview': {'type': 'bars', 'cap': '四类分布', 'rows': bars}}
        if key == 'context':
            with_context = [row for row in rows if text(column(row, 'context')).strip()]
            dialogue = [row for row in rows if text(column(row, '分类')).strip() == '活动对白']
            empty_dialogue = sum(1 for row in dialogue if not text(column(row, 'context')).strip())
            scenes = task.get('dialogue_scenes', [])
            cluster = max((r.get('related_count', 0) for r in task.get('records', [])), default=0)
            stats = [['有 Context', len(with_context), ' 行'], ['无 Context', total - len(with_context), ' 行'],
                     ['活动对白', len(dialogue), ' 行'], ['场景分析', len(scenes), ' 组']]
            sample = next((row for row in with_context), None)
            if sample:
                preview_text = f"{text(column(sample, 'key'))}\n{text(column(sample, 'context'))[:900]}"
            else:
                preview_text = '这一批没有生成 Context。'
            checks = [f'最大关联组 {cluster} 条，请核对是否跨活动或跨段误关联',
                      f'对白 {len(dialogue) - empty_dialogue} 条有 Context，{empty_dialogue} 条为空；空值需核对完整 Key 和同段原文',
                      '对白 Context 是模型场景判断，请展开对白和人物资料依据核对语气、指代及叙事方式']
            if scenes:
                labels = [m['label'] for m in scenes[0].get('materials', [])]
                checks.append('本次全文读取资料：' + ('、'.join(labels) if labels else '未提供人物或背景资料'))
            preview = ({'type': 'dialogue_context', 'cap': '对白场景与原文依据', 'scenes': scenes}
                       if scenes else {'type': 'code', 'cap': '一条 Context 样例', 'text': preview_text})
            notices = [f"{r['sheet']} 第{r['row']}行：{r['warning']}" for r in task.get('records', [])
                       if r['category'] == '活动对白' and r.get('status') == '待确认']
            checks.extend(notices[:5])
            return {'stats': stats, 'preview': preview,
                    'checks': checks, 'review': checks}
        if key == 'term':
            hit = [row for row in rows if text(column(row, '关联术语')).strip()]
            terms = sum(len([part for part in text(column(row, '关联术语')).split(' | ') if part.strip()])
                        for row in hit)
            coverage = (len(hit) / total * 100) if total else 0
            stats = [['命中行数', len(hit), ''], ['命中术语', terms, ''], ['未命中', total - len(hit), ''],
                     ['覆盖率', f'{coverage:.1f}', '%']]
            sample = [[snippet(column(row, *pair.aliases('source')), 50), snippet(column(row, '关联术语'), 80)]
                      for row in hit[:4]]
            preview = {'type': 'table', 'cap': '命中样例（关联术语列）',
                       'head': ['源文片段', '关联术语'], 'rows': sample}
            return {'stats': stats, 'preview': preview}
        if key == 'term_filter':
            active = [row for row in rows if text(column(row, '生效术语')).strip()]
            excluded = [row for row in rows if text(column(row, '无效术语')).strip()]
            active_terms = sum(len([part for part in text(column(row, '生效术语')).split(' | ') if part.strip()])
                               for row in active)
            excluded_terms = sum(len([part for part in text(column(row, '无效术语')).split(' | ') if part.strip()])
                                 for row in excluded)
            considered = active_terms + excluded_terms
            ratio = (excluded_terms / considered * 100) if considered else 0
            stats = [['生效术语', active_terms, ''], ['无效术语', excluded_terms, ''],
                     ['判掉比例', f'{ratio:.1f}', '%'], ['耗时', f'{seconds:.1f}', 's']]
            sample = [[snippet(column(row, *pair.aliases('source')), 50), snippet(column(row, '生效术语'), 55),
                       snippet(column(row, '无效术语'), 55)] for row in (active + excluded)[:4]]
            preview = {'type': 'table', 'cap': '筛选结果（两列分开存放）',
                       'head': ['源文片段', '生效术语', '无效术语'], 'rows': sample}
            checks = [f'排除 {excluded_terms} 个（{ratio:.1f}%），请结合本批语境核对',
                      '抽 2 行看有没有把真术语判掉',
                      '若有整行候选全被判掉，需要回看原文']
            return {'stats': stats, 'preview': preview, 'checks': checks, 'review': checks}
        if key == 'tm':
            counts = task.get('counts', {})
            db = task.get('tm_db', {})
            if db.get('enabled') is False:
                checks = ['本项目已选择不使用 TM，参考列留空，未执行检索']
                return {'stats': [['未使用 TM', counts.get('未使用', total), ' 行'],
                                  ['待确认', counts.get('待确认', 0), ' 行']],
                        'desc': '本项目不使用 TM，保留空的 TM参考 列供后续翻译使用。',
                        'preview': {'type': 'code', 'cap': 'TM 未启用', 'text': checks[0]},
                        'checks': checks, 'review': checks}
            stats = [['有参考', task.get('matched_count', 0), ' 行'],
                     ['原文一致', counts.get('原文一致', 0), ' 行'],
                     ['未命中', counts.get('未命中', 0), ' 行'],
                     ['记忆库', f'{db.get("tu_count", 0):,}' if db.get('tu_count') else '—', ' 对']]
            sample = next((row for row in rows if '原文一致' in text(column(row, 'TM参考'))), None)
            if sample is None:
                sample = next((row for row in rows if text(column(row, 'TM参考')).strip()), None)
            preview = {'type': 'code', 'cap': '一条 TM 参考样例',
                       'text': text(column(sample, 'TM参考'))[:900] if sample else '这一批没有 TM 参考。'}
            exact = counts.get('原文一致', 0)
            matched = task.get('matched_count', 0)
            checks = [f'{exact} 行原文一致，仍需核对语境、术语和客户要求',
                      f'{matched - exact} 行有其他历史参考；语义相近不代表译文可直接照抄',
                      f'{counts.get("待确认", 0)} 行存在差异或冲突，请检查 TM参考 中的提示',
                      '当前检索依据源文原句；Context、分类和生效术语由后续翻译环节结合判断']
            return {'stats': stats, 'preview': preview, 'checks': checks, 'review': checks}
        return {'stats': [], 'preview': {}}

    # ------------------------------------------------------------ 产物

    def final_artifact(self, project_id: str, batch_id: str):
        """最终交付物：翻译前准备完成的 05-TM.xlsx，而不是自检报告。"""
        with self.lock:
            state = self._load_batch(project_id, batch_id)
            if not state:
                raise PipelineError('批次不存在')
            path = self._batch_dir(project_id, batch_id) / ARTIFACTS['tm']
            if state.get('stages', {}).get('tm', {}).get('state') not in ('review', 'done'):
                raise PipelineError('TM 节点尚未成功完成，暂不能下载最终产物')
            if not path.is_file():
                raise PipelineError('最终产物还没有生成，请先完成 TM 参考节点')
            stem = Path(state.get('name') or 'batch').stem
            return path, f'{stem}_ready.xlsx'

    def artifact_path(self, project_id: str, batch_id: str, stage_key: str) -> Path:
        if stage_key not in STAGE_BY_KEY:
            raise PipelineError('未知节点')
        with self.lock:
            state = self._load_batch(project_id, batch_id)
            if not state:
                raise PipelineError('批次不存在')
            path = self._batch_dir(project_id, batch_id) / ARTIFACTS[stage_key]
            if not path.is_file():
                raise PipelineError('产物还没有生成')
            return path


service = PipelineService()
