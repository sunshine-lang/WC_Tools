import json
import shutil
import tempfile
import time
import unittest
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook

import pipeline_service
from pipeline_service import PipelineService
from core_tools.term_scanner import TermScanner
from core_tools.tm_matcher import TMMatcher


ROOT = Path(__file__).resolve().parents[1]


class Upload:
    def __init__(self, path, name=None):
        self.path = Path(path)
        self.filename = name or self.path.name

    def save(self, destination):
        shutil.copyfile(self.path, destination)


class PipelineServiceStateTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.workspace = Path(self.tmp.name)
        self.termbase = self.workspace / 'terms.csv'
        self.termbase.write_text('EN,DE\nHello,Hallo\nWorld,Welt\n', encoding='utf-8')
        self.tmx = self.workspace / 'memory.tmx'
        self.tmx.write_text('<tmx version="1.4"><header srclang="en-US"/><body><tu>'
                            '<tuv xml:lang="en-US"><seg>Hello</seg></tuv>'
                            '<tuv xml:lang="de-DE"><seg>Hallo</seg></tuv>'
                            '</tu></body></tmx>', encoding='utf-8')
        store = {
            'version': 1,
            'active_project': 'p',
            'projects': {
                'p': {
                    'id': 'p', 'name': '测试项目', 'src': 'en', 'tgt': 'de', 'game': 'wc',
                    'created_at': '2026-10-20 20:00:00',
                    'library': {
                        'termbase': {'path': str(self.termbase), 'name': self.termbase.name, 'meta': ''},
                        'tm': {'path': str(self.tmx), 'name': self.tmx.name, 'meta': ''},
                        'cast': None, 'world': None, 'style': None,
                    },
                    'active_batch': None,
                }
            },
        }
        (self.workspace / 'projects.json').write_text(json.dumps(store, ensure_ascii=False), 'utf-8')
        self.service = PipelineService(workspace=self.workspace)

        source = self.workspace / 'source.xlsx'
        book = Workbook()
        sheet = book.active
        sheet.append(('limit', 'key', 'game_text_en', 'game_text_de', 'speaker'))
        sheet.append((20, 'a', 'Hello', None, ''))
        sheet.append((20, 'b', 'World', None, ''))
        book.save(source)
        book.close()
        self.batch = self.service.create_batch('p', Upload(source))['batch']

        self.calls = {'term_filter': 0}
        self.original_runners = dict(pipeline_service.RUNNERS)
        for key in pipeline_service.ORDER:
            pipeline_service.RUNNERS[key] = self._runner(key)

    def tearDown(self):
        pipeline_service.RUNNERS.clear()
        pipeline_service.RUNNERS.update(self.original_runners)
        self.tmp.cleanup()

    def _runner(self, key):
        def run(ctx, source, target, task):
            target = Path(target)
            if key == 'preflight':
                target.write_text('# 自检\n', encoding='utf-8')
                task.update(state='completed', status='自检通过', stats=[['断言', 9, ' 条']],
                            preview={'type': 'checks', 'cap': '断言结果', 'rows': []},
                            checks=[], review=['阻断项 0，可以交付'])
                return
            shutil.copyfile(source, target)
            if key == 'term_filter' and self.calls['term_filter'] == 0:
                self.calls['term_filter'] += 1
                task.update(state='error', status='模型连接失败')
                return
            task.update(state='completed', status='完成')
        return run

    def wait(self, timeout=10):
        deadline = time.time() + timeout
        while time.time() < deadline:
            state = self.service.batch_state('p', self.batch['id'])
            if not state['running']:
                return state
            time.sleep(0.03)
        self.fail('批次执行超时')

    def test_review_retry_auto_confirm_and_completion(self):
        started = self.service.start('p', self.batch['id'])
        self.assertTrue(started['running'])
        state = self.wait()
        self.assertEqual([s['state'] for s in state['stages'][:3]], ['done', 'done', 'review'])
        self.assertEqual(state['active_stage'], 'context')

        self.service.confirm('p', self.batch['id'])
        state = self.wait()
        self.assertEqual(state['stages'][3]['state'], 'done')
        self.assertEqual(state['stages'][4]['state'], 'failed')
        self.assertIn('模型连接失败', state['stages'][4]['message'])

        self.service.start('p', self.batch['id'], stage='term_filter', mode='retry', then_continue=True)
        state = self.wait()
        self.assertEqual(state['stages'][4]['state'], 'done')
        self.assertEqual(state['stages'][5]['state'], 'review')
        self.assertEqual(state['active_stage'], 'tm')
        self.assertTrue(state['final_download_url'])
        self.assertEqual(state['final_name'], 'source_ready.xlsx')
        final_path, final_name = self.service.final_artifact('p', self.batch['id'])
        self.assertTrue(final_path.is_file())
        self.assertEqual(final_name, state['final_name'])

        self.assertEqual(len(state['stages']), 6)
        self.service.confirm('p', self.batch['id'])
        state = self.wait()
        self.assertTrue(all(s['state'] == 'done' for s in state['stages']))
        self.assertFalse(state['running'])

        old_id = self.batch['id']
        cloned = self.service.clone_batch('p', old_id)
        new_id = cloned['batch']['id']
        self.assertNotEqual(new_id, old_id)
        self.assertEqual(len(cloned['batch']['stages']), 6)
        self.assertTrue(all(stage['state'] == 'pending' for stage in cloned['batch']['stages']))
        self.assertTrue((self.workspace / 'projects/p/batches' / new_id / 'source.xlsx').is_file())
        old = self.service.batch_state('p', old_id)
        self.assertTrue(all(stage['state'] == 'done' for stage in old['stages']))

    def test_rerun_invalidates_downstream_and_reset_clears_artifacts(self):
        self.service.start('p', self.batch['id'])
        self.wait()
        self.service.confirm('p', self.batch['id'])
        self.wait()
        state = self.service.batch_state('p', self.batch['id'])
        self.assertEqual(state['stages'][3]['state'], 'done')

        self.service.start('p', self.batch['id'], stage='classify', mode='rerun')
        state = self.wait()
        self.assertEqual(state['stages'][1]['state'], 'done')
        self.assertTrue(all(s['state'] == 'pending' for s in state['stages'][2:]))
        self.assertFalse((self.workspace / 'projects/p/batches' / self.batch['id'] / '02-上下文.xlsx').exists())

        state = self.service.reset('p', self.batch['id'])
        self.assertTrue(all(s['state'] == 'pending' for s in state['stages']))
        self.assertFalse((self.workspace / 'projects/p/batches' / self.batch['id'] / '01-分类.xlsx').exists())

    def test_fresh_store_requires_explicit_glossary_and_tm_uploads(self):
        fresh = PipelineService(root=self.workspace, workspace=self.workspace / 'fresh')
        data = fresh.bootstrap()
        slots = {item['key']: item for item in data['project']['library']}
        self.assertFalse(slots['termbase']['ready'])
        self.assertFalse(slots['tm']['ready'])
        self.assertEqual(slots['cast']['name'], '人物人设')
        self.assertEqual(slots['relationships']['name'], '人物关系')
        self.assertFalse(slots['cast']['ready'])
        self.assertFalse(slots['relationships']['ready'])
        self.assertFalse((self.workspace / 'uploads/knowledge_db').exists())

    def test_classification_rules_are_available_without_upload_history(self):
        self.assertTrue(self.service._server_status()['key_rules_ready'])
        project = self.service._read_store()['projects']['p']
        context = self.service._ctx(project, self.batch)
        self.assertEqual(Path(context['game']['key_rules']), ROOT / 'profiles/game/wc_key_rules.json')

    def test_invalid_library_upload_preserves_both_active_libraries(self):
        before = self.service.projects_file.read_bytes()
        for slot, contents, name in (
            ('termbase', 'CN,EN\nwrong,columns\n', 'invalid.csv'),
            ('tm', '<tmx><body/></tmx>', 'empty.tmx'),
        ):
            path = self.workspace / name
            path.write_text(contents, encoding='utf-8')
            with self.assertRaises(pipeline_service.PipelineError):
                self.service.upload_library('p', slot, Upload(path))
            self.assertEqual(self.service.projects_file.read_bytes(), before)
            self.assertFalse(list((self.workspace / 'projects/p/library' / slot).iterdir()))

    def test_library_replacement_is_persisted_and_keeps_loaded_snapshots(self):
        old_term = TermScanner(self.termbase)
        old_tm = TMMatcher(self.tmx)
        replacement_term = self.workspace / 'replacement.csv'
        replacement_term.write_text('EN,DE\nHello,Guten Tag\n', encoding='utf-8')
        replacement_tm = self.workspace / 'replacement.tmx'
        replacement_tm.write_text(self.tmx.read_text().replace('Hallo', 'Guten Tag'), encoding='utf-8')
        self.service.upload_library('p', 'termbase', Upload(replacement_term))
        self.service.upload_library('p', 'tm', Upload(replacement_tm))
        reloaded = PipelineService(workspace=self.workspace)
        project = reloaded._read_store()['projects']['p']
        new_term = TermScanner(reloaded._project_term(project))
        new_tm = TMMatcher(reloaded._project_tm(project))
        self.assertEqual(old_term.scan_text('Hello'), 'Hello (Hallo)')
        self.assertEqual(new_term.scan_text('Hello'), 'Hello (Guten Tag)')
        self.assertEqual(old_tm.find_matches('Hello')[0]['german'], 'Hallo')
        self.assertEqual(new_tm.find_matches('Hello')[0]['german'], 'Guten Tag')

    def test_pipeline_http_upload_and_bootstrap_use_project_library(self):
        import io
        import app as server
        import pipeline_api
        with patch.object(pipeline_api, 'service', self.service):
            client = server.app.test_client()
            response = client.post('/api/pipeline/projects/p/library/termbase',
                                   data={'file': (io.BytesIO(self.termbase.read_bytes()), 'terms.csv')})
            self.assertEqual(response.status_code, 200, response.json)
            response = client.get('/api/pipeline/bootstrap?project=p')
            self.assertEqual(response.status_code, 200)
            self.assertEqual(response.json['batch']['id'], self.batch['id'])
            slots = {item['key']: item for item in response.json['project']['library']}
            self.assertTrue(slots['termbase']['ready'])
            self.assertTrue(slots['tm']['ready'])
            self.assertEqual(slots['termbase']['file'], 'terms.csv')

    def test_failed_tm_rerun_cannot_download_old_ready(self):
        batch = self.service._load_batch('p', self.batch['id'])
        for stage in batch['stages'].values():
            stage['state'] = 'done'
        folder = self.service._batch_dir('p', self.batch['id'])
        prior = folder / pipeline_service.ARTIFACTS['term_filter']
        old_ready = folder / pipeline_service.ARTIFACTS['tm']
        prior.write_bytes(b'KEEP_PREVIOUS_STAGE')
        old_ready.write_bytes(b'OLD_READY')
        self.service._invalidate_after(batch, 'tm')
        self.service._save_batch('p', batch)
        self.assertEqual(prior.read_bytes(), b'KEEP_PREVIOUS_STAGE')
        self.assertFalse(old_ready.exists())
        # Even a stale file left externally must not bypass the stage state.
        old_ready.write_bytes(b'STALE_READY')
        with self.assertRaisesRegex(pipeline_service.PipelineError, '尚未成功'):
            self.service.final_artifact('p', self.batch['id'])

    def test_context_rerun_removes_old_summary_but_keeps_classification(self):
        batch = self.service._load_batch('p', self.batch['id'])
        folder = self.service._batch_dir('p', self.batch['id'])
        prior = folder / pipeline_service.ARTIFACTS['classify']
        old_context = folder / pipeline_service.ARTIFACTS['context']
        prior.write_bytes(b'KEEP_CLASSIFICATION')
        old_context.write_bytes(b'OLD_SCENE_SUMMARY')
        self.service._invalidate_after(batch, 'context')
        self.service._save_batch('p', batch)
        self.assertEqual(prior.read_bytes(), b'KEEP_CLASSIFICATION')
        self.assertFalse(old_context.exists())
        with self.assertRaises(pipeline_service.PipelineError):
            self.service.artifact_path('p', batch['id'], 'context')

    def test_context_preview_keeps_full_scene_and_both_kinds_of_evidence(self):
        folder = self.service._batch_dir('p', self.batch['id'])
        source = folder / 'source.xlsx'
        scene = {'group': 'forestconcert/dialogue/3',
                 'context': '【场景结构】\n' + '完整场景。' * 200 + '\n\n【情境与基调】\n可核对的语气说明。',
                 'sources': [{'key': 'scene_1', 'english': 'Hello', 'speaker': 'Max', 'emotion': '', 'bubble_type': ''}],
                 'references': [], 'evidence': [{'key': 'scene_1', 'quote': 'Hello'}],
                 'materials': [{'key': 'cast', 'label': '人物人设'}],
                 'material_evidence': [{'material': 'cast', 'quote': '人物资料。'}], 'warnings': []}
        summary = self.service._summarize('context', 'p', self.batch, 'context',
                                         {'dialogue_scenes': [scene]}, source, source, 1)
        self.assertEqual(summary['preview']['type'], 'dialogue_context')
        self.assertEqual(summary['preview']['scenes'][0], scene)
        self.assertTrue(any('人物人设' in c for c in summary['checks']))

    def test_context_retry_keeps_cache_but_explicit_rerun_clears_it(self):
        batch = self.service._load_batch('p', self.batch['id'])
        folder = self.service._batch_dir('p', self.batch['id'])
        shutil.copyfile(folder / 'source.xlsx', folder / pipeline_service.ARTIFACTS['classify'])
        for key in ('normalize', 'classify'):
            batch['stages'][key]['state'] = 'done'
        batch['stages']['context']['state'] = 'failed'
        self.service._save_batch('p', batch)
        cache = folder / 'context_scenes'
        cache.mkdir()
        saved = cache / 'scene.json'
        saved.write_text('{}')
        self.service.start('p', batch['id'], stage='context', mode='retry')
        self.wait()
        self.assertTrue(saved.exists())
        self.service.start('p', batch['id'], stage='context', mode='rerun')
        self.wait()
        self.assertFalse(cache.exists())

    def test_legacy_combined_cast_is_preserved_and_migration_is_idempotent(self):
        old_file = self.workspace / '旧合并资料.md'
        old_file.write_text('# 原合并人物资料\n保持完整内容。', encoding='utf-8')
        store = self.service._read_store()
        store['version'] = 1
        library = store['projects']['p']['library']
        old_item = {'path': str(old_file), 'name': old_file.name, 'meta': '2 行'}
        library['cast'] = dict(old_item)
        library.pop('relationships', None)
        self.service._write_store(store)
        migrated = PipelineService(workspace=self.workspace)
        after = migrated._read_store()
        self.assertEqual(after['version'], pipeline_service.STORE_VERSION)
        self.assertEqual(after['projects']['p']['library']['cast'], old_item | {'legacy_combined': True})
        self.assertIsNone(after['projects']['p']['library']['relationships'])
        self.assertEqual(old_file.read_text(), '# 原合并人物资料\n保持完整内容。')
        slots = {item['key']: item for item in migrated.bootstrap('p')['project']['library']}
        self.assertIn('合并文件', slots['cast']['notice'])
        self.assertFalse(slots['relationships']['ready'])
        before_reload = migrated.projects_file.read_bytes()
        PipelineService(workspace=self.workspace)
        self.assertEqual(migrated.projects_file.read_bytes(), before_reload)
        replacement = self.workspace / '新人物人设.md'
        replacement.write_text('角色的新人物人设。', encoding='utf-8')
        data = migrated.upload_library('p', 'cast', Upload(replacement))
        self.assertFalse(next(x for x in data['project']['library'] if x['key'] == 'cast')['notice'])
        self.assertTrue(old_file.exists())

    def test_character_uploads_are_independent_persisted_and_reach_pipeline(self):
        import io
        import app as server
        import pipeline_api
        profiles = '# 人物人设\n' + '完整人设内容。' * 300 + '人物人设末尾'
        relationships = '# 人物关系\n' + '完整关系内容。' * 300 + '人物关系末尾'
        with patch.object(pipeline_api, 'service', self.service):
            client = server.app.test_client()
            for slot, name, content in (('cast', '人物人设.md', profiles),
                                        ('relationships', '人物关系.txt', relationships)):
                response = client.post(f'/api/pipeline/projects/p/library/{slot}',
                                       data={'file': (io.BytesIO(content.encode()), name)})
                self.assertEqual(response.status_code, 200, response.json)
            slots = {item['key']: item for item in response.json['project']['library']}
            self.assertEqual(slots['cast']['file'], '人物人设.md')
            self.assertEqual(slots['relationships']['file'], '人物关系.txt')
            reloaded = PipelineService(workspace=self.workspace)
            project = reloaded._read_store()['projects']['p']
            materials = reloaded._ctx(project, self.batch)['materials']
            self.assertEqual(materials['cast'], profiles)
            self.assertEqual(materials['relationships'], relationships)
            before = self.service.projects_file.read_bytes()
            invalid = client.post('/api/pipeline/projects/p/library/relationships',
                                  data={'file': (io.BytesIO(b'\xff'), 'broken.txt')})
            self.assertEqual(invalid.status_code, 400)
            self.assertEqual(self.service.projects_file.read_bytes(), before)
            removed = client.delete('/api/pipeline/projects/p/library/relationships')
            self.assertEqual(removed.status_code, 200, removed.json)
            slots = {item['key']: item for item in removed.json['project']['library']}
            self.assertTrue(slots['cast']['ready'])
            self.assertFalse(slots['relationships']['ready'])
            self.assertEqual(self.service._materials(self.service._read_store()['projects']['p'])['cast'], profiles)


if __name__ == '__main__':
    unittest.main()
