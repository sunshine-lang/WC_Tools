import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

from core_tools.preflight import preflight, process


HEADER = ('limit', 'key', 'game_text_en', 'game_text_de', 'speaker')
FINAL_HEADER = HEADER + ('分类', 'Context', '生效术语', '无效术语', 'TM参考')


def write_book(path, rows, header=FINAL_HEADER):
    book = Workbook()
    sheet = book.active
    sheet.append(header)
    for row in rows:
        sheet.append(row)
    book.save(path)
    book.close()


class PreflightTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)
        self.source = self.root / '00-规范化.xlsx'
        self.final = self.root / '05-TM.xlsx'
        write_book(self.source, [
            (100, 'a', 'Complete the event!', None, ''),
            (14, 'b', 'Feline Frenzy!', None, ''),
        ], HEADER)
        write_book(self.final, [
            (100, 'a', 'Complete the event!', None, '', '活动 UI 与商店文案', 'ctx', 'event (Event)', '', ''),
            (14, 'b', 'Feline Frenzy!', None, '', '集卡名称', '', '', '',
             '[TM参考 1｜原文一致]\n【英】Feline Frenzy!\n【德】Katzenwahn!'),
        ])

    def tearDown(self):
        self.tmp.cleanup()

    def test_passes_complete_preparation_and_reports_length_warning(self):
        result = preflight(self.source, self.final, 'batch')
        checks = {check['name']: check for check in result['checks']}
        self.assertEqual(len(result['checks']), 9)
        self.assertTrue(checks['行数守恒']['ok'])
        self.assertTrue(checks['源文未改写']['ok'])
        self.assertTrue(checks['无悬空引用']['ok'])
        self.assertEqual(result['exact'], 1)
        self.assertEqual(result['blocked'], 0)
        self.assertEqual(result['passed'], 8)
        self.assertEqual(result['warnings'], 1)
        self.assertIn('长度告警 1 条', '\n'.join(result['review']))
        self.assertIn('| 告警 | 长度余量', result['report'])

    def test_blocks_when_target_or_source_changed(self):
        write_book(self.final, [
            (100, 'a', 'Complete the event!', 'Fertig!', '', '活动 UI 与商店文案', '', '', '', ''),
            (14, 'b', 'Feline Frenzy!', None, '', '集卡名称', '', '', '', ''),
        ])
        result = preflight(self.source, self.final, 'batch')
        checks = {check['name']: check for check in result['checks']}
        self.assertFalse(checks['目标列未写入']['ok'])
        self.assertGreaterEqual(result['blocked'], 1)

    def test_process_writes_markdown_and_stops_on_block(self):
        task = {}
        process(self.source, self.final, {'t': task}, 't', 'batch')
        self.assertEqual(task['state'], 'completed')
        report = self.root / '06-自检报告.md'
        self.assertTrue(report.is_file())
        self.assertIn('# 准备自检报告', report.read_text('utf-8'))

        write_book(self.final, [(80, 'a', 'Changed', 'Fertig', '', '', '', '', '', '')])
        bad = {}
        process(self.source, self.final, {'t': bad}, 't', 'batch')
        self.assertEqual(bad['state'], 'error')
        self.assertIn('阻断', bad['status'])


if __name__ == '__main__':
    unittest.main()
