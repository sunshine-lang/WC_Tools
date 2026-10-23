import unittest

from core_tools.text_metrics import PARAM_PATTERN, TAG_PATTERN, display_length
from pipeline_service import display_len


class TextMetricsTests(unittest.TestCase):
    def test_display_length_preserves_preparation_counting_rules(self):
        for value, expected in [('Hi <b>there</b>\\n!', 9),
                                ('<i>äöüß</i>\n {0}', 8), ('', 0)]:
            with self.subTest(value=value):
                self.assertEqual(display_length(value), expected)
                self.assertEqual(display_len(value), expected)
        self.assertEqual(display_len(None), 0)
        self.assertEqual(display_len(20), 2)

    def test_placeholders_and_tags_still_support_preflight(self):
        self.assertEqual(PARAM_PATTERN.findall('100%ige 200% increase {0} {name} %s %1$s %(count)d %i %f'),
                         ['{0}', '{name}', '%s', '%1$s', '%(count)d', '%i', '%f'])
        self.assertEqual(TAG_PATTERN.findall('<b>Hi</b> <sprite name="star">'),
                         ['<b>', '</b>', '<sprite name="star">'])


if __name__ == '__main__':
    unittest.main()
