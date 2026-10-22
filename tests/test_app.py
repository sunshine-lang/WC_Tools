import re
import unittest

import app as server


class AppNavigationTests(unittest.TestCase):
    def setUp(self):
        self.client = server.app.test_client()

    def test_three_pages_link_to_each_other_and_load_local_assets(self):
        for path in ('/pipeline', '/tag_clean', '/qa'):
            response = self.client.get(path)
            self.assertEqual(response.status_code, 200)
            html = response.get_data(as_text=True)
            for destination in ('/pipeline', '/tag_clean', '/qa'):
                if destination != path:
                    self.assertIn(f'href="{destination}"', html)
            for asset in re.findall(r'(?:src|href)="(/static/[^\"]+)"', html):
                with self.client.get(asset) as result:
                    self.assertEqual(result.status_code, 200, asset)
            for old in ('key_classifier', 'context', 'term', 'term_filter', 'tm'):
                self.assertNotIn(f'href="/{old}"', html)

    def test_old_bookmarks_redirect_to_pipeline(self):
        for path in ('/', '/key_classifier', '/context', '/term', '/term_filter', '/tm'):
            response = self.client.get(path)
            self.assertEqual(response.status_code, 302)
            self.assertEqual(response.headers['Location'], '/pipeline')

    def test_old_preparation_apis_are_removed(self):
        for path in ('/upload/key_classifier', '/upload/context', '/upload/term',
                     '/upload/term_filter', '/upload/tm', '/update/term_db',
                     '/update/tm_db', '/api/key_rules'):
            self.assertEqual(self.client.post(path).status_code, 404, path)
        for path in ('/api/classification/missing', '/api/context/missing',
                     '/api/term/missing', '/api/term_filter/missing', '/api/tm/missing',
                     '/context/assets/missing/image.png', '/download/context_bundle/missing'):
            self.assertEqual(self.client.get(path).status_code, 404, path)


if __name__ == '__main__':
    unittest.main()
