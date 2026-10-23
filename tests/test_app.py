import unittest

import app as server


class AppNavigationTests(unittest.TestCase):
    def setUp(self):
        self.client = server.app.test_client()

    def test_workbench_keeps_library_without_retired_navigation(self):
        response = self.client.get('/pipeline')
        self.assertEqual(response.status_code, 200)
        self.assertIn('no-store', response.headers['Cache-Control'])
        html = response.get_data(as_text=True)
        self.assertIn('id="libBtn"', html)
        self.assertIn('id="stageList"', html)
        for old in ('key_classifier', 'context', 'term', 'term_filter', 'tm', 'tag_clean', 'qa'):
            self.assertNotIn(f'href="/{old}"', html)
        self.assertNotIn('/static/', html)

    def test_home_redirects_to_workbench(self):
        response = self.client.get('/')
        self.assertEqual(response.status_code, 302)
        self.assertEqual(response.headers['Location'], '/pipeline')

    def test_retired_pages_and_assets_are_removed(self):
        for path in ('/key_classifier', '/context', '/term', '/term_filter', '/tm',
                     '/tag_clean', '/qa', '/static/css/style.css', '/static/css/tool-common.css',
                     '/static/css/qa.css', '/static/css/tag_clean.css',
                     '/static/js/qa.js', '/static/js/tag_clean.js'):
            self.assertEqual(self.client.get(path).status_code, 404, path)

    def test_retired_apis_are_removed(self):
        for path in ('/upload/key_classifier', '/upload/context', '/upload/term',
                     '/upload/term_filter', '/upload/tm', '/update/term_db',
                     '/update/tm_db', '/api/key_rules', '/upload/tag_clean', '/upload/qa'):
            self.assertEqual(self.client.post(path).status_code, 404, path)
        for path in ('/api/classification/missing', '/api/context/missing',
                     '/api/term/missing', '/api/term_filter/missing', '/api/tm/missing',
                     '/context/assets/missing/image.png', '/download/context_bundle/missing',
                     '/api/tag_clean/missing', '/api/qa/missing',
                     '/progress/missing', '/download/missing'):
            self.assertEqual(self.client.get(path).status_code, 404, path)


if __name__ == '__main__':
    unittest.main()
