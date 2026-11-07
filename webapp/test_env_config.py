"""Secret-free configuration tests."""
import unittest
from pathlib import Path
from unittest.mock import patch
from env_config import load_project_env

class ProjectEnvTests(unittest.TestCase):
    def load_text(self, text, env=None):
        target = {} if env is None else env
        with patch.object(Path, 'is_file', return_value=True), patch.object(Path, 'read_text', return_value=text):
            load_project_env(Path('example-project'), target)
        return target

    def test_saved_project_configuration_is_present(self):
        project = Path(__file__).resolve().parent.parent
        if not (project / '.env').is_file():
            self.skipTest('Local secrets are not part of the repository')
        env = {}
        load_project_env(project, env)
        self.assertTrue(bool(env.get('GEMINI_API_KEY')), 'Project key is missing')
        self.assertTrue(env.get('GEMINI_MODEL') == 'gemini-3.8-flash')
        self.assertTrue(env.get('GEMINI_PRO_MODEL') == 'gemini-4-argon')

    def test_loads_project_key_and_models(self):
        env = self.load_text('# comment\nGEMINI_API_KEY=test-key\nGEMINI_MODEL=gemini-3.8-flash\nGEMINI_PRO_MODEL=gemini-4-argon\n')
        self.assertEqual(env, {'GEMINI_API_KEY': 'test-key', 'GEMINI_MODEL': 'gemini-3.8-flash', 'GEMINI_PRO_MODEL': 'gemini-4-argon'})

    def test_explicit_environment_takes_precedence(self):
        env = self.load_text('GEMINI_API_KEY=local-example', {'GEMINI_API_KEY': 'deployment-example'})
        self.assertEqual(env['GEMINI_API_KEY'], 'deployment-example')

    def test_missing_file_is_optional(self):
        env = {}
        with patch.object(Path, 'is_file', return_value=False):
            load_project_env(Path('example-project'), env)
        self.assertEqual(env, {})

    def test_quotes_comments_and_invalid_lines(self):
        env = self.load_text('export GEMINI_API_KEY="example=key"\nGEMINI_MODEL=\'example-model\'\nLIMIT=400 # comment\ninvalid line\n9INVALID=no\n')
        self.assertEqual(env, {'GEMINI_API_KEY': 'example=key', 'GEMINI_MODEL': 'example-model', 'LIMIT': '400'})

if __name__ == '__main__':
    unittest.main()
