import json
from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'program'))
from mathtype_discovery import discover, save_location, configure_conversion


class DiscoveryTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)

    def tearDown(self):
        self.temp.cleanup()

    def interface(self, root):
        path = root / 'MathPage/64/MathPage.wll'
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(b'test fixture, not a real DLL')
        return path

    def test_common_install(self):
        path = self.interface(self.root / 'Apps/MathType')
        result = discover(env={'ProgramFiles(x86)': str(self.root / 'Apps')}, config=None, registry=[])
        self.assertEqual(result.require(), path)

    def test_registry_nondefault(self):
        path = self.interface(self.root / '自定义 安装/数学工具')
        result = discover(env={}, config=None, registry=[(path.parents[2] / 'MathType.exe', 'test registry')])
        self.assertEqual(result.require(), path)

    def test_deduplicated_sources(self):
        path = self.interface(self.root / 'Apps/MathType')
        result = discover(env={'ProgramFiles': str(self.root / 'Apps')}, config=None, registry=[(path.parents[2], 'registry')])
        self.assertEqual(len(result.candidates), 1)
        self.assertEqual(len(result.candidates[0].sources), 2)

    def test_multiple_candidates_require_choice(self):
        roots = [self.root / 'One', self.root / 'Two']
        for root in roots:
            self.interface(root)
        result = discover(env={}, config=None, registry=[(root, 'registry') for root in roots])
        self.assertEqual(result.status, 'selection_required')
        with self.assertRaises(RuntimeError):
            result.require()

    def test_manual_file(self):
        path = self.interface(self.root / 'App')
        self.assertEqual(discover(manual=path, config=None, env={}, registry=[]).require(), path)

    def test_manual_directory_saved_and_reloaded(self):
        root = self.root / 'App'; path = self.interface(root)
        config = self.root / 'settings.json'
        save_location(root, config)
        self.assertEqual(discover(env={}, config=config, registry=[]).require(), path)

    def test_invalid_explicit_does_not_fall_back(self):
        path = self.interface(self.root / 'App')
        result = discover(manual=self.root / 'missing', env={}, config=None, registry=[(path.parents[2], 'registry')])
        self.assertEqual(result.status, 'invalid_override')

    def test_stale_saved_path_is_reported(self):
        config = self.root / 'settings.json'
        config.write_text(json.dumps({'schema_version': 1, 'mathtype_location': str(self.root / 'missing')}))
        self.assertEqual(discover(env={}, config=config, registry=[]).status, 'invalid_override')

    def test_corrupt_settings(self):
        config = self.root / 'settings.json'; config.write_text('bad json')
        self.assertEqual(discover(env={}, config=config, registry=[]).status, 'config_error')

    def test_interface_missing(self):
        root = self.root / 'App'; root.mkdir()
        self.assertEqual(discover(env={}, config=None, registry=[(root, 'registry')]).status, 'interface_missing')

    def test_no_software_found(self):
        self.assertEqual(discover(env={}, config=None, registry=[]).status, 'not_found')

    def test_blank_environment_path(self):
        self.assertEqual(discover(env={'MATHTYPE_DLL': '   '}, config=None, registry=[]).status, 'invalid_override')

    def test_manual_wins_over_environment(self):
        path = self.interface(self.root / 'App')
        self.assertEqual(discover(manual=path, env={'MATHTYPE_DLL': str(self.root / 'missing')}, config=None, registry=[]).require(), path)

    def test_versioned_folder(self):
        path = self.interface(self.root / 'Apps/MathType/7.8')
        self.assertEqual(discover(env={'ProgramFiles': str(self.root / 'Apps')}, config=None, registry=[]).require(), path)

    def test_environment_priority(self):
        path = self.interface(self.root / 'App')
        result = discover(env={'MATHTYPE_DLL': str(path)}, config=self.root / 'unused.json', registry=[])
        self.assertEqual(result.require(), path)

    def test_shared_conversion_result(self):
        path = self.interface(self.root / 'App')
        env, result = configure_conversion(env={'MATHTYPE_DLL': str(path)}, config=None)
        self.assertEqual(Path(env['MATHTYPE_DLL']), result.require())



if __name__ == '__main__':
    unittest.main()
