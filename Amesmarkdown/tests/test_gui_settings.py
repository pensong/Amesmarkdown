import unittest

from mdconvert_app import gui


class GuiSettingsTests(unittest.TestCase):
    def test_load_default_output_folder_returns_empty_for_missing_file(self) -> None:
        from tempfile import TemporaryDirectory

        with TemporaryDirectory() as tmp:
            settings_dir = gui.SETTINGS_DIR
            settings_file = gui.SETTINGS_FILE
            try:
                gui.SETTINGS_DIR = gui.Path(tmp)
                gui.SETTINGS_FILE = gui.Path(tmp) / "gui_settings.json"
                self.assertEqual(gui.load_default_output_folder(), "")
            finally:
                gui.SETTINGS_DIR = settings_dir
                gui.SETTINGS_FILE = settings_file

    def test_save_and_load_default_output_folder_round_trip(self) -> None:
        from tempfile import TemporaryDirectory

        with TemporaryDirectory() as tmp:
            settings_dir = gui.SETTINGS_DIR
            settings_file = gui.SETTINGS_FILE
            try:
                gui.SETTINGS_DIR = gui.Path(tmp)
                gui.SETTINGS_FILE = gui.Path(tmp) / "gui_settings.json"
                gui.save_default_output_folder("/tmp/markdown-output")
                self.assertTrue(gui.SETTINGS_FILE.exists())
                self.assertEqual(gui.load_default_output_folder(), "/tmp/markdown-output")
            finally:
                gui.SETTINGS_DIR = settings_dir
                gui.SETTINGS_FILE = settings_file


if __name__ == "__main__":
    unittest.main()
