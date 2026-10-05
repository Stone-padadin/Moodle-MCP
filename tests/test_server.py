import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import server


class UtilityTests(unittest.TestCase):
    def test_extension_filter_and_validation(self):
        self.assertEqual(server._extension_filter(["PDF", ".pptx"]), {".pdf", ".pptx"})
        with self.assertRaises(ValueError):
            server._extension_filter(["../secret"])

    def test_safe_filename_components(self):
        self.assertNotIn("..", server._safe_component("../../escape.pdf", 120))
        self.assertEqual(server._safe_component("CON.txt", 120), "item")

    def test_config_requires_https_and_token(self):
        with patch.object(server, "_config_path", return_value=Path("missing-config.json")):
            with self.assertRaisesRegex(ValueError, "not configured"):
                server._config()
        with tempfile.TemporaryDirectory() as directory:
            config = Path(directory) / "config.json"
            config.write_text('{"base_url":"http://moodle.example.edu","token":"x"}', encoding="utf-8")
            with patch.object(server, "_config_path", return_value=config):
                with self.assertRaisesRegex(ValueError, "HTTPS"):
                    server._config()

    def test_saves_without_overwriting(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory) / "Moodle"
            target = root / "course" / "week" / "slides.pdf"
            first_status, first = server._save_without_overwrite(root, target, b"one")
            second_status, second = server._save_without_overwrite(root, target, b"two")
            same_status, same = server._save_without_overwrite(root, target, b"one")
            self.assertEqual(first_status, "downloaded")
            self.assertEqual(second_status, "downloaded")
            self.assertEqual(first.read_bytes(), b"one")
            self.assertEqual(second.read_bytes(), b"two")
            self.assertEqual(same_status, "unchanged")
            self.assertEqual(same, first)

    def test_public_url_rejects_other_hosts_and_strips_tokens(self):
        with patch.object(server, "_config", return_value={"base_url": "https://moodle.example.edu", "host": "moodle.example.edu", "token": "secret"}):
            self.assertEqual(server._public_url("https://moodle.example.edu/mod/page?id=1&wstoken=secret"), "https://moodle.example.edu/mod/page?id=1")
            self.assertIsNone(server._public_url("https://elsewhere.example/file"))


if __name__ == "__main__":
    unittest.main()
