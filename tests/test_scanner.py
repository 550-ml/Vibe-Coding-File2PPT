from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from src.scanner import ScanError, scan_project


class ScannerTests(unittest.TestCase):
    def test_scan_project_supports_files_directly_under_selected_folder(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir) / "root"
            root.mkdir()
            (root / "one.pdf").write_bytes(b"%PDF-1.4")
            (root / "two.jpg").write_bytes(b"jpg")

            project = scan_project(root)

            self.assertEqual(len(project.sections), 1)
            self.assertEqual(project.sections[0].name, "当前文件夹")
            self.assertEqual(project.sections[0].topics[0].name, "root")
            self.assertEqual(len(project.sections[0].topics[0].items), 2)

    def test_scan_project_supports_nested_directories(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir) / "root"
            topic = root / "会议" / "2026" / "NLP"
            topic.mkdir(parents=True)
            (topic / "paper.pdf").write_bytes(b"%PDF-1.4")

            project = scan_project(root)

            self.assertEqual(len(project.sections), 1)
            self.assertEqual(project.sections[0].name, "会议")
            self.assertEqual(project.sections[0].topics[0].name, "2026 / NLP")
            self.assertEqual(len(project.sections[0].topics[0].items), 1)

    def test_scan_project_collects_supported_files_and_skips_others(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir) / "root"
            topic = root / "ACL" / "NLP"
            topic.mkdir(parents=True)
            (topic / "one.jpg").write_bytes(b"jpg")
            (topic / "two.pdf").write_bytes(b"%PDF-1.4")
            (topic / "skip.txt").write_text("ignore", encoding="utf-8")

            project = scan_project(root)

            self.assertEqual(project.root_name, "root")
            self.assertEqual(len(project.sections), 1)
            self.assertEqual(len(project.sections[0].topics[0].items), 2)
            self.assertEqual(len(project.skipped_files), 1)

    def test_scan_project_rejects_empty_root(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir) / "root"
            root.mkdir()

            with self.assertRaises(ScanError):
                scan_project(root)

    def test_scan_project_rejects_root_without_supported_files(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir) / "root"
            topic = root / "ACL" / "NLP"
            topic.mkdir(parents=True)
            (topic / "skip.txt").write_text("ignore", encoding="utf-8")

            with self.assertRaises(ScanError):
                scan_project(root)


if __name__ == "__main__":
    unittest.main()
