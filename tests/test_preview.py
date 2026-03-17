from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

import fitz
from PIL import Image

from src.preview import generate_previews
from src.scanner import scan_project


class PreviewTests(unittest.TestCase):
    def test_generate_previews_supports_mixed_image_and_pdf(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir) / "root"
            topic = root / "ACL" / "NLP"
            topic.mkdir(parents=True)

            Image.new("RGB", (320, 200), "green").save(topic / "one.png")

            document = fitz.open()
            page = document.new_page()
            page.insert_text((72, 72), "Preview PDF")
            document.save(topic / "two.pdf")
            document.close()

            project = scan_project(root)
            previews_dir = Path(temp_dir) / "previews"
            project = generate_previews(project, previews_dir)

            items = project.sections[0].topics[0].items
            self.assertEqual(len(items), 2)
            self.assertTrue(all(item.preview_path and item.preview_path.exists() for item in items))


if __name__ == "__main__":
    unittest.main()
