"""验证本地 Office 输出包的实际关系完整性。"""

import tempfile
import unittest
import zipfile
from pathlib import Path

from k12.converters import build_pptx
from k12.local_client import _macos_office_output_validation_error


class OfficeOutputTests(unittest.TestCase):
    def test_internal_fragment_resolves_part_without_fragment(self) -> None:
        """书签片段指向现有部件，纯片段指向当前源文档。"""
        with tempfile.TemporaryDirectory() as temporary:
            target = Path(temporary) / "source.docx"
            for reference in ("document.xml#bookmark", "#bookmark", "/word/document.xml#bookmark"):
                with self.subTest(reference=reference):
                    with zipfile.ZipFile(target, "w") as archive:
                        archive.writestr("[Content_Types].xml", '<Types/>')
                        archive.writestr("word/document.xml", '<document/>')
                        archive.writestr("word/_rels/document.xml.rels", f'<Relationships><Relationship Id="link" Target="{reference}"/></Relationships>')
                    self.assertEqual(_macos_office_output_validation_error(target, ".docx"), "")

    def test_internal_address_cannot_reference_external_or_missing_part(self) -> None:
        """带片段仍须引用包内实际部件，内部关系不能指向网络地址。"""
        with tempfile.TemporaryDirectory() as temporary:
            target = Path(temporary) / "source.docx"
            for reference in ("missing.xml#bookmark", "https://example.com/file", "//example.com/file", "../../outside.xml"):
                with self.subTest(reference=reference):
                    with zipfile.ZipFile(target, "w") as archive:
                        archive.writestr("[Content_Types].xml", '<Types/>')
                        archive.writestr("word/document.xml", '<document/>')
                        archive.writestr("word/_rels/document.xml.rels", f'<Relationships><Relationship Id="link" Target="{reference}"/></Relationships>')
                    self.assertIn("内部关系", _macos_office_output_validation_error(target, ".docx"))

    def test_missing_slide_layout_is_rejected(self) -> None:
        """删除实际版式后即使 ZIP 和演示文稿仍存在，也不能判为成功。"""
        with tempfile.TemporaryDirectory() as temporary:
            target = Path(temporary) / "source.pptx"
            build_pptx([{"title": "标题", "body": "正文"}], target)
            self.assertEqual(_macos_office_output_validation_error(target, ".pptx"), "")
            with zipfile.ZipFile(target) as archive:
                parts = {name: archive.read(name) for name in archive.namelist() if name != "ppt/slideLayouts/slideLayout7.xml"}
            with zipfile.ZipFile(target, "w") as archive:
                for name, data in parts.items():
                    archive.writestr(name, data)
            self.assertIn("内部关系", _macos_office_output_validation_error(target, ".pptx"))

    def test_external_link_does_not_require_embedded_target(self) -> None:
        """合法外部超链接不应被误判为丢失的内部部件。"""
        with tempfile.TemporaryDirectory() as temporary:
            target = Path(temporary) / "source.docx"
            with zipfile.ZipFile(target, "w") as archive:
                archive.writestr("[Content_Types].xml", '<Types/>')
                archive.writestr("word/document.xml", '<document/>')
                archive.writestr("word/_rels/document.xml.rels", '<Relationships><Relationship Id="link" TargetMode="External" Target="https://example.com"/></Relationships>')
            self.assertEqual(_macos_office_output_validation_error(target, ".docx"), "")
