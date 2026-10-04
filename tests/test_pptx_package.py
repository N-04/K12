"""验证演示文稿的母版、版式和主题关系完整。"""

import posixpath
import tempfile
import unittest
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from k12.converters import build_pptx, extract_pptx_slides


class PresentationPackageTests(unittest.TestCase):
    def test_wrong_slide_relationship_type_rejected(self) -> None:
        """已有页面文件不能掩盖错误的关系类型。"""
        with tempfile.TemporaryDirectory() as temporary:
            target = Path(temporary) / "invalid.pptx"
            build_pptx([{"title": "标题", "body": ["正文"]}], target)
            with zipfile.ZipFile(target) as archive:
                parts = {name: archive.read(name) for name in archive.namelist()}
            name = "ppt/_rels/presentation.xml.rels"
            parts[name] = parts[name].replace(b'relationships/slide"', b'relationships/theme"')
            with zipfile.ZipFile(target, "w") as archive:
                for name, data in parts.items():
                    archive.writestr(name, data)
            with self.assertRaisesRegex(ValueError, "关系类型"):
                extract_pptx_slides(target)

    def test_external_hyperlink_does_not_break_slide_extraction(self) -> None:
        """普通外部链接不作为 ZIP 部件读取，页面正文仍完整。"""
        with tempfile.TemporaryDirectory() as temporary:
            target = Path(temporary) / "linked.pptx"
            build_pptx([{"title": "带链接标题", "body": ["带链接正文"]}], target)
            with zipfile.ZipFile(target) as archive:
                parts = {name: archive.read(name) for name in archive.namelist()}
            name = "ppt/slides/_rels/slide1.xml.rels"
            relations = ET.fromstring(parts[name])
            ET.SubElement(relations, "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship",
                          {"Id": "link", "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
                           "Target": "https://example.com/course", "TargetMode": "External"})
            parts[name] = ET.tostring(relations)
            with zipfile.ZipFile(target, "w") as archive:
                for name, data in parts.items():
                    archive.writestr(name, data)
            slides = extract_pptx_slides(target)
            self.assertEqual(slides[0]["title"], "带链接标题")
            self.assertIn("带链接正文", slides[0]["body"])

    def test_slide_layout_master_theme_relationships_resolve(self) -> None:
        """每页有版式，版式有母版，母版有主题，内部引用均可解析。"""
        with tempfile.TemporaryDirectory() as temporary:
            target = Path(temporary) / "result.pptx"
            build_pptx([{"title": "标题", "body": "正文"}, {"title": "表格", "table": [["甲", "乙"]]}], target)
            with zipfile.ZipFile(target) as archive:
                names = archive.namelist()
                self.assertEqual(len(names), len(set(names)))
                relationships = {}
                for name in names:
                    if not name.endswith(".rels"):
                        continue
                    directory, filename = posixpath.split(name)
                    source = posixpath.join(posixpath.dirname(directory), filename[:-5])
                    resolved = {}
                    for relation in ET.fromstring(archive.read(name)):
                        if relation.get("TargetMode") == "External":
                            continue
                        destination = posixpath.normpath(posixpath.join(posixpath.dirname(source), relation.get("Target")))
                        self.assertIn(destination, names, (name, destination))
                        resolved[relation.get("Type").rsplit("/", 1)[-1]] = destination
                    relationships[source] = resolved
                for slide in ("ppt/slides/slide1.xml", "ppt/slides/slide2.xml"):
                    layout = relationships[slide]["slideLayout"]
                    master = relationships[layout]["slideMaster"]
                    self.assertIn("theme", relationships[master])
                self.assertIn("slideMaster", relationships["ppt/presentation.xml"])
