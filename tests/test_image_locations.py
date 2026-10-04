"""验证图片检索使用真实正文与页眉引用。"""

import tempfile
import unittest
import zipfile
from pathlib import Path

from k12.processor import TaskProcessor
from k12.store import AppStore
from test_image_conversion import make_source, image_bytes


class ImageLocationTests(unittest.TestCase):
    def test_preview_contains_searchable_image_source_location(self) -> None:
        """正文搜索预览提供图片来源记录，不把媒体位置当作正文内容。"""
        from k12.previews import build_file_preview
        with tempfile.TemporaryDirectory() as temporary:
            source = Path(temporary) / "source.docx"
            make_source(source)
            preview = build_file_preview({"id": "source", "file_type": "Word", "storage_path": str(source)})
            locations = [page for page in preview["pages"] if page["kind"] == "图片来源记录"]
            self.assertEqual(len(locations), 1)
            self.assertIn("第 2 段", locations[0]["text"])
            self.assertIn(locations[0]["location"], locations[0]["text"])
            self.assertEqual(locations[0]["source_name"], "word/media/source.png")

    def test_word_shared_image_paragraph_locations(self) -> None:
        """同一媒体在两个段落的引用分别定位，不合并成单一正文位置。"""
        with tempfile.TemporaryDirectory() as temporary:
            source = Path(temporary) / "source.docx"
            make_source(source)
            with zipfile.ZipFile(source) as archive:
                parts = {name: archive.read(name) for name in archive.namelist()}
            xml = parts["word/document.xml"].decode()
            start = xml.index('<w:p><w:r><w:drawing>')
            end = xml.index('</w:p>', start) + len('</w:p>')
            parts["word/document.xml"] = xml.replace('</w:body>', xml[start:end] + '</w:body>').encode()
            with zipfile.ZipFile(source, "w") as archive:
                for name, data in parts.items():
                    archive.writestr(name, data)
            store = AppStore(Path(temporary) / "data")
            items = TaskProcessor(store)._extract_ooxml_small_images({"id": "source", "file_type": "Word"}, {"id": "task"}, source)
            self.assertEqual(len(items), 1)
            self.assertIn("第 2 段", items[0]["location"])
            self.assertIn("第 4 段", items[0]["location"])
            self.assertEqual(items[0]["page_index"], 0)

    def test_ppt_slide_order_uses_presentation_relationships(self) -> None:
        """文件编号与实际页序不同，共用媒体保留所有实际页码。"""
        with tempfile.TemporaryDirectory() as temporary:
            source = Path(temporary) / "source.pptx"
            relation_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
            namespace = f'xmlns:r="{relation_type[:-1]}"'
            with zipfile.ZipFile(source, "w") as archive:
                archive.writestr("ppt/presentation.xml", f'<presentation {namespace}><sldIdLst><sldId r:id="second"/><sldId r:id="first"/></sldIdLst></presentation>')
                archive.writestr("ppt/_rels/presentation.xml.rels", f'<Relationships><Relationship Id="first" Type="{relation_type}slide" Target="slides/slide1.xml"/><Relationship Id="second" Type="{relation_type}slide" Target="slides/custom.xml"/></Relationships>')
                for name in ("slide1.xml", "custom.xml", "slide_unused.xml"):
                    archive.writestr(f"ppt/slides/{name}", f'<sld {namespace}><pic><blip r:embed="image"/></pic></sld>')
                    archive.writestr(f"ppt/slides/_rels/{name}.rels", f'<Relationships><Relationship Id="image" Type="{relation_type}image" Target="../media/image.png"/></Relationships>')
                archive.writestr("ppt/media/image.png", image_bytes())
            store = AppStore(Path(temporary) / "data")
            items = TaskProcessor(store)._extract_ooxml_small_images({"id": "source", "file_type": "PPT"}, {"id": "task"}, source)
            self.assertEqual(len(items), 1)
            self.assertIn("幻灯片第 1 页 · ppt/slides/custom.xml", items[0]["location"])
            self.assertIn("幻灯片第 2 页 · ppt/slides/slide1.xml", items[0]["location"])
            self.assertNotIn("unused", items[0]["location"])
            self.assertEqual(items[0]["page_indexes"], [1, 2])
            self.assertEqual(items[0]["page_index"], 1)

    def test_excel_sheet_and_shared_image_anchors(self) -> None:
        """绘图文件名不决定工作表，同一媒体的不同锚点完整保留。"""
        with tempfile.TemporaryDirectory() as temporary:
            source = Path(temporary) / "source.xlsx"
            relation_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
            namespace = f'xmlns:r="{relation_type[:-1]}"'
            with zipfile.ZipFile(source, "w") as archive:
                archive.writestr("xl/workbook.xml", f'<workbook {namespace}><sheets><sheet name="成绩表" r:id="sheet"/></sheets></workbook>')
                archive.writestr("xl/_rels/workbook.xml.rels", f'<Relationships><Relationship Id="sheet" Type="{relation_type}worksheet" Target="worksheets/custom.xml"/></Relationships>')
                archive.writestr("xl/worksheets/custom.xml", f'<worksheet {namespace}><drawing r:id="drawing"/></worksheet>')
                archive.writestr("xl/worksheets/_rels/custom.xml.rels", f'<Relationships><Relationship Id="drawing" Type="{relation_type}drawing" Target="../drawings/custom.xml"/></Relationships>')
                anchors = ''.join(f'<oneCellAnchor><from><col>{column}</col><row>{row}</row></from><pic><blipFill><blip r:embed="image"/></blipFill></pic></oneCellAnchor>' for column, row in [(1, 2), (27, 9)])
                archive.writestr("xl/drawings/custom.xml", f'<wsDr {namespace}>{anchors}</wsDr>')
                archive.writestr("xl/drawings/_rels/custom.xml.rels", f'<Relationships><Relationship Id="image" Type="{relation_type}image" Target="../media/image.png"/></Relationships>')
                archive.writestr("xl/media/image.png", image_bytes())
            store = AppStore(Path(temporary) / "data")
            items = TaskProcessor(store)._extract_ooxml_small_images({"id": "source", "file_type": "Excel"}, {"id": "task"}, source)
            self.assertEqual(len(items), 1)
            self.assertEqual(items[0]["location"], "工作表 成绩表 · B3；工作表 成绩表 · AB10")
            self.assertEqual(items[0]["page_index"], 0)

    def test_header_filter_and_shared_body_reference(self) -> None:
        """关闭页眉筛选时仅排除页眉独占图片，共享正文图片仍可检索。"""
        with tempfile.TemporaryDirectory() as temporary:
            source = Path(temporary) / "source.docx"
            make_source(source)
            with zipfile.ZipFile(source) as archive:
                parts = {name: archive.read(name) for name in archive.namelist()}
            parts["word/document.xml"] = parts["word/document.xml"].replace(b'</w:body>', b'<w:sectPr><w:headerReference xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="header"/></w:sectPr></w:body>')
            parts["word/_rels/document.xml.rels"] = parts["word/_rels/document.xml.rels"].replace(b'</Relationships>', b'<Relationship Id="header" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>')
            parts["word/header1.xml"] = b'<hdr xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><blip r:embed="shared"/><blip r:embed="only"/></hdr>'
            parts["word/_rels/header1.xml.rels"] = b'<Relationships><Relationship Id="shared" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/source.png"/><Relationship Id="only" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/header.png"/></Relationships>'
            parts["word/media/header.png"] = parts["word/media/source.png"]
            with zipfile.ZipFile(source, "w") as archive:
                for name, data in parts.items():
                    archive.writestr(name, data)
            store = AppStore(Path(temporary) / "data")
            processor = TaskProcessor(store)
            file = {"id": "source", "file_type": "Word"}
            enabled = processor._extract_ooxml_small_images(file, {"id": "enabled"}, source)
            self.assertEqual(len(enabled), 2)
            shared = next(item for item in enabled if item["source_name"].endswith("source.png"))
            self.assertIn("正文", shared["location"])
            self.assertIn("页眉", shared["location"])
            self.assertEqual(shared["page_index"], 0)
            store.update_settings({"includeHeaderFooterImages": False})
            disabled = processor._extract_ooxml_small_images(file, {"id": "disabled"}, source)
            self.assertEqual(len(disabled), 1)
            self.assertIn("正文", disabled[0]["location"])
            self.assertFalse(disabled[0]["is_header_footer"])

    def test_conflicting_story_reference_rejected_before_deduplication(self) -> None:
        """共享部件去重不能掩盖错误的页眉与页脚混用。"""
        from k12.media import image_locations, word_story_images
        from xml.etree import ElementTree as ET
        with tempfile.TemporaryDirectory() as temporary:
            source = Path(temporary) / "source.docx"
            relation_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
            document = f'<document xmlns:r="{relation_type[:-1]}"><headerReference r:id="header"/><footerReference r:id="footer"/></document>'
            with zipfile.ZipFile(source, "w") as archive:
                archive.writestr("word/document.xml", document)
                archive.writestr("word/shared.xml", '<hdr/>')
                archive.writestr("word/_rels/document.xml.rels", f'<Relationships><Relationship Id="header" Type="{relation_type}header" Target="shared.xml"/><Relationship Id="footer" Type="{relation_type}footer" Target="shared.xml"/></Relationships>')
            with zipfile.ZipFile(source) as archive, self.assertRaisesRegex(ValueError, "同时"):
                image_locations(archive)
            with self.assertRaisesRegex(ValueError, "类型不匹配"):
                word_story_images(source, ET.fromstring(document))
