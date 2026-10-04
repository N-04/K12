"""验证 Word 内嵌图片写入真实 PPT 媒体部件与关系。"""

import struct
import tempfile
import unittest
import zipfile
import zlib
from pathlib import Path
from xml.etree import ElementTree as ET

from k12.converters import build_pptx_from_docx, extract_docx_blocks, extract_pptx_slides


def image_bytes() -> bytes:
    """生成两像素 PNG，避免依赖图像库。"""
    def chunk(kind: bytes, data: bytes) -> bytes:
        """生成含校验值的 PNG 数据块。"""
        return struct.pack(">I", len(data)) + kind + data + struct.pack(">I", zlib.crc32(kind + data))
    return b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", 2, 1, 8, 2, 0, 0, 0)) + chunk(b"IDAT", zlib.compress(b"\0\xff\0\0\0\xff\0")) + chunk(b"IEND", b"")


def make_source(path: Path, external: bool = False) -> None:
    """写入带明确 DrawingML 引用及显示尺寸的 Word 样本。"""
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("[Content_Types].xml", "<Types/>")
        archive.writestr("word/document.xml", '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t>标题</w:t></w:r></w:p><w:p><w:r><w:drawing><wp:inline><wp:extent cx="2000000" cy="1000000"/><a:blip r:embed="image"/></wp:inline></w:drawing></w:r></w:p><w:p><w:r><w:t>图片后的正文</w:t></w:r></w:p></w:body></w:document>')
        mode = ' TargetMode="External"' if external else ''
        archive.writestr("word/_rels/document.xml.rels", f'<Relationships><Relationship Id="image" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/source.png"{mode}/></Relationships>')
        archive.writestr("word/media/source.png", image_bytes())


class ImageConversionTests(unittest.TestCase):
    """覆盖图片内容、关系、显示比例与禁用保留。"""

    def test_embedded_image_parts_relationships_and_ratio(self) -> None:
        """真实媒体内容不变，形状通过关系引用，显示比例保留。"""
        with tempfile.TemporaryDirectory() as tmp:
            source, target = Path(tmp) / "source.docx", Path(tmp) / "output.pptx"
            make_source(source)
            blocks = extract_docx_blocks(source)
            build_pptx_from_docx(blocks, target)
            slides = extract_pptx_slides(target)
            self.assertEqual(sum(slide["image_count"] for slide in slides), 1)
            self.assertIn("图片后的正文", slides[-1]["body"])
            with zipfile.ZipFile(target) as archive:
                media = next(name for name in archive.namelist() if name.startswith("ppt/media/"))
                self.assertEqual(archive.read(media), image_bytes())
                root = ET.fromstring(archive.read("ppt/slides/slide2.xml"))
                extent = root.find(".//{*}pic/{*}spPr/{*}xfrm/{*}ext")
                self.assertEqual(int(extent.get("cx")), 2 * int(extent.get("cy")))
            build_pptx_from_docx(blocks, target, retain_images=False)
            with zipfile.ZipFile(target) as archive:
                self.assertFalse(any(name.startswith("ppt/media/") for name in archive.namelist()))

    def test_external_image_rejected(self) -> None:
        """外部图片不下载也不静默省略，返回明确错误。"""
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "source.docx"
            make_source(source, True)
            with self.assertRaisesRegex(ValueError, "外部链接"):
                extract_docx_blocks(source)
            self.assertFalse(any("image" in block for block in extract_docx_blocks(source, retain_images=False)))

    def test_crop_rotation_and_flip_preserved(self) -> None:
        """源图片裁剪、旋转和翻转写入实际 PPT 图片形状。"""
        with tempfile.TemporaryDirectory() as tmp:
            source, target = Path(tmp) / "source.docx", Path(tmp) / "output.pptx"
            make_source(source)
            with zipfile.ZipFile(source) as archive:
                parts = {name: archive.read(name) for name in archive.namelist()}
            xml = parts["word/document.xml"].decode().replace('<a:blip r:embed="image"/>', '<a:pic><a:blipFill><a:blip r:embed="image"/><a:srcRect l="10000" t="5000" r="20000" b="0"/></a:blipFill><a:spPr><a:xfrm rot="5400000" flipH="1"><a:ext cx="2000000" cy="1000000"/></a:xfrm></a:spPr></a:pic>')
            parts["word/document.xml"] = xml.encode()
            with zipfile.ZipFile(source, "w") as archive:
                for name, content in parts.items():
                    archive.writestr(name, content)
            build_pptx_from_docx(extract_docx_blocks(source), target)
            with zipfile.ZipFile(target) as archive:
                root = ET.fromstring(archive.read("ppt/slides/slide2.xml"))
            crop = root.find(".//{*}pic/{*}blipFill/{*}srcRect")
            self.assertEqual(crop.attrib, {"l": "10000", "t": "5000", "r": "20000", "b": "0"})
            transform = root.find(".//{*}pic/{*}spPr/{*}xfrm")
            self.assertEqual(transform.get("rot"), "5400000")
            self.assertEqual(transform.get("flipH"), "1")
            self.assertLessEqual(int(transform.find("{*}ext").get("cx")), 4937760)

    def test_each_image_uses_own_geometry(self) -> None:
        """同一容器的多张图片分别使用自身尺寸，不复用首图尺寸。"""
        from k12.media import word_images
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "source.docx"
            make_source(source)
            root = ET.fromstring('<container xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><pic><blip r:embed="image"/><xfrm><ext cx="200" cy="100"/></xfrm></pic><pic><blip r:embed="image"/><xfrm><ext cx="100" cy="200"/></xfrm></pic></container>')
            images = word_images(source, root)
            self.assertEqual([(image["width"], image["height"]) for image in images], [(200, 100), (100, 200)])

    def test_ppt_image_roundtrip_into_word(self) -> None:
        """PPT 原生图片写回 Word 媒体、关系及可解析的行内图片。"""
        from k12.converters import build_docx_from_slides
        with tempfile.TemporaryDirectory() as tmp:
            source, ppt, word = [Path(tmp) / name for name in ["source.docx", "image.pptx", "result.docx"]]
            make_source(source)
            build_pptx_from_docx(extract_docx_blocks(source), ppt)
            slides = extract_pptx_slides(ppt)
            self.assertEqual(sum(len(slide["images"]) for slide in slides), 1)
            build_docx_from_slides(slides, word)
            blocks = extract_docx_blocks(word)
            self.assertEqual(sum("image" in block for block in blocks), 1)
            with zipfile.ZipFile(word) as archive:
                media = next(name for name in archive.namelist() if name.startswith("word/media/"))
                self.assertEqual(archive.read(media), image_bytes())
            build_docx_from_slides(slides, word, retain_images=False)
            with zipfile.ZipFile(word) as archive:
                self.assertFalse(any(name.startswith("word/media/") for name in archive.namelist()))

    def test_vml_physical_units_and_intrinsic_fallback(self) -> None:
        """旧式图片读取物理尺寸，缺少一边时按原图比例补齐。"""
        from k12.media import _vml_length, word_images
        self.assertEqual(_vml_length("2.54cm"), 914400)
        self.assertEqual(_vml_length("72pt"), 914400)
        self.assertEqual(_vml_length("96px"), 914400)
        self.assertEqual(_vml_length("50%"), 0)
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "source.docx"
            make_source(source)
            for style, expected in [("width:72pt;height:36pt", (914400, 457200)),
                                    ("width:72pt", (914400, 457200)),
                                    ("", (19050, 9525))]:
                root = ET.fromstring(f'<pict xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><shape style="{style}"><imagedata r:id="image"/></shape></pict>')
                image = word_images(source, root)[0]
                self.assertEqual((image["width"], image["height"]), expected)

    def test_raster_header_dimensions(self) -> None:
        """光栅头解析保持高宽顺序，损坏头不伪造尺寸。"""
        from k12.media import _raster_dimensions
        self.assertEqual(_raster_dimensions(image_bytes()), (2, 1))
        self.assertEqual(_raster_dimensions(b'GIF89a' + struct.pack('<HH', 20, 10)), (20, 10))
        jpeg = b'\xff\xd8\xff\xc0\x00\x07\x08' + struct.pack('>HH', 10, 20)
        self.assertEqual(_raster_dimensions(jpeg), (20, 10))
        self.assertEqual(_raster_dimensions(b'\xff\xd8\xff\xc0\x00\x01'), (0, 0))

    def test_vml_crop_round_trip_preserves_media(self) -> None:
        """旧式裁剪各类数值转换后往返保持比例，图片字节不变。"""
        from k12.converters import build_docx_from_slides
        from k12.media import word_images
        with tempfile.TemporaryDirectory() as tmp:
            source, presentation, word = (Path(tmp) / name for name in ("source.docx", "output.pptx", "roundtrip.docx"))
            make_source(source)
            root = ET.fromstring('<v:shape xmlns:v="urn:schemas-microsoft-com:vml" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" style="width:72pt;height:36pt"><v:imagedata r:id="image" cropleft=".125" croptop="16384f" cropright="10%" cropbottom="0"/></v:shape>')
            image = word_images(source, root)[0]
            expected = {"l": 12500, "t": 25000, "r": 10000, "b": 0}
            self.assertEqual(image["crop"], expected)
            build_pptx_from_docx([{"text": "", "image": image}], presentation)
            slides = extract_pptx_slides(presentation)
            build_docx_from_slides(slides, word)
            with zipfile.ZipFile(word) as archive:
                drawing = ET.fromstring(archive.read("word/document.xml"))
                self.assertEqual({key: int(value) for key, value in drawing.find(".//{*}srcRect").attrib.items()}, expected)
                media = next(name for name in archive.namelist() if name.startswith("word/media/"))
                self.assertEqual(archive.read(media), image_bytes())

    def test_vml_crop_invalid_and_empty_region(self) -> None:
        """无效数值和完全裁掉的图片明确失败，不静默丢失裁剪。"""
        from k12.media import _vml_crop, picture_xml
        self.assertEqual(_vml_crop("-32768f"), -50000)
        for value in ("NaN", "Inf", "1.1", "1.5f", "", "100001%"):
            with self.subTest(value=value), self.assertRaises(ValueError):
                _vml_crop(value)
        with self.assertRaisesRegex(ValueError, "可见区域"):
            picture_xml({"width": 100, "height": 100, "crop": {"l": _vml_crop(".5"), "r": _vml_crop("32768f")}})

    def test_vml_rotation_and_flip_round_trip(self) -> None:
        """旧式角度及两轴翻转在实际 PPT 与 Word 图片对象间保持。"""
        from k12.converters import build_docx_from_slides
        from k12.media import word_images
        with tempfile.TemporaryDirectory() as tmp:
            source, target, word = (Path(tmp) / name for name in ("source.docx", "target.pptx", "target.docx"))
            make_source(source)
            for rotation, flip, expected in [("-90", "x", {"rot": "16200000", "flipH": "1"}),
                                             ("450deg", "y", {"rot": "5400000", "flipV": "1"}),
                                             ("22.5", "y x", {"rot": "1350000", "flipH": "1", "flipV": "1"})]:
                with self.subTest(rotation=rotation, flip=flip):
                    root = ET.fromstring(f'<v:shape xmlns:v="urn:schemas-microsoft-com:vml" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" style="width:72pt;height:36pt; rotation : {rotation}; flip : {flip}"><v:imagedata r:id="image"/></v:shape>')
                    image = word_images(source, root)[0]
                    build_pptx_from_docx([{"text": "", "image": image}], target)
                    build_docx_from_slides(extract_pptx_slides(target), word)
                    with zipfile.ZipFile(word) as archive:
                        document = ET.fromstring(archive.read("word/document.xml"))
                        self.assertEqual(document.find(".//{*}xfrm").attrib, expected)
                        media = next(name for name in archive.namelist() if name.startswith("word/media/"))
                        self.assertEqual(archive.read(media), image_bytes())

    def test_vml_transform_rejects_invalid_values(self) -> None:
        """非法变换明确失败，避免静默显示未旋转图片。"""
        from k12.media import _vml_transform
        for style in ({"rotation": "NaN"}, {"rotation": "90rad"}, {"flip": "z"}, {"flip": "xx"}):
            with self.subTest(style=style), self.assertRaises(ValueError):
                _vml_transform(style)

    def test_word_image_source_size_and_rotated_layout_space(self) -> None:
        """小图不放大，旋转后的占用区域不超出正文宽高。"""
        from k12.media import word_picture_xml
        for width, height, rotation in [(200000, 100000, 0), (1000000, 2000000, 5400000),
                                        (2000000, 12000000, 5400000), (9000000, 9000000, 2700000)]:
            with self.subTest(width=width, height=height, rotation=rotation):
                drawing = ET.fromstring('<root xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
                                        word_picture_xml({"width": width, "height": height, "transform": {"rot": rotation}}, 1) + '</root>')
                extent = drawing.find(".//{*}inline/{*}extent")
                effects = drawing.find(".//{*}effectExtent")
                result_width, result_height = int(extent.get("cx")), int(extent.get("cy"))
                self.assertLessEqual(result_width + int(effects.get("l")) + int(effects.get("r")), 5715001)
                self.assertLessEqual(result_height + int(effects.get("t")) + int(effects.get("b")), 8863331)
                self.assertAlmostEqual(result_width / result_height, width / height, places=5)
                self.assertLessEqual(result_width, width)
                self.assertLessEqual(result_height, height)
                if width == 200000:
                    self.assertEqual((result_width, result_height), (width, height))
                if width == 1000000:
                    self.assertEqual(int(effects.get("l")), 500000)
                    self.assertEqual(int(effects.get("t")), 0)

    def test_referenced_header_images_and_duplicate_sections(self) -> None:
        """页眉图片按关系提取，共享页眉只保留一次，未引用部件不混入。"""
        with tempfile.TemporaryDirectory() as tmp:
            source, target = Path(tmp) / "source.docx", Path(tmp) / "output.pptx"
            make_source(source)
            with zipfile.ZipFile(source) as archive:
                parts = {name: archive.read(name) for name in archive.namelist()}
            document = parts["word/document.xml"].decode()
            document = document.replace('</w:body>', '<w:sectPr><w:headerReference r:id="header1"/><w:headerReference r:id="header2"/></w:sectPr></w:body>')
            parts["word/document.xml"] = document.encode()
            relations = parts["word/_rels/document.xml.rels"].decode().replace('</Relationships>', '<Relationship Id="header1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/><Relationship Id="header2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>')
            parts["word/_rels/document.xml.rels"] = relations.encode()
            header = '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:p><w:r><w:drawing><a:pic><a:blip r:embed="image"/><a:xfrm><a:ext cx="1000000" cy="1000000"/></a:xfrm></a:pic></w:drawing></w:r></w:p></w:hdr>'
            parts["word/header1.xml"] = header.encode()
            parts["word/header_unused.xml"] = header.encode()
            parts["word/_rels/header1.xml.rels"] = b'<Relationships><Relationship Id="image" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/source.png"/></Relationships>'
            with zipfile.ZipFile(source, "w") as archive:
                for name, content in parts.items():
                    archive.writestr(name, content)
            blocks = extract_docx_blocks(source)
            self.assertEqual(sum("image" in block for block in blocks), 2)
            self.assertEqual(sum(block.get("image_label") == "页眉图片" for block in blocks), 1)
            build_pptx_from_docx(blocks, target)
            slides = extract_pptx_slides(target)
            self.assertTrue(any(slide["title"] == "页眉图片" and slide["image_count"] == 1 for slide in slides))
            self.assertFalse(any("image" in block for block in extract_docx_blocks(source, retain_images=False)))

    def test_image_layout_limits_reach_artifact_and_text_report(self) -> None:
        """真实图片产物可下载，布局限制同时进入任务消息和导出报告。"""
        from k12.store import AppStore
        from k12.processor import TaskProcessor
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "source.docx"
            make_source(source)
            store = AppStore(Path(tmp) / "data")
            processor = TaskProcessor(store)
            file = processor.create_uploaded_file("source.docx", source.read_bytes())[0]
            task = processor.create_task({"task_type": "word_to_ppt", "file_ids": [file["id"]]})
            self.assertEqual(task["status"], "成功")
            report = store.list_reports()[0]
            artifact = report["analysis"]["artifacts"][0]
            self.assertIn("图片按独立页保留", artifact["message"])
            self.assertIn("原始位置和文字环绕未还原", Path(report["txt_path"]).read_text(encoding="utf-8"))
            output = Path(artifact["path"])
            file = processor.create_uploaded_file("result.pptx", output.read_bytes())[0]
            task = processor.create_task({"task_type": "ppt_to_word", "file_ids": [file["id"]]})
            self.assertEqual(task["status"], "成功")
            report = next(item for item in store.list_reports() if item["task_id"] == task["id"])
            self.assertIn("图片按行内绘图保留", report["analysis"]["artifacts"][0]["message"])
            self.assertIn("分组布局未还原", Path(report["txt_path"]).read_text(encoding="utf-8"))

    def test_alternate_image_representations_are_not_duplicated(self) -> None:
        """同一图片的 DrawingML/VML 兼容表示只输出一次。"""
        with tempfile.TemporaryDirectory() as tmp:
            source, target = Path(tmp) / "source.docx", Path(tmp) / "output.pptx"
            make_source(source)
            with zipfile.ZipFile(source) as archive:
                parts = {name: archive.read(name) for name in archive.namelist()}
            xml = parts["word/document.xml"].decode()
            xml = xml.replace('<w:drawing>', '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="a"><w:drawing>')
            xml = xml.replace('</w:drawing>', '</w:drawing></mc:Choice><mc:Fallback><w:pict><v:shape xmlns:v="urn:schemas-microsoft-com:vml" style="width:72pt;height:36pt"><v:imagedata r:id="image"/></v:shape></w:pict></mc:Fallback></mc:AlternateContent>')
            parts["word/document.xml"] = xml.encode()
            with zipfile.ZipFile(source, "w") as archive:
                for name, content in parts.items():
                    archive.writestr(name, content)
            blocks = extract_docx_blocks(source)
            self.assertEqual(sum("image" in block for block in blocks), 1)
            build_pptx_from_docx(blocks, target)
            self.assertEqual(sum(len(slide["images"]) for slide in extract_pptx_slides(target)), 1)
