"""验证正文局部加粗和斜体在分页转换后仍保持。"""

import tempfile
import unittest
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from k12.converters import extract_docx_blocks, build_pptx_from_docx, extract_pptx_slides, build_docx_from_slides
from k12.text_styles import WordTextStyles


class TextStyleTests(unittest.TestCase):
    def test_pptx_to_word_direct_format_roundtrip(self) -> None:
        """PPT 转讲义和大纲保留各片段显式格式及软换行。"""
        with tempfile.TemporaryDirectory() as temporary:
            source, ppt = Path(temporary) / "source.docx", Path(temporary) / "source.pptx"
            with zipfile.ZipFile(source, "w") as archive:
                archive.writestr("word/document.xml", '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>标题</w:t></w:r></w:p><w:p><w:r><w:rPr><w:b/><w:i/><w:sz w:val="32"/><w:rFonts w:ascii="Times New Roman" w:eastAsia="宋体"/><w:color w:val="cc3300"/></w:rPr><w:t>格式文字</w:t><w:br/><w:t>第二行</w:t></w:r><w:r><w:rPr><w:b w:val="0"/></w:rPr><w:t>关闭加粗</w:t></w:r></w:p></w:body></w:document>')
            build_pptx_from_docx(extract_docx_blocks(source), ppt)
            slides = extract_pptx_slides(ppt)
            for mode in ("逐页讲义模式", "大纲模式"):
                target = Path(temporary) / (mode + ".docx")
                build_docx_from_slides(slides, target, mode=mode, generate_toc=False)
                body = next(block for block in extract_docx_blocks(target) if block["text"].startswith("格式文字"))
                styled = [run for run in body["runs"] if run.get("b")]
                self.assertEqual("".join(run["text"] for run in styled), "格式文字第二行")
                self.assertTrue(all(run["i"] and run["size"] == 1600 and run["font"] == "Times New Roman" and run["east_asia"] == "宋体" and run["color"] == "CC3300" for run in styled))
                self.assertEqual(body["text"], "格式文字\n第二行关闭加粗")
                self.assertFalse(body["runs"][-1]["b"])

    def test_theme_underline_color_and_direct_override(self) -> None:
        """主题下划线继承实际颜色，直接 RGB 覆盖会清除旧主题引用。"""
        theme = ET.fromstring('<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:clrScheme><a:accent1><a:srgbClr val="4F81BD"/></a:accent1></a:clrScheme></a:themeElements></a:theme>')
        styles = ET.fromstring('<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:styleId="Colored"><w:rPr><w:u w:val="single" w:color="123456" w:themeColor="accent1"/></w:rPr></w:style></w:styles>')
        paragraph = ET.fromstring('<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:pPr><w:pStyle w:val="Colored"/></w:pPr><w:r><w:t>主题下划线</w:t></w:r><w:r><w:rPr><w:u w:val="single" w:color="FF0000"/></w:rPr><w:t>直接红色</w:t></w:r></w:p>')
        resolver = WordTextStyles(styles, theme)
        runs = [{"text": node.findtext("{*}t"), **resolver.run(paragraph, node)} for node in paragraph.findall("{*}r")]
        self.assertEqual([run["underline_color"] for run in runs], ["4F81BD", "FF0000"])
        self.assertTrue(all("underline_color_unresolved" not in run for run in runs))
        with tempfile.TemporaryDirectory() as temporary:
            target = Path(temporary) / "target.pptx"
            self.assertFalse(build_pptx_from_docx([{"text": "标题", "level": 1}, {"text": "主题下划线直接红色", "level": 0, "runs": runs}], target))
            with zipfile.ZipFile(target) as archive:
                root = ET.fromstring(archive.read("ppt/slides/slide1.xml"))
            self.assertEqual([color.get("val") for color in root.findall(".//{*}uFill/{*}solidFill/{*}srgbClr")], ["4F81BD", "FF0000"])

    def test_theme_base_color_and_unresolved_adjustments(self) -> None:
        """主题基础颜色优先于回退值，明暗调整尚未验证则不声明解析。"""
        theme = ET.fromstring('<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:clrScheme><a:accent2><a:srgbClr val="C0504D"/></a:accent2><a:accent1><a:srgbClr val="4F81BD"/></a:accent1><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1></a:clrScheme></a:themeElements></a:theme>')
        resolver = WordTextStyles(None, theme)
        self.assertIsNone(resolver.theme_color({"themeColor": "accent2", "themeShade": "BF"}))
        self.assertIsNone(resolver.theme_color({"themeColor": "accent1", "themeTint": "99"}))
        self.assertIsNone(resolver.theme_color({"themeColor": "accent1", "themeTint": "99", "themeShade": "00"}))
        self.assertEqual(resolver.theme_color({"themeColor": "text1"}), "000000")
        paragraph = ET.fromstring('<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:rPr><w:color w:val="123456" w:themeColor="accent2"/></w:rPr><w:t>主题文字</w:t></w:r></w:p>')
        result = resolver.run(paragraph, paragraph.find("{*}r"))
        self.assertEqual(result["color"], "C0504D")
        self.assertNotIn("theme_color_unresolved", result)
        with tempfile.TemporaryDirectory() as temporary:
            target = Path(temporary) / "target.pptx"
            warnings = build_pptx_from_docx([{"text": "标题", "level": 1}, {"text": "主题文字", "level": 0, "runs": [{"text": "主题文字", **result}]}], target)
            self.assertFalse(warnings)
            with zipfile.ZipFile(target) as archive:
                self.assertIn(b'val="C0504D"', archive.read("ppt/slides/slide1.xml"))

    def test_underline_inheritance_and_paginated_output(self) -> None:
        """下划线随样式继承与分页保留，显式 none 只关闭指定片段。"""
        with tempfile.TemporaryDirectory() as temporary:
            source, target = Path(temporary) / "source.docx", Path(temporary) / "target.pptx"
            with zipfile.ZipFile(source, "w") as archive:
                archive.writestr("word/styles.xml", '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:styleId="Underlined"><w:rPr><w:u w:val="double" w:color="cc3300"/></w:rPr></w:style></w:styles>')
                archive.writestr("word/document.xml", '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>标题</w:t></w:r></w:p><w:p><w:pPr><w:pStyle w:val="Underlined"/></w:pPr><w:r><w:t>继承双下划线文字</w:t></w:r><w:r><w:rPr><w:u w:val="none"/></w:rPr><w:t>关闭</w:t></w:r></w:p></w:body></w:document>')
            build_pptx_from_docx(extract_docx_blocks(source), target, max_chars=5)
            with zipfile.ZipFile(target) as archive:
                runs = [run for name in sorted(archive.namelist()) if name.startswith("ppt/slides/slide") and name.endswith(".xml")
                        for shape in ET.fromstring(archive.read(name)).findall(".//{*}sp")
                        if shape.find("{*}nvSpPr/{*}cNvPr").get("name") == "Content"
                        for run in shape.findall(".//{*}r")]
            self.assertEqual("".join(run.findtext("{*}t") for run in runs if run.find("{*}rPr").get("u") == "dbl"), "继承双下划线文字")
            self.assertEqual("".join(run.findtext("{*}t") for run in runs if run.find("{*}rPr").get("u") == "none"), "关闭")
            for run in runs:
                color = run.find("{*}rPr/{*}uFill/{*}solidFill/{*}srgbClr")
                if run.find("{*}rPr").get("u") == "dbl":
                    self.assertEqual(color.get("val"), "CC3300")
                else:
                    self.assertIsNone(color)

    def test_underline_mapping_and_color_limitation(self) -> None:
        """波浪线和粗点划线转换为 PPT 枚举，主题颜色明确报告限制。"""
        for source, expected in (("wave", "wavy"), ("dashDotHeavy", "dotDashHeavy"), ("wavyDouble", "wavyDbl")):
            paragraph = ET.fromstring(f'<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:rPr><w:u w:val="{source}" w:color="FF0000" w:themeColor="accent1"/></w:rPr><w:t>文字</w:t></w:r></w:p>')
            result = WordTextStyles(None).run(paragraph, paragraph.find("{*}r"))
            self.assertEqual(result["underline"], expected)
            with tempfile.TemporaryDirectory() as temporary:
                warnings = build_pptx_from_docx([{"text": "标题", "level": 1}, {"text": "文字", "level": 0, "runs": [{"text": "文字", **result}]}], Path(temporary) / "target.pptx")
            self.assertTrue(any("下划线主题颜色" in warning for warning in warnings))

    def test_theme_language_selects_supplemental_font(self) -> None:
        """主题语言同时控制东亚与西文引用，缺失语言不猜测。"""
        theme = ET.fromstring('<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:fontScheme><a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:font script="Hans" typeface="宋体"/><a:font script="Hant" typeface="新細明體"/><a:font script="Jpan" typeface="MS Mincho"/><a:font script="Hang" typeface="Batang"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>')
        for language, expected in (("zh-CN", "宋体"), ("zh-SG", "宋体"), ("zh-TW", "新細明體"), ("zh-HK", "新細明體"), ("zh-Hant-CN", "新細明體"), ("ja-JP", "MS Mincho"), ("ko-KR", "Batang")):
            with self.subTest(language=language):
                settings = ET.fromstring(f'<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:themeFontLang w:eastAsia="{language}" w:val="{language}"/></w:settings>')
                resolver = WordTextStyles(None, theme, settings)
                self.assertEqual(resolver.theme_font("minorEastAsia"), expected)
                self.assertEqual(resolver.theme_font("minorAscii"), expected)
        self.assertIsNone(WordTextStyles(None, theme).theme_font("minorEastAsia"))

    def test_theme_fonts_follow_custom_relationship_and_direct_override(self) -> None:
        """主题关系可以使用自定义部件，主题字体优先于同层回退字体。"""
        with tempfile.TemporaryDirectory() as temporary:
            source, target = Path(temporary) / "source.docx", Path(temporary) / "target.pptx"
            with zipfile.ZipFile(source, "w") as archive:
                archive.writestr("word/document.xml", '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>标题</w:t></w:r></w:p><w:p><w:r><w:rPr><w:rFonts w:ascii="Fallback" w:asciiTheme="majorAscii" w:eastAsiaTheme="minorEastAsia"/></w:rPr><w:t>主题文字</w:t></w:r></w:p></w:body></w:document>')
                archive.writestr("word/_rels/document.xml.rels", '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="theme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/custom.xml"/></Relationships>')
                archive.writestr("word/theme/custom.xml", '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:fontScheme><a:majorFont><a:latin typeface="Times New Roman"/></a:majorFont><a:minorFont><a:ea typeface="宋体"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>')
            blocks = extract_docx_blocks(source)
            self.assertEqual(blocks[1]["runs"][0]["font"], "Times New Roman")
            self.assertEqual(blocks[1]["runs"][0]["east_asia"], "宋体")
            self.assertFalse(build_pptx_from_docx(blocks, target))
            with zipfile.ZipFile(target) as archive:
                root = ET.fromstring(archive.read("ppt/slides/slide1.xml"))
            self.assertTrue(any(node.get("typeface") == "宋体" for node in root.findall(".//{*}ea")))

    def test_missing_theme_font_returns_limitation(self) -> None:
        """主题字体缺失不能被当作完成格式保留。"""
        paragraph = ET.fromstring('<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:rPr><w:rFonts w:asciiTheme="minorAscii"/></w:rPr><w:t>文字</w:t></w:r></w:p>')
        style = WordTextStyles(None).run(paragraph, paragraph.find("{*}r"))
        self.assertTrue(style["theme_font_unresolved"])
        with tempfile.TemporaryDirectory() as temporary:
            warnings = build_pptx_from_docx([{"text": "标题", "level": 1}, {"text": "文字", "runs": [{"text": "文字", **style}], "level": 0}], Path(temporary) / "target.pptx")
            self.assertTrue(any("主题字体" in warning for warning in warnings))

    def test_inherited_styles_and_direct_override(self) -> None:
        """文档、父样式和字符样式逐级生效，直接格式最终覆盖。"""
        styles = ET.fromstring('<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="24"/><w:rFonts w:ascii="Arial" w:eastAsia="宋体"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:styleId="Base" w:default="1"><w:rPr><w:b/><w:color w:val="112233"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Child"><w:basedOn w:val="Base"/><w:rPr><w:sz w:val="32"/></w:rPr></w:style><w:style w:type="character" w:styleId="Emphasis"><w:rPr><w:i/><w:rFonts w:ascii="Times New Roman"/></w:rPr></w:style></w:styles>')
        paragraph = ET.fromstring('<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:pPr><w:pStyle w:val="Child"/></w:pPr><w:r><w:rPr><w:rStyle w:val="Emphasis"/><w:b w:val="0"/><w:color w:val="auto"/></w:rPr><w:t>文字</w:t></w:r></w:p>')
        resolver = WordTextStyles(styles)
        result = resolver.run(paragraph, paragraph.find("{*}r"))
        self.assertEqual(result, {"size": 1600, "font": "Times New Roman", "east_asia": "宋体", "b": False, "i": True})
        plain = ET.fromstring("<p><r/></p>")
        self.assertEqual(resolver.run(plain, plain.find("r"))["size"], 1200)
        self.assertTrue(resolver.run(plain, plain.find("r"))["b"])

    def test_style_toggle_and_cycle(self) -> None:
        """父子样式连续加粗会切换，循环继承不能无限执行。"""
        styles = ET.fromstring('<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:styleId="Base"><w:rPr><w:b/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Child"><w:basedOn w:val="Base"/><w:rPr><w:b/></w:rPr></w:style></w:styles>')
        paragraph = ET.fromstring('<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:pPr><w:pStyle w:val="Child"/></w:pPr><w:r/></w:p>')
        self.assertFalse(WordTextStyles(styles).run(paragraph, paragraph.find("{*}r"))["b"])
        ET.SubElement(styles[0], "basedOn", {"{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val": "Child"})
        with self.assertRaisesRegex(ValueError, "循环"):
            WordTextStyles(styles).run(paragraph, paragraph.find("{*}r"))

    def test_soft_break_is_native_and_preserves_style(self) -> None:
        """Word 软换行成为原生 PPT 换行，前后文字保留格式。"""
        with tempfile.TemporaryDirectory() as temporary:
            source, target = Path(temporary) / "source.docx", Path(temporary) / "target.pptx"
            with zipfile.ZipFile(source, "w") as archive:
                archive.writestr("word/document.xml", '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>标题</w:t></w:r></w:p><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>第一行</w:t><w:br/><w:t>第二行</w:t></w:r></w:p></w:body></w:document>')
            build_pptx_from_docx(extract_docx_blocks(source), target)
            with zipfile.ZipFile(target) as archive:
                root = ET.fromstring(archive.read("ppt/slides/slide1.xml"))
            shape = next(shape for shape in root.findall(".//{*}sp") if shape.find("{*}nvSpPr/{*}cNvPr").get("name") == "Content")
            self.assertEqual(len(shape.findall(".//{*}br")), 1)
            runs = shape.findall(".//{*}r")
            self.assertEqual([run.findtext("{*}t") for run in runs], ["第一行", "第二行"])
            self.assertTrue(all(run.find("{*}rPr").get("b") == "1" for run in runs))

    def test_theme_color_fallback_reports_limitation(self) -> None:
        """主题颜色回退成功也明确报告尚未完成主题解析。"""
        with tempfile.TemporaryDirectory() as temporary:
            source, target = Path(temporary) / "source.docx", Path(temporary) / "target.pptx"
            with zipfile.ZipFile(source, "w") as archive:
                archive.writestr("word/document.xml", '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>标题</w:t></w:r></w:p><w:p><w:r><w:rPr><w:color w:val="123456" w:themeColor="accent1" w:themeTint="80"/></w:rPr><w:t>主题文字</w:t></w:r></w:p></w:body></w:document>')
            warnings = build_pptx_from_docx(extract_docx_blocks(source), target)
            self.assertTrue(any("主题文字颜色" in warning for warning in warnings))
            with zipfile.ZipFile(target) as archive:
                self.assertIn(b'val="123456"', archive.read("ppt/slides/slide1.xml"))

    def test_mixed_styles_across_page_boundary(self) -> None:
        """分页切开格式片段后文字不丢失，显式关闭加粗不会扩散样式。"""
        with tempfile.TemporaryDirectory() as temporary:
            source, target = Path(temporary) / "source.docx", Path(temporary) / "target.pptx"
            with zipfile.ZipFile(source, "w") as archive:
                archive.writestr("word/document.xml", '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>标题</w:t></w:r></w:p><w:p><w:r><w:t>普通</w:t></w:r><w:r><w:rPr><w:b/><w:i/><w:color w:val="cc3300"/><w:sz w:val="32"/><w:rFonts w:ascii="Times New Roman" w:eastAsia="宋体"/></w:rPr><w:t>加粗斜体文字</w:t></w:r><w:r><w:rPr><w:b w:val="0"/></w:rPr><w:t>结尾</w:t></w:r></w:p></w:body></w:document>')
            build_pptx_from_docx(extract_docx_blocks(source), target, max_chars=5)
            with zipfile.ZipFile(target) as archive:
                runs = [run for name in sorted(archive.namelist()) if name.startswith("ppt/slides/slide") and name.endswith(".xml")
                        for shape in ET.fromstring(archive.read(name)).findall(".//{*}sp")
                        if shape.find("{*}nvSpPr/{*}cNvPr").get("name") == "Content"
                        for run in shape.findall(".//{*}r")]
            self.assertEqual("".join(run.findtext("{*}t") for run in runs), "普通加粗斜体文字结尾")
            styled = [run.findtext("{*}t") for run in runs if run.find("{*}rPr").get("b") == "1"]
            self.assertEqual("".join(styled), "加粗斜体文字")
            self.assertTrue(all(run.find("{*}rPr").get("i") == "1" for run in runs if run.find("{*}rPr").get("b") == "1"))
            for run in runs:
                if run.find("{*}rPr").get("b") == "1":
                    self.assertEqual(run.find("{*}rPr").get("sz"), "1600")
                    self.assertEqual(run.find("{*}rPr/{*}latin").get("typeface"), "Times New Roman")
                    self.assertEqual(run.find("{*}rPr/{*}ea").get("typeface"), "宋体")
                    self.assertEqual(run.find("{*}rPr/{*}solidFill/{*}srgbClr").get("val"), "CC3300")
