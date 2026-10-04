"""验证正文局部加粗和斜体在分页转换后仍保持。"""

import tempfile
import unittest
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from k12.converters import extract_docx_blocks, build_pptx_from_docx, extract_pptx_slides, build_docx_from_slides, _pptx_styled_paragraphs
from k12.text_styles import WordTextStyles


class TextStyleTests(unittest.TestCase):
    def test_distributed_alignment_round_trip(self) -> None:
        """中文及泰文分散对齐在正文双向转换时保留原生属性。"""
        with tempfile.TemporaryDirectory() as temporary:
            source, ppt, output = (Path(temporary) / name for name in ("source.docx", "target.pptx", "result.docx"))
            paragraphs = ''.join(f'<w:p><w:pPr><w:jc w:val="{value}"/></w:pPr><w:r><w:t>{value}</w:t></w:r></w:p>' for value in ("distribute", "thaiDistribute"))
            with zipfile.ZipFile(source, "w") as archive:
                archive.writestr("word/document.xml", f'<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>标题</w:t></w:r></w:p>{paragraphs}</w:body></w:document>')
            build_pptx_from_docx(extract_docx_blocks(source), ppt)
            slides = extract_pptx_slides(ppt)
            self.assertEqual([runs[0]["alignment"] for slide in slides for runs in slide["body_runs"]], ["distribute", "thaiDistribute"])
            build_docx_from_slides(slides, output, generate_toc=False)
            with zipfile.ZipFile(output) as archive:
                root = ET.fromstring(archive.read("word/document.xml"))
            for value in ("distribute", "thaiDistribute"):
                paragraph = next(node for node in root.findall(".//{*}p") if node.findtext("{*}r/{*}t") == value)
                self.assertEqual(paragraph.find("{*}pPr/{*}jc").get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val"), value)

    def test_word_alignment_inheritance(self) -> None:
        """默认值、父样式和直接段落对齐按优先级覆盖。"""
        namespace = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
        styles = ET.fromstring(f'<w:styles {namespace}><w:docDefaults><w:pPrDefault><w:pPr><w:jc w:val="left"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:styleId="Base"><w:pPr><w:jc w:val="right"/></w:pPr></w:style><w:style w:type="paragraph" w:styleId="Child"><w:basedOn w:val="Base"/></w:style></w:styles>')
        resolver = WordTextStyles(styles)
        for properties, expected in (("", "left"), ('<w:pStyle w:val="Child"/>', "right"), ('<w:pStyle w:val="Child"/><w:jc w:val="center"/>', "center"), ('<w:pStyle w:val="Child"/><w:jc w:val="numTab"/>', None)):
            paragraph = ET.fromstring(f'<w:p {namespace}><w:pPr>{properties}</w:pPr><w:r/></w:p>')
            self.assertEqual(resolver.run(paragraph, paragraph.find("{*}r")).get("alignment"), expected)

    def test_word_alignment_survives_pagination(self) -> None:
        """分页后的每个正文段落都写入源段落对齐。"""
        with tempfile.TemporaryDirectory() as temporary:
            source, target = Path(temporary) / "source.docx", Path(temporary) / "target.pptx"
            with zipfile.ZipFile(source, "w") as archive:
                archive.writestr("word/document.xml", '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>标题</w:t></w:r></w:p><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>分页正文保留段落对齐</w:t></w:r></w:p></w:body></w:document>')
            build_pptx_from_docx(extract_docx_blocks(source), target, max_chars=5)
            with zipfile.ZipFile(target) as archive:
                paragraphs = [paragraph for name in archive.namelist() if name.startswith("ppt/slides/slide") and name.endswith(".xml")
                              for shape in ET.fromstring(archive.read(name)).findall(".//{*}sp")
                              if shape.find("{*}nvSpPr/{*}cNvPr").get("name") == "Content"
                              for paragraph in shape.findall("{*}txBody/{*}p")]
            self.assertGreater(len(paragraphs), 1)
            self.assertTrue(all(paragraph.find("{*}pPr").get("algn") == "ctr" for paragraph in paragraphs))
            self.assertTrue(all(runs[0]["alignment"] == "center" for slide in extract_pptx_slides(target) for runs in slide["body_runs"]))

    def test_ppt_alignment_inheritance_and_word_output(self) -> None:
        """段落对齐在无 defRPr 时也能继承，直接对齐覆盖母版。"""
        namespace = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
        master = ET.fromstring(f'<p:sldMaster {namespace}><p:txStyles><p:otherStyle><a:lvl1pPr algn="r"/></p:otherStyle></p:txStyles></p:sldMaster>')
        xml = f'<p:sld {namespace}><p:sp><p:txBody><a:p><a:r><a:t>正文</a:t></a:r></a:p></p:txBody></p:sp></p:sld>'
        runs = _pptx_styled_paragraphs(xml, master=master)[0]
        self.assertEqual(runs[0]["alignment"], "right")
        centered = xml.replace('<a:p>', '<a:p><a:pPr algn="ctr"/>')
        runs = _pptx_styled_paragraphs(centered, master=master)[0]
        self.assertEqual(runs[0]["alignment"], "center")
        with tempfile.TemporaryDirectory() as temporary:
            target = Path(temporary) / "target.docx"
            build_docx_from_slides([{"title": "标题", "body": ["正文"], "body_runs": [runs]}], target, generate_toc=False)
            with zipfile.ZipFile(target) as archive:
                root = ET.fromstring(archive.read("word/document.xml"))
            paragraph = next(node for node in root.findall(".//{*}p") if ''.join(node.itertext()) == "正文")
            self.assertEqual(paragraph.find("{*}pPr/{*}jc").get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val"), "center")

    def test_ppt_theme_color_mapping_and_transform_limitation(self) -> None:
        """色板覆盖和母版重置遵循来源层次，未解析变换不冒充基础颜色。"""
        namespace = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
        theme = ET.fromstring(f'<a:theme {namespace}><a:themeElements><a:clrScheme><a:accent1><a:srgbClr val="112233"/></a:accent1><a:accent2><a:srgbClr val="CC3300"/></a:accent2></a:clrScheme></a:themeElements></a:theme>')
        master = ET.fromstring(f'<p:sldMaster {namespace}><p:clrMap tx1="accent1"/></p:sldMaster>')
        layout = ET.fromstring(f'<p:sldLayout {namespace}><p:clrMapOvr><a:overrideClrMapping tx1="accent2"/></p:clrMapOvr></p:sldLayout>')
        slide = f'<p:sld {namespace}><p:sp><p:txBody><a:p><a:r><a:rPr><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:rPr><a:t>文字</a:t></a:r></a:p></p:txBody></p:sp></p:sld>'
        self.assertEqual(_pptx_styled_paragraphs(slide, master=master, theme=theme)[0][0]["color"], "112233")
        self.assertEqual(_pptx_styled_paragraphs(slide, layout, master, theme)[0][0]["color"], "CC3300")
        reset = slide.replace('</p:sld>', '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>')
        self.assertEqual(_pptx_styled_paragraphs(reset, layout, master, theme)[0][0]["color"], "112233")
        transformed = slide.replace('<a:schemeClr val="tx1"/>', '<a:schemeClr val="tx1"><a:lumMod val="50000"/></a:schemeClr>')
        run = _pptx_styled_paragraphs(transformed, layout, master, theme)[0][0]
        self.assertNotIn("color", run)
        self.assertTrue(run["theme_color_unresolved"])
        with tempfile.TemporaryDirectory() as temporary:
            warnings = build_docx_from_slides([{"title": "文字", "title_runs": [run], "body": []}], Path(temporary) / "target.docx", generate_toc=False)
            self.assertTrue(any("颜色及填充变换" in warning for warning in warnings))

    def test_ppt_language_selects_east_asian_theme_font(self) -> None:
        """语言属性随段落继承和局部覆盖，选择相应东亚补充字体。"""
        theme = ET.fromstring('<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:fontScheme><a:minorFont><a:ea typeface=""/><a:font script="Hans" typeface="宋体"/><a:font script="Hant" typeface="新細明體"/><a:font script="Jpan" typeface="MS Mincho"/><a:font script="Hang" typeface="Batang"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>')
        xml = '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:sp><p:txBody><a:p><a:pPr><a:defRPr lang="zh-CN"><a:ea typeface="+mn-ea"/></a:defRPr></a:pPr><a:r><a:t>中文</a:t></a:r><a:r><a:rPr lang="zh-TW"/><a:t>繁體</a:t></a:r><a:r><a:rPr lang="ja-JP"/><a:t>日文</a:t></a:r><a:r><a:rPr lang="ko-KR"/><a:t>韩文</a:t></a:r></a:p></p:txBody></p:sp></p:sld>'
        runs = _pptx_styled_paragraphs(xml, theme=theme)[0]
        self.assertEqual([run["east_asia"] for run in runs], ["宋体", "新細明體", "MS Mincho", "Batang"])
        self.assertTrue(all("theme_font_unresolved" not in run for run in runs))
        with tempfile.TemporaryDirectory() as temporary:
            target = Path(temporary) / "target.docx"
            build_docx_from_slides([{"title": "标题", "body": ["中文繁體日文韩文"], "body_runs": [runs]}], target, generate_toc=False)
            result = next(block for block in extract_docx_blocks(target) if block["text"] == "中文繁體日文韩文")
            self.assertEqual([run["east_asia"] for run in result["runs"]], ["宋体", "新細明體", "MS Mincho", "Batang"])

    def test_ppt_theme_fonts_and_direct_override(self) -> None:
        """PPT 主题引用解析最终字体，直接字体覆盖会清除旧引用。"""
        theme = ET.fromstring('<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:fontScheme><a:majorFont><a:latin typeface="Times New Roman"/><a:ea typeface="宋体"/></a:majorFont><a:minorFont><a:latin typeface="Arial"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>')
        xml = '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:sp><p:txBody><a:p><a:pPr><a:defRPr><a:latin typeface="+mj-lt"/><a:ea typeface="+mj-ea"/></a:defRPr></a:pPr><a:r><a:t>主题</a:t></a:r><a:r><a:rPr><a:latin typeface="Arial"/></a:rPr><a:t>覆盖</a:t></a:r></a:p></p:txBody></p:sp></p:sld>'
        runs = _pptx_styled_paragraphs(xml, theme=theme)[0]
        self.assertEqual([run["font"] for run in runs], ["Times New Roman", "Arial"])
        self.assertEqual([run["east_asia"] for run in runs], ["宋体", "宋体"])
        unresolved = _pptx_styled_paragraphs(xml)[0]
        self.assertTrue(unresolved[0]["theme_font_unresolved"])
        with tempfile.TemporaryDirectory() as temporary:
            warnings = build_docx_from_slides([{"title": "标题", "body": ["主题覆盖"], "body_runs": [unresolved]}], Path(temporary) / "target.docx", generate_toc=False)
            self.assertTrue(any("PPT 主题字体" in warning for warning in warnings))

    def test_master_placeholder_matches_type_before_layout_override(self) -> None:
        """母版按类型匹配，不要求 idx 相同，版式及直接格式仍优先。"""
        namespace = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
        master = ET.fromstring(f'<p:sldMaster {namespace}><p:cSld><p:spTree><p:sp><p:nvSpPr><p:nvPr><p:ph type="body" idx="2"/></p:nvPr></p:nvSpPr><p:txBody><a:p><a:pPr><a:defRPr b="1" sz="2000"><a:latin typeface="Times New Roman"/></a:defRPr></a:pPr></a:p></p:txBody></p:sp></p:spTree></p:cSld><p:txStyles><p:bodyStyle><a:lvl1pPr><a:defRPr sz="1600"/></a:lvl1pPr></p:bodyStyle></p:txStyles></p:sldMaster>')
        slide = f'<p:sld {namespace}><p:sp><p:nvSpPr><p:nvPr><p:ph type="body" idx="7"/></p:nvPr></p:nvSpPr><p:txBody><a:p><a:r><a:t>继承</a:t></a:r><a:r><a:rPr b="0"/><a:t>关闭</a:t></a:r></a:p></p:txBody></p:sp></p:sld>'
        runs = _pptx_styled_paragraphs(slide, master=master)[0]
        self.assertEqual(runs[0]["size"], 2000)
        self.assertEqual(runs[0]["font"], "Times New Roman")
        self.assertTrue(runs[0]["b"])
        self.assertFalse(runs[1]["b"])
        layout = ET.fromstring(f'<p:sldLayout {namespace}><p:sp><p:nvSpPr><p:nvPr><p:ph type="body" idx="7"/></p:nvPr></p:nvSpPr><p:txBody><a:p><a:pPr><a:defRPr sz="1800"/></a:pPr></a:p></p:txBody></p:sp></p:sldLayout>')
        self.assertEqual(_pptx_styled_paragraphs(slide, layout, master)[0][0]["size"], 1800)

    def test_master_text_styles_and_slide_override(self) -> None:
        """母版标题、正文和其他样式各自应用，幻灯片直接格式优先。"""
        master = ET.fromstring('<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:txStyles><p:titleStyle><a:lvl1pPr><a:defRPr b="1" sz="3200"/></a:lvl1pPr></p:titleStyle><p:bodyStyle><a:lvl1pPr><a:defRPr i="1" sz="1800"/></a:lvl1pPr></p:bodyStyle><p:otherStyle><a:lvl1pPr><a:defRPr sz="1200"/></a:lvl1pPr></p:otherStyle></p:txStyles></p:sldMaster>')
        for kind, size in (("title", 3200), ("body", 1800), ("", 1200)):
            placeholder = f'<p:ph type="{kind}"/>' if kind else ''
            xml = f'<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:sp><p:nvSpPr><p:nvPr>{placeholder}</p:nvPr></p:nvSpPr><p:txBody><a:p><a:r><a:t>继承</a:t></a:r><a:r><a:rPr sz="1400"/><a:t>覆盖</a:t></a:r></a:p></p:txBody></p:sp></p:sld>'
            runs = _pptx_styled_paragraphs(xml, master=master)[0]
            self.assertEqual(runs[0]["size"], size)
            self.assertEqual(runs[1]["size"], 1400)
            if kind == "title":
                self.assertTrue(runs[0]["b"])
            if kind == "body":
                self.assertTrue(runs[0]["i"])

    def test_layout_placeholder_matches_index(self) -> None:
        """版式按 idx 匹配，不把邻近占位符格式应用到正文。"""
        with tempfile.TemporaryDirectory() as temporary:
            source, target = Path(temporary) / "source.pptx", Path(temporary) / "target.docx"
            with zipfile.ZipFile(source, "w") as archive:
                archive.writestr("ppt/slides/slide1.xml", '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:sp><p:nvSpPr><p:nvPr><p:ph idx="7"/></p:nvPr></p:nvSpPr><p:txBody><a:p><a:r><a:t>正文</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>')
                archive.writestr("ppt/slides/_rels/slide1.xml.rels", '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="layout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../custom/layout.xml"/></Relationships>')
                archive.writestr("ppt/custom/layout.xml", '<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:sp><p:nvSpPr><p:nvPr><p:ph idx="8"/></p:nvPr></p:nvSpPr><p:txBody><a:p><a:pPr><a:defRPr sz="4000"/></a:pPr></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:nvPr><p:ph idx="7"/></p:nvPr></p:nvSpPr><p:txBody><a:p><a:pPr><a:defRPr b="1" sz="1600"><a:latin typeface="Times New Roman"/></a:defRPr></a:pPr></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sldLayout>')
            slides = extract_pptx_slides(source)
            self.assertEqual(slides[0]["title_runs"][0]["size"], 1600)
            self.assertTrue(slides[0]["title_runs"][0]["b"])
            build_docx_from_slides(slides, target, generate_toc=False)
            run = next(block for block in extract_docx_blocks(target) if block["text"] == "正文")["runs"][0]
            self.assertEqual(run["size"], 1600)
            self.assertEqual(run["font"], "Times New Roman")

    def test_ppt_textbox_list_level_hierarchy(self) -> None:
        """文本框默认、当前层级、段落和片段按顺序逐属性覆盖。"""
        xml = '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:sp><p:txBody><a:lstStyle><a:defPPr><a:defRPr b="1" sz="1200"><a:latin typeface="Arial"/></a:defRPr></a:defPPr><a:lvl1pPr><a:defRPr sz="1600"/></a:lvl1pPr><a:lvl2pPr><a:defRPr sz="2000" i="1"/></a:lvl2pPr></a:lstStyle><a:p><a:r><a:t>一级</a:t></a:r></a:p><a:p><a:pPr lvl="1"><a:defRPr sz="1800"/></a:pPr><a:r><a:t>二级</a:t></a:r><a:r><a:rPr b="0" sz="1400"/><a:t>局部</a:t></a:r></a:p><a:p><a:pPr lvl="2"/><a:r><a:t>默认</a:t></a:r></a:p></p:txBody></p:sp></p:sld>'
        groups = _pptx_styled_paragraphs(xml)
        self.assertEqual(groups[0][0]["size"], 1600)
        self.assertEqual([run["size"] for run in groups[1]], [1800, 1400])
        self.assertTrue(groups[1][0]["i"])
        self.assertFalse(groups[1][1]["b"])
        self.assertEqual(groups[2][0]["size"], 1200)
        self.assertNotIn("i", groups[2][0])
        with tempfile.TemporaryDirectory() as temporary:
            target = Path(temporary) / "target.docx"
            build_docx_from_slides([{"title": "标题", "body": ["一级", "二级局部", "默认"], "body_runs": groups}], target, generate_toc=False)
            blocks = extract_docx_blocks(target)
            self.assertEqual(next(block for block in blocks if block["text"] == "二级局部")["runs"][1]["size"], 1400)
        with self.assertRaisesRegex(ValueError, "级别"):
            _pptx_styled_paragraphs(xml.replace('lvl="2"', 'lvl="9"'))

    def test_ppt_paragraph_defaults_and_local_override(self) -> None:
        """段落默认格式逐属性继承，局部关闭不影响后续片段。"""
        xml = '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:sp><p:txBody><a:p><a:pPr><a:defRPr b="1" i="1" sz="1600" u="dbl"><a:solidFill><a:srgbClr val="CC3300"/></a:solidFill><a:uFill><a:solidFill><a:srgbClr val="00AA00"/></a:solidFill></a:uFill><a:latin typeface="Times New Roman"/></a:defRPr></a:pPr><a:r><a:t>继承</a:t></a:r><a:r><a:rPr b="0" u="none"><a:uFillTx/></a:rPr><a:t>关闭</a:t></a:r><a:r><a:t>恢复</a:t></a:r></a:p></p:txBody></p:sp></p:sld>'
        runs = _pptx_styled_paragraphs(xml)[0]
        self.assertTrue(runs[0]["b"] and runs[2]["b"])
        self.assertFalse(runs[1]["b"])
        self.assertEqual(runs[1]["size"], 1600)
        self.assertEqual(runs[1]["color"], "CC3300")
        self.assertNotIn("underline_color", runs[1])
        self.assertEqual(runs[2]["underline_color"], "00AA00")
        with tempfile.TemporaryDirectory() as temporary:
            target = Path(temporary) / "target.docx"
            build_docx_from_slides([{"title": "标题", "body": ["继承关闭恢复"], "body_runs": [runs]}], target, generate_toc=False)
            result = next(block for block in extract_docx_blocks(target) if block["text"] == "继承关闭恢复")
            self.assertEqual([run["b"] for run in result["runs"]], [True, False, True])
            self.assertTrue(all(run["size"] == 1600 and run["font"] == "Times New Roman" for run in result["runs"]))

    def test_pptx_to_word_direct_format_roundtrip(self) -> None:
        """PPT 转讲义和大纲保留各片段显式格式及软换行。"""
        with tempfile.TemporaryDirectory() as temporary:
            source, ppt = Path(temporary) / "source.docx", Path(temporary) / "source.pptx"
            with zipfile.ZipFile(source, "w") as archive:
                archive.writestr("word/document.xml", '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>标题</w:t></w:r></w:p><w:p><w:r><w:rPr><w:b/><w:i/><w:u w:val="wavyDouble" w:color="00AA00"/><w:sz w:val="32"/><w:rFonts w:ascii="Times New Roman" w:eastAsia="宋体"/><w:color w:val="cc3300"/></w:rPr><w:t>格式文字</w:t><w:br/><w:t>第二行</w:t></w:r><w:r><w:rPr><w:b w:val="0"/><w:u w:val="none"/></w:rPr><w:t>关闭加粗</w:t></w:r></w:p></w:body></w:document>')
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
                self.assertTrue(all(run["underline"] == "wavyDbl" and run["underline_color"] == "00AA00" for run in styled))
                self.assertEqual(body["runs"][-1]["underline"], "none")

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
