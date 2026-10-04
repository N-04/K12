"""验证 Word 表格转换为可编辑 PPT 表格。"""

import tempfile
import unittest
import zipfile
from xml.etree import ElementTree as ET
from pathlib import Path

from k12.converters import build_docx, build_pptx_from_docx, extract_docx_blocks, extract_pptx_slides


class WordTableTests(unittest.TestCase):
    """覆盖表格顺序、分页以及关闭保留后的内容排除。"""

    def test_leading_table_does_not_promote_following_body(self) -> None:
        """表格后的普通正文不误判为首段标题。"""
        with tempfile.TemporaryDirectory() as tmp:
            source, target = Path(tmp) / "source.docx", Path(tmp) / "result.pptx"
            build_docx([{"table": [["开头表格"]]}, {"text": "表格后的普通正文"}], source)
            blocks = extract_docx_blocks(source)
            self.assertEqual(blocks[-1]["level"], 0)
            build_pptx_from_docx(blocks, target)
            slides = extract_pptx_slides(target)
            self.assertIn("表格后的普通正文", slides[-1]["body"])

    def test_native_table_preserves_cells_and_order_without_duplicate_text(self) -> None:
        """单元格只出现一次，表格前后正文顺序不变。"""
        with tempfile.TemporaryDirectory() as tmp:
            source, target = Path(tmp) / "table.docx", Path(tmp) / "table.pptx"
            rows = [["姓名", "成绩"], ["甲", "98"], ["乙", "0"]]
            build_docx([{"text": "课程标题", "style": "Heading1"}, {"text": "表格之前"}, {"table": rows}, {"text": "表格之后"}], source)
            blocks = extract_docx_blocks(source)
            self.assertEqual(sum("table" in block for block in blocks), 1)
            build_pptx_from_docx(blocks, target)
            slides = extract_pptx_slides(target)
            table_slide = next(slide for slide in slides if slide["tables"])
            self.assertEqual(table_slide["tables"], [rows])
            self.assertNotIn("姓名", table_slide["body"])
            self.assertIn("表格之前", slides[0]["body"])
            self.assertIn("表格之后", slides[-1]["body"])
            disabled = Path(tmp) / "disabled.pptx"
            build_pptx_from_docx(blocks, disabled, retain_tables=False)
            with zipfile.ZipFile(disabled) as archive:
                xml = b"".join(archive.read(name) for name in archive.namelist() if name.startswith("ppt/slides/"))
            self.assertNotIn("姓名".encode(), xml)
            self.assertNotIn(b"<a:tbl>", xml)

    def test_wide_long_table_pagination_preserves_every_cell(self) -> None:
        """超过一页的行列分段后，末行末列及全部内容仍然完整。"""
        with tempfile.TemporaryDirectory() as tmp:
            rows = [[f"CELL_{row}_{column}" for column in range(15)] for row in range(25)]
            target = Path(tmp) / "long.pptx"
            build_pptx_from_docx([{"text": "表格", "level": 0, "table": rows}], target)
            slides = extract_pptx_slides(target)
            cells = [cell for slide in slides for table in slide["tables"] for row in table for cell in row]
            self.assertEqual(len(slides), 6)
            self.assertEqual(len(cells), 375)
            self.assertEqual(set(cells), {cell for row in rows for cell in row})

    def test_combined_merge_grid_and_ppt_attributes(self) -> None:
        """横向和纵向合并保持位置、跨度和唯一正文。"""
        from k12.converters import _docx_table_grid
        xml = '<w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/><w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>合并内容</w:t></w:r></w:p></w:tc><w:tc><w:p/></w:tc></w:tr><w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/><w:vMerge/></w:tcPr><w:p/></w:tc><w:tc><w:p/></w:tc></w:tr></w:tbl>'
        rows, merges, _alignments = _docx_table_grid(ET.fromstring(xml))
        self.assertEqual(rows, [["合并内容", "", ""], ["", "", ""]])
        self.assertEqual(merges, [[0, 0, 2, 2]])
        with tempfile.TemporaryDirectory() as tmp:
            target = Path(tmp) / "merged.pptx"
            build_pptx_from_docx([{"text": "", "table": rows, "table_merges": merges}], target)
            with zipfile.ZipFile(target) as archive:
                root = ET.fromstring(archive.read("ppt/slides/slide1.xml"))
            cells = root.findall(".//{*}tc")
            self.assertEqual(cells[0].get("rowSpan"), "2")
            self.assertEqual(cells[0].get("gridSpan"), "2")
            self.assertEqual(cells[1].get("rowSpan"), "2")
            self.assertEqual(cells[3].get("gridSpan"), "2")
            self.assertEqual(cells[4].get("hMerge"), "1")
            self.assertEqual(cells[4].get("vMerge"), "1")
            rows = [[""] for _ in range(13)]
            build_pptx_from_docx([{"text": "", "table": rows, "table_merges": [[11, 0, 2, 1]]}], target)
            slides = extract_pptx_slides(target)
            self.assertEqual([len(slide["tables"][0]) for slide in slides], [11, 2])

    def test_merge_aware_page_ranges(self) -> None:
        """分页边界不切合并区，连续重叠和超大区域完整保留。"""
        from k12.converters import _table_page_ranges
        for total, merges, expected in [
            (25, [], [(0, 12), (12, 24), (24, 25)]),
            (25, [(11, 14)], [(0, 11), (11, 23), (23, 25)]),
            (30, [(0, 20), (18, 25)], [(0, 25), (25, 30)]),
        ]:
            with self.subTest(merges=merges):
                pages = _table_page_ranges(total, merges)
                self.assertEqual(pages, expected)
                self.assertFalse(any(first < end < last for _, end in pages for first, last in merges))

    def test_ppt_to_word_preserves_combined_merge(self) -> None:
        """PPT 转 Word 反向保留横纵合并，正文只写入起点。"""
        from k12.converters import build_docx_from_slides
        with tempfile.TemporaryDirectory() as tmp:
            ppt, word = Path(tmp) / "merged.pptx", Path(tmp) / "merged.docx"
            rows = [["合并内容", "", "末列"], ["", "", "0"]]
            build_pptx_from_docx([{"text": "", "table": rows, "table_merges": [[0, 0, 2, 2]]}], ppt)
            slides = extract_pptx_slides(ppt)
            self.assertEqual(slides[0]["table_merges"], [[[0, 0, 2, 2]]])
            build_docx_from_slides(slides, word)
            table = next(block for block in extract_docx_blocks(word) if "table" in block)
            self.assertEqual(table["table"], rows)
            self.assertEqual(table["table_merges"], [[0, 0, 2, 2]])

    def test_ppt_column_proportions_and_merged_width(self) -> None:
        """非等宽列比例进入 Word 网格，合并格宽度为覆盖列之和。"""
        from k12.converters import _pptx_table_widths, build_docx_from_slides
        xml = '<root><tbl><tblGrid><gridCol w="100"/><gridCol w="200"/><gridCol w="300"/></tblGrid></tbl></root>'
        widths = _pptx_table_widths(xml)
        self.assertEqual(widths, [[100, 200, 300]])
        with tempfile.TemporaryDirectory() as tmp:
            target = Path(tmp) / "widths.docx"
            build_docx_from_slides([{"title": "表格", "body": [], "tables": [[["合并", "", "末列"]]],
                                    "table_merges": [[[0, 0, 1, 2]]], "table_widths": widths}], target)
            with zipfile.ZipFile(target) as archive:
                document = ET.fromstring(archive.read("word/document.xml"))
            value_key = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}w"
            grid = document.findall(".//{*}tblGrid/{*}gridCol")
            self.assertEqual([int(column.get(value_key)) for column in grid], [1500, 3000, 4500])
            cells = document.findall(".//{*}tcPr/{*}tcW")
            self.assertEqual([int(cell.get(value_key)) for cell in cells], [4500, 4500])

    def test_word_column_widths_survive_roundtrip(self) -> None:
        """非等宽 Word 表格往返转换仍保持比例和合并结构。"""
        from k12.converters import build_docx_from_slides
        with tempfile.TemporaryDirectory() as tmp:
            word, ppt, result = [Path(tmp) / name for name in ["source.docx", "table.pptx", "result.docx"]]
            build_docx([{"table": [[{"text": "合并", "span": 2}, {"skip": True}, "末列"]],
                         "column_widths": [100, 200, 300]}], word)
            blocks = extract_docx_blocks(word)
            self.assertEqual(blocks[0]["column_widths"], [1500, 3000, 4500])
            build_pptx_from_docx(blocks, ppt)
            slides = extract_pptx_slides(ppt)
            self.assertEqual(slides[0]["table_widths"], [[1371600, 2743200, 4114800]])
            build_docx_from_slides(slides, result)
            table = next(block for block in extract_docx_blocks(result) if "table" in block)
            self.assertEqual(table["column_widths"], [1500, 3000, 4500])
            self.assertEqual(table["table_merges"], [[0, 0, 1, 2]])

    def test_row_height_proportion_and_word_minimum(self) -> None:
        """显式 Word 行高比例进入 PPT，反向使用最小行高避免裁剪正文。"""
        from k12.converters import build_docx_from_slides
        with tempfile.TemporaryDirectory() as tmp:
            word, ppt, result = [Path(tmp) / name for name in ["source.docx", "table.pptx", "result.docx"]]
            build_docx([{"table": [["短行"], ["长行"]], "row_heights": [360, 720]}], word)
            blocks = extract_docx_blocks(word)
            self.assertEqual(blocks[0]["row_heights"], [360, 720])
            build_pptx_from_docx(blocks, ppt)
            slides = extract_pptx_slides(ppt)
            self.assertEqual(slides[0]["table_heights"], [[2592, 5184]])
            build_docx_from_slides(slides, result)
            with zipfile.ZipFile(result) as archive:
                root = ET.fromstring(archive.read("word/document.xml"))
            heights = root.findall(".//{*}trPr/{*}trHeight")
            key = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
            self.assertEqual([height.get(key + "val") for height in heights], ["2592", "5184"])
            self.assertTrue(all(height.get(key + "hRule") == "atLeast" for height in heights))

    def test_mixed_automatic_row_heights(self) -> None:
        """自动行按换行及中文折行估算，不覆盖已指定行高。"""
        from k12.converters import _pptx_row_weights
        rows = [["显式高度"], ["第一行\n第二行"], ["中" * 20]]
        weights = _pptx_row_weights(rows, [1143000], [600, 0, 0])
        self.assertEqual(weights, [600, 800, 1600])
        self.assertEqual(_pptx_row_weights([["单行"]], [8229600], None), [400])
        with tempfile.TemporaryDirectory() as tmp:
            target = Path(tmp) / "mixed.pptx"
            build_pptx_from_docx([{"text": "", "table": rows, "row_heights": [600, 0, 0]}], target)
            with zipfile.ZipFile(target) as archive:
                root = ET.fromstring(archive.read("ppt/slides/slide1.xml"))
            heights = [int(row.get("h")) for row in root.findall(".//{*}tbl/{*}tr")]
            self.assertEqual(sum(heights), 4937760)
            self.assertGreater(heights[1], heights[0])

    def test_merged_auto_height_uses_total_width_and_rows(self) -> None:
        """合并单元格使用总列宽折行，所需高度均摊且显式行不变。"""
        from k12.converters import _pptx_row_weights
        rows = [["中" * 40, ""], ["", ""]]
        widths = [1143000, 1143000]
        self.assertEqual(_pptx_row_weights(rows, widths, None, [[0, 0, 2, 2]]), [800, 800])
        self.assertEqual(_pptx_row_weights(rows, widths, [600, 0], [[0, 0, 2, 2]]), [600, 1000])

    def test_cell_blank_paragraphs_and_spaces_survive_roundtrip(self) -> None:
        """单元格前后空段落、空行和空格在双向转换后保留。"""
        from k12.converters import _docx_visible_text, build_docx_from_slides
        cell = ET.fromstring('<tc><p/><p><r><t>  正文  </t></r></p><p/><p/></tc>')
        text = _docx_visible_text(cell)
        self.assertEqual(text, "\n  正文  \n\n")
        with tempfile.TemporaryDirectory() as tmp:
            ppt, word = Path(tmp) / "spacing.pptx", Path(tmp) / "spacing.docx"
            build_pptx_from_docx([{"text": "", "table": [[text]]}], ppt)
            slides = extract_pptx_slides(ppt)
            self.assertEqual(slides[0]["tables"], [[[text]]])
            build_docx_from_slides(slides, word)
            table = next(block for block in extract_docx_blocks(word) if "table" in block)
            self.assertEqual(table["table"], [[text]])

    def test_mixed_paragraph_alignment_roundtrip(self) -> None:
        """同一单元格多段的对齐方式通过 Word/PPT 往返保留。"""
        from k12.converters import build_docx_from_slides
        with tempfile.TemporaryDirectory() as tmp:
            word, ppt, result = [Path(tmp) / name for name in ["source.docx", "table.pptx", "result.docx"]]
            alignments = [[["center", "right", "both", "distribute", "thaiDistribute"]]]
            build_docx([{"table": [["居中\n右对齐\n两端对齐\n中文分散\n泰文分散"]], "cell_alignments": alignments}], word)
            blocks = extract_docx_blocks(word)
            self.assertEqual(blocks[0]["cell_alignments"], alignments)
            build_pptx_from_docx(blocks, ppt)
            slides = extract_pptx_slides(ppt)
            self.assertEqual(slides[0]["table_alignments"], [alignments])
            build_docx_from_slides(slides, result)
            table = next(block for block in extract_docx_blocks(result) if "table" in block)
            self.assertEqual(table["cell_alignments"], alignments)
