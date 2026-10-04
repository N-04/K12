"""验证 Excel 合并区域转换为原生 Word 表格。"""

import re
import tempfile
import unittest
import zipfile
from pathlib import Path
import xml.etree.ElementTree as ET

from k12.converters import build_pdf_from_xlsx, extract_pdf_text_blocks, build_docx_from_xlsx, build_pptx_from_xlsx, extract_pptx_slides, extract_xlsx_sheets
from k12.excel import excel_cell_refs, select_excel_sheets


class ExcelMergedTableTests(unittest.TestCase):
    """覆盖命名空间、横向和纵向合并及容量边界。"""

    def test_merged_region_survives_xlsx_to_docx(self) -> None:
        """实际 XLSX 合并区域应生成 Word 网格跨度和纵向合并。"""
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "merged.xlsx"
            with zipfile.ZipFile(source, "w") as archive:
                archive.writestr("xl/workbook.xml", '<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="合并表" r:id="rId1"/></sheets></workbook>')
                archive.writestr("xl/_rels/workbook.xml.rels", '<Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>')
                archive.writestr("xl/worksheets/sheet1.xml", '''<s:worksheet xmlns:s="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><s:sheetData>
                <s:row r="2"><s:c r="B2" t="inlineStr"><s:is><s:t>合并标题</s:t></s:is></s:c></s:row>
                <s:row r="4"><s:c r="D4" t="inlineStr"><s:is><s:t>末行</s:t></s:is></s:c></s:row>
                </s:sheetData><s:mergeCells><s:mergeCell ref="B2:C3"/></s:mergeCells></s:worksheet>''')
            sheets = extract_xlsx_sheets(source)
            self.assertEqual(sheets[0]["merge_ranges"], ["B2:C3"])
            self.assertEqual(sheets[0]["merged_cells"], 1)
            target = Path(tmp) / "merged.docx"
            build_docx_from_xlsx(sheets, target, "合并表格")
            with zipfile.ZipFile(target) as archive:
                root = ET.fromstring(archive.read("word/document.xml"))
            ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
            spans = root.findall(".//w:gridSpan", ns)
            self.assertEqual([node.get(f'{{{ns["w"]}}}val') for node in spans], ["2", "2"])
            merges = root.findall(".//w:vMerge", ns)
            self.assertEqual([node.get(f'{{{ns["w"]}}}val') for node in merges], ["restart", "continue"])
            text = "".join(root.itertext())
            self.assertEqual(text.count("合并标题"), 1)
            self.assertIn("末行", text)

    def test_oversized_merge_fails_explicitly(self) -> None:
        """超宽合并区域应明确失败，避免悄悄拆散结构。"""
        with tempfile.TemporaryDirectory() as tmp:
            with self.assertRaisesRegex(ValueError, "容量"):
                build_docx_from_xlsx([{"cells": [{"ref": "A1", "value": "标题"}],
                                       "merge_ranges": ["A1:ZZ1"]}], Path(tmp) / "wide.docx", "宽表")

    def test_selection_ranges_preserve_only_complete_merges(self) -> None:
        """矩形选区保留完整合并，单点选区不引入其他行列。"""
        sheet = {"name": "表格", "index": 1, "cells": [{"ref": "B2", "value": "标题"},
                 {"ref": "D4", "value": "末行"}], "merge_ranges": ["B2:C3"], "merged_cells": 1}
        full = select_excel_sheets([sheet], {"conversionRange": "选区", "cellRefs": "B2:C3"}, {})[0]
        self.assertEqual(full["merge_ranges"], ["B2:C3"])
        self.assertEqual([cell["ref"] for cell in full["cells"]], ["B2"])
        partial = select_excel_sheets([sheet], {"conversionRange": "选区", "cellRefs": ["B2"]}, {})[0]
        self.assertEqual(partial["merge_ranges"], [])
        self.assertEqual(partial["merged_cells"], 0)
        with tempfile.TemporaryDirectory() as tmp:
            target = Path(tmp) / "selection.docx"
            build_docx_from_xlsx([partial], target, "选区")
            with zipfile.ZipFile(target) as archive:
                xml = archive.read("word/document.xml").decode()
            self.assertNotIn("gridSpan", xml)
            self.assertNotIn("vMerge", xml)
            self.assertNotIn("末行", xml)
        self.assertEqual(sheet["merge_ranges"], ["B2:C3"])

    def test_selection_validates_coordinates_and_capacity(self) -> None:
        """错误坐标和巨型范围应明确失败，不能静默输出空文件。"""
        self.assertEqual(excel_cell_refs(["a1:b2", "B2"]), {"A1", "B1", "A2", "B2"})
        for value in ["B2:A1", "XFE1", "A1048577", "A0", "A1:XFD1048576"]:
            with self.subTest(value=value), self.assertRaises(ValueError):
                excel_cell_refs([value])

    def test_chart_cache_exports_series_with_sparse_points(self) -> None:
        """图表缓存按点索引对齐，Word 和 PPT 保留类别及零值。"""
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "charts.xlsx"
            with zipfile.ZipFile(source, "w") as archive:
                archive.writestr("xl/workbook.xml", '<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="成绩" r:id="rId1"/></sheets></workbook>')
                archive.writestr("xl/_rels/workbook.xml.rels", '<Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>')
                archive.writestr("xl/worksheets/sheet1.xml", '<worksheet><sheetData/></worksheet>')
                archive.writestr("xl/worksheets/_rels/sheet1.xml.rels", '<Relationships><Relationship Target="../drawings/drawing1.xml"/></Relationships>')
                archive.writestr("xl/drawings/drawing1.xml", '<drawing/>')
                archive.writestr("xl/drawings/_rels/drawing1.xml.rels", '<Relationships><Relationship Target="../charts/chart1.xml"/></Relationships>')
                archive.writestr("xl/charts/chart1.xml", '<c:chartSpace xmlns:c="urn:chart" xmlns:a="urn:text"><c:chart><c:title><a:t>成绩对比</a:t></c:title><c:plotArea><c:barChart><c:ser><c:tx><c:v>语文</c:v></c:tx><c:cat><c:strRef><c:f>成绩!A1:A3</c:f><c:strCache><c:pt idx="2"><c:v>乙</c:v></c:pt><c:pt idx="0"><c:v>甲</c:v></c:pt></c:strCache></c:strRef></c:cat><c:val><c:numRef><c:f>成绩!B1:B3</c:f><c:numCache><c:pt idx="0"><c:v>0</c:v></c:pt><c:pt idx="2"><c:v>98</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>')
            sheets = extract_xlsx_sheets(source)
            chart = sheets[0]["charts"][0]
            self.assertEqual(chart["title"], "成绩对比")
            self.assertEqual(chart["series"][0]["values"], {0: "0", 2: "98"})
            self.assertEqual(chart["series"][0]["value_formula"], "成绩!B1:B3")
            word, ppt = Path(tmp) / "chart.docx", Path(tmp) / "chart.pptx"
            build_docx_from_xlsx(sheets, word, "成绩")
            with zipfile.ZipFile(word) as archive:
                text = "".join(ET.fromstring(archive.read("word/document.xml")).itertext())
            self.assertIn("语文甲0", text)
            self.assertIn("语文乙98", text)
            pdf = Path(tmp) / "chart.pdf"
            warnings = build_pdf_from_xlsx(sheets, pdf, "成绩")
            pdf_text = "".join(item["text"] for item in extract_pdf_text_blocks(pdf))
            self.assertIn("成绩对比", pdf_text)
            self.assertIn("语文 · 甲: 0", pdf_text)
            self.assertIn("语文 · 乙: 98", pdf_text)
            self.assertEqual(warnings, [])
            build_pptx_from_xlsx(sheets, ppt, "成绩")
            slides = extract_pptx_slides(ppt)
            self.assertIn("语文 · 甲: 0", slides[-1]["body"])
            self.assertTrue(any(slide["chart_count"] == 1 for slide in slides))
            with zipfile.ZipFile(ppt) as archive:
                chart_name = next(name for name in archive.namelist() if name.startswith("ppt/charts/chart"))
                chart_xml = archive.read(chart_name).decode()
                self.assertIn("externalData", chart_xml)
                self.assertIn('ptCount val="3"', chart_xml)
                self.assertNotIn('<c:pt idx="1"><c:v>0', chart_xml)
                self.assertIn("0", chart_xml)
                workbook_name = next(name for name in archive.namelist() if name.endswith(".xlsx"))
                workbook_path = Path(tmp) / "embedded.xlsx"
                workbook_path.write_bytes(archive.read(workbook_name))
            embedded_cells = extract_xlsx_sheets(workbook_path)[0]["cells"]
            self.assertIn("语文", [cell["value"] for cell in embedded_cells])
            self.assertIn("0", [cell["value"] for cell in embedded_cells])

            selected = select_excel_sheets(sheets, {"conversionRange": "选区", "cellRefs": "A1"}, {})
            self.assertEqual(selected[0]["charts"], [])
            with zipfile.ZipFile(source) as archive:
                parts = {name: archive.read(name) for name in archive.namelist()}
            parts["xl/charts/chart1.xml"] = re.sub(rb"<c:(strCache|numCache)>.*?</c:\1>", b"", parts["xl/charts/chart1.xml"])
            parts["xl/worksheets/sheet1.xml"] = '<worksheet><sheetData><row><c r="A1" t="inlineStr"><is><t>甲</t></is></c><c r="B1"><v>0</v></c><c r="A3" t="inlineStr"><is><t>乙</t></is></c><c r="B3"><f>SUM(B2)</f><v>98</v></c></row></sheetData></worksheet>'.encode()
            with zipfile.ZipFile(source, "w") as archive:
                for name, content in parts.items():
                    archive.writestr(name, content)
            recovered = extract_xlsx_sheets(source)[0]["charts"][0]
            self.assertEqual(recovered["series"][0]["values"], {0: "0", 1: "", 2: "98"})
            self.assertEqual(recovered["series"][0]["categories"], {0: "甲", 1: "", 2: "乙"})
            self.assertNotIn("warning", recovered)


    def test_chart_source_references_and_uncalculated_formulas(self) -> None:
        """引用带引号表名可读取，外部引用和未计算公式不能当成数值。"""
        from k12.converters import _xlsx_chart_source_points
        sources = {"教'学": {"A1": {"value": "0", "result": "0"},
                              "B1": {"value": "=SUM(A1)", "formula": "SUM(A1)", "result": ""}}}
        self.assertEqual(_xlsx_chart_source_points("'教''学'!$A$1", sources), {0: "0"})
        for formula in ["'[外部.xlsx]教''学'!A1", "不存在!A1", "'教''学'!B1", "'教''学'!A1:B2"]:
            with self.subTest(formula=formula), self.assertRaises(ValueError):
                _xlsx_chart_source_points(formula, sources)

    def test_comments_keep_author_text_and_selected_cell(self) -> None:
        """批注正文按保留开关及选区筛选，富文本合并且作者不丢失。"""
        from k12.converters import _xlsx_sheet_objects
        from io import BytesIO
        buffer = BytesIO()
        with zipfile.ZipFile(buffer, "w") as archive:
            archive.writestr("xl/worksheets/_rels/sheet1.xml.rels", '<Relationships><Relationship Target="../comments1.xml"/></Relationships>')
            archive.writestr("xl/comments1.xml", '<s:comments xmlns:s="urn:sheet"><s:authors><s:author>老师</s:author></s:authors><s:commentList><s:comment ref="A1" authorId="0"><s:text><s:r><s:t>需要</s:t></s:r><s:r><s:t>复核 &amp; 确认</s:t></s:r></s:text></s:comment><s:comment ref="B2" authorId="8"><s:text><s:t>不在选区</s:t></s:text></s:comment></s:commentList></s:comments>')
        with zipfile.ZipFile(buffer) as archive:
            objects = _xlsx_sheet_objects(archive, "xl/worksheets/sheet1.xml", "<worksheet/>")
        sheet = {"name": "成绩", "index": 1, "cells": [{"ref": "A1", "value": "标题"}], **objects}
        self.assertEqual(objects["comment_count"], 2)
        self.assertEqual(objects["comments"][0], {"ref": "A1", "author": "老师", "text": "需要复核 & 确认"})
        selected = select_excel_sheets([sheet], {"conversionRange": "选区", "cellRefs": "A1", "retainComments": True}, {})
        self.assertEqual(selected[0]["comment_count"], 1)
        self.assertEqual(select_excel_sheets([sheet], {}, {})[0]["comments"], [])
        self.assertEqual(len(select_excel_sheets([sheet], {"retainComments": True}, {"retainComments": False})[0]["comments"]), 2)
        with tempfile.TemporaryDirectory() as tmp:
            target = Path(tmp) / "comments.docx"
            build_docx_from_xlsx(selected, target, "成绩")
            with zipfile.ZipFile(target) as archive:
                text = "".join(ET.fromstring(archive.read("word/document.xml")).itertext())
            self.assertIn("A1老师需要复核 & 确认", text)
            self.assertNotIn("不在选区", text)
            ppt = Path(tmp) / "comments.pptx"
            build_pptx_from_xlsx(selected, ppt, "成绩")
            self.assertIn("批注 A1 · 老师: 需要复核 & 确认", extract_pptx_slides(ppt)[-1]["body"])
            pdf = Path(tmp) / "comments.pdf"
            build_pdf_from_xlsx(selected, pdf, "成绩")
            pdf_text = "".join(item["text"] for item in extract_pdf_text_blocks(pdf))
            self.assertIn("需要复核 & 确认", pdf_text)
            self.assertNotIn("不在选区", pdf_text)
        self.assertEqual(len(sheet["comments"]), 2)

    def test_threaded_comments_inherit_cell_and_keep_reply_details(self) -> None:
        """线程回复继承根批注位置，保留作者、正文及解决状态。"""
        from k12.converters import _xlsx_sheet_objects
        from io import BytesIO
        buffer = BytesIO()
        with zipfile.ZipFile(buffer, "w") as archive:
            archive.writestr("xl/worksheets/_rels/sheet1.xml.rels", '<Relationships><Relationship Target="../threadedComments/threadedComment1.xml"/><Relationship Target="../comments1.xml"/></Relationships>')
            archive.writestr("xl/comments1.xml", '<comments><authors><author>tc={C}</author></authors><commentList><comment ref="C3" authorId="0"><text><t>兼容占位正文</t></text></comment></commentList></comments>')
            archive.writestr("xl/_rels/workbook.xml.rels", '<Relationships><Relationship Type="http://schemas.microsoft.com/office/2017/10/relationships/person" Target="persons/person.xml"/></Relationships>')
            archive.writestr("xl/persons/person.xml", '<personList xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"><person id="{P1}" displayName="老师" userId="不应导出"/><person id="{P2}" displayName="助教"/></personList>')
            archive.writestr("xl/threadedComments/threadedComment1.xml", '<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"><threadedComment id="{R}" parentId="{C}" personId="{P2}"><text>已复核 &amp; 确认</text></threadedComment><threadedComment id="{C}" ref="C3" personId="{P1}" dT="2026-10-03T00:00:00Z" done="1"><text>请复核</text></threadedComment></ThreadedComments>')
        with zipfile.ZipFile(buffer) as archive:
            objects = _xlsx_sheet_objects(archive, "xl/worksheets/sheet1.xml", "<worksheet/>")
        self.assertEqual(objects["comment_count"], 2)
        self.assertEqual([item["ref"] for item in objects["comments"]], ["C3", "C3"])
        self.assertEqual([item["author"] for item in objects["comments"]], ["助教", "老师"])
        self.assertTrue(objects["comments"][1]["resolved"])
        sheet = {"name": "成绩", "index": 1, "cells": [], **objects}
        selected = select_excel_sheets([sheet], {"conversionRange": "选区", "cellRefs": "C3", "retainComments": True}, {})
        with tempfile.TemporaryDirectory() as tmp:
            target = Path(tmp) / "thread.docx"
            build_docx_from_xlsx(selected, target, "批注")
            with zipfile.ZipFile(target) as archive:
                text = "".join(ET.fromstring(archive.read("word/document.xml")).itertext())
            self.assertIn("助教回复：已复核 & 确认", text)
            self.assertIn("老师请复核（已解决）", text)
            self.assertNotIn("不应导出", text)
            self.assertNotIn("兼容占位正文", text)
        outside = select_excel_sheets([sheet], {"conversionRange": "选区", "cellRefs": "A1", "retainComments": True}, {})
        self.assertEqual(outside[0]["comments"], [])

    def test_threaded_comment_cycle_fails_explicitly(self) -> None:
        """回复关系循环不能使转换无限等待或丢弃内容。"""
        from k12.converters import _xlsx_threaded_comments
        from io import BytesIO
        buffer = BytesIO()
        with zipfile.ZipFile(buffer, "w") as archive:
            archive.writestr("xl/threadedComments/threadedComment1.xml", '<ThreadedComments><threadedComment id="a" parentId="b"/><threadedComment id="b" parentId="a"/></ThreadedComments>')
        with zipfile.ZipFile(buffer) as archive, self.assertRaisesRegex(ValueError, "回复关系"):
            _xlsx_threaded_comments(archive, ["xl/threadedComments/threadedComment1.xml"])

    def test_legacy_thread_placeholder_is_deduplicated_by_id(self) -> None:
        """仅删除匹配线程标识的兼容批注，整条线程继承兼容坐标。"""
        from k12.converters import _xlsx_merge_comments
        legacy = [{"ref": "D4", "author": "tc={ROOT}", "text": "兼容占位"},
                  {"ref": "D4", "author": "老师", "text": "独立传统批注"},
                  {"ref": "A1", "author": "tc={MISSING}", "text": "无线程，须保留"}]
        threaded = [{"id": "{root}", "parent_id": "", "ref": "C3", "text": "线程正文"},
                    {"id": "{reply}", "parent_id": "{ROOT}", "ref": "C3", "text": "回复正文"}]
        combined = _xlsx_merge_comments(legacy, threaded)
        self.assertEqual([item["text"] for item in combined], ["独立传统批注", "无线程，须保留", "线程正文", "回复正文"])
        self.assertEqual([item["ref"] for item in combined[-2:]], ["D4", "D4"])

    def test_task_comment_override_matches_artifact_and_exported_report(self) -> None:
        """任务级批注设置应同时决定正文产物、设置快照和导出报告。"""
        from k12.processor import TaskProcessor
        from k12.store import AppStore
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "comments.xlsx"
            with zipfile.ZipFile(source, "w") as archive:
                archive.writestr("xl/workbook.xml", '<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="成绩" r:id="sheet"/></sheets></workbook>')
                archive.writestr("xl/_rels/workbook.xml.rels", '<Relationships><Relationship Id="sheet" Target="worksheets/sheet1.xml"/></Relationships>')
                archive.writestr("xl/worksheets/sheet1.xml", '<worksheet><sheetData><row><c r="A1" t="inlineStr"><is><t>标题</t></is></c></row></sheetData></worksheet>')
                archive.writestr("xl/worksheets/_rels/sheet1.xml.rels", '<Relationships><Relationship Target="../comments1.xml"/></Relationships>')
                archive.writestr("xl/comments1.xml", '<comments><authors><author>老师</author></authors><commentList><comment ref="A1" authorId="0"><text><t>必须复核的批注正文</t></text></comment></commentList></comments>')
            store = AppStore(Path(tmp) / "data")
            processor = TaskProcessor(store)
            file = processor.create_uploaded_file(source.name, source.read_bytes())[0]
            for retain in [True, False]:
                store.update_settings({"retainComments": not retain})
                task = processor.create_task({"task_type": "excel_to_word", "file_ids": [file["id"]], "options": {"excel": {"retainComments": retain}}})
                self.assertEqual(task["status"], "成功")
                report = next(item for item in store.list_reports() if item["task_id"] == task["id"])
                artifact = report["analysis"]["artifacts"][0]
                self.assertEqual(artifact["conversion_settings"]["excel_retain_comments"], retain)
                with zipfile.ZipFile(artifact["path"]) as archive:
                    xml = archive.read("word/document.xml").decode()
                self.assertEqual("必须复核的批注正文" in xml, retain)
                text = Path(report["txt_path"]).read_text()
                self.assertIn("Excel批注 " + ("是" if retain else "否"), text)
