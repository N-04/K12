"""验证原生 PPT 图表类型与内嵌数据一致。"""

from io import BytesIO
from pathlib import Path
import tempfile
import unittest
import zipfile
import xml.etree.ElementTree as ET

from k12.charts import build_chart
from k12.converters import build_pptx_from_xlsx


class NativeChartTests(unittest.TestCase):
    """覆盖折线、饼图及不支持的图表结构。"""

    def test_line_and_pie_use_original_type_and_workbook(self) -> None:
        """图表应使用对应绘图区，饼图无坐标轴，数据工作簿可读取。"""
        namespace = {"c": "http://schemas.openxmlformats.org/drawingml/2006/chart"}
        series = [{"name": "收入", "categories": {0: "一月", 1: "二月"}, "values": {0: "0", 1: "12345678901234567890.123"}}]
        for kind in ["lineChart", "pieChart"]:
            with self.subTest(kind=kind), tempfile.TemporaryDirectory() as tmp:
                target = Path(tmp) / "chart.pptx"
                build_pptx_from_xlsx([{"name": "统计", "cells": [], "charts": [{"chart_type": kind, "series": series}]}], target, "统计")
                with zipfile.ZipFile(target) as archive:
                    chart_name = next(name for name in archive.namelist() if name.startswith("ppt/charts/chart"))
                    root = ET.fromstring(archive.read(chart_name))
                    self.assertIsNotNone(root.find(f".//c:{kind}", namespace))
                    self.assertEqual(len(root.findall(".//c:catAx", namespace)), 0 if kind == "pieChart" else 1)
                    embedded = next(name for name in archive.namelist() if name.endswith(".xlsx"))
                    with zipfile.ZipFile(BytesIO(archive.read(embedded))) as workbook:
                        xml = workbook.read("xl/worksheets/sheet1.xml").decode()
                    self.assertIn("12345678901234567890.123", xml)
                    self.assertIn("一月", xml)

    def test_invalid_pie_and_nonfinite_values_are_rejected(self) -> None:
        """多个饼图系列与非有限数值不能被悄悄丢弃或写入。"""
        series = {"name": "收入", "categories": {0: "一月"}, "values": {0: "1"}}
        with self.assertRaisesRegex(ValueError, "单个系列"):
            build_chart({"chart_type": "pieChart", "series": [series, series]})
        for value in ["NaN", "Infinity", "不是数值"]:
            with self.subTest(value=value), self.assertRaises(ValueError):
                build_chart({"chart_type": "lineChart", "series": [{**series, "values": {0: value}}]})

    def test_each_line_series_retains_markers_and_smoothing(self) -> None:
        """源图表中的系列标记、大小和平滑开关应分别保留。"""
        from k12.converters import _xlsx_chart_data
        buffer = BytesIO()
        with zipfile.ZipFile(buffer, "w") as archive:
            archive.writestr("xl/charts/chart1.xml", '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:lineChart><c:grouping val="standard"/><c:ser><c:tx><c:v>平滑系列</c:v></c:tx><c:marker><c:symbol val="diamond"/><c:size val="9"/></c:marker><c:cat><c:strLit><c:pt idx="0"><c:v>甲</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>10</c:v></c:pt></c:numLit></c:val><c:smooth val="1"/></c:ser><c:ser><c:tx><c:v>直线系列</c:v></c:tx><c:marker><c:symbol val="none"/></c:marker><c:val><c:numLit><c:pt idx="0"><c:v>0</c:v></c:pt></c:numLit></c:val><c:smooth val="0"/></c:ser></c:lineChart></c:plotArea></c:chart></c:chartSpace>')
        with zipfile.ZipFile(buffer) as archive:
            chart = _xlsx_chart_data(archive, "xl/charts/chart1.xml")
        xml, workbook = build_chart(chart)
        root = ET.fromstring(xml)
        ns = {"c": "http://schemas.openxmlformats.org/drawingml/2006/chart"}
        series = root.findall(".//c:ser", ns)
        self.assertEqual([item.find("c:marker/c:symbol", ns).get("val") for item in series], ["diamond", "none"])
        self.assertEqual(series[0].find("c:marker/c:size", ns).get("val"), "9")
        self.assertEqual([item.find("c:smooth", ns).get("val") for item in series], ["1", "0"])
        self.assertTrue(zipfile.is_zipfile(BytesIO(workbook)))

    def test_empty_source_values_keep_warning_and_do_not_generate_chart(self) -> None:
        """全空源范围不能被标记恢复成功，零值仍视为有效数据。"""
        from k12.converters import _xlsx_resolve_chart_sources
        def make_sheet(value: str) -> dict:
            """构造未保存图表缓存的源工作表。"""
            return {"name": "统计", "cells": [{"ref": "A1", "value": value, "result": value}],
                    "charts": [{"chart_type": "lineChart", "series": [{"name": "收入", "name_cached": True,
                                "categories": {}, "values": {}, "value_formula": "统计!A1:A2"}]}]}
        blank = make_sheet("")
        _xlsx_resolve_chart_sources([blank])
        self.assertIn("warning", blank["charts"][0])
        with self.assertRaisesRegex(ValueError, "可用数据"):
            build_chart(blank["charts"][0])
        zero = make_sheet("0")
        _xlsx_resolve_chart_sources([zero])
        self.assertNotIn("warning", zero["charts"][0])
        self.assertEqual(zero["charts"][0]["series"][0]["values"], {0: "0", 1: ""})

    def test_missing_series_name_keeps_recovery_warning(self) -> None:
        """引用不存在的系列名时保留默认名称，但不能声称完整恢复。"""
        from k12.converters import _xlsx_resolve_chart_sources
        chart = {"series": [{"name": "系列 1", "name_formula": "统计!A1", "name_cached": False,
                            "categories": {0: "一月"}, "values": {0: "10"}}]}
        _xlsx_resolve_chart_sources([{"name": "统计", "cells": [], "charts": [chart]}])
        self.assertIn("warning", chart)
        self.assertEqual(chart["series"][0]["name"], "系列 1")

    def test_chart_warning_is_visible_in_task_and_exported_report(self) -> None:
        """图表部件无效时仍输出数据页，但任务报告必须提示限制。"""
        from test_processor import make_xlsx_bytes, make_xlsx_object_files
        from k12.processor import TaskProcessor
        from k12.store import AppStore
        for task_type in ["excel_to_pdf", "excel_to_word", "excel_to_ppt"]:
            with self.subTest(task_type=task_type), tempfile.TemporaryDirectory() as tmp:
                store = AppStore(tmp)
                processor = TaskProcessor(store)
                file = processor.create_uploaded_file("chart.xlsx", make_xlsx_bytes(make_xlsx_object_files()))[0]
                task = processor.create_task({"task_type": task_type, "file_ids": [file["id"]]})
                self.assertEqual(task["status"], "成功")
                report = store.list_reports()[0]
                artifact = report["analysis"]["artifacts"][0]
                self.assertIn("图表部件缺失或格式无效", artifact["message"])
                self.assertIn("图表部件缺失或格式无效", Path(report["txt_path"]).read_text())

    def test_unsupported_chart_returns_warning_with_data_page(self) -> None:
        """未知图表类型保留数据并返回可进入报告的提示。"""
        with tempfile.TemporaryDirectory() as tmp:
            warnings = build_pptx_from_xlsx([{"name": "统计", "cells": [], "charts": [{"chart_type": "radarChart", "series": [{"name": "收入", "categories": {0: "一月"}, "values": {0: "10"}}]}]}], Path(tmp) / "chart.pptx", "统计")
            self.assertTrue(any("未生成原生图表" in warning for warning in warnings))
