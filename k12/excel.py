"""共享 Excel 范围与公式设置，保证网页和本地转换一致。"""

import re
from typing import Any


def select_excel_sheets(sheets: list[dict[str, Any]], options: dict[str, Any], settings: dict[str, Any]) -> list[dict[str, Any]]:
    """根据工作表、选区和公式偏好筛选转换内容。"""
    return _ExcelSelection(settings).select(sheets, options)


def effective_excel_settings(options: dict[str, Any], settings: dict[str, Any]) -> dict[str, Any]:
    """统一计算任务覆盖后的 Excel 范围、公式及分表设置。"""
    range_mode = str(options.get("conversionRange") or settings.get("excelConversionRange") or "全部工作表")
    if range_mode not in {"全部工作表", "当前工作表", "选区"}:
        range_mode = "全部工作表"
    split = options.get("splitSheets", options.get("splitWorksheets", settings.get("excelSplitSheets", False)))
    return {
        "excel_conversion_range": range_mode,
        "excel_formula_mode": _ExcelSelection._excel_formula_mode(options, settings),
        "excel_split_sheets": bool(split),
        "excel_retain_comments": bool(options.get("retainComments", settings.get("retainComments", False))),
    }


def excel_cell_refs(values: list[str]) -> set[str]:
    """展开单元格或矩形选区，校验 Excel 行列及选区容量。"""
    refs: set[str] = set()
    for value in values:
        match = re.fullmatch(r"([A-Z]+)([1-9][0-9]*)(?::([A-Z]+)([1-9][0-9]*))?", value.strip().upper())
        if not match:
            raise ValueError("单元格选区无效，请使用 A1 或 A1:B3 格式")
        left, top, right, bottom = match.groups()
        first, last = excel_column_number(left), excel_column_number(right or left)
        top, bottom = int(top), int(bottom or top)
        if first > last or top > bottom or last > 16384 or bottom > 1048576:
            raise ValueError("单元格选区超出 Excel 范围")
        if (last - first + 1) * (bottom - top + 1) > 200000:
            raise ValueError("单元格选区过大，请分批转换")
        for column in range(first, last + 1):
            letters = excel_column_letters(column)
            refs.update(f"{letters}{row}" for row in range(top, bottom + 1))
        if len(refs) > 200000:
            raise ValueError("单元格选区过大，请分批转换")
    return refs


def excel_column_number(letters: str) -> int:
    """将 Excel 列名换算为列序号。"""
    number = 0
    for letter in letters:
        number = number * 26 + ord(letter) - ord("A") + 1
    return number


def excel_column_letters(number: int) -> str:
    """将列序号换算为 Excel 列名。"""
    letters = ""
    while number:
        number, remainder = divmod(number - 1, 26)
        letters = chr(ord("A") + remainder) + letters
    return letters


class _ExcelSelection:
    """封装工作表选择及公式结果转换。"""

    def __init__(self, settings: dict[str, Any]):
        """保存本轮转换的设置快照。"""
        self.settings = settings

    def select(self, sheets: list[dict[str, Any]], excel_options: dict[str, Any]) -> list[dict[str, Any]]:
        """按转换范围筛选并标注 Excel 工作表。"""
        settings = self.settings
        effective = effective_excel_settings(excel_options, settings)
        range_mode = effective["excel_conversion_range"]
        formula_mode = effective["excel_formula_mode"]
        selected_names = self._excel_selected_sheet_names(excel_options)
        selected_indexes = self._excel_selected_sheet_indexes(excel_options)
        selected = sheets
        if selected_names or selected_indexes:
            selected = [
                sheet
                for sheet in sheets
                if str(sheet.get("name", "")).strip().lower() in selected_names or int(sheet.get("index", 0) or 0) in selected_indexes
            ]
            if not selected:
                raise ValueError("指定的工作表不存在，请检查名称或序号")
            label = "指定工作表"
        elif range_mode == "当前工作表":
            selected = sheets[:1]
            label = "当前工作表"
        elif range_mode == "选区":
            selected = sheets[:1]
            label = "选区"
        else:
            selected = sheets
            label = "全部工作表"
        if range_mode == "选区":
            selected = [self._excel_selection_sheet(sheet, excel_options) for sheet in selected]
        return [self._excel_sheet_with_range(sheet, label, formula_mode, effective["excel_retain_comments"]) for sheet in selected]

    @staticmethod
    def _excel_formula_mode(options: dict[str, Any], settings: dict[str, Any]) -> str:
        """确定保留 Excel 公式还是计算结果。"""
        if "retainFormulas" in options:
            return "保留公式" if bool(options.get("retainFormulas")) else "仅保留计算结果"
        mode = str(options.get("formulaMode") or options.get("excelFormulaMode") or settings.get("excelFormulaMode", "保留公式") or "保留公式")
        return mode if mode in {"保留公式", "仅保留计算结果"} else "保留公式"

    @staticmethod
    def _excel_selected_sheet_names(options: dict[str, Any]) -> set[str]:
        """从任务选项解析选中的工作表名称。"""
        raw = options.get("sheetNames") or options.get("selectedSheets") or []
        if isinstance(raw, str):
            values = re.split(r"[,，;；\n]+", raw)
        elif isinstance(raw, list):
            values = [str(item) for item in raw]
        else:
            values = []
        return {value.strip().lower() for value in values if value.strip()}

    @staticmethod
    def _excel_selected_sheet_indexes(options: dict[str, Any]) -> set[int]:
        """从任务选项解析选中的工作表序号。"""
        raw = options.get("sheetIndexes") or options.get("selectedSheetIndexes") or []
        if isinstance(raw, str):
            values = re.split(r"[,，;；\n]+", raw)
        elif isinstance(raw, list):
            values = raw
        else:
            values = []
        indexes: set[int] = set()
        for value in values:
            try:
                index = int(value)
            except (TypeError, ValueError):
                continue
            if index > 0:
                indexes.add(index)
        return indexes

    @staticmethod
    def _excel_selection_sheet(sheet: dict[str, Any], options: dict[str, Any]) -> dict[str, Any]:
        """返回限制在选中单元格或安全预览范围内的工作表副本。"""
        selected = dict(sheet)
        cells = list(selected.get("cells", []))
        # 图表缓存可能引用选区之外的数据，选区转换暂不附带整表图表。
        selected["charts"] = []
        selected["chart_count"] = 0
        cell_refs = options.get("cellRefs") or options.get("selectedCells") or []
        if isinstance(cell_refs, str):
            values = [value for value in re.split(r"[,，;；\s]+", cell_refs) if value]
        elif isinstance(cell_refs, list):
            values = [str(value) for value in cell_refs if str(value).strip()]
        else:
            values = []
        refs = excel_cell_refs(values)
        if refs:
            selected["cells"] = [cell for cell in cells if str(cell.get("ref", "")).upper() in refs]
            # 仅完整选择的合并区域保留结构，避免带入未选中的行列。
            selected["merge_ranges"] = [region for region in selected.get("merge_ranges", [])
                                        if region.split(":")[0].upper() in refs and excel_cell_refs([region]).issubset(refs)]
            selected["merged_cells"] = len(selected["merge_ranges"])
            selected["comments"] = [item for item in selected.get("comments", []) if item["ref"].upper() in refs]
            selected["comment_count"] = len(selected["comments"])
            selected["conversion_range"] = f"选区 {'、'.join(value.strip().upper() for value in values)}"
            return selected
        try:
            limit = max(1, min(200, int(options.get("cellLimit", 12) or 12)))
        except (TypeError, ValueError):
            limit = 12
        selected["cells"] = cells[:limit]
        selected_refs = {cell["ref"].upper() for cell in selected["cells"]}
        selected["comments"] = [item for item in selected.get("comments", []) if item["ref"].upper() in selected_refs]
        selected["comment_count"] = len(selected["comments"])
        selected["merge_ranges"] = []
        selected["merged_cells"] = 0
        selected["conversion_range"] = f"选区 前 {min(limit, len(cells))} 个单元格"
        return selected

    @staticmethod
    def _excel_sheet_with_range(sheet: dict[str, Any], label: str, formula_mode: str, retain_comments: bool = False) -> dict[str, Any]:
        """为工作表附加转换范围和公式模式。"""
        selected = dict(sheet)
        selected["conversion_range"] = selected.get("conversion_range") or label
        selected["formula_count"] = sum(bool(cell.get("formula")) for cell in selected.get("cells", []))
        selected["formula_mode"] = formula_mode
        if not retain_comments:
            selected["comments"] = []
        if formula_mode == "仅保留计算结果":
            selected["cells"] = [_ExcelSelection._excel_cell_result_only(cell) for cell in selected.get("cells", [])]
        return selected

    @staticmethod
    def _excel_cell_result_only(cell: dict[str, Any]) -> dict[str, Any]:
        """返回以已存结果替代公式的单元格副本。"""
        converted = dict(cell)
        if converted.get("formula"):
            converted["value"] = converted.get("result") if converted.get("result") is not None and converted.get("result") != "" else "计算结果未保存"
        return converted

