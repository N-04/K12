"""Export helpers for formula, image, macro, and failure report artifacts.

The PRD asks for downloadable reports rather than opaque task logs. This module
builds deterministic CSV, JSON, TeX, MathML, and XLSX-compatible packages from
stored report data without reaching back into local source paths.
"""

from __future__ import annotations

import html
import json
import zipfile
from io import BytesIO
from pathlib import Path
from typing import Any


def build_formula_zip(report: dict[str, Any], annotations: list[dict[str, Any]] | None = None) -> bytes:
    """Package formulas as manifest data plus TeX and MathML files."""
    formulas = list(report.get("analysis", {}).get("formulas", []))
    files_by_id = {file.get("id"): file.get("file_name", "") for file in report.get("files", [])}
    annotations_by_formula = {item.get("formula_id"): item for item in annotations or []}
    buffer = BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("manifest.csv", _formula_manifest(formulas, files_by_id, annotations_by_formula))
        archive.writestr("formulas.json", _formula_json(report, formulas, files_by_id, annotations_by_formula))
        archive.writestr("formulas.tex", _formula_tex_document(formulas, files_by_id, annotations_by_formula))
        archive.writestr("formulas.mml", _formula_mml_document(formulas, files_by_id, annotations_by_formula))
        for index, formula in enumerate(formulas, start=1):
            stem = _formula_file_stem(index, formula)
            archive.writestr(f"tex/{stem}.tex", _formula_tex_item(formula, files_by_id, annotations_by_formula))
            archive.writestr(f"mml/{stem}.mml", _formula_mathml(formula, annotations_by_formula))
    return buffer.getvalue()


def build_formula_xlsx(report: dict[str, Any], annotations: list[dict[str, Any]] | None = None) -> bytes:
    """Build a spreadsheet-compatible formula report from normalized items."""
    formulas = list(report.get("analysis", {}).get("formulas", []))
    files_by_id = {file.get("id"): file.get("file_name", "") for file in report.get("files", [])}
    annotations_by_formula = {item.get("formula_id"): item for item in annotations or []}
    rows: list[list[Any]] = [
        ["报告 ID", report.get("id", "")],
        ["任务 ID", report.get("task_id", "")],
        ["报告类型", report.get("report_type", "")],
        [],
        [
            "formula_id",
            "file_id",
            "file_name",
            "page_index",
            "position",
            "position_status",
            "position_issue",
            "source_type",
            "original_image_ref",
            "latex",
            "mathml",
            "mathtype_preview",
            "format_status",
            "confidence",
            "status",
            "annotation_note",
            "retry_recognition",
            "recognition_status",
            "recognition_next_step",
        ],
    ]
    for formula in formulas:
        annotation = annotations_by_formula.get(formula.get("id"), {})
        rows.append(
            [
                formula.get("id", ""),
                formula.get("file_id", ""),
                files_by_id.get(formula.get("file_id"), formula.get("file_id", "")),
                formula.get("page_index", ""),
                formula.get("position", ""),
                formula.get("position_status", ""),
                formula.get("position_issue", ""),
                formula.get("source_type", ""),
                formula.get("original_image_ref") or formula.get("original_image_path", ""),
                annotation.get("latex") or formula.get("latex", ""),
                annotation.get("mathml") or formula.get("mathml", ""),
                formula.get("mathtype_preview") or formula.get("mathtype_data", ""),
                formula.get("format_status", ""),
                formula.get("confidence", ""),
                annotation.get("status") or formula.get("status", ""),
                annotation.get("note", ""),
                annotation.get("retry_recognition", ""),
                annotation.get("recognition_status", ""),
                annotation.get("next_step", ""),
            ]
        )
    return _xlsx_bytes(rows, "Formula Results")


def build_image_manifest_xlsx(report: dict[str, Any], annotations: list[dict[str, Any]] | None = None) -> bytes:
    """Build a spreadsheet-compatible manifest for extracted small images."""
    images = list(report.get("analysis", {}).get("smallImages", []))
    files_by_id = {file.get("id"): file.get("file_name", "") for file in report.get("files", [])}
    annotations_by_image = {item.get("image_id"): item for item in annotations or []}
    rows: list[list[Any]] = [
        ["报告 ID", report.get("id", "")],
        ["任务 ID", report.get("task_id", "")],
        ["报告类型", report.get("report_type", "")],
        [],
        [
            "image_id",
            "file_id",
            "file_name",
            "source_name",
            "location",
            "width",
            "height",
            "area",
            "image_format",
            "suspected_type",
            "confirmed_type",
            "confidence",
            "header_footer",
            "watermark",
            "transparent",
            "duplicate",
            "duplicate_check_status",
            "duplicate_fallback",
            "export_status",
            "export_message",
            "reexported_at",
            "export_format",
            "annotation",
            "note",
            "replacement_file",
            "asset_url",
        ],
    ]
    for image in images:
        annotation = annotations_by_image.get(image.get("id"), {})
        rows.append(
            [
                image.get("id", ""),
                image.get("file_id", ""),
                files_by_id.get(image.get("file_id"), image.get("file_id", "")),
                image.get("source_name", ""),
                image.get("location", ""),
                image.get("width", ""),
                image.get("height", ""),
                image.get("area", ""),
                image.get("image_type", ""),
                _small_image_kind(image),
                annotation.get("label", ""),
                image.get("confidence", ""),
                image.get("is_header_footer", ""),
                image.get("is_watermark", ""),
                image.get("is_transparent", ""),
                image.get("is_duplicate", ""),
                image.get("duplicate_check_status", ""),
                image.get("duplicate_fallback", ""),
                image.get("export_status", ""),
                image.get("export_message", ""),
                image.get("reexported_at", ""),
                image.get("export_format", ""),
                annotation.get("status", ""),
                annotation.get("note", ""),
                annotation.get("replacement_file_name", ""),
                image.get("asset_url", ""),
            ]
        )
    return _xlsx_bytes(rows, "Image Manifest")


def build_omml_failure_csv(report: dict[str, Any], annotations: list[dict[str, Any]] | None = None, expose_paths: bool = False) -> str:
    """Build the OMML failure CSV while honoring the path exposure policy."""
    files = report.get("files", [])
    files_by_id = {file.get("id"): file for file in files}
    annotations_by_dependency = {item.get("dependency_id"): item for item in annotations or []}
    rows: list[dict[str, Any]] = []
    covered_file_ids: set[Any] = set()
    for dependency in report.get("analysis", {}).get("ommlDependencies", []):
        annotation = annotations_by_dependency.get(dependency.get("id"), {})
        if not _omml_failure_row_needed(dependency, annotation):
            continue
        file = files_by_id.get(dependency.get("file_id"), {})
        covered_file_ids.add(dependency.get("file_id"))
        rows.append(_omml_failure_row(report, dependency, file, annotation, expose_paths))
    for file in files:
        if not file.get("missing_omml_dependency") or file.get("id") in covered_file_ids:
            continue
        rows.append(
            _omml_failure_row(
                report,
                {
                    "file_id": file.get("id", ""),
                    "found_status": "未找到",
                    "copy_status": "失败",
                    "error_message": "文件含 OMML 公式，但未检测到可用 OMML 依赖文件",
                },
                file,
                {},
                expose_paths,
            )
        )
    return _dict_csv(
        rows,
        [
            "report_id",
            "dependency_id",
            "file_id",
            "file_name",
            "document_path",
            "omml_file_name",
            "omml_source_path",
            "omml_target_path",
            "found_status",
            "copy_status",
            "annotation_status",
            "keep_omml",
            "retry_conversion",
            "manual_omml_path",
            "message",
            "recommendation",
            "note",
        ],
    )


def build_macro_failure_csv(report: dict[str, Any], expose_paths: bool = False) -> str:
    """Build the macro failure CSV without exposing backup paths by default."""
    files_by_id = {file.get("id"): file for file in report.get("files", [])}
    rows = [
        _macro_failure_row(report, macro, files_by_id.get(macro.get("file_id"), {}), expose_paths)
        for macro in report.get("analysis", {}).get("macros", [])
        if _macro_failure_row_needed(macro)
    ]
    return _dict_csv(
        rows,
        [
            "report_id",
            "macro_id",
            "file_id",
            "file_name",
            "macro_name",
            "macro_source",
            "execute_order",
            "execute_timing",
            "execute_status",
            "failure_strategy",
            "failure_action",
            "requires_user_decision",
            "backup_path",
            "message",
            "recommendation",
        ],
    )


def _macro_failure_row_needed(macro: dict[str, Any]) -> bool:
    """Return whether one macro should appear in the macro failure CSV."""
    return macro.get("execute_status") in {"失败", "未授权", "已禁用"}


def _macro_failure_row(report: dict[str, Any], macro: dict[str, Any], file: dict[str, Any], expose_paths: bool) -> dict[str, Any]:
    """Build one macro failure export row with optional local path redaction."""
    policy = macro.get("failure_policy") or {}
    return {
        "report_id": report.get("id", ""),
        "macro_id": macro.get("id", ""),
        "file_id": macro.get("file_id") or file.get("id", ""),
        "file_name": file.get("file_name", "") or macro.get("file_name", ""),
        "macro_name": macro.get("macro_name", ""),
        "macro_source": macro.get("macro_source", ""),
        "execute_order": macro.get("execute_order", ""),
        "execute_timing": macro.get("execute_timing", ""),
        "execute_status": macro.get("execute_status", ""),
        "failure_strategy": macro.get("failure_strategy", ""),
        "failure_action": policy.get("action", ""),
        "requires_user_decision": policy.get("requires_user_decision", ""),
        "backup_path": _export_path(macro.get("backup_path", ""), expose_paths),
        "message": macro.get("error_message", ""),
        "recommendation": _macro_failure_recommendation(macro),
    }


def _macro_failure_recommendation(macro: dict[str, Any]) -> str:
    """Return the next action for a failed, blocked, or disabled macro."""
    status = str(macro.get("execute_status") or "")
    if status == "未授权":
        return "检查宏来源授权、白名单和当前用户宏执行权限后重试"
    if status == "已禁用":
        return "在设置中启用宏执行队列并重新创建宏任务"
    policy = macro.get("failure_policy") or {}
    action = str(policy.get("action") or "")
    if action == "stop_sequence":
        return "检查失败宏后重跑该宏，必要时从备份恢复原文档"
    if action == "ask_user":
        return "根据本地客户端提示选择停止、跳过或继续后重新执行"
    return "检查宏日志、备份和失败策略后重试该宏"


def _omml_failure_row_needed(dependency: dict[str, Any], annotation: dict[str, Any]) -> bool:
    """Return whether one OMML dependency or annotation needs export attention."""
    return (
        dependency.get("found_status") == "未找到"
        or dependency.get("copy_status") == "失败"
        or annotation.get("status") in {"转换失败", "重新转换"}
    )


def _omml_failure_row(
    report: dict[str, Any],
    dependency: dict[str, Any],
    file: dict[str, Any],
    annotation: dict[str, Any],
    expose_paths: bool,
) -> dict[str, Any]:
    """Build one OMML dependency failure row with annotation context."""
    message = dependency.get("error_message") or annotation.get("note") or f"检索：{dependency.get('found_status', '')}，复制：{dependency.get('copy_status', '')}"
    recommendation = _omml_failure_recommendation(dependency, annotation)
    return {
        "report_id": report.get("id", ""),
        "dependency_id": dependency.get("id", ""),
        "file_id": dependency.get("file_id") or file.get("id", ""),
        "file_name": file.get("file_name", "") or dependency.get("file_name", ""),
        "document_path": _export_path(dependency.get("document_path", ""), expose_paths),
        "omml_file_name": dependency.get("omml_file_name", ""),
        "omml_source_path": _export_path(dependency.get("omml_source_path", ""), expose_paths),
        "omml_target_path": _export_path(dependency.get("omml_target_path", ""), expose_paths),
        "found_status": dependency.get("found_status", ""),
        "copy_status": dependency.get("copy_status", ""),
        "annotation_status": annotation.get("status", ""),
        "keep_omml": annotation.get("keep_omml", ""),
        "retry_conversion": annotation.get("retry_conversion", ""),
        "manual_omml_path": _export_path(annotation.get("manual_omml_path", ""), expose_paths),
        "message": message,
        "recommendation": recommendation,
        "note": annotation.get("note", ""),
    }


def _omml_failure_recommendation(dependency: dict[str, Any], annotation: dict[str, Any]) -> str:
    """Return the next action for OMML lookup, copy, or conversion failures."""
    if annotation.get("retry_conversion"):
        return "重新执行 OMML 转 MathType，并保留原 OMML 兜底"
    if annotation.get("manual_omml_path"):
        return "使用手动指定的 OMML 依赖文件后重新转换"
    if dependency.get("found_status") == "未找到":
        return "配置 OMML 检索路径或手动选择 OMML2MML.XSL"
    if dependency.get("copy_status") == "失败":
        return "检查目标目录写入权限或调整 OMML 复制策略"
    return "复核 OMML 转换结果，必要时重新转换或保留原公式"


def _dict_csv(rows: list[dict[str, Any]], fieldnames: list[str]) -> str:
    """Render dictionaries to CSV using an explicit field order."""
    lines = [",".join(fieldnames)]
    for row in rows:
        lines.append(",".join(_csv_cell(row.get(name, "")) for name in fieldnames))
    return "\n".join(lines) + "\n"


def _export_path(value: object, expose_paths: bool) -> str:
    """Hide local paths unless the caller explicitly allows path exposure."""
    text = str(value or "")
    if not text or expose_paths:
        return text
    name = Path(text).name
    return f"本地路径已隐藏/{name}" if name else "本地路径已隐藏"


def _formula_json(report: dict[str, Any], formulas: list[dict[str, Any]], files_by_id: dict[str, str], annotations: dict[str, dict[str, Any]]) -> str:
    """Render formula export metadata as path-free JSON."""
    items: list[dict[str, Any]] = []
    for formula in formulas:
        annotation = annotations.get(formula.get("id"), {})
        items.append(
            {
                "formula_id": formula.get("id", ""),
                "file_id": formula.get("file_id", ""),
                "file_name": files_by_id.get(formula.get("file_id"), formula.get("file_id", "")),
                "page_index": formula.get("page_index", ""),
                "position": formula.get("position", ""),
                "position_status": formula.get("position_status", ""),
                "position_issue": formula.get("position_issue", ""),
                "fallback_position": formula.get("fallback_position", ""),
                "source_type": formula.get("source_type", ""),
                "original_image_ref": formula.get("original_image_ref") or formula.get("original_image_path", ""),
                "confidence": formula.get("confidence", ""),
                "status": annotation.get("status") or formula.get("status", ""),
                "format_status": formula.get("format_status", ""),
                "latex": annotation.get("latex") or formula.get("latex", ""),
                "mathml": annotation.get("mathml") or formula.get("mathml", ""),
                "mathtype_preview": formula.get("mathtype_preview") or formula.get("mathtype_data", ""),
                "annotation_note": annotation.get("note", ""),
                "retry_recognition": annotation.get("retry_recognition", False),
                "recognition_status": annotation.get("recognition_status", ""),
                "recognition_next_step": annotation.get("next_step", ""),
                "recognition_request": annotation.get("recognition_request", {}),
            }
        )
    return json.dumps({"report_id": report.get("id", ""), "formula_count": len(items), "formulas": items}, ensure_ascii=False, indent=2) + "\n"


def _formula_manifest(formulas: list[dict[str, Any]], files_by_id: dict[str, str], annotations: dict[str, dict[str, Any]]) -> str:
    """Render formula export metadata as a manifest CSV."""
    rows = ["formula_id,file_name,page,position,position_status,position_issue,source_type,confidence,status,format_status,latex,mathml,retry_recognition,recognition_status,recognition_next_step"]
    for formula in formulas:
        annotation = annotations.get(formula.get("id"), {})
        rows.append(
            ",".join(
                _csv_cell(value)
                for value in [
                    formula.get("id", ""),
                    files_by_id.get(formula.get("file_id"), formula.get("file_id", "")),
                    formula.get("page_index", ""),
                    formula.get("position", ""),
                    formula.get("position_status", ""),
                    formula.get("position_issue", ""),
                    formula.get("source_type", ""),
                    formula.get("confidence", ""),
                    annotation.get("status") or formula.get("status", ""),
                    formula.get("format_status", ""),
                    annotation.get("latex") or formula.get("latex", ""),
                    annotation.get("mathml") or formula.get("mathml", ""),
                    annotation.get("retry_recognition", ""),
                    annotation.get("recognition_status", ""),
                    annotation.get("next_step", ""),
                ]
            )
        )
    return "\n".join(rows) + "\n"


def _formula_tex_document(formulas: list[dict[str, Any]], files_by_id: dict[str, str], annotations: dict[str, dict[str, Any]]) -> str:
    """Combine all formulas into one TeX document for report download."""
    lines = ["% K12 formula export", ""]
    for formula in formulas:
        lines.append(_formula_tex_item(formula, files_by_id, annotations).rstrip())
        lines.append("")
    return "\n".join(lines).rstrip() + "\n"


def _formula_tex_item(formula: dict[str, Any], files_by_id: dict[str, str], annotations: dict[str, dict[str, Any]]) -> str:
    """Render one formula as a TeX display block with audit metadata."""
    annotation = annotations.get(formula.get("id"), {})
    source_file = files_by_id.get(formula.get("file_id"), formula.get("file_id", ""))
    status = annotation.get("status") or formula.get("status", "")
    latex = annotation.get("latex") or formula.get("latex") or ""
    header = (
        f"% id={formula.get('id', '')} file={source_file} "
        f"position={formula.get('position', '')} source={formula.get('source_type', '')} "
        f"confidence={formula.get('confidence', '')} status={status}"
    )
    return f"{header}\n\\[\n{latex}\n\\]\n"


def _formula_mml_document(formulas: list[dict[str, Any]], files_by_id: dict[str, str], annotations: dict[str, dict[str, Any]]) -> str:
    """Combine all formulas into one XML-style MathML export document."""
    rows = ["<?xml version=\"1.0\" encoding=\"UTF-8\"?>", "<formulas>"]
    for formula in formulas:
        annotation = annotations.get(formula.get("id"), {})
        source_file = files_by_id.get(formula.get("file_id"), formula.get("file_id", ""))
        rows.append(
            f'  <formula id="{html.escape(str(formula.get("id", "")))}" '
            f'file="{html.escape(str(source_file))}" '
            f'page="{html.escape(str(formula.get("page_index", "")))}" '
            f'source="{html.escape(str(formula.get("source_type", "")))}" '
            f'confidence="{html.escape(str(formula.get("confidence", "")))}" '
            f'status="{html.escape(str(annotation.get("status") or formula.get("status", "")))}">'
        )
        rows.append(f"    <position>{html.escape(str(formula.get('position', '')))}</position>")
        rows.append(f"    <latex>{html.escape(str(annotation.get('latex') or formula.get('latex') or ''))}</latex>")
        rows.append(f"    {_formula_mathml(formula, annotations).strip()}")
        rows.append("  </formula>")
    rows.append("</formulas>")
    return "\n".join(rows) + "\n"


def _formula_mathml(formula: dict[str, Any], annotations: dict[str, dict[str, Any]]) -> str:
    """Return stored MathML or a safe MathML text fallback from LaTeX."""
    annotation = annotations.get(formula.get("id"), {})
    mathml = annotation.get("mathml") or formula.get("mathml") or ""
    if mathml.strip():
        return mathml if mathml.endswith("\n") else mathml + "\n"
    latex = annotation.get("latex") or formula.get("latex") or ""
    return f"<math><mtext>{html.escape(str(latex))}</mtext></math>\n"


def _formula_file_stem(index: int, formula: dict[str, Any]) -> str:
    """Create a stable safe filename stem for one exported formula."""
    raw = str(formula.get("id") or f"formula-{index}")
    safe = "".join(char if char.isalnum() or char in {"-", "_"} else "-" for char in raw)
    return f"{index:03d}-{safe[:80]}"


def _small_image_kind(image: dict[str, Any]) -> str:
    """Classify a small image for manifest and workbook exports."""
    if image.get("is_formula_like"):
        return "公式"
    if image.get("is_qrcode_like"):
        return "二维码"
    if image.get("is_stamp_like"):
        return "印章"
    if image.get("is_signature_like"):
        return "签名"
    return "图标"


def _xlsx_bytes(rows: list[list[Any]], sheet_name: str) -> bytes:
    """Package rows into a minimal XLSX workbook."""
    buffer = BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", _xlsx_content_types())
        archive.writestr("_rels/.rels", _xlsx_package_rels())
        archive.writestr("xl/workbook.xml", _xlsx_workbook(sheet_name))
        archive.writestr("xl/_rels/workbook.xml.rels", _xlsx_workbook_rels())
        archive.writestr("xl/worksheets/sheet1.xml", _xlsx_sheet(rows))
    return buffer.getvalue()


def _xlsx_sheet(rows: list[list[Any]]) -> str:
    """Build worksheet XML with inline string cells."""
    row_xml = []
    for row_index, row in enumerate(rows, start=1):
        cells = []
        for column_index, value in enumerate(row, start=1):
            ref = f"{_column_name(column_index)}{row_index}"
            cells.append(f'<c r="{ref}" t="inlineStr"><is><t>{html.escape(str(value))}</t></is></c>')
        row_xml.append(f'<row r="{row_index}">{"".join(cells)}</row>')
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>{"".join(row_xml)}</sheetData></worksheet>"""


def _column_name(index: int) -> str:
    """Convert a 1-based column index to an Excel column name."""
    name = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        name = chr(65 + remainder) + name
    return name or "A"


def _xlsx_content_types() -> str:
    """Return the content-types XML for the generated XLSX package."""
    return """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>"""


def _xlsx_package_rels() -> str:
    """Return the package relationships XML for the workbook entry."""
    return """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"""


def _xlsx_workbook(sheet_name: str) -> str:
    """Return workbook XML with a safe single-sheet name."""
    safe_name = html.escape(sheet_name[:31] or "Sheet1", quote=True)
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="{safe_name}" sheetId="1" r:id="rId1"/></sheets></workbook>"""


def _xlsx_workbook_rels() -> str:
    """Return workbook relationships XML for the single worksheet."""
    return """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"""


def _csv_cell(value: object) -> str:
    """Quote one CSV cell for deterministic export output."""
    text = str(value).replace('"', '""')
    return f'"{text}"'
