"""使用标准库提取文档结构并提供轻量转换。支持本地优先的基础功能，为预览、报告和最小 OOXML 产物提取结构；完整 Office 渲染、MathType、OMML 写回和宏处理交由本地客户端执行。"""

from __future__ import annotations

import base64
import html
import posixpath
import re
import unicodedata
import zipfile
import zlib
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Any

from .text_styles import PARAGRAPH_ALIGNMENTS, UNDERLINES, WordTextStyles
from .charts import OFFICE_REL, REL_NS, build_chart, chart_frame, chart_has_values
from .ooxml import parse_compatible_xml
from .media import IMAGE_TYPES, picture_xml, word_images, word_picture_xml, word_story_images
from .excel import excel_cell_refs, excel_column_letters, excel_column_number


PDF_TEXT_SOURCE_LIMIT_BYTES = 50 * 1024 * 1024
PDF_TEXT_STREAM_LIMIT_BYTES = 10 * 1024 * 1024
PDF_TEXT_TOTAL_STREAM_LIMIT_BYTES = 25 * 1024 * 1024
PDF_TEXT_STREAM_COUNT_LIMIT = 1000


def extract_docx_blocks(path: Path, retain_images: bool = True) -> list[dict[str, Any]]:
    """提取 DOCX 可见段落，用于预览和 Word 转 PPT。"""
    with zipfile.ZipFile(path) as archive:
        document_xml = archive.read("word/document.xml").decode("utf-8", errors="ignore")
        styles = WordTextStyles.from_archive(archive)
    # 兼容历史无命名空间最小文件；真实 OOXML 按 XML 节点读取。
    try:
        root = parse_compatible_xml(document_xml)
    except ET.ParseError:
        root = None
    blocks: list[dict[str, Any]] = []
    if root is not None:
        tables = root.findall(".//{*}tbl")
        table_nodes = {id(node) for table in tables for node in table.iter() if node is not table}
        paragraph_index = 0
        for node in root.iter():
            tag = node.tag.rsplit("}", 1)[-1]
            if tag == "p":
                paragraph_index += 1
            if id(node) in table_nodes:
                continue
            if retain_images and tag in {"drawing", "pict"}:
                blocks.extend({"text": "", "level": 0, "image": image} for image in word_images(path, node))
                continue
            if tag == "tbl":
                rows, merges, alignments = _docx_table_grid(node)
                if rows:
                    blocks.append({"text": "\n".join("\t".join(row) for row in rows), "table": rows, "table_merges": merges, "cell_alignments": alignments, "row_heights": _docx_row_heights(node), "column_widths": [int(column.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}w", "0")) for column in node.findall("{*}tblGrid/{*}gridCol")], "level": 0})
                    if retain_images:
                        blocks.extend({"text": "", "level": 0, "image": image} for image in word_images(path, node))
            elif tag == "p":
                text = _docx_visible_text(node)
                if text:
                    runs = []
                    for run in node.findall(".//{*}r"):
                        style = styles.run(node, run)
                        runs.append({"text": _docx_visible_text(run), **style})
                    blocks.append({"text": text, "runs": runs, "level": _heading_level(ET.tostring(node, encoding="unicode"), paragraph_index)})
        if retain_images:
            blocks.extend(word_story_images(path, root))
    else:
        for index, paragraph in enumerate(re.findall(r"<w:p\b[\s\S]*?</w:p>", document_xml), 1):
            text = _extract_text(paragraph, "w:t") or _strip_xml(paragraph)
            if text:
                blocks.append({"text": text, "level": _heading_level(paragraph, index)})
    return blocks or [{"text": path.stem, "level": 1}]


def _docx_row_heights(table: ET.Element) -> list[int]:
    """读取 Word 显式行高；缺失行高用零表示自动布局。"""
    heights = []
    for row in table.findall("{*}tr"):
        height = row.find("{*}trPr/{*}trHeight")
        heights.append(int(height.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val", "0")) if height is not None else 0)
    return heights


def _docx_table_grid(table: ET.Element) -> tuple[list[list[str]], list[list[int]], list[list[list[str]]]]:
    """展开 Word 合并网格，记录起点及行列跨度。"""
    rows, merges, alignments, active = [], [], [], {}
    for row_index, row in enumerate(table.findall("{*}tr")):
        before = row.find("{*}trPr/{*}gridBefore")
        values = [""] * (int(before.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val", "0")) if before is not None else 0)
        row_alignments = [[] for _ in values]
        next_active = {}
        for cell in row.findall("{*}tc"):
            column = len(values)
            span_node = cell.find("{*}tcPr/{*}gridSpan")
            span = int(span_node.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val", "1")) if span_node is not None else 1
            vertical = cell.find("{*}tcPr/{*}vMerge")
            continuation = vertical is not None and vertical.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val") != "restart"
            if continuation and column in active and active[column][3] == span:
                merge = active[column]
                merge[2] += 1
                next_active[column] = merge
                text = ""
            else:
                text = _docx_visible_text(cell)
                if span > 1 or vertical is not None:
                    merge = [row_index, column, 1, span]
                    merges.append(merge)
                    if vertical is not None:
                        next_active[column] = merge
            values.extend([text, *([""] * (span - 1))])
            row_alignments.extend([_cell_line_alignments(cell, True), *([[]] * (span - 1))])
        rows.append(values)
        alignments.append(row_alignments)
        active = next_active
    return rows, merges, alignments


def _cell_line_alignments(cell: ET.Element, word: bool = False) -> list[str]:
    """把每段对齐扩展到文字行，供两种格式共同传递。"""
    result = []
    for paragraph in cell.findall(".//{*}p"):
        properties = paragraph.find("{*}pPr")
        alignment = ""
        if properties is not None:
            if word:
                node = properties.find("{*}jc")
                alignment = node.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val", "") if node is not None else ""
            else:
                alignment = {value: key for key, value in PARAGRAPH_ALIGNMENTS.items()}.get(properties.get("algn"), "")
        text = _docx_visible_text(paragraph) if word else "\n".join(_pptx_texts(ET.tostring(paragraph, encoding="unicode"), preserve_whitespace=True))
        result.extend([alignment] * (text.count("\n") + 1))
    return result


def _docx_visible_text(element: ET.Element) -> str:
    """提取文字及换行、制表符，供正文和表格单元格共用。"""
    pieces = []
    paragraph_seen = False
    for node in element.iter():
        tag = node.tag.rsplit("}", 1)[-1]
        if tag == "p":
            if paragraph_seen:
                pieces.append("\n")
            paragraph_seen = True
        elif tag == "t":
            pieces.append(node.text or "")
        elif tag in {"br", "cr", "tab"}:
            pieces.append("\t" if tag == "tab" else "\n")
    return "".join(pieces)


def count_xml_elements(xml: str, name: str) -> int:
    """按完整节点名计数，兼容不同前缀及历史最小 XML。"""
    return len(re.findall(r"<(?:[\w.-]+:)?" + re.escape(name) + r"(?=[\s/>])", xml))


def extract_docx_object_summary(path: Path) -> dict[str, Any]:
    """统计 DOCX 中需要保留或交接桌面端的对象。"""
    with zipfile.ZipFile(path) as archive:
        names = archive.namelist()
        try:
            document_xml = archive.read("word/document.xml").decode("utf-8", errors="ignore")
        except KeyError:
            document_xml = ""
    image_parts = sorted(name for name in names if name.startswith("word/media/"))
    embedded_parts = sorted(name for name in names if name.startswith("word/embeddings/"))
    omml_formulas = count_xml_elements(document_xml, "oMath")
    mathtype_objects = len(embedded_parts)
    if mathtype_objects == 0 and ("Equation Native" in document_xml or "MathType" in document_xml):
        mathtype_objects = 1
    return {
        "paragraphs": count_xml_elements(document_xml, "p"),
        "headings": len(re.findall(r"<w:pStyle[^>]+w:val=\"Heading", document_xml)),
        "images": len(image_parts),
        "tables": count_xml_elements(document_xml, "tbl"),
        "omml_formulas": omml_formulas,
        "mathtype_objects": mathtype_objects,
        "embedded_objects": len(embedded_parts),
        "formulas": omml_formulas + mathtype_objects,
        "image_parts": image_parts[:12],
        "embedded_parts": embedded_parts[:12],
    }


def extract_pptx_slides(path: Path, retain_images: bool = True) -> list[dict[str, Any]]:
    """提取 PPTX 的文本、备注与对象数量，用于转换报告。"""
    with zipfile.ZipFile(path) as archive:
        names = _pptx_slide_order(archive)
        slides: list[dict[str, Any]] = []
        for index, name in enumerate(names, start=1):
            xml = archive.read(name).decode("utf-8", errors="ignore")
            try:
                root = parse_compatible_xml(xml)
                if any(node.tag.endswith("}AlternateContent") for node in root.iter()):
                    xml = ET.tostring(root, encoding="unicode")
            except ET.ParseError:
                root = None
            cleaned = _pptx_texts(xml, exclude_tables=True)
            styled = _pptx_styled_paragraphs(xml, _pptx_related_root(archive, name, "slideLayout"), _pptx_related_root(archive, name, "slideLayout", "slideMaster"), _pptx_related_root(archive, name, "slideLayout", ("slideMaster", "theme")))
            title = cleaned[0] if cleaned else f"幻灯片 {index}"
            body = cleaned[1:] if len(cleaned) > 1 else []
            related_parts = _pptx_related_parts(archive, name)
            notes = _pptx_notes(archive, related_parts)
            formula_count = _pptx_formula_count(xml)
            images = word_images(path, ET.fromstring(xml), name) if retain_images and "blip" in xml else []
            image_count = len({part for part in related_parts if "/media/" in f"/{part.lower()}"})
            if root is not None and root.find(".//{http://schemas.microsoft.com/office/drawing/2016/SVG/main}svgBlip") is not None:
                # SVG 原图与备用位图按绘图对象计数，避免共享关系重复扣减。
                image_count = len(root.findall(".//{http://schemas.openxmlformats.org/drawingml/2006/main}blip"))
            slides.append(
                {
                    "title": title,
                    "body": body,
                    "title_runs": styled[0] if styled else [],
                    "body_runs": styled[1:] if len(styled) > 1 else [],
                    "notes": notes,
                    "index": index,
                    "text_box_count": len(root.findall(".//{http://schemas.openxmlformats.org/presentationml/2006/main}txBody")) if root is not None else len(re.findall(r"<p:txBody\b", xml)),
                    "shape_count": len(root.findall(".//{http://schemas.openxmlformats.org/presentationml/2006/main}sp")) if root is not None else len(re.findall(r"<p:sp\b", xml)),
                    "table_count": len(re.findall(r"<(?:\w+:)?tbl\b", xml)),
                    "tables": _pptx_tables(xml),
                    "table_merges": _pptx_table_merges(xml),
                    "table_widths": _pptx_table_widths(xml),
                    "table_heights": _pptx_table_heights(xml),
                    "table_alignments": _pptx_table_alignments(xml),
                    "images": images,
                    "image_count": image_count,
                    "chart_count": len({part for part in related_parts if "/charts/" in f"/{part.lower()}"}),
                    "formula_count": formula_count,
                }
            )
    return slides or [{"title": path.stem, "body": [], "index": 1}]


def _pptx_tables(xml: str) -> list[list[list[str]]]:
    """提取 PPT 表格行列文本，供 Word 输出可编辑表格。"""
    try:
        root = ET.fromstring(xml)
    except ET.ParseError:
        return []
    return [[["\n".join(_pptx_texts(ET.tostring(cell, encoding="unicode"), preserve_whitespace=True)) for cell in row.findall("{*}tc")]
             for row in table.findall("{*}tr")] for table in root.findall(".//{*}tbl")]


def _pptx_table_alignments(xml: str) -> list[list[list[list[str]]]]:
    """读取 PPT 各单元格文字行的显式对齐方式。"""
    try:
        root = ET.fromstring(xml)
    except ET.ParseError:
        return []
    return [[[_cell_line_alignments(cell) for cell in row.findall("{*}tc")] for row in table.findall("{*}tr")]
            for table in root.findall(".//{*}tbl")]


def _pptx_table_heights(xml: str) -> list[list[int]]:
    """读取 PPT 行高并将 EMU 转换为 Word 的 twips。"""
    try:
        root = ET.fromstring(xml)
    except ET.ParseError:
        return []
    return [[max(1, (int(row.get("h", "0")) + 317) // 635) for row in table.findall("{*}tr")]
            for table in root.findall(".//{*}tbl")]


def _pptx_table_widths(xml: str) -> list[list[int]]:
    """读取 PPT 表格网格列宽，供 Word 保留相对比例。"""
    try:
        root = ET.fromstring(xml)
    except ET.ParseError:
        return []
    return [[int(column.get("w", "0")) for column in table.findall("{*}tblGrid/{*}gridCol")]
            for table in root.findall(".//{*}tbl")]


def _pptx_table_merges(xml: str) -> list[list[list[int]]]:
    """读取原生 PPT 表格合并起点及行列跨度。"""
    try:
        root = ET.fromstring(xml)
    except ET.ParseError:
        return []
    tables = []
    for table in root.findall(".//{*}tbl"):
        merges = []
        for row_index, row in enumerate(table.findall("{*}tr")):
            for column, cell in enumerate(row.findall("{*}tc")):
                if cell.get("hMerge") in {"1", "true"} or cell.get("vMerge") in {"1", "true"}:
                    continue
                height, width = int(cell.get("rowSpan", "1")), int(cell.get("gridSpan", "1"))
                if height > 1 or width > 1:
                    merges.append([row_index, column, height, width])
        tables.append(merges)
    return tables


def _slide_docx_tables(slide: dict[str, Any]) -> list[dict[str, Any]]:
    """将 PPT 合并网格映射为 Word 跨列及纵向续接单元格。"""
    tables = []
    merge_groups = slide.get("table_merges", [])
    for index, source in enumerate(slide.get("tables", [])):
        rows = [list(row) for row in source]
        for top, left, height, width in merge_groups[index] if index < len(merge_groups) else []:
            if top < 0 or left < 0 or height < 1 or width < 1 or top + height > len(rows) or any(left + width > len(row) for row in rows[top:top + height]):
                raise ValueError("PPT 表格合并范围超出网格")
            text = str(rows[top][left])
            for row in range(top, top + height):
                for column in range(left, left + width):
                    item = {"text": text if row == top and column == left else "", "span": width, "skip": column != left}
                    if height > 1:
                        item["vertical"] = "restart" if row == top else "continue"
                    rows[row][column] = item
        width_groups = slide.get("table_widths", [])
        tables.append({"table": rows, "column_widths": width_groups[index] if index < len(width_groups) else [],
                       "cell_alignments": slide.get("table_alignments", [])[index] if index < len(slide.get("table_alignments", [])) else [],
                       "row_heights": slide.get("table_heights", [])[index] if index < len(slide.get("table_heights", [])) else []})
    return tables


def _pptx_slide_order(archive: zipfile.ZipFile) -> list[str]:
    """按演示文稿关系读取页面顺序，防止重排后正文和备注错位。"""
    names = archive.namelist()
    if "ppt/presentation.xml" in names:
        try:
            root = ET.fromstring(archive.read("ppt/presentation.xml"))
            pages = root.findall(".//{*}sldId")
            relation_key = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
            if any(page.get(relation_key) for page in pages):
                relations = ET.fromstring(archive.read("ppt/_rels/presentation.xml.rels"))
                by_id = {item.get("Id"): item for item in relations}
                ordered = []
                for page in pages:
                    relation = by_id.get(page.get(relation_key))
                    if relation is None or relation.get("TargetMode") == "External":
                        raise ValueError("幻灯片关系缺失或引用外部资源")
                    if not relation.get("Type", "").endswith("/slide"):
                        raise ValueError("幻灯片关系类型无效")
                    path = _resolve_xlsx_target("ppt/presentation.xml", relation.get("Target", ""))
                    if not path.startswith("ppt/slides/") or path not in names:
                        raise ValueError("幻灯片引用不存在或超出页面目录")
                    ordered.append(path)
                return ordered
        except (KeyError, ET.ParseError) as exc:
            raise ValueError("演示文稿关系文件缺失或无法解析") from exc
    # 兼容未声明页面关系的历史最小文件。
    return sorted([name for name in names if name.startswith("ppt/slides/slide") and name.endswith(".xml")], key=_slide_number)


def extract_xlsx_sheets(path: Path) -> list[dict[str, Any]]:
    """提取 XLSX 工作表预览与公式数量。"""
    with zipfile.ZipFile(path) as archive:
        names = archive.namelist()
        shared_strings = _xlsx_shared_strings(archive)
        worksheets = _xlsx_worksheet_order(archive, names)
        sheets: list[dict[str, Any]] = []
        for index, (title, name) in enumerate(worksheets, start=1):
            xml = archive.read(name).decode("utf-8", errors="ignore")
            cells = _xlsx_cells(xml, shared_strings)
            formulas = sum(bool(cell.get("formula")) for cell in cells)
            merge_ranges = [node.get("ref", "") for node in ET.fromstring(xml).findall(".//{*}mergeCell")]
            merged = len(merge_ranges)
            objects = _xlsx_sheet_objects(archive, name, xml)
            sheets.append(
                {
                    "name": title,
                    "index": index,
                    "cells": cells,
                    "formula_count": formulas,
                    "merged_cells": merged,
                    "merge_ranges": merge_ranges,
                    **objects,
                }
            )
    _xlsx_resolve_chart_sources(sheets)
    return sheets or [
        {
            "name": path.stem,
            "index": 1,
            "cells": [],
            "formula_count": 0,
            "merged_cells": 0,
            "chart_count": 0,
            "drawing_count": 0,
            "image_count": 0,
            "comment_count": 0,
            "table_count": 0,
        }
    ]


def build_pptx_from_docx(
    blocks: list[dict[str, Any]],
    target: Path,
    max_chars: int = 320,
    auto_pagination: bool = True,
    generate_toc: bool = False,
    object_preservation: dict[str, Any] | None = None,
    retain_tables: bool = True,
    retain_images: bool = True,
) -> list[str]:
    """根据 DOCX 段落及对象保留信息生成轻量 PPTX。"""
    if max_chars <= 0:
        raise ValueError("每页最大字数必须大于零")
    slide_limit = max_chars if auto_pagination else None
    slides = _slides_from_blocks([block for block in blocks if (retain_tables or "table" not in block) and (retain_images or "image" not in block)], slide_limit)
    if generate_toc and slides:
        toc_items = [slide.get("title", f"幻灯片 {index}") for index, slide in enumerate(slides, start=1)]
        toc_slides = [{"title": "目录" if start == 0 else f"目录 · 第 {start // 18 + 1} 页",
                       "body": toc_items[start:start + 18]} for start in range(0, len(toc_items), 18)]
        slides = [*toc_slides, *slides]
    preservation_lines = _docx_object_preservation_lines(object_preservation or {})
    if preservation_lines:
        slides.append({"title": "对象保留清单", "body": preservation_lines})
    build_pptx(slides, target)
    warnings = []
    if any(run.get("underline_color_unresolved") for block in blocks for run in block.get("runs", [])):
        warnings.append("下划线主题颜色尚未解析，已使用显式 RGB 回退值或文字颜色")
    if any(run.get("theme_font_unresolved") for block in blocks for run in block.get("runs", [])):
        warnings.append("部分主题字体尚未解析，已使用显式字体回退值或默认字体")
    if any(run.get("theme_color_unresolved") for block in blocks for run in block.get("runs", [])):
        warnings.append("主题文字颜色及明暗调整尚未解析，已使用显式 RGB 回退值或默认颜色")
    if retain_images and any("image" in block for block in blocks):
        warnings.append("图片按独立页保留，原始位置和文字环绕未还原")
    if retain_images and any(block.get("image_label") for block in blocks):
        warnings.append("页眉页脚图片按部件保留，原分页重复及显示条件未还原")
    return warnings


def build_docx_from_slides(
    slides: list[dict[str, Any]],
    target: Path,
    mode: str = "逐页讲义模式",
    generate_toc: bool = True,
    template_name: str = "",
    include_notes: bool = True,
    retain_images: bool = True,
    retain_formulas: bool = True,
    retain_tables: bool = True,
) -> list[str]:
    """根据提取的幻灯片结构生成轻量 DOCX 讲义。"""
    warnings = ["图片按行内绘图保留，原幻灯片位置和分组布局未还原"] if retain_images and any(slide.get("images") for slide in slides) else []
    if any(run.get("theme_font_unresolved") for slide in slides for group in [slide.get("title_runs", []), *slide.get("body_runs", [])] for run in group):
        warnings.append("部分 PPT 主题字体尚未解析，已使用 Word 默认字体")
    if any(run.get("theme_color_unresolved") for slide in slides for group in [slide.get("title_runs", []), *slide.get("body_runs", [])] for run in group):
        warnings.append("部分 PPT 文字颜色及填充变换尚未解析，已使用 Word 默认颜色")
    paragraphs: list[dict[str, Any]] = []
    if mode:
        paragraphs.append({"text": f"PPT 转 Word 模式：{mode}", "style": "Heading1"})
    if template_name:
        paragraphs.append({"text": f"Word 模板：{template_name}", "style": "Normal"})
    if generate_toc and slides:
        paragraphs.append({"text": "目录", "style": "Heading1"})
        for index, slide in enumerate(slides, start=1):
            paragraphs.append({"text": f"{index}. {slide.get('title', f'幻灯片 {index}')}", "style": "Normal"})
    if mode == "大纲模式":
        paragraphs.extend(_docx_slide_outline(slides, retain_tables, retain_images))
        build_docx(paragraphs, target)
        return warnings
    for slide in slides:
        paragraphs.append({"text": slide["title"], "style": "Heading1", "runs": slide.get("title_runs")})
        if mode == "图文混排模式":
            paragraphs.append({"text": f"幻灯片缩略图占位：第 {slide.get('index', 1)} 页", "style": "Normal"})
        summary = _pptx_slide_summary(
            slide,
            include_notes=include_notes,
            retain_images=retain_images,
            retain_formulas=retain_formulas,
        )
        if summary:
            paragraphs.append({"text": f"对象摘要：{summary}", "style": "Normal"})
        notes = [item for item in slide.get("notes", []) if item] if include_notes else []
        if mode == "备注优先模式" and notes:
            paragraphs.append({"text": "讲者备注", "style": "Normal"})
            for item in notes:
                paragraphs.append({"text": item, "style": "Normal"})
        for index, item in enumerate(slide.get("body", [])):
            if item:
                runs = slide.get("body_runs", [])
                paragraphs.append({"text": item, "style": "Normal", "runs": runs[index] if index < len(runs) else None})
        if retain_images:
            paragraphs.extend({"image": image} for image in slide.get("images", []))
        if retain_tables:
            paragraphs.extend(_slide_docx_tables(slide))
        if mode != "备注优先模式" and notes:
            paragraphs.append({"text": "备注", "style": "Normal"})
            for item in notes:
                paragraphs.append({"text": item, "style": "Normal"})
    build_docx(paragraphs, target)
    return warnings


def extract_pdf_text_blocks(path: Path) -> list[dict[str, Any]]:
    """从普通或 Flate 压缩的 PDF 内容流中提取文本操作数。"""
    if path.stat().st_size > PDF_TEXT_SOURCE_LIMIT_BYTES:
        return []
    data = path.read_bytes()
    streams = [data]
    expanded_total = 0
    for index, match in enumerate(re.finditer(rb"<<(.*?)>>\s*stream\r?\n(.*?)\r?\nendstream", data, re.DOTALL), start=1):
        if index > PDF_TEXT_STREAM_COUNT_LIMIT:
            return []
        dictionary, stream = match.groups()
        if b"/FlateDecode" in dictionary:
            try:
                stream = _bounded_flate_decode(stream, PDF_TEXT_STREAM_LIMIT_BYTES)
            except zlib.error:
                continue
            if stream is None:
                return []
        elif len(stream) > PDF_TEXT_STREAM_LIMIT_BYTES:
            return []
        expanded_total += len(stream)
        if expanded_total > PDF_TEXT_TOTAL_STREAM_LIMIT_BYTES:
            return []
        streams.append(stream)
    lines: list[str] = []
    for stream in streams:
        text_objects = re.findall(rb"BT(.*?)ET", stream, re.DOTALL) or [stream]
        for text_object in text_objects:
            values: list[str] = []
            for array, literal, hexadecimal in re.findall(
                rb"(\[(?:.|\r|\n)*?\])\s*TJ|((?:\((?:\\.|[^\\()])*\)\s*)+)Tj|<([0-9A-Fa-f\s]+)>\s*Tj",
                text_object,
            ):
                if array:
                    values.extend(_pdf_array_strings(array))
                elif literal:
                    values.extend(_pdf_literal_strings(literal))
                elif hexadecimal:
                    if b"/K12UnicodeText true" in data and b"/F1" in text_object:
                        values.append(bytes.fromhex(hexadecimal.decode()).decode("utf-16-be", errors="replace"))
                    else:
                        values.append(_decode_pdf_hex_string(hexadecimal))
            line = " ".join(value for value in values if value).strip()
            if line:
                lines.append(re.sub(r"\s+", " ", line))
    unique_lines = list(dict.fromkeys(lines))
    return [{"text": line, "style": "Normal"} for line in unique_lines]


def build_docx_from_pdf_text(path: Path, target: Path, title: str = "") -> dict[str, Any]:
    """将可提取的 PDF 文本层转为 DOCX，并返回转换证据。"""
    blocks = extract_pdf_text_blocks(path)
    if not blocks:
        raise ValueError("PDF 文本层未提取到可写入内容")
    paragraphs = [{"text": title or path.stem, "style": "Heading1"}, *blocks]
    build_docx(paragraphs, target)
    return {"paragraph_count": len(blocks), "character_count": sum(len(item["text"]) for item in blocks)}


def build_pdf_from_xlsx(sheets: list[dict[str, Any]], target: Path, title: str) -> list[str]:
    """将 Excel 工作表摘要写为文本 PDF。"""
    lines = [f"Excel to PDF: {title}", ""]
    for sheet in sheets:
        lines.append(f"Sheet {sheet.get('index')}: {sheet.get('name')}")
        if sheet.get("conversion_range"):
            lines.append(f"Range: {sheet.get('conversion_range')}")
        if sheet.get("formula_mode"):
            lines.append(f"Formula mode: {sheet.get('formula_mode')}")
        lines.append(
            "  ".join(
                [
                    f"Cells: {len(sheet.get('cells', []))}",
                    f"Formulas: {sheet.get('formula_count', 0)}",
                    f"Merged: {sheet.get('merged_cells', 0)}",
                    f"Charts: {sheet.get('chart_count', 0)}",
                    f"Images: {sheet.get('image_count', 0)}",
                    f"Comments: {sheet.get('comment_count', 0)}",
                ]
            )
        )
        for cell in sheet.get("cells", []):
            lines.append(f"{cell.get('ref', '')}: {cell.get('value', '')}")
        for index, chart in enumerate(sheet.get("charts", []), 1):
            lines.append(chart.get("title") or f"图表 {index}")
            lines.extend(f"{name} · {category}: {value}" for name, category, value in _xlsx_chart_rows(chart)[1:])
            if chart.get("warning"):
                lines.append(chart["warning"])
        lines.extend(f'批注 {item["ref"]} · {item["author"]}: {item.get("display_text", item["text"])}' for item in sheet.get("comments", []))
        lines.append("")
    build_text_pdf(lines, target, title)
    return _xlsx_chart_warnings(sheets)


def build_docx_from_xlsx(sheets: list[dict[str, Any]], target: Path, title: str) -> list[str]:
    """将规范化的 Excel 工作表预览写为 DOCX 摘要。"""
    paragraphs = [{"text": f"Excel 表格提取：{title}", "style": "Heading1"}]
    for sheet in sheets:
        paragraphs.append({"text": f"Sheet {sheet.get('index')}: {sheet.get('name')}", "style": "Heading1"})
        if sheet.get("conversion_range"):
            paragraphs.append({"text": f"转换范围：{sheet.get('conversion_range')}", "style": "Normal"})
        if sheet.get("formula_mode"):
            paragraphs.append({"text": f"公式模式：{sheet.get('formula_mode')}", "style": "Normal"})
        paragraphs.append(
            {
                "text": f"单元格 {len(sheet.get('cells', []))} 个；公式 {sheet.get('formula_count', 0)} 个；合并单元格 {sheet.get('merged_cells', 0)} 个",
                "style": "Normal",
            }
        )
        paragraphs.append(
            {
                "text": f"图表 {sheet.get('chart_count', 0)} 个；图片 {sheet.get('image_count', 0)} 个；批注 {sheet.get('comment_count', 0)} 个；表格对象 {sheet.get('table_count', 0)} 个",
                "style": "Normal",
            }
        )
        for rows in _xlsx_word_tables(sheet.get("cells", []), sheet.get("merge_ranges", [])):
            paragraphs.append({"table": rows})
        if sheet.get("comments"):
            paragraphs.append({"text": "单元格批注", "style": "Heading2"})
            paragraphs.append({"table": [["单元格", "作者", "批注"], *[[item["ref"], item["author"], item.get("display_text", item["text"])] for item in sheet["comments"]]]})
        for index, chart in enumerate(sheet.get("charts", []), 1):
            paragraphs.append({"text": chart.get("title") or f"图表 {index}", "style": "Heading2"})
            paragraphs.append({"table": _xlsx_chart_rows(chart)})
            if chart.get("warning"):
                paragraphs.append({"text": chart["warning"], "style": "Normal"})
    build_docx(paragraphs, target)
    return _xlsx_chart_warnings(sheets)


def _xlsx_word_tables(cells: list[dict[str, Any]], merge_ranges: list[str] | None = None) -> list[list[list[Any]]]:
    """按实际行列组织单元格，宽表分段，稀疏坐标不展开为空白巨表。"""
    positions = {}
    columns = {}
    for cell in cells:
        ref = str(cell.get("ref", "")).upper()
        match = re.fullmatch(r"([A-Z]+)([1-9][0-9]*)", ref)
        if not match:
            return [[["单元格", "值"], *[[str(item.get("ref", "")), str(item.get("value", ""))] for item in cells]]]
        letters, row = match.groups()
        column = excel_column_number(letters)
        columns[column] = letters
        positions[(int(row), column)] = f'{ref}: {cell.get("value", "")}'
    merges = []
    for ref in merge_ranges or []:
        match = re.fullmatch(r"([A-Z]+)([1-9][0-9]*):([A-Z]+)([1-9][0-9]*)", ref.upper())
        if not match:
            continue
        left, top, right, bottom = match.groups()
        first = excel_column_number(left)
        last = excel_column_number(right)
        top, bottom = int(top), int(bottom)
        if (top, first) not in positions:
            continue
        if last < first or bottom < top:
            raise ValueError("合并单元格范围无效")
        if (last - first + 1) * (bottom - top + 1) > 200000 or last - first >= 60:
            raise ValueError("合并单元格超出 Word 表格容量，请缩小转换范围")
        merges.append((top, bottom, first, last))
        for column in range(first, last + 1):
            columns[column] = excel_column_letters(column)
            for row in range(top, bottom + 1):
                positions.setdefault((row, column), "")
    row_numbers = sorted({row for row, column in positions})
    column_numbers = sorted(columns)
    tables = []
    # Word 表格列数有限制；坐标标签保留原始行列身份。
    for start in range(0, len(column_numbers), 60):
        group = column_numbers[start:start + 60]
        rows = [["行号", *[columns[column] for column in group]]]
        for row in row_numbers:
            if any((row, column) in positions for column in group):
                values = []
                for column in group:
                    value: Any = positions.get((row, column), "")
                    for top, bottom, first, last in merges:
                        if top <= row <= bottom and first <= column <= last:
                            if first not in group or last not in group:
                                raise ValueError("合并单元格跨越 Word 分表边界，请缩小转换范围")
                            value = {"text": value if row == top and column == first else "",
                                     "skip": column != first, "span": last - first + 1,
                                     "vertical": "restart" if row == top else "continue"} if bottom > top else {
                                         "text": value if column == first else "", "skip": column != first, "span": last - first + 1}
                            break
                    values.append(value)
                rows.append([str(row), *values])
        tables.append(rows)
    return tables


def build_pptx_from_xlsx(sheets: list[dict[str, Any]], target: Path, title: str) -> list[str]:
    """将规范化的 Excel 工作表预览写为 PPTX 摘要。"""
    slides: list[dict[str, Any]] = []
    warnings = []
    for sheet in sheets:
        body = [
            f"转换范围 {sheet.get('conversion_range', '全部工作表')}",
            f"公式模式 {sheet.get('formula_mode', '保留公式')}",
            f"单元格 {len(sheet.get('cells', []))} 个",
            f"公式 {sheet.get('formula_count', 0)} 个",
            f"合并单元格 {sheet.get('merged_cells', 0)} 个",
            f"图表 {sheet.get('chart_count', 0)} 个",
            f"图片 {sheet.get('image_count', 0)} 个",
            f"批注 {sheet.get('comment_count', 0)} 个",
        ]
        body.extend(f"{cell.get('ref', '')}: {cell.get('value', '')}" for cell in sheet.get("cells", []))
        for start in range(0, len(body), 12):
            page = start // 12 + 1
            slides.append({"title": f"{title} · {sheet.get('name')} · 第 {page} 页", "body": body[start:start + 12]})
        comments = [f'批注 {item["ref"]} · {item["author"]}: {item.get("display_text", item["text"])}' for item in sheet.get("comments", [])]
        for start in range(0, len(comments), 12):
            slides.append({"title": f"{sheet.get('name')} · 单元格批注 {start // 12 + 1}", "body": comments[start:start + 12]})
        for index, chart in enumerate(sheet.get("charts", []), 1):
            warning = str(chart.get("warning") or "")
            if not warning and chart.get("chart_type") not in {"barChart", "lineChart", "pieChart"}:
                warning = "此图表类型暂以数据页保留，未生成原生图表"
            if chart.get("chart_type") in {"barChart", "lineChart", "pieChart"} and not chart.get("warning"):
                try:
                    native_chart = build_chart(chart)
                    slides.append({"title": f"{sheet.get('name')} · {chart.get('title') or f'图表 {index}'}", "body": [], "native_chart": native_chart})
                except ValueError as exc:
                    warning = str(exc)
            rows = _xlsx_chart_rows(chart)[1:]
            lines = [f"{name} · {category}: {value}" for name, category, value in rows]
            if warning:
                lines.append(warning)
                warnings.append(f"{sheet.get('name')} · {chart.get('title') or f'图表 {index}'}：{warning}")
            for start in range(0, max(1, len(lines)), 12):
                slides.append({"title": f"{sheet.get('name')} · {chart.get('title') or f'图表 {index}'} · 数据 {start // 12 + 1}",
                               "body": lines[start:start + 12]})
    build_pptx(slides or [{"title": title, "body": ["未提取到工作表"]}], target)

    return warnings


def build_text_pdf(lines: list[str], target: Path, title: str = "K12 Export") -> None:
    """根据纯文本行生成符合基本规范的小型 PDF。"""
    target.parent.mkdir(parents=True, exist_ok=True)
    unicode_text = any(ord(char) > 127 for line in lines or [title] for char in str(line))
    pages = _pdf_pages(lines or [title], 42)
    font = "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>"
    if unicode_text:
        font = "<< /Type /Font /Subtype /Type0 /K12UnicodeText true /BaseFont /STSong-Light /Encoding /UniGB-UCS2-H /DescendantFonts [<< /Type /Font /Subtype /CIDFontType0 /BaseFont /STSong-Light /CIDSystemInfo << /Registry (Adobe) /Ordering (GB1) /Supplement 4 >> >>] >>"
    objects: list[bytes] = []
    objects.append(b"<< /Type /Catalog /Pages 2 0 R >>")
    page_refs = " ".join(f"{3 + index * 2} 0 R" for index in range(len(pages)))
    objects.append(f"<< /Type /Pages /Kids [{page_refs}] /Count {len(pages)} >>".encode("ascii"))
    for index, page_lines in enumerate(pages):
        page_obj = 3 + index * 2
        content_obj = page_obj + 1
        objects.append(
            f"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Resources << /Font << /F1 {font} >> >> /Contents {content_obj} 0 R >>".encode(
                "ascii"
            )
        )
        stream = _pdf_stream(page_lines, unicode_text)
        objects.append(f"<< /Length {len(stream)} >>\nstream\n".encode("ascii") + stream + b"\nendstream")
    _write_pdf(objects, target)


def build_docx(paragraphs: list[dict[str, Any]], target: Path) -> None:
    """将段落打包为可下载的最小 DOCX。"""
    target.parent.mkdir(parents=True, exist_ok=True)
    images = []
    body_parts = []
    for item in paragraphs:
        if "image" in item:
            images.append(item["image"])
            body_parts.append(word_picture_xml(item["image"], len(images)))
        elif "table" in item:
            body_parts.append(_docx_table(item["table"], item.get("column_widths"), item.get("row_heights"), item.get("cell_alignments")))
        else:
            body_parts.append(_docx_paragraph(item["text"], item.get("style", "Normal"), alignment=item.get("alignment", ""), runs=item.get("runs")))
    body = "".join(body_parts)
    document = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>"""
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", _docx_content_types())
        archive.writestr("_rels/.rels", _package_rels("word/document.xml"))
        archive.writestr("word/document.xml", document)
        archive.writestr("word/styles.xml", _docx_styles())
        relationships = ['<Relationship Id="styles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>']
        for index, image in enumerate(images, 1):
            part = f'image{index}.{image["extension"]}'
            archive.writestr(f"word/media/{part}", base64.b64decode(image["data"], validate=True))
            relationships.append(f'<Relationship Id="image{index}" Type="{OFFICE_REL}/image" Target="media/{part}"/>')
            if image.get("svg"):
                archive.writestr(f"word/media/vector{index}.svg", base64.b64decode(image["svg"]["data"], validate=True))
                relationships.append(f'<Relationship Id="svg{index}" Type="{OFFICE_REL}/image" Target="media/vector{index}.svg"/>')
        archive.writestr("word/_rels/document.xml.rels", f'<Relationships xmlns="{REL_NS}">{"".join(relationships)}</Relationships>')


def build_pptx(slides: list[dict[str, Any]], target: Path) -> None:
    """将幻灯片数据打包为可下载的最小 PPTX。"""
    target.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as archive:
        with zipfile.ZipFile(Path(__file__).with_name("pptx_template.pptx")) as template:
            for name in template.namelist():
                if name != "[Content_Types].xml":
                    archive.writestr(name, template.read(name))
        archive.writestr("[Content_Types].xml", _pptx_content_types(len(slides), [index for index, slide in enumerate(slides, 1) if slide.get("native_chart")]))
        archive.writestr("_rels/.rels", _package_rels("ppt/presentation.xml"))
        archive.writestr("ppt/presentation.xml", _presentation_xml(len(slides)))
        archive.writestr("ppt/_rels/presentation.xml.rels", _presentation_rels(len(slides)))
        for index, slide in enumerate(slides, start=1):
            relationships = [f'<Relationship Id="layout" Type="{OFFICE_REL}/slideLayout" Target="../slideLayouts/slideLayout7.xml"/>']
            xml = _slide_xml(slide.get("title", f"幻灯片 {index}"), slide.get("body", []), slide.get("title_runs"), slide.get("body_runs"))
            if slide.get("native_table"):
                xml = xml.replace("</p:spTree>", _pptx_table_frame(slide["native_table"], slide.get("table_merges", []), slide.get("column_widths"), slide.get("row_heights"), slide.get("cell_alignments")) + "</p:spTree>")
            if slide.get("native_image"):
                image = slide["native_image"]
                extension = image["extension"]
                archive.writestr(f"ppt/media/image{index}.{extension}", base64.b64decode(image["data"], validate=True))
                relationships.append(f'<Relationship Id="image" Type="{OFFICE_REL}/image" Target="../media/image{index}.{extension}"/>')
                if image.get("svg"):
                    archive.writestr(f"ppt/media/vector{index}.svg", base64.b64decode(image["svg"]["data"], validate=True))
                    relationships.append(f'<Relationship Id="svg" Type="{OFFICE_REL}/image" Target="../media/vector{index}.svg"/>')
                xml = xml.replace("</p:spTree>", picture_xml(image) + "</p:spTree>")
            if slide.get("native_chart"):
                chart_xml, workbook = slide["native_chart"]
                xml = xml.replace("</p:spTree>", chart_frame() + "</p:spTree>")
                archive.writestr(f"ppt/charts/chart{index}.xml", chart_xml)
                archive.writestr(f"ppt/embeddings/chart{index}.xlsx", workbook)
                relationships.append(f'<Relationship Id="chart" Type="{OFFICE_REL}/chart" Target="../charts/chart{index}.xml"/>')
                archive.writestr(f"ppt/charts/_rels/chart{index}.xml.rels", f'<Relationships xmlns="{REL_NS}"><Relationship Id="workbook" Type="{OFFICE_REL}/package" Target="../embeddings/chart{index}.xlsx"/></Relationships>')
            archive.writestr(f"ppt/slides/_rels/slide{index}.xml.rels", f'<Relationships xmlns="{REL_NS}">{"".join(relationships)}</Relationships>')
            archive.writestr(f"ppt/slides/slide{index}.xml", xml)


def _table_page_ranges(total: int, merges: list[tuple[int, int]]) -> list[tuple[int, int]]:
    """优先按十二格分页，避让合并区；超大合并区整块保留。"""
    pages, start = [], 0
    while start < total:
        end = min(start + 12, total)
        while end > start and any(first < end < last for first, last in merges):
            end -= 1
        if end == start:
            end = min(start + 12, total)
            while True:
                extended = max([end, *(last for first, last in merges if first < end < last)])
                if extended == end:
                    break
                end = extended
        pages.append((start, end))
        start = end
    return pages


def _slides_from_blocks(blocks: list[dict[str, Any]], max_chars: int | None) -> list[dict[str, Any]]:
    """将 DOCX 段落分组为基础 PPTX 幻灯片。保留完整标题，并按正文长度分拆长段落。"""
    slides: list[dict[str, Any]] = []
    current: dict[str, Any] | None = None
    resume_title: str | None = None
    used = 0
    for block in blocks:
        if "image" in block:
            title = block.get("image_label") or (current["title"] if current else (resume_title or "图片"))
            if current:
                slides.append(current)
            slides.append({"title": title, "body": [], "native_image": block["image"]})
            current, resume_title, used = None, title, 0
            continue
        if "table" in block:
            title = current["title"] if current else (resume_title or "表格")
            if current:
                slides.append(current)
                current = None
            rows = block["table"]
            merges = block.get("table_merges", [])
            widths = block.get("column_widths", [])
            heights = block.get("row_heights", [])
            alignments = block.get("cell_alignments", [])
            columns = max((len(row) for row in rows), default=0)
            column_pages = _table_page_ranges(columns, [(left, left + width) for _, left, _, width in merges])
            row_pages = _table_page_ranges(len(rows), [(top, top + height) for top, _, height, _ in merges])
            for column_page, (column, column_end) in enumerate(column_pages, 1):
                for row_page, (row, row_end) in enumerate(row_pages, 1):
                    page_merges = [[top - row, left - column, height, width] for top, left, height, width in merges
                                   if row <= top < row_end and column <= left < column_end]
                    slides.append({"title": f"{title} · 表格 {row_page}-{column_page}", "table_merges": page_merges, "column_widths": widths[column:column_end], "row_heights": heights[row:row_end], "cell_alignments": [values[column:column_end] for values in alignments[row:row_end]], "body": [],
                                   "native_table": [values[column:column_end] for values in rows[row:row_end]]})
            used = 0
            resume_title = title
            continue
        text = str(block["text"])
        if current is None and resume_title and block.get("level", 0) not in {1, 2}:
            current = {"title": resume_title, "body": []}
        resume_title = None
        if block.get("level", 0) in {1, 2} or current is None:
            if current:
                slides.append(current)
            current = {"title": text, "title_runs": block.get("runs", []), "body": [], "body_runs": []}
            used = 0
            continue
        if max_chars is None:
            current["body"].append(text)
            current.setdefault("body_runs", []).append(block.get("runs", []))
            continue
        offset = 0
        while text:
            if used == max_chars:
                slides.append(current)
                current = {"title": current["title"], "title_runs": current.get("title_runs", []), "body": [], "body_runs": []}
                used = 0
            piece = text[:max_chars - used]
            current["body"].append(piece)
            current.setdefault("body_runs", []).append(_slice_text_runs(block.get("runs", []), offset, len(piece)))
            offset += len(piece)
            used += len(piece)
            text = text[len(piece):]
    if current:
        slides.append(current)
    return slides or [{"title": "内容", "body": ["未提取到正文"]}]


def _slice_text_runs(runs: list[dict], start: int, length: int) -> list[dict]:
    """按文字偏移切分格式片段，分页边界不丢失局部样式。"""
    result, offset = [], 0
    for run in runs:
        text = run.get("text", "")
        left, right = max(0, start - offset), min(len(text), start + length - offset)
        if left < right:
            result.append({**run, "text": text[left:right]})
        offset += len(text)
    return result


def _docx_object_preservation_lines(plan: dict[str, Any]) -> list[str]:
    """格式化 DOCX 对象保留证据，不执行原生写回。"""
    source = plan.get("source") or {}
    statuses = plan.get("statuses") or {}
    object_total = sum(int(source.get(key, 0) or 0) for key in ("images", "tables", "formulas"))
    if not object_total:
        return []
    lines = [
        f"图片 {source.get('images', 0)} 个：{statuses.get('images') or '已记录源图片对象'}",
        f"表格 {source.get('tables', 0)} 个：{statuses.get('tables') or '已记录源表格结构'}",
        f"公式 {source.get('formulas', 0)} 个：{statuses.get('formulas') or '已记录源公式对象'}",
        f"OMML {source.get('omml_formulas', 0)} 个；MathType {source.get('mathtype_objects', 0)} 个",
    ]
    notes = plan.get("notes") or []
    if isinstance(notes, list):
        lines.extend(str(item) for item in notes[:4] if item)
    return lines


def _extract_text(xml: str, tag: str) -> str:
    """提取指定 XML 标签的文本并清理内容。"""
    pieces = re.findall(fr"<{tag}[^>]*>([\s\S]*?)</{tag}>", xml)
    return _clean_text("".join(pieces))


def _xlsx_shared_strings(archive: zipfile.ZipFile) -> list[str]:
    """读取 XLSX 共享字符串供工作表预览使用。"""
    try:
        xml = archive.read("xl/sharedStrings.xml").decode("utf-8", errors="ignore")
    except KeyError:
        return []
    try:
        root = ET.fromstring(xml)
    except ET.ParseError as exc:
        raise ValueError("共享字符串 XML 无法解析") from exc
    return ["".join(node.text or "" for node in item.findall(".//{*}t")) for item in root.findall("{*}si")]


def _xlsx_worksheet_order(archive: zipfile.ZipFile, names: list[str]) -> list[tuple[str, str]]:
    """通过工作簿关系读取真实顺序，拒绝缺失或外部工作表引用。"""
    try:
        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        sheets = workbook.findall(".//{*}sheet")
        relation_key = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
        if any(sheet.get(relation_key) for sheet in sheets):
            relationships = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
            by_id = {item.get("Id"): item for item in relationships}
            ordered = []
            for sheet in sheets:
                relation = by_id.get(sheet.get(relation_key))
                if relation is None or relation.get("TargetMode") == "External":
                    raise ValueError("工作表关系缺失或引用外部资源")
                target = relation.get("Target", "")
                path = posixpath.normpath(target.lstrip("/") if target.startswith("/") else "xl/" + target)
                if not path.startswith("xl/worksheets/") or path not in names:
                    raise ValueError("工作表引用不存在或超出工作表目录")
                ordered.append((sheet.get("name") or f"Sheet{len(ordered) + 1}", path))
            return ordered
    except KeyError as exc:
        raise ValueError("工作簿缺少必需的关系文件") from exc
    except ET.ParseError as exc:
        raise ValueError("工作簿 XML 无法解析") from exc
    # 兼容未声明关系的旧版最小工作簿。
    titles = _xlsx_sheet_names(archive)
    paths = sorted([name for name in names if name.startswith("xl/worksheets/sheet") and name.endswith(".xml")], key=_worksheet_number)
    return [(titles.get(index, f"Sheet{index}"), path) for index, path in enumerate(paths, 1)]


def _xlsx_sheet_names(archive: zipfile.ZipFile) -> dict[int, str]:
    """按工作表顺序读取名称。"""
    try:
        xml = archive.read("xl/workbook.xml").decode("utf-8", errors="ignore")
    except KeyError:
        return {}
    names: dict[int, str] = {}
    for index, match in enumerate(re.finditer(r"<sheet[^>]*name=\"([^\"]+)\"", xml), start=1):
        names[index] = html.unescape(match.group(1))
    return names


def _xlsx_cells(xml: str, shared_strings: list[str]) -> list[dict[str, str]]:
    """从工作表 XML 提取可见单元格值与公式。"""
    try:
        root = ET.fromstring(xml)
    except ET.ParseError as exc:
        raise ValueError("工作表 XML 无法解析") from exc
    cells: list[dict[str, str]] = []
    for cell in root.findall(".//{*}c"):
        value = cell.findtext("{*}v", default="")
        inline = cell.find("{*}is")
        if inline is not None:
            value = "".join(node.text or "" for node in inline.findall(".//{*}t"))
        elif cell.get("t") == "s":
            if not value.isdigit() or int(value) >= len(shared_strings):
                raise ValueError("单元格共享字符串索引无效")
            value = shared_strings[int(value)]
        result = value
        formula = cell.findtext("{*}f", default="")
        if formula:
            value = f"{value} (={formula})" if value else f"={formula}"
        if value:
            cells.append({"ref": cell.get("r", ""), "value": value, "result": result, "formula": formula})
    return cells


def _pptx_related_root(archive: zipfile.ZipFile, part: str, kind: str, following: str | tuple[str, ...] = "") -> ET.Element | None:
    """按关系类型加载版式或母版，不依赖固定文件名。"""
    relation_part = f"{posixpath.dirname(part)}/_rels/{posixpath.basename(part)}.rels"
    if relation_part not in archive.namelist():
        return None
    for relation in ET.fromstring(archive.read(relation_part)):
        if not relation.get("Type", "").endswith("/" + kind):
            continue
        if relation.get("TargetMode") == "External":
            raise ValueError("PPT 样式部件不能使用外部关系")
        target = _resolve_xlsx_target(part, relation.get("Target", ""))
        if following:
            chain = (following,) if isinstance(following, str) else following
            return _pptx_related_root(archive, target, chain[0], chain[1:])
        return parse_compatible_xml(archive.read(target))
    return None


def _pptx_related_parts(archive: zipfile.ZipFile, part_path: str) -> list[str]:
    """解析幻灯片或备注部件引用的关联部件。"""
    rels_path = f"{posixpath.dirname(part_path)}/_rels/{posixpath.basename(part_path)}.rels"
    try:
        relations = ET.fromstring(archive.read(rels_path))
    except KeyError:
        return []
    return [_resolve_xlsx_target(part_path, relation.get("Target", ""))
            for relation in relations if relation.get("TargetMode") != "External"]


def _pptx_notes(archive: zipfile.ZipFile, related_parts: list[str]) -> list[str]:
    """从关联的备注幻灯片提取演讲者备注。"""
    notes: list[str] = []
    for part in related_parts:
        if "/notesSlides/" not in f"/{part}":
            continue
        try:
            notes_xml = archive.read(part).decode("utf-8", errors="ignore")
        except KeyError:
            continue
        try:
            root = ET.fromstring(notes_xml)
            shapes = root.findall(".//{*}sp")
        except ET.ParseError:
            shapes = []
        if not shapes:
            notes.extend(_pptx_texts(notes_xml))
            continue
        for shape in shapes:
            placeholder = shape.find("{*}nvSpPr/{*}nvPr/{*}ph")
            if placeholder is not None and placeholder.get("type") in {"hdr", "ftr", "dt", "sldNum", "sldImg"}:
                continue
            notes.extend(_pptx_texts(ET.tostring(shape, encoding="unicode")))
    return notes


def _pptx_text_body_defaults(body: ET.Element | None, level: int) -> list[ET.Element | None]:
    """读取占位符文本框及同级段落的默认格式，供版式和母版共用。"""
    if body is None:
        return []
    defaults = [body.find(path) for path in ("{*}lstStyle/{*}defPPr", "{*}lstStyle/{*}defPPr/{*}defRPr", "{*}lstStyle/{*}lvl" + str(level + 1) + "pPr", "{*}lstStyle/{*}lvl" + str(level + 1) + "pPr/{*}defRPr")]
    for paragraph in body.findall("{*}p"):
        properties = paragraph.find("{*}pPr")
        if properties is not None and int(properties.get("lvl", "0")) == level:
            defaults.extend([properties, properties.find("{*}defRPr")])
            break
    return defaults


def _pptx_color_mapping(slide: ET.Element, layout: ET.Element | None, master: ET.Element | None) -> dict[str, str]:
    """按母版、版式、幻灯片读取色板映射，显式母版映射会重置覆盖。"""
    mapping = {"tx1": "dk1", "tx2": "dk2", "bg1": "lt1", "bg2": "lt2"}
    master_mapping = master.find("{*}clrMap") if master is not None else None
    if master_mapping is not None:
        mapping.update(master_mapping.attrib)
    result = mapping.copy()
    for root in (layout, slide):
        if root is None:
            continue
        override = root.find("{*}clrMapOvr")
        if override is not None:
            if override.find("{*}masterClrMapping") is not None:
                result = mapping.copy()
            explicit = override.find("{*}overrideClrMapping")
            if explicit is not None:
                result.update(explicit.attrib)
    return result


def _pptx_styled_paragraphs(xml: str, layout: ET.Element | None = None, master: ET.Element | None = None, theme: ET.Element | None = None) -> list[list[dict]]:
    """读取文本框、段落和片段格式，空白裁剪保持文字与片段对应。"""
    try:
        root = ET.fromstring(xml)
    except ET.ParseError:
        return []
    theme_fonts = WordTextStyles(None, theme)
    color_mapping = _pptx_color_mapping(root, layout, master)
    for parent in root.iter():
        for child in list(parent):
            if child.tag.rsplit("}", 1)[-1] == "tbl":
                parent.remove(child)
    parents = {child: parent for parent in root.iter() for child in parent}
    paragraphs = []
    for paragraph in root.findall(".//{*}p"):
        body = parents.get(paragraph)
        list_style = body.find("{*}lstStyle") if body is not None else None
        paragraph_properties = paragraph.find("{*}pPr")
        level = int(paragraph_properties.get("lvl", "0")) if paragraph_properties is not None else 0
        if not 0 <= level <= 8:
            raise ValueError("PPT 段落级别超出可转换范围")
        defaults = [list_style.find(path) if list_style is not None else None
                    for path in ("{*}defPPr", "{*}defPPr/{*}defRPr", "{*}lvl" + str(level + 1) + "pPr", "{*}lvl" + str(level + 1) + "pPr/{*}defRPr")]
        shape = parents.get(body)
        placeholder = shape.find("{*}nvSpPr/{*}nvPr/{*}ph") if shape is not None else None
        placeholder_type = placeholder.get("type", "obj") if placeholder is not None else ""
        if placeholder is not None and layout is not None:
            for candidate in layout.findall(".//{*}sp"):
                reference = candidate.find("{*}nvSpPr/{*}nvPr/{*}ph")
                if reference is None or reference.get("idx", "0") != placeholder.get("idx", "0"):
                    continue
                placeholder_type = reference.get("type", "obj")
                inherited = candidate.find("{*}txBody")
                if inherited is not None:
                    defaults = [*_pptx_text_body_defaults(inherited, level), *defaults]
                break
        if master is not None:
            style_kind = "titleStyle" if placeholder_type in {"title", "ctrTitle"} else "bodyStyle" if placeholder_type in {"body", "obj", "subTitle"} else "otherStyle"
            inherited_defaults = []
            master_type = {"obj": "body", "ctrTitle": "title"}.get(placeholder_type, placeholder_type)
            if placeholder is not None:
                for candidate in master.findall(".//{*}sp"):
                    reference = candidate.find("{*}nvSpPr/{*}nvPr/{*}ph")
                    if reference is not None and reference.get("type", "obj") == master_type:
                        inherited_defaults = _pptx_text_body_defaults(candidate.find("{*}txBody"), level)
                        break
            defaults = [*inherited_defaults, *defaults]
            master_style = master.find("{*}txStyles/{*}" + style_kind)
            if master_style is not None:
                defaults = [master_style.find(path) for path in ("{*}defPPr", "{*}defPPr/{*}defRPr", "{*}lvl" + str(level + 1) + "pPr", "{*}lvl" + str(level + 1) + "pPr/{*}defRPr")] + defaults
        defaults.extend([paragraph_properties, paragraph.find("{*}pPr/{*}defRPr")])
        runs = []
        for node in paragraph:
            tag = node.tag.rsplit("}", 1)[-1]
            if tag == "br":
                runs.append({"text": "\n"})
            elif tag in {"r", "fld"}:
                run = {"text": node.findtext("{*}t", "")}
                for properties in [*defaults, node.find("{*}rPr")]:
                    if properties is not None:
                        if properties.get("algn"):
                            run["alignment"] = {value: key for key, value in PARAGRAPH_ALIGNMENTS.items()}.get(properties.get("algn"), "")
                        if properties.get("lang"):
                            run["lang"] = properties.get("lang")
                        if properties.get("u"):
                            if properties.get("u") not in UNDERLINES.values():
                                raise ValueError("PPT 下划线类型无效")
                            run["underline"] = properties.get("u")
                        underline_color = properties.find("{*}uFill/{*}solidFill/{*}srgbClr")
                        if properties.find("{*}uFillTx") is not None or properties.find("{*}uFill") is not None:
                            run.pop("underline_color", None)
                        if underline_color is not None and re.fullmatch(r"[0-9a-fA-F]{6}", underline_color.get("val", "")):
                            run["underline_color"] = underline_color.get("val").upper()
                        for key in ("b", "i"):
                            if key in properties.attrib:
                                run[key] = properties.get(key) not in {"0", "false", "off"}
                        if properties.get("sz"):
                            run["size"] = int(properties.get("sz"))
                        for key, font in (("font", "latin"), ("east_asia", "ea")):
                            value = properties.find("{*}" + font)
                            if value is not None:
                                run.pop(key, None)
                            if value is not None:
                                run.pop(key + "_theme", None)
                                if value.get("typeface", "").startswith("+"):
                                    run[key + "_theme"] = value.get("typeface")
                                elif value.get("typeface"):
                                    run[key] = value.get("typeface")
                        color = properties.find("{*}solidFill/{*}srgbClr")
                        if any(properties.find("{*}" + fill) is not None for fill in ("solidFill", "noFill", "gradFill", "blipFill", "pattFill", "grpFill")):
                            run.pop("color", None)
                            run.pop("theme_color_unresolved", None)
                        if color is not None and not len(color) and re.fullmatch(r"[0-9a-fA-F]{6}", color.get("val", "")):
                            run["color"] = color.get("val").upper()
                        else:
                            scheme = properties.find("{*}solidFill/{*}schemeClr")
                            value = theme_fonts.theme_color({"themeColor": color_mapping.get(scheme.get("val"), scheme.get("val"))}) if scheme is not None and not len(scheme) else None
                            if value:
                                run["color"] = value
                            elif properties.find("{*}solidFill") is not None:
                                run["theme_color_unresolved"] = True
                for key in ("font", "east_asia"):
                    reference = run.pop(key + "_theme", None)
                    if reference:
                        match = re.fullmatch(r"\+(mj|mn)-(lt|ea)", reference)
                        font = theme_fonts.theme_font(("major" if match[1] == "mj" else "minor") + ("Ascii" if match[2] == "lt" else "EastAsia"), language=run.get("lang", "") if match[2] == "ea" else None) if match else None
                        if font:
                            run[key] = font
                        else:
                            run["theme_font_unresolved"] = True
                runs.append(run)
        text = "".join(run["text"] for run in runs)
        if text.strip():
            paragraphs.append(_slice_text_runs(runs, len(text) - len(text.lstrip()), len(text.strip())))
    return paragraphs


def _pptx_texts(xml: str, exclude_tables: bool = False, preserve_whitespace: bool = False) -> list[str]:
    """提取 PPTX XML 中可见的 DrawingML 文本。"""
    try:
        root = ET.fromstring(xml)
    except ET.ParseError:
        texts = re.findall(r"<a:t\b[^>]*>([\s\S]*?)</a:t>", xml)
        return [_clean_text(item) for item in texts if _clean_text(item)]
    if exclude_tables:
        for parent in root.iter():
            for child in list(parent):
                if child.tag.rsplit("}", 1)[-1] == "tbl":
                    parent.remove(child)
    paragraphs = [root] if root.tag.rsplit("}", 1)[-1] == "p" else root.findall(".//{*}p")
    if not paragraphs:
        return [node.text for node in root.findall(".//{*}t") if node.text]
    texts = []
    for paragraph in paragraphs:
        pieces = []
        for node in paragraph.iter():
            tag = node.tag.rsplit("}", 1)[-1]
            if tag == "t":
                pieces.append(node.text or "")
            elif tag == "br":
                pieces.append("\n")
        text = "".join(pieces)
        if not preserve_whitespace:
            text = text.strip()
        if text or preserve_whitespace:
            texts.append(text)
    return texts


def _docx_slide_outline(slides: list[dict[str, Any]], retain_tables: bool = True, retain_images: bool = True) -> list[dict[str, Any]]:
    """将幻灯片摘要转为 DOCX 大纲段落。"""
    paragraphs: list[dict[str, Any]] = []
    for slide in slides:
        paragraphs.append({"text": slide["title"], "style": "Heading1", "runs": slide.get("title_runs")})
        for index, item in enumerate(slide.get("body", [])):
            runs = slide.get("body_runs", [])
            paragraphs.append({"text": item, "style": "Normal", "runs": runs[index] if index < len(runs) else None})
        if retain_images:
            paragraphs.extend({"image": image} for image in slide.get("images", []))
        if retain_tables:
            paragraphs.extend(_slide_docx_tables(slide))
    return paragraphs


def _pptx_slide_summary(
    slide: dict[str, Any],
    include_notes: bool = True,
    retain_images: bool = True,
    retain_formulas: bool = True,
) -> str:
    """汇总讲义所需的 PPTX 对象，不声称完成真实渲染。"""
    parts = [
        ("文本框", slide.get("text_box_count", 0)),
        ("图片", slide.get("image_count", 0) if retain_images else 0),
        ("表格", slide.get("table_count", 0)),
        ("公式", slide.get("formula_count", 0) if retain_formulas else 0),
        ("图表", slide.get("chart_count", 0)),
        ("形状", slide.get("shape_count", 0)),
        ("备注", len(slide.get("notes", []) or []) if include_notes else 0),
    ]
    return "；".join(f"{label} {int(count or 0)} 个" for label, count in parts if int(count or 0) > 0)


def _pptx_formula_count(xml: str) -> int:
    """统计 PPTX 公式标记、MathType 引用与 TeX 片段。"""
    math_objects = len(re.findall(r"<(?:\w+:)?oMath(?:Para)?\b", xml))
    mathtype_refs = len(re.findall(r"MathType|Equation Native", xml, flags=re.IGNORECASE))
    latex_refs = len(re.findall(r"\$[^$]{1,120}\$|\\(?:frac|sqrt|sum|int)\b", xml))
    return math_objects + mathtype_refs + latex_refs


def _xlsx_sheet_objects(archive: zipfile.ZipFile, sheet_path: str, sheet_xml: str) -> dict[str, Any]:
    """统计工作表中的绘图、图表、图片、批注和表格。"""
    rels_path = f"{posixpath.dirname(sheet_path)}/_rels/{posixpath.basename(sheet_path)}.rels"
    try:
        rels_xml = archive.read(rels_path).decode("utf-8", errors="ignore")
    except KeyError:
        rels_xml = ""
    relations = ET.fromstring(rels_xml) if rels_xml else ET.Element("Relationships")
    internal = [node for node in relations if node.get("Target") and node.get("TargetMode") != "External"]
    related_parts = [_resolve_xlsx_target(sheet_path, node.get("Target")) for node in internal]
    thread_paths = [_resolve_xlsx_target(sheet_path, node.get("Target")) for node in internal
                    if node.get("Type", "").endswith("/threadedComment") or "/threadedcomments/" in node.get("Target", "").lower()]
    drawing_paths = sorted({path for path in related_parts if "/drawings/" in f"/{path.lower()}"})
    comment_paths = sorted({path for path in related_parts if posixpath.basename(path).lower().startswith("comments")})
    table_paths = {path for path in related_parts if "/tables/" in f"/{path.lower()}"}
    table_count = max(len(re.findall(r"<(?:\w+:)?tablePart\b", sheet_xml)), len(table_paths))
    chart_count = 0
    image_count = 0
    chart_paths: set[str] = set()
    for drawing_path in drawing_paths:
        try:
            drawing_xml = archive.read(drawing_path).decode("utf-8", errors="ignore")
        except KeyError:
            continue
        chart_tags = len(re.findall(r"<(?:\w+:)?chart\b", drawing_xml))
        image_tags = len(re.findall(r"<(?:\w+:)?pic\b", drawing_xml))
        drawing_rels = f"{posixpath.dirname(drawing_path)}/_rels/{posixpath.basename(drawing_path)}.rels"
        try:
            drawing_rels_xml = archive.read(drawing_rels).decode("utf-8", errors="ignore")
        except KeyError:
            drawing_rels_xml = ""
        drawing_related_parts = [
            _resolve_xlsx_target(drawing_path, target)
            for target in re.findall(r'Target="([^"]+)"', drawing_rels_xml)
        ]
        chart_paths.update(path for path in drawing_related_parts if "/charts/" in f"/{path.lower()}")
        chart_count += max(chart_tags, len({path for path in drawing_related_parts if "/charts/" in f"/{path.lower()}"}))
        image_count += max(image_tags, len({path for path in drawing_related_parts if "/media/" in f"/{path.lower()}"}))
    comments = []
    for comment_path in comment_paths:
        try:
            comments_xml = archive.read(comment_path).decode("utf-8", errors="ignore")
        except KeyError:
            continue
        root = ET.fromstring(comments_xml)
        authors = [node.text or "" for node in root.findall(".//{*}authors/{*}author")]
        for node in root.findall(".//{*}comment"):
            author_id = node.get("authorId", "")
            author = authors[int(author_id)] if author_id.isdigit() and int(author_id) < len(authors) else ""
            text = "".join(part.text or "" for part in node.findall(".//{*}t"))
            comments.append({"ref": node.get("ref", ""), "author": author, "text": text})
    comments = _xlsx_merge_comments(comments, _xlsx_threaded_comments(archive, thread_paths))
    return {
        "drawing_count": len(drawing_paths),
        "chart_count": chart_count,
        "charts": [_xlsx_chart_data(archive, path) for path in sorted(chart_paths)],
        "image_count": image_count,
        "comment_count": len(comments),
        "comments": comments,
        "table_count": table_count,
    }


def _xlsx_merge_comments(legacy: list[dict[str, Any]], threaded: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """按兼容批注的线程标识去重，并采用兼容批注的实际单元格位置。"""
    by_id = {item["id"].casefold(): item for item in threaded if item["id"]}
    roots = {key: item for key, item in by_id.items() if not item["parent_id"]}
    positions = {}
    retained = []
    for item in legacy:
        marker = re.search(r"tc=(\{[^}]+\})", item["author"], flags=re.IGNORECASE)
        key = marker.group(1).casefold() if marker else ""
        if key in roots:
            if item["ref"]:
                positions[key] = item["ref"]
        else:
            retained.append(item)
    for item in threaded:
        root = item
        visited = set()
        while root["parent_id"]:
            parent_id = root["parent_id"].casefold()
            if parent_id in visited or parent_id not in by_id:
                raise ValueError("线程批注回复关系无效")
            visited.add(parent_id)
            root = by_id[parent_id]
        if root["id"].casefold() in positions:
            item["ref"] = positions[root["id"].casefold()]
    return retained + threaded


def _xlsx_threaded_comments(archive: zipfile.ZipFile, related_parts: list[str]) -> list[dict[str, Any]]:
    """读取线程批注及人员显示名，回复继承父批注位置。"""
    paths = list(dict.fromkeys(related_parts))
    if not paths:
        return []
    people = {}
    try:
        relations = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    except KeyError:
        relations = ET.Element("Relationships")
    for relation in relations:
        if relation.get("TargetMode") == "External" or not relation.get("Type", "").endswith("/person"):
            continue
        target = _resolve_xlsx_target("xl/workbook.xml", relation.get("Target", ""))
        try:
            persons = ET.fromstring(archive.read(target))
        except KeyError:
            continue
        people.update({node.get("id", "").casefold(): node.get("displayName", "") for node in persons.findall(".//{*}person")})
    comments = []
    for path in paths:
        try:
            root = ET.fromstring(archive.read(path))
        except KeyError:
            raise ValueError("工作表线程批注部件缺失") from None
        for node in root.findall(".//{*}threadedComment"):
            comments.append({"id": node.get("id", ""), "parent_id": node.get("parentId", ""),
                             "ref": node.get("ref", ""), "author": people.get(node.get("personId", "").casefold(), ""),
                             "text": node.findtext("{*}text", ""), "date": node.get("dT", ""),
                             "resolved": node.get("done") in {"1", "true"}, "threaded": True})
    by_id = {item["id"].casefold(): item for item in comments if item["id"]}
    for item in comments:
        parent = item
        visited = set()
        while not parent["ref"] and parent["parent_id"]:
            parent_id = parent["parent_id"].casefold()
            if parent_id in visited or parent_id not in by_id:
                raise ValueError("线程批注回复关系无效，无法确定单元格位置")
            visited.add(parent_id)
            parent = by_id[parent_id]
        if not parent["ref"]:
            raise ValueError("线程批注缺少单元格位置")
        item["ref"] = parent["ref"]
        prefix = "回复：" if item["parent_id"] else ""
        suffix = "（已解决）" if item["resolved"] else ""
        item["display_text"] = prefix + item["text"] + suffix
    return comments


def _xlsx_chart_data(archive: zipfile.ZipFile, path: str) -> dict[str, Any]:
    """提取图表缓存数据及源公式，缺失缓存时明确标注。"""
    chart: dict[str, Any] = {"part": path, "title": "", "series": []}
    try:
        root = ET.fromstring(archive.read(path))
    except (KeyError, ET.ParseError):
        chart["warning"] = "图表部件缺失或格式无效"
        return chart
    plot = root.find(".//{*}plotArea")
    types = [node.tag.rsplit("}", 1)[-1] for node in plot if node.tag.endswith("Chart")] if plot is not None else []
    chart["chart_type"] = types[0] if len(types) == 1 else "combination"
    marker = root.find(".//{*}lineChart/{*}marker")
    chart["line_marker"] = marker is not None and marker.get("val") in {"1", "true"}
    chart["bar_direction"] = "col"
    grouping = root.find(".//{*}grouping")
    chart["grouping"] = grouping.get("val", "clustered") if grouping is not None else "clustered"
    direction = root.find(".//{*}barDir")
    if direction is not None:
        chart["bar_direction"] = direction.get("val", "col")
    chart["title"] = "".join(node.text or "" for node in root.findall(".//{*}title//{*}t"))
    for index, series in enumerate(root.findall(".//{*}ser"), 1):
        cached_name = series.findtext("{*}tx/{*}v") or series.findtext("{*}tx/{*}strRef/{*}strCache/{*}pt/{*}v")
        name = cached_name or f"系列 {index}"
        category = series.find("{*}cat")
        if category is None:
            category = series.find("{*}xVal")
        value = series.find("{*}val")
        if value is None:
            value = series.find("{*}yVal")
        categories = _chart_cached_points(category)
        values = _chart_cached_points(value)
        symbol = series.find("{*}marker/{*}symbol")
        marker_size = series.find("{*}marker/{*}size")
        smooth = series.find("{*}smooth")
        if smooth is None:
            smooth = root.find(".//{*}lineChart/{*}smooth")
        chart["series"].append({"name": name, "categories": categories, "values": values,
                                "category_formula": category.findtext(".//{*}f", "") if category is not None else "",
                                "value_formula": value.findtext(".//{*}f", "") if value is not None else "",
                                "name_formula": series.findtext("{*}tx/{*}strRef/{*}f", ""), "name_cached": bool(cached_name),
                                "marker_symbol": symbol.get("val") if symbol is not None else None,
                                "marker_size": marker_size.get("val") if marker_size is not None else None,
                                "smooth": smooth is not None and smooth.get("val", "1") in {"1", "true"}})
    if not chart["series"] or any(not chart_has_values(item) for item in chart["series"]):
        chart["warning"] = "图表未保存完整数据缓存，需用 Office 更新源数据"
    return chart


def _xlsx_resolve_chart_sources(sheets: list[dict[str, Any]]) -> None:
    """用同一工作簿的单元格补齐未保存的图表数据缓存。"""
    sources = {sheet["name"].casefold(): {cell["ref"].upper(): cell for cell in sheet["cells"]} for sheet in sheets}
    for sheet in sheets:
        for chart in sheet.get("charts", []):
            unresolved = not chart.get("series")
            for series in chart.get("series", []):
                for field, formula in [("categories", "category_formula"), ("values", "value_formula")]:
                    if not series[field] and series.get(formula):
                        try:
                            series[field] = _xlsx_chart_source_points(series[formula], sources)
                        except ValueError:
                            unresolved = True
                if series.get("name_formula") and not series.get("name_cached"):
                    try:
                        names = _xlsx_chart_source_points(series["name_formula"], sources)
                        if len(names) == 1 and names.get(0):
                            series["name"] = names[0]
                        else:
                            unresolved = True
                    except ValueError:
                        unresolved = True
                unresolved = unresolved or not chart_has_values(series)
            if unresolved:
                chart["warning"] = chart.get("warning") or "图表源数据无法完整解析，需用 Office 更新源数据"
            else:
                chart.pop("warning", None)


def _xlsx_chart_source_points(formula: str, sources: dict[str, dict[str, Any]]) -> dict[int, str]:
    """读取内部单行或单列引用，拒绝外部工作簿和未计算公式。"""
    match = re.fullmatch(r"=?('(?:[^']|'')+'|[^'!\[\]]+)!(\$?[A-Za-z]+\$?[1-9][0-9]*(?::\$?[A-Za-z]+\$?[1-9][0-9]*)?)", formula.strip())
    if not match:
        raise ValueError("图表源引用格式不受支持")
    name, region = match.groups()
    name = name[1:-1].replace("''", "'") if name.startswith("'") else name
    if "[" in name or "]" in name or name.casefold() not in sources:
        raise ValueError("图表源工作表不存在或属于外部工作簿")
    refs = excel_cell_refs([region.replace("$", "")])
    positions = sorted((int(re.search(r"[0-9]+", ref).group()), excel_column_number(re.match(r"[A-Z]+", ref).group()), ref) for ref in refs)
    if len({row for row, column, ref in positions}) > 1 and len({column for row, column, ref in positions}) > 1:
        raise ValueError("图表系列引用必须是单行或单列")
    points = {}
    for index, (row, column, ref) in enumerate(positions):
        cell = sources[name.casefold()].get(ref, {})
        result = cell.get("result", cell.get("value", ""))
        if cell.get("formula") and result in (None, ""):
            raise ValueError("图表源公式尚未保存计算结果")
        points[index] = str(result) if result is not None else ""
    return points


def _chart_cached_points(node: ET.Element | None) -> dict[int, str]:
    """按原始索引读取缓存点，保留空洞、零值和文字。"""
    if node is None:
        return {}
    points = {}
    for point in node.findall(".//{*}pt"):
        try:
            index = int(point.get("idx", ""))
        except ValueError:
            continue
        if index >= 0:
            points[index] = point.findtext("{*}v", "")
    return points


def _xlsx_chart_warnings(sheets: list[dict[str, Any]]) -> list[str]:
    """收集源图表解析提示，供转换报告与本地同步共用。"""
    return [f"{sheet.get('name')} · {chart.get('title') or f'图表 {index}'}：{chart['warning']}"
            for sheet in sheets for index, chart in enumerate(sheet.get("charts", []), 1) if chart.get("warning")]


def _xlsx_chart_rows(chart: dict[str, Any]) -> list[list[str]]:
    """将图表系列按点索引组织为可读数据表。"""
    rows = [["系列", "类别", "数值"]]
    for series in chart.get("series", []):
        categories, values = series["categories"], series["values"]
        for index in sorted(set(categories) | set(values)):
            rows.append([series["name"], categories.get(index, str(index + 1)), values.get(index, "")])
    return rows


def _resolve_xlsx_target(source_path: str, target: str) -> str:
    """相对于源部件解析 OOXML 关系目标。"""
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(posixpath.dirname(source_path), target))


def _strip_xml(xml: str) -> str:
    """移除 XML 标签并规范剩余文本。"""
    return _clean_text(re.sub(r"<[^>]+>", "", xml))


def _clean_text(text: str) -> str:
    """合并空白并还原 XML 或 HTML 实体。"""
    return html.unescape(re.sub(r"\s+", " ", text)).strip()


def _heading_level(paragraph: str, index: int) -> int:
    """根据 DOCX 样式标记推断标题层级。"""
    style_match = re.search(r'<(?:\w+:)?pStyle\b[^>]*\b(?:\w+:)?val=["\']([^"\']+)', paragraph)
    style = style_match.group(1).lower().replace(" ", "") if style_match else ""
    heading = re.fullmatch(r"heading([1-9])", style)
    if heading:
        return int(heading.group(1))
    return 1 if style in {"title", "1"} or index == 1 else 0


def _slide_number(path: str) -> int:
    """返回 PPTX 幻灯片路径中的序号。"""
    match = re.search(r"slide(\d+)\.xml$", path)
    return int(match.group(1)) if match else 0


def _worksheet_number(path: str) -> int:
    """返回 XLSX 工作表路径中的序号。"""
    match = re.search(r"sheet(\d+)\.xml$", path)
    return int(match.group(1)) if match else 0


def _pdf_pages(lines: list[str], page_size: int) -> list[list[str]]:
    """将文本行分页供轻量 PDF 写入器使用。"""
    pages: list[list[str]] = []
    current: list[str] = []
    for line in lines:
        for chunk in _pdf_line_chunks(line):
            current.append(chunk)
            if len(current) >= page_size:
                pages.append(current)
                current = []
    if current:
        pages.append(current)
    return pages or [["K12 Export"]]


def _pdf_array_strings(value: bytes) -> list[str]:
    """解码 PDF TJ 数组中的字面量和十六进制字符串。"""
    fragments: list[str] = []
    for match in re.finditer(rb"\((?:\\.|[^\\()])*\)|<([0-9A-Fa-f\s]+)>", value):
        token = match.group(0)
        if token.startswith(b"("):
            decoded = _pdf_literal_strings(token)
            fragments.extend(decoded)
        else:
            fragments.append(_decode_pdf_hex_string(match.group(1)))
    text = "".join(fragment for fragment in fragments if fragment)
    return [text] if text else []


def _bounded_flate_decode(value: bytes, limit: int) -> bytes | None:
    """解码 Flate 流，并限制展开后的数据大小。"""
    decoder = zlib.decompressobj()
    output = decoder.decompress(value, limit + 1)
    if len(output) > limit or decoder.unconsumed_tail:
        return None
    remaining = limit + 1 - len(output)
    output += decoder.flush(remaining)
    if len(output) > limit or not decoder.eof:
        return None
    return output


def _pdf_literal_strings(value: bytes) -> list[str]:
    """解码成对括号中的 PDF 字符串，处理转义与八进制字节。"""
    results: list[str] = []
    index = 0
    while index < len(value):
        if value[index] != 40:
            index += 1
            continue
        index += 1
        depth = 1
        decoded = bytearray()
        while index < len(value) and depth:
            byte = value[index]
            if byte == 92:
                index += 1
                if index >= len(value):
                    break
                escaped = value[index]
                escape_map = {110: 10, 114: 13, 116: 9, 98: 8, 102: 12, 40: 40, 41: 41, 92: 92}
                if 48 <= escaped <= 55:
                    digits = bytes([escaped])
                    while len(digits) < 3 and index + 1 < len(value) and 48 <= value[index + 1] <= 55:
                        index += 1
                        digits += bytes([value[index]])
                    decoded.append(int(digits, 8))
                elif escaped in (10, 13):
                    if escaped == 13 and index + 1 < len(value) and value[index + 1] == 10:
                        index += 1
                else:
                    decoded.append(escape_map.get(escaped, escaped))
            elif byte == 40:
                depth += 1
                decoded.append(byte)
            elif byte == 41:
                depth -= 1
                if depth:
                    decoded.append(byte)
            else:
                decoded.append(byte)
            index += 1
        results.append(_decode_pdf_text_bytes(bytes(decoded)))
    return results


def _decode_pdf_hex_string(value: bytes) -> str:
    """解码 PDF 十六进制字符串，支持 UTF-16 字节序标记。"""
    compact = re.sub(rb"\s+", b"", value)
    if len(compact) % 2:
        compact += b"0"
    try:
        return _decode_pdf_text_bytes(bytes.fromhex(compact.decode("ascii")))
    except (ValueError, UnicodeDecodeError):
        return ""


def _decode_pdf_text_bytes(value: bytes) -> str:
    """解码常见 PDF 文本字节，无需外部 PDF 引擎。"""
    if value.startswith((b"\xfe\xff", b"\xff\xfe")):
        encoding = "utf-16-be" if value.startswith(b"\xfe\xff") else "utf-16-le"
        return value[2:].decode(encoding, errors="replace").strip()
    try:
        return value.decode("utf-8").strip()
    except UnicodeDecodeError:
        return value.decode("latin-1", errors="replace").strip()


def _pdf_line_chunks(line: str, width: int = 92) -> list[str]:
    """按中文和西文字宽拆分 PDF 文本行。"""
    unicode_text = any(ord(char) > 127 for char in str(line))
    clean = str(line) if unicode_text else _pdf_ascii(line)
    if unicode_text:
        width = min(width, 44)
    if not clean:
        return [""]
    return [clean[index : index + width] for index in range(0, len(clean), width)]


def _pdf_ascii(value: str) -> str:
    """将文本转为简单 PDF 写入器支持的 ASCII 子集。"""
    text = html.unescape(str(value))
    return "".join(char if 32 <= ord(char) <= 126 else "?" for char in text)


def _pdf_escape(value: str) -> str:
    """转义 PDF 字面量字符串。"""
    return value.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")


def _pdf_stream(lines: list[str], unicode_text: bool = False) -> bytes:
    """根据已换行文本生成 ASCII 或 Unicode PDF 内容流。"""
    commands = ["BT", "/F1 11 Tf", "50 792 Td", "14 TL"]
    for index, line in enumerate(lines):
        if index:
            commands.append("T*")
        if unicode_text:
            commands.append(f"<{str(line).encode('utf-16-be').hex()}> Tj")
        else:
            commands.append(f"({_pdf_escape(line)}) Tj")
    commands.append("ET")
    return "\n".join(commands).encode("ascii", errors="replace")


def _write_pdf(objects: list[bytes], target: Path) -> None:
    """写入带交叉引用表的最小 PDF。"""
    offsets: list[int] = []
    content = bytearray(b"%PDF-1.4\n%\xe2\xe3\xcf\xd3\n")
    for index, obj in enumerate(objects, start=1):
        offsets.append(len(content))
        content.extend(f"{index} 0 obj\n".encode("ascii"))
        content.extend(obj)
        content.extend(b"\nendobj\n")
    xref_offset = len(content)
    content.extend(f"xref\n0 {len(objects) + 1}\n".encode("ascii"))
    content.extend(b"0000000000 65535 f \n")
    for offset in offsets:
        content.extend(f"{offset:010d} 00000 n \n".encode("ascii"))
    content.extend(f"trailer\n<< /Size {len(objects) + 1} /Root 1 0 R >>\nstartxref\n{xref_offset}\n%%EOF\n".encode("ascii"))
    target.write_bytes(bytes(content))


def _docx_paragraph(text: str, style: str, alignment: str = "", runs: list[dict] | None = None) -> str:
    """生成 Word 段落，文字片段对应时保留显式格式。"""
    style = style if style in {"Normal", "Heading1", "Heading2"} else "Normal"
    if not runs or "".join(run.get("text", "") for run in runs) != str(text):
        runs = [{"text": str(text)}]
    output = []
    for run in runs:
        properties = [f'<w:{key} w:val="{int(run[key])}"/>' for key in ("b", "i") if key in run]
        pieces = []
        for part in re.split(r"(\n|\t)", run.get("text", "").replace("\r\n", "\n").replace("\r", "\n")):
            if part == "\n":
                pieces.append("<w:br/>")
            elif part == "\t":
                pieces.append("<w:tab/>")
            elif part:
                pieces.append(f'<w:t xml:space="preserve">{html.escape(part)}</w:t>')
        if "size" in run:
            properties.append(f'<w:sz w:val="{max(1, (int(run["size"]) + 25) // 50)}"/>')
        fonts = " ".join(f'w:{attribute}="{html.escape(run[key], quote=True)}"' for key, attribute in (("font", "ascii"), ("font", "hAnsi"), ("east_asia", "eastAsia")) if run.get(key))
        if fonts:
            properties.append(f'<w:rFonts {fonts}/>')
        if run.get("color"):
            properties.append(f'<w:color w:val="{html.escape(run["color"], quote=True)}"/>')
        if "underline" in run:
            value = next((word for word, drawing in UNDERLINES.items() if drawing == run["underline"]), None)
            if value is None:
                raise ValueError("PPT 下划线类型无效")
            color = f' w:color="{html.escape(run["underline_color"], quote=True)}"' if run.get("underline_color") else ""
            properties.append(f'<w:u w:val="{value}"{color}/>')
        rpr = f'<w:rPr>{"".join(properties)}</w:rPr>' if properties else ""
        output.append(f'<w:r>{rpr}{"".join(pieces)}</w:r>')
    alignment = alignment or runs[0].get("alignment", "")
    alignment_xml = f'<w:jc w:val="{alignment}"/>' if alignment in PARAGRAPH_ALIGNMENTS else ""
    return f'<w:p><w:pPr>{alignment_xml}<w:pStyle w:val="{style}"/></w:pPr>{"".join(output)}</w:p>'


def _scaled_column_widths(columns: int, column_widths: list[int] | None, extent: int) -> list[int]:
    """按源列宽比例分配目标总宽度，累计取整避免总宽度漂移。"""
    weights = column_widths if column_widths and len(column_widths) == columns and all(width > 0 for width in column_widths) else [1] * columns
    total = sum(weights)
    boundaries = [0]
    for weight in weights:
        boundaries.append(boundaries[-1] + weight)
    return [extent * boundaries[index + 1] // total - extent * boundaries[index] // total for index in range(columns)]

def _docx_table(rows: list[list[Any]], column_widths: list[int] | None = None, row_heights: list[int] | None = None, cell_alignments: list[list[list[str]]] | None = None) -> str:
    """将行列文本写成带边框的原生 Word 表格。"""
    columns = max((len(row) for row in rows), default=0)
    if not columns:
        return ""
    widths = _scaled_column_widths(columns, column_widths, 9000)
    grid = "".join(f'<w:gridCol w:w="{width}"/>' for width in widths)
    borders = "".join(f'<w:{side} w:val="single" w:sz="4"/>' for side in ["top", "left", "bottom", "right", "insideH", "insideV"])
    body = []
    for row_index, row in enumerate(rows):
        cells = [*row, *([""] * (columns - len(row)))]
        rendered = []
        for column, cell in enumerate(cells):
            item = cell if isinstance(cell, dict) else {"text": cell}
            if item.get("skip"):
                continue
            span = int(item.get("span", 1))
            properties = f'<w:tcW w:w="{sum(widths[column:column + span])}" w:type="dxa"/>'
            if span > 1:
                properties += f'<w:gridSpan w:val="{span}"/>'
            if item.get("vertical"):
                properties += f'<w:vMerge w:val="{item["vertical"]}"/>'
            alignments = cell_alignments[row_index][column] if cell_alignments and row_index < len(cell_alignments) and column < len(cell_alignments[row_index]) else []
            text = str(item.get("text", ""))
            paragraphs = "".join(_docx_paragraph(line, "Normal", alignments[index] if index < len(alignments) else "") for index, line in enumerate(text.split("\n"))) if alignments else _docx_paragraph(text, "Normal")
            rendered.append(f'<w:tc><w:tcPr>{properties}</w:tcPr>{paragraphs}</w:tc>')
        height = row_heights[row_index] if row_heights and row_index < len(row_heights) else 0
        properties = f'<w:trPr><w:trHeight w:val="{height}" w:hRule="atLeast"/></w:trPr>' if height > 0 else ""
        body.append('<w:tr>' + properties + "".join(rendered) + '</w:tr>')
    return f'<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:tblBorders>{borders}</w:tblBorders></w:tblPr><w:tblGrid>{grid}</w:tblGrid>{"".join(body)}</w:tbl>'


def _docx_styles() -> str:
    """定义正文和两级标题的中文字体、字号与大纲级别。"""
    styles = []
    for name, size, level in [("Normal", 22, None), ("Heading1", 32, 0), ("Heading2", 26, 1)]:
        outline = f'<w:outlineLvl w:val="{level}"/><w:keepNext/>' if level is not None else ""
        bold = "<w:b/>" if level is not None else ""
        styles.append(f'<w:style w:type="paragraph" w:styleId="{name}"><w:name w:val="{name}"/><w:pPr>{outline}<w:spacing w:after="120"/></w:pPr><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Microsoft YaHei"/><w:sz w:val="{size}"/>{bold}</w:rPr></w:style>')
    return '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' + "".join(styles) + '</w:styles>'


def _docx_content_types() -> str:
    """返回单文档部件的 DOCX 内容类型声明。"""
    defaults = "".join(f'<Default Extension="{extension}" ContentType="{mime}"/>' for extension, mime in IMAGE_TYPES.items())
    return f"""<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>{defaults}<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>"""


def _package_rels(target: str) -> str:
    """返回指向 OOXML 主部件的包关系 XML。"""
    return f"""<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="{target}"/></Relationships>"""


def _pptx_content_types(slide_count: int, chart_indexes: list[int] | None = None) -> str:
    """根据幻灯片数量生成 PPTX 内容类型 XML。"""
    with zipfile.ZipFile(Path(__file__).with_name("pptx_template.pptx")) as template:
        template_types = ET.fromstring(template.read("[Content_Types].xml"))
        foundation_types = "".join(ET.tostring(item, encoding="unicode") for item in template_types
                                   if any(item.get("PartName", "").startswith(prefix) for prefix in ("/ppt/slideMasters/", "/ppt/slideLayouts/", "/ppt/theme/")))
    overrides = "".join(
        f'<Override PartName="/ppt/slides/slide{index}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
        for index in range(1, slide_count + 1)
    )
    overrides += "".join(f'<Override PartName="/ppt/charts/chart{index}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>' for index in chart_indexes or [])
    if chart_indexes:
        overrides += '<Default Extension="xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/>'
    overrides += "".join(f'<Default Extension="{extension}" ContentType="{mime}"/>' for extension, mime in IMAGE_TYPES.items())
    return f"""<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>{foundation_types}{overrides}</Types>"""


def _presentation_xml(slide_count: int) -> str:
    """返回含幻灯片标识的最小 PPTX 演示文稿 XML。"""
    slide_ids = "".join(f'<p:sldId id="{255 + index}" r:id="rId{index}"/>' for index in range(1, slide_count + 1))
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="master"/></p:sldMasterIdLst><p:sldIdLst>{slide_ids}</p:sldIdLst><p:sldSz cx="9144000" cy="6858000" type="screen4x3"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>"""


def _presentation_rels(slide_count: int) -> str:
    """返回所有生成幻灯片的演示文稿关系。"""
    rels = "".join(
        f'<Relationship Id="rId{index}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide{index}.xml"/>'
        for index in range(1, slide_count + 1)
    )
    return f"""<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="master" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>{rels}</Relationships>"""


def _pptx_text_shape(shape_id: int, name: str, lines: list[str], y: int, height: int, font_size: int, line_runs: list[list[dict]] | None = None) -> str:
    """生成具有明确位置、尺寸和字体的可编辑文本框。"""
    paragraphs = []
    for index, line in enumerate(lines):
        runs = line_runs[index] if line_runs and index < len(line_runs) else []
        if "".join(run.get("text", "") for run in runs) != str(line):
            runs = [{"text": str(line)}]
        content = []
        for run in runs:
            style = " ".join(f'{key}="{int(run[key])}"' for key in ("b", "i") if key in run)
            if "underline" in run:
                style += f' u="{html.escape(run["underline"], quote=True)}"'
            size = run.get("size", font_size)
            latin = html.escape(run.get("font", "Arial"), quote=True)
            east_asia = html.escape(run.get("east_asia", "Microsoft YaHei"), quote=True)
            color = f'<a:solidFill><a:srgbClr val="{html.escape(run["color"], quote=True)}"/></a:solidFill>' if run.get("color") else ""
            underline_color = f'<a:uFill><a:solidFill><a:srgbClr val="{html.escape(run["underline_color"], quote=True)}"/></a:solidFill></a:uFill>' if run.get("underline_color") else ""
            for line_index, piece in enumerate(run.get("text", "").split("\n")):
                if line_index:
                    content.append('<a:br/>')
                content.append(f'<a:r><a:rPr lang="zh-CN" sz="{size}" {style}>{color}{underline_color}<a:latin typeface="{latin}"/><a:ea typeface="{east_asia}"/></a:rPr><a:t>{html.escape(piece)}</a:t></a:r>')
        alignment = PARAGRAPH_ALIGNMENTS.get(runs[0].get("alignment")) if runs else None
        properties = f'<a:pPr algn="{alignment}"/>' if alignment else ""
        paragraphs.append(f'<a:p>{properties}{"".join(content)}<a:endParaRPr lang="zh-CN" sz="{font_size}"/></a:p>')
    paragraphs = "".join(paragraphs) or '<a:p><a:endParaRPr lang="zh-CN"/></a:p>'
    return f'<p:sp><p:nvSpPr><p:cNvPr id="{shape_id}" name="{name}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="457200" y="{y}"/><a:ext cx="8229600" cy="{height}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr wrap="square"><a:spAutoFit/></a:bodyPr><a:lstStyle/>{paragraphs}</p:txBody></p:sp>'


def _pptx_row_weights(rows: list[list[str]], widths: list[int], heights: list[int] | None, merges: list[list[int]] | None = None) -> list[int]:
    """按合并宽度估算文字高度，分摊至自动行并保留显式高度。"""
    explicit = [heights[index] if heights and index < len(heights) else 0 for index in range(len(rows))]
    weights = [height if height > 0 else 400 for height in explicit]
    requirements = []
    for index, row in enumerate(rows):
        for column, text in enumerate(row):
            height, span = 1, 1
            covered = False
            for top, left, merge_height, merge_span in merges or []:
                if top <= index < top + merge_height and left <= column < left + merge_span:
                    covered = index != top or column != left
                    height, span = merge_height, merge_span
                    break
            if covered:
                continue
            capacity = max(1, sum(widths[column:column + span]) // 114300)
            count = 0
            for line in text.split("\n"):
                units = sum(2 if unicodedata.east_asian_width(character) in {"W", "F"} else 1 for character in line)
                count += max(1, (units + capacity - 1) // capacity)
            requirements.append((index, height, 400 * count))
    # 先满足普通行，再按合并区已分配的高度补足缺口。
    for start, height, required in sorted(requirements, key=lambda item: item[1]):
        automatic = [index for index in range(start, start + height) if explicit[index] <= 0]
        deficit = required - sum(weights[start:start + height])
        if automatic and deficit > 0:
            quotient, remainder = divmod(deficit, len(automatic))
            for offset, index in enumerate(automatic):
                weights[index] += quotient + (offset < remainder)
    return weights


def _pptx_table_frame(rows: list[list[str]], merges: list[list[int]], column_widths: list[int] | None = None, row_heights: list[int] | None = None, cell_alignments: list[list[list[str]]] | None = None) -> str:
    """生成可编辑的原生 PPT 行列表格，补齐不规则行。"""
    columns = max(len(row) for row in rows)
    widths = _scaled_column_widths(columns, column_widths, 8229600)
    heights = _scaled_column_widths(len(rows), _pptx_row_weights(rows, widths, row_heights, merges), 4937760)
    grid = "".join(f'<a:gridCol w="{width}"/>' for width in widths)
    body = []
    for row_index, row in enumerate(rows):
        cells = []
        for column, text in enumerate([*row, *([""] * (columns - len(row)))]):
            attributes = ""
            for top, left, height, span in merges:
                if top <= row_index < top + height and left <= column < left + span:
                    attributes = ((f' rowSpan="{height}"' if row_index == top and height > 1 else "") +
                                  (f' gridSpan="{span}"' if column == left and span > 1 else "") +
                                  (' hMerge="1"' if column > left else "") + (' vMerge="1"' if row_index > top else ""))
                    break
            alignments = cell_alignments[row_index][column] if cell_alignments and row_index < len(cell_alignments) and column < len(cell_alignments[row_index]) else []
            paragraphs = []
            for index, line in enumerate(text.split("\n")):
                alignment = PARAGRAPH_ALIGNMENTS.get(alignments[index] if index < len(alignments) else "")
                properties = f'<a:pPr algn="{alignment}"/>' if alignment else ""
                paragraphs.append(f'<a:p>{properties}<a:r><a:rPr lang="zh-CN" sz="1800"/><a:t>{html.escape(line)}</a:t></a:r></a:p>')
            paragraphs = "".join(paragraphs)
            cells.append(f'<a:tc{attributes}><a:txBody><a:bodyPr/><a:lstStyle/>{paragraphs}</a:txBody><a:tcPr/></a:tc>')
        body.append(f'<a:tr h="{heights[row_index]}">{"".join(cells)}</a:tr>')
    return f'<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="4" name="Table"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="457200" y="1463040"/><a:ext cx="8229600" cy="4937760"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr firstRow="1" bandRow="1"/><a:tblGrid>{grid}</a:tblGrid>{"".join(body)}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>'


def _slide_xml(title: str, body: list[str], title_runs: list[dict] | None = None, body_runs: list[list[dict]] | None = None) -> str:
    """生成含标题和正文的最小 PPTX 幻灯片。"""
    title_shape = _pptx_text_shape(2, "Title", [title], 365760, 914400, 2800, [title_runs or []])
    body_shape = _pptx_text_shape(3, "Content", body, 1463040, 4937760, 1800, body_runs)
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>{title_shape}{body_shape}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>"""
