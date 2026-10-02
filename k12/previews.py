"""构建受管文件的结构化预览。展示内容、OCR 需求和本地客户端要求，不编辑用户文档，也不暴露敏感本地路径。"""

from __future__ import annotations

import html
import re
import zipfile
from pathlib import Path
from typing import Any

from .converters import extract_docx_blocks, extract_pptx_slides, extract_xlsx_sheets


def build_file_preview(file: dict[str, Any]) -> dict[str, Any]:
    """为已保存文件生成安全预览载荷。"""
    path = Path(file.get("storage_path") or file.get("file_path") or "")
    preview = {
        "file_id": file.get("id", ""),
        "file_name": file.get("file_name", ""),
        "file_type": file.get("file_type", ""),
        "status": "ready",
        "mode": "structured",
        "pages": [],
        "objects": _base_objects(file),
        "warnings": list(file.get("validation_errors") or []),
        "metrics": {
            "page_count": file.get("page_count", 0),
            "slide_count": file.get("slide_count", 0),
            "sheet_count": file.get("sheet_count", 0),
            "file_size": file.get("file_size", 0),
        },
    }
    if not path.exists() or not path.is_file():
        preview["status"] = "metadata_only"
        preview["mode"] = "metadata"
        preview["warnings"].append("未找到本地源文件，当前仅展示元数据预览")
        preview["pages"] = _metadata_pages(file)
        return _attach_preview_markers(preview, file)
    try:
        if file.get("file_type") == "Word":
            _word_preview(path, file, preview)
        elif file.get("file_type") == "PPT":
            _ppt_preview(path, file, preview)
        elif file.get("file_type") == "Excel":
            _excel_preview(path, file, preview)
        elif file.get("file_type") == "PDF":
            _pdf_preview(path, file, preview)
        elif file.get("file_type") == "图片":
            _image_preview(path, file, preview)
        elif file.get("file_type") == "压缩包":
            _zip_preview(path, file, preview)
        else:
            preview["status"] = "unsupported"
            preview["warnings"].append("当前文件类型暂不支持正文预览")
            preview["pages"] = _metadata_pages(file)
    except (OSError, KeyError, zipfile.BadZipFile, ValueError) as exc:
        preview["status"] = "error"
        preview["warnings"].append(f"预览生成失败：{exc}")
        preview["pages"] = _metadata_pages(file)
    return _attach_preview_markers(preview, file)


def _word_preview(path: Path, file: dict[str, Any], preview: dict[str, Any]) -> None:
    """填充 OOXML Word 的预览页与对象计数。"""
    if path.suffix.lower() not in {".docx", ".docm", ".dotx", ".dotm"}:
        preview["status"] = "local_required"
        preview["mode"] = "local-client"
        preview["warnings"].append("旧版 Word 格式需要本地 Office 客户端生成预览")
        preview["pages"] = _metadata_pages(file)
        return
    blocks = extract_docx_blocks(path)
    preview["pages"] = [
        {
            "index": index,
            "title": block["text"][:48] if block.get("level") == 1 else f"段落 {index}",
            "kind": "标题" if block.get("level") == 1 else "正文",
            "text": block["text"][:600],
        }
        for index, block in enumerate(blocks[:30], start=1)
    ]
    summary = file.get("content_summary") or {}
    _add_count_object(preview, "段落", summary.get("paragraphs", len(blocks)), "可预览")
    _add_count_object(preview, "表格", summary.get("tables", 0), "已检测")
    _add_count_object(preview, "图片", summary.get("images", 0), "已检测")
    _add_count_object(preview, "OMML 公式", summary.get("ommlFormulas", 0), "需预检" if file.get("has_omml") else "无")
    _add_count_object(preview, "页眉页脚", int(summary.get("headers", 0) or 0) + int(summary.get("footers", 0) or 0), "可选保留")
    _add_count_object(preview, "脚注尾注", int(summary.get("footnotes", 0) or 0) + int(summary.get("endnotes", 0) or 0), "可转备注")
    _add_count_object(preview, "批注", summary.get("comments", 0), "可选保留")
    _add_count_object(preview, "修订", summary.get("revisions", 0), "需确认")
    if int(summary.get("revisions", 0) or 0) > 0:
        preview["warnings"].append("检测到修订标记，转换前建议确认保留或接受修订")


def _ppt_preview(path: Path, file: dict[str, Any], preview: dict[str, Any]) -> None:
    """填充 OOXML 演示文稿的预览页与对象计数。"""
    if path.suffix.lower() not in {".pptx", ".pptm"}:
        preview["status"] = "local_required"
        preview["mode"] = "local-client"
        preview["warnings"].append("旧版 PPT 格式需要本地 Office 客户端生成预览")
        preview["pages"] = _metadata_pages(file)
        return
    slides = extract_pptx_slides(path)
    preview["pages"] = [
        {
            "index": slide.get("index", index),
            "title": slide.get("title", f"幻灯片 {index}")[:64],
            "kind": "幻灯片",
            "text": "\n".join(slide.get("body") or [])[:600] or "无正文",
        }
        for index, slide in enumerate(slides[:40], start=1)
    ]
    summary = file.get("content_summary") or {}
    _add_count_object(preview, "幻灯片", summary.get("slides", len(slides)), "可预览")
    _add_count_object(preview, "备注", summary.get("notes", 0), "已检测")
    _add_count_object(preview, "图片", summary.get("images", 0), "已检测")
    _add_count_object(preview, "表格", summary.get("tables", 0), "已检测")
    _add_count_object(preview, "公式", summary.get("formulas", sum(int(slide.get("formula_count", 0) or 0) for slide in slides)), "可识别")
    _add_count_object(preview, "图表", summary.get("charts", sum(int(slide.get("chart_count", 0) or 0) for slide in slides)), "可提取")
    _add_count_object(preview, "文本框", summary.get("textBoxes", sum(int(slide.get("text_box_count", 0) or 0) for slide in slides)), "已提取")
    _add_count_object(preview, "形状", summary.get("shapes", sum(int(slide.get("shape_count", 0) or 0) for slide in slides)), "已检测")
    _add_count_object(preview, "母版版式", int(summary.get("masters", 0) or 0) + int(summary.get("layouts", 0) or 0), "可选解析")


def _excel_preview(path: Path, file: dict[str, Any], preview: dict[str, Any]) -> None:
    """填充 OOXML 工作簿的预览页与对象计数。"""
    if path.suffix.lower() not in {".xlsx", ".xlsm"}:
        preview["status"] = "local_required"
        preview["mode"] = "local-client"
        preview["warnings"].append("旧版 Excel 格式需要本地 Office 客户端生成预览")
        preview["pages"] = _metadata_pages(file)
        return
    sheets = extract_xlsx_sheets(path)
    preview["pages"] = [
        {
            "index": sheet.get("index", index),
            "title": sheet.get("name", f"Sheet{index}"),
            "kind": "工作表",
            "text": _sheet_preview_text(sheet),
        }
        for index, sheet in enumerate(sheets[:24], start=1)
    ]
    _add_count_object(preview, "工作表", len(sheets), "可预览")
    _add_count_object(preview, "单元格", sum(len(sheet.get("cells", [])) for sheet in sheets), "已提取")
    _add_count_object(preview, "公式", sum(int(sheet.get("formula_count", 0) or 0) for sheet in sheets), "已提取")
    _add_count_object(preview, "合并单元格", sum(int(sheet.get("merged_cells", 0) or 0) for sheet in sheets), "已检测")
    _add_count_object(preview, "图表", sum(int(sheet.get("chart_count", 0) or 0) for sheet in sheets), "可转图表页")
    _add_count_object(preview, "图片", sum(int(sheet.get("image_count", 0) or 0) for sheet in sheets), "可提取")
    _add_count_object(preview, "批注", sum(int(sheet.get("comment_count", 0) or 0) for sheet in sheets), "可选提取")


def _pdf_preview(path: Path, file: dict[str, Any], preview: dict[str, Any]) -> None:
    """填充以元数据为主的 PDF 预览及 OCR、Mathpix 线索。"""
    data = path.read_bytes()[:3_000_000]
    text = data.decode("latin-1", errors="ignore")
    snippets = _pdf_text_snippets(text)
    image_count = int((file.get("content_summary") or {}).get("imageObjects", text.count("/Subtype /Image")) or 0)
    pdf_type = (file.get("content_summary") or {}).get("pdfType") or ("混合型 PDF" if snippets and image_count else ("文本型 PDF" if snippets else "扫描型 PDF"))
    page_count = max(1, int(file.get("page_count") or text.count("/Type /Page") - text.count("/Type /Pages") or 1))
    preview["metrics"]["pdf_type"] = pdf_type
    preview["pages"] = [
        {
            "index": index,
            "title": f"第 {index} 页",
            "kind": pdf_type,
            "text": snippets[index - 1] if index - 1 < len(snippets) else "等待 OCR 或 Mathpix 识别",
        }
        for index in range(1, min(page_count, 20) + 1)
    ]
    _add_count_object(preview, "PDF 页面", page_count, "已检测")
    _add_count_object(preview, "图片对象", image_count, "可进入 OCR")
    _add_count_object(preview, "表格线索", (file.get("content_summary") or {}).get("tableHints", 0), "建议表格 OCR")
    if file.get("has_formula"):
        _add_count_object(preview, "疑似公式", 1, "建议 Mathpix 识别")
    recommendation = (file.get("content_summary") or {}).get("ocrRecommendation")
    if recommendation and recommendation != "可解析文本层":
        preview["warnings"].append(f"PDF 识别建议：{recommendation}")
    if file.get("encrypted"):
        preview["warnings"].append("PDF 已加密，需要密码后才能完整预览")


def _image_preview(path: Path, file: dict[str, Any], preview: dict[str, Any]) -> None:
    """填充图片预览元数据，不解码完整像素内容。"""
    dimensions = _image_dimensions(path.read_bytes()[:2_000_000])
    text = "图片尺寸未识别"
    if dimensions:
        width, height, image_type = dimensions
        text = f"{width} x {height} px / {image_type}"
        preview["metrics"].update({"width": width, "height": height, "image_type": image_type})
    preview["pages"] = [{"index": 1, "title": file.get("file_name", "图片"), "kind": "图片", "text": text}]
    _add_count_object(preview, "图片", 1, "可预览元数据")


def _zip_preview(path: Path, file: dict[str, Any], preview: dict[str, Any]) -> None:
    """根据安全 ZIP 条目元数据填充压缩包预览。"""
    with zipfile.ZipFile(path) as archive:
        entries = [info for info in archive.infolist() if not info.is_dir()]
    preview["pages"] = [
        {
            "index": index,
            "title": Path(info.filename).name or info.filename,
            "kind": "压缩包条目",
            "text": f"{info.filename} / {info.file_size} bytes",
        }
        for index, info in enumerate(entries[:40], start=1)
    ]
    _add_count_object(preview, "压缩包条目", len(entries), "已扫描")
    _add_count_object(preview, "支持条目", int((file.get("content_summary") or {}).get("supportedEntries", 0) or 0), "可导入")


def _base_objects(file: dict[str, Any]) -> list[dict[str, Any]]:
    """根据文件能力标记生成预览对象徽标。"""
    objects: list[dict[str, Any]] = []
    flags = [
        ("公式", file.get("has_formula"), "需预检"),
        ("MathType", file.get("has_mathtype"), "需本地客户端"),
        ("OMML", file.get("has_omml"), "缺少依赖" if file.get("missing_omml_dependency") else "可检索依赖"),
        ("OMML 依赖", file.get("missing_omml_dependency"), "缺少"),
        ("宏", file.get("has_macro"), "需授权执行"),
        ("图片", file.get("has_image"), "可预览"),
        ("微小图片", file.get("has_small_image"), "可生成图片报告"),
    ]
    for label, value, status in flags:
        if value:
            objects.append({"type": label, "label": label, "count": 1, "status": status})
    return objects


def _add_count_object(preview: dict[str, Any], label: str, count: Any, status: str) -> None:
    """数量为正时附加计数预览对象。"""
    amount = int(count or 0)
    if amount <= 0:
        return
    preview["objects"].append({"type": label, "label": label, "count": amount, "status": status})


def _metadata_pages(file: dict[str, Any]) -> list[dict[str, Any]]:
    """仅用已存元数据生成兜底预览页。"""
    summary = file.get("content_summary") or {}
    details = [f"{key}: {value}" for key, value in list(summary.items())[:12]]
    if not details:
        details = ["暂无可预览正文，已展示文件元数据"]
    return [{"index": 1, "title": file.get("file_name", "文件"), "kind": "元数据", "text": "\n".join(details)}]


def _attach_preview_markers(preview: dict[str, Any], file: dict[str, Any]) -> dict[str, Any]:
    """为预览页附加公式、OMML、宏、图片及置信度标记。"""
    markers = _preview_markers(file)
    if not markers:
        return preview
    pages = preview.get("pages") or []
    if not pages:
        pages = _metadata_pages(file)
        preview["pages"] = pages
    for index, page in enumerate(pages):
        text = f"{page.get('title', '')} {page.get('kind', '')} {page.get('text', '')}".lower()
        matched = [marker for marker in markers if str(marker.get("term", "")).lower() in text]
        if index == 0:
            merged = [*matched, *[marker for marker in markers if marker not in matched]]
        else:
            merged = matched
        if merged:
            existing = list(page.get("markers") or [])
            page["markers"] = [*existing, *merged]
    return preview


def _preview_markers(file: dict[str, Any]) -> list[dict[str, str]]:
    """返回解释敏感预览发现的标记标签。"""
    markers: list[dict[str, str]] = []
    marker_specs = [
        ("公式", file.get("has_formula"), "warn", "公式高亮", "检测到公式对象，建议进入公式预检"),
        ("OMML", file.get("has_omml"), "warn", "OMML 公式标记", "检测到 Word 自带公式，需确认是否转 MathType"),
        ("MathType", file.get("has_mathtype"), "blue", "MathType 公式标记", "检测到 MathType 或嵌入公式对象"),
        ("宏", file.get("has_macro"), "bad", "宏执行结果标记", "检测到宏，执行前需确认风险和备份"),
        ("微小图片", file.get("has_small_image"), "blue", "微小图片高亮", "检测到微小图片，可进入图片检索报告"),
        ("低置信度", _has_low_confidence_hint(file), "warn", "低置信度标记", "存在需人工确认的低置信度线索"),
    ]
    for term, enabled, level, label, message in marker_specs:
        if enabled:
            markers.append({"term": term, "level": level, "label": label, "message": message})
    return markers


def _has_low_confidence_hint(file: dict[str, Any]) -> bool:
    """判断预览是否应显示低置信度警告。"""
    summary = file.get("content_summary") or {}
    if file.get("validation_errors"):
        return True
    return any(int(summary.get(key, 0) or 0) > 0 for key in ("lowConfidenceFormulas", "formulaWarnings", "uncertainObjects"))


def _sheet_preview_text(sheet: dict[str, Any]) -> str:
    """将工作表前几个单元格格式化为简短预览文本。"""
    cells = sheet.get("cells", [])
    if not cells:
        return "未提取到可显示单元格"
    return "\n".join(f"{cell.get('ref', '')}: {cell.get('value', '')}" for cell in cells[:16])


def _pdf_text_snippets(text: str) -> list[str]:
    """从简单 PDF 文本操作符提取有大小限制的文本片段。"""
    chunks = []
    for pattern in [r"\(([^()]{3,160})\)\s*Tj", r"\(([^()]{3,160})\)\s*'"]:
        for value in re.findall(pattern, text):
            clean = _clean_preview_text(value)
            if clean:
                chunks.append(clean)
    return chunks[:20]


def _clean_preview_text(value: str) -> str:
    """规范从 XML、PDF 或压缩包元数据提取的预览文本。"""
    return html.unescape(re.sub(r"\s+", " ", value)).strip()


def _image_dimensions(data: bytes) -> tuple[int, int, str] | None:
    """从常见图片头读取宽高。"""
    if data.startswith(b"\x89PNG\r\n\x1a\n") and len(data) >= 24:
        return int.from_bytes(data[16:20], "big"), int.from_bytes(data[20:24], "big"), "png"
    if data.startswith(b"GIF87a") or data.startswith(b"GIF89a"):
        return int.from_bytes(data[6:8], "little"), int.from_bytes(data[8:10], "little"), "gif"
    if data.startswith(b"BM") and len(data) >= 26:
        return int.from_bytes(data[18:22], "little"), int.from_bytes(data[22:26], "little"), "bmp"
    if data.startswith(b"\xff\xd8"):
        return _jpeg_dimensions(data)
    return None


def _jpeg_dimensions(data: bytes) -> tuple[int, int, str] | None:
    """从 SOF 标记读取 JPEG 宽高。"""
    index = 2
    while index + 9 < len(data):
        if data[index] != 0xFF:
            index += 1
            continue
        marker = data[index + 1]
        if marker in {0xC0, 0xC1, 0xC2, 0xC3, 0xC5, 0xC6, 0xC7, 0xC9, 0xCA, 0xCB, 0xCD, 0xCE, 0xCF}:
            height = int.from_bytes(data[index + 5 : index + 7], "big")
            width = int.from_bytes(data[index + 7 : index + 9], "big")
            return width, height, "jpg"
        segment_length = int.from_bytes(data[index + 2 : index + 4], "big")
        if segment_length < 2:
            return None
        index += 2 + segment_length
    return None
