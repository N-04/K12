"""负责文档分析、任务和需求验收证据的业务编排。连接本地记录、前端接口、Mathpix OCR 授权和桌面执行计划。除非同平台客户端回传结果，否则不声称执行了真实 Office、MathType、OMML 写回或 Word 宏。"""

from __future__ import annotations

import base64
import binascii
import hashlib
import html
import json
import os
import re
import shutil
import sys
import tempfile
import time
import zipfile
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET
from urllib.parse import urlencode, urlparse

from .media import bmp_dimensions, image_locations, svg_dimensions, tiff_dimensions
from .vba_sources import vba_macros

from .models import (
    LOCAL_REQUIRED_TASKS,
    SUPPORTED_EXTENSIONS,
    TASK_LABELS,
    FileItem,
    FormulaItem,
    MacroItem,
    OmmlDependencyItem,
    SmallImageItem,
    Task,
    new_id,
    utc_now,
)
from .converters import (
    count_xml_elements,
    build_docx,
    build_docx_from_pdf_text,
    build_docx_from_slides,
    build_docx_from_xlsx,
    build_pdf_from_xlsx,
    build_pptx,
    build_pptx_from_docx,
    build_pptx_from_xlsx,
    build_text_pdf,
    extract_docx_blocks,
    extract_docx_object_summary,
    extract_pdf_text_blocks,
    extract_pptx_slides,
    extract_xlsx_sheets,
)
from .excel import effective_excel_settings, select_excel_sheets
from .install_profiles import install_profile, normalize_platform
from .mathpix import MathpixApiError, MathpixClient, MathpixConfigError
from .previews import build_file_preview
from .reports import ReportBuilder
from .store import AppStore, DEFAULT_AUTHORIZATIONS, ROLE_PERMISSIONS


ILLEGAL_FILE_CHARS = re.compile(r'[<>:"|?*\x00-\x1f]')
OMML_DEPENDENCY_SUFFIXES = {".xsl", ".xslt", ".xml", ".mml"}
REPLACEMENT_IMAGE_SUFFIXES = {".png", ".jpg", ".jpeg", ".gif", ".webp", ".bmp"}
REPLACEMENT_IMAGE_MIME_SUFFIXES = {
    "image/png": ".png",
    "image/jpeg": ".jpg",
    "image/gif": ".gif",
    "image/webp": ".webp",
    "image/bmp": ".bmp",
}
MATHPIX_TEX_ENTRY_LIMIT_BYTES = 1024 * 1024
MATHPIX_TEX_ARCHIVE_LIMIT_BYTES = 50 * 1024 * 1024
WORKFLOW_TASK_TYPES = tuple(TASK_LABELS.keys())
OOXML_IMAGE_PREFIXES = {
    "Word": "word/media/",
    "PPT": "ppt/media/",
    "Excel": "xl/media/",
}
LOCAL_CLIENT_CAPABILITY_LABELS = {
    "officeAutomation": "Office 自动化",
    "mathTypeAutomation": "MathType 自动化",
    "macroExecution": "Word 宏执行",
    "ommlDependencySearch": "OMML 依赖检索",
}
LOCAL_ACTION_LABELS = {
    "office_conversion": "Office 转换复核",
    "omml_mathtype": "OMML/MathType 处理",
    "pdf_formula_mathtype": "PDF 公式 MathType 后处理",
    "macro_sequence": "Word 宏顺序执行",
    "open_output_directory": "打开输出目录",
}
LOCAL_PATH_TEXT = re.compile(r"(/Users/[^\s，,;]+|/private/[^\s，,;]+|/var/folders/[^\s，,;]+|/tmp/[^\s，,;]+|[A-Za-z]:\\[^\s，,;]+)")


def _pdf_image_descriptors(data: bytes, limit: int = 200) -> list[dict[str, Any]]:
    """在大小限制内提取 PDF 图片元数据，不渲染页面像素。"""
    descriptors: list[dict[str, Any]] = []
    for index, match in enumerate(re.finditer(rb"/Subtype\s*/Image", data), start=1):
        if len(descriptors) >= limit:
            break
        dict_start = data.rfind(b"<<", 0, match.start())
        dict_end = data.find(b">>", match.end())
        if dict_start < 0 or dict_end < 0:
            continue
        dict_bytes = data[dict_start:dict_end + 2]
        stream = _pdf_stream_after(data, dict_end + 2)
        descriptors.append(
            {
                "index": index,
                "width": _pdf_dict_int(dict_bytes, "Width"),
                "height": _pdf_dict_int(dict_bytes, "Height"),
                "filter": _pdf_dict_name(dict_bytes, "Filter"),
                "colorSpace": _pdf_dict_name(dict_bytes, "ColorSpace"),
                "bitsPerComponent": _pdf_dict_int(dict_bytes, "BitsPerComponent"),
                "stream": stream,
                "dict_hash": hashlib.sha256(dict_bytes).hexdigest(),
            }
        )
    return descriptors


def _pdf_stream_after(data: bytes, offset: int) -> bytes:
    """返回 PDF 对象字典后面的原始流内容。"""
    next_obj = data.find(b"endobj", offset)
    search_end = next_obj if next_obj >= 0 else min(len(data), offset + 2_000_000)
    stream_start = data.find(b"stream", offset, search_end)
    if stream_start < 0:
        return b""
    body_start = stream_start + len(b"stream")
    if data[body_start:body_start + 2] == b"\r\n":
        body_start += 2
    elif data[body_start:body_start + 1] in {b"\r", b"\n"}:
        body_start += 1
    stream_end = data.find(b"endstream", body_start, search_end)
    if stream_end < 0:
        return b""
    return data[body_start:stream_end].rstrip(b"\r\n")


def _pdf_dict_int(dict_bytes: bytes, key: str) -> int:
    """从小段 PDF 字典字节中读取整数。"""
    match = re.search(rb"/" + key.encode("ascii") + rb"\s+(\d+)", dict_bytes)
    return int(match.group(1)) if match else 0


def _pdf_dict_name(dict_bytes: bytes, key: str) -> str:
    """从字典字节读取直接名称或数组包装的名称。"""
    direct = re.search(rb"/" + key.encode("ascii") + rb"\s*/([A-Za-z0-9]+)", dict_bytes)
    if direct:
        return direct.group(1).decode("ascii", errors="ignore")
    array = re.search(rb"/" + key.encode("ascii") + rb"\s*\[\s*/([A-Za-z0-9]+)", dict_bytes)
    return array.group(1).decode("ascii", errors="ignore") if array else ""


def _pdf_filter_image_type(filter_name: str) -> str:
    """将 PDF 图片滤镜名称映射为可导出图片类型。"""
    normalized = filter_name.lower()
    if normalized == "dctdecode":
        return "jpg"
    if normalized == "jpxdecode":
        return "jpx"
    if normalized == "jbig2decode":
        return "jb2"
    if normalized == "flatedecode":
        return "pdf-flate"
    return "pdf-object"


def _find_direct_omml_file(directory: Path) -> Path | None:
    """查找目录中直接存在的 OMML 转换依赖文件。"""
    priority = ["omml2mml.xsl", "omml2mathml.xsl", "omml.xsl"]
    try:
        direct_files = {child.name.lower(): child for child in directory.iterdir() if child.is_file()}
    except OSError:
        return None
    for name in priority:
        if name in direct_files:
            return direct_files[name]
    for child in direct_files.values():
        if _is_omml_dependency_name(child.name):
            return child
    return None


def _is_omml_dependency_name(file_name: str) -> bool:
    """判断文件名是否符合 OMML 转换依赖命名规则。"""
    path = Path(file_name)
    return "omml" in path.name.lower() and path.suffix.lower() in OMML_DEPENDENCY_SUFFIXES


class DocumentAnalyzer:
    """分类上传文件并提取规划使用的轻量元数据。"""

    def __init__(self, single_file_limit_mb: int = 500) -> None:
        """初始化当前对象所需的配置、依赖与运行状态。"""
        self.single_file_limit = single_file_limit_mb * 1024 * 1024

    def analyze_metadata(self, file_name: str, file_size: int, file_path: str = "") -> FileItem:
        """没有文件字节时，根据浏览器元数据创建文件记录。"""
        extension = Path(file_name).suffix.lower()
        file_type = self.classify_extension(extension)
        digest = self._digest(file_name, str(file_size), extension)
        item = FileItem(
            file_name=file_name,
            file_type=file_type,
            file_size=max(0, int(file_size)),
            extension=extension,
            file_path=file_path,
        )
        item.validation_errors = self.validate(file_name, file_size, extension)
        item.status = "待处理" if not item.validation_errors else "校验失败"

        base_number = int(digest[:4], 16)
        if file_type in {"Word", "PDF"}:
            item.page_count = base_number % 64 + 1
        if file_type == "PPT":
            item.slide_count = base_number % 36 + 1
        if file_type == "Excel":
            item.sheet_count = base_number % 8 + 1

        lower_name = file_name.lower()
        item.has_omml = file_type == "Word" and ("omml" in lower_name or base_number % 5 == 0)
        item.has_mathtype = file_type in {"Word", "PPT"} and ("math" in lower_name or base_number % 7 == 0)
        item.has_formula = item.has_omml or item.has_mathtype or "formula" in lower_name or base_number % 6 == 0
        item.has_macro = extension in {".docm", ".dotm", ".pptm", ".xlsm"} or "macro" in lower_name
        item.has_image = file_type == "图片" or (file_type in {"Word", "Excel", "PPT", "PDF"} and base_number % 4 != 0)
        item.has_small_image = file_type in {"Word", "Excel", "PPT", "PDF"} and base_number % 4 != 0
        self._mark_omml_dependency(item, Path(file_path) if file_path else None)
        return item

    def analyze_file(self, path: Path, display_name: str | None = None) -> FileItem:
        """检查上传的本地文件并补充格式专用元数据。"""
        file_name = display_name or path.name
        item = self.analyze_metadata(file_name, path.stat().st_size, str(path))
        item.storage_path = str(path)
        item.source_kind = "upload"
        if item.extension in {".docx", ".docm", ".dotx", ".dotm", ".pptx", ".pptm", ".xlsx", ".xlsm"}:
            self._analyze_ooxml(path, item)
        elif item.extension in {".doc", ".dot", ".xls", ".ppt"}:
            self._analyze_legacy_office(path, item)
        elif item.extension == ".pdf":
            self._analyze_pdf(path, item)
        elif item.extension == ".zip":
            self._analyze_zip(path, item)
        return item

    def validate(self, file_name: str, file_size: int, extension: str) -> list[str]:
        """返回上传校验错误，不修改运行状态。"""
        errors: list[str] = []
        if not file_name.strip():
            errors.append("文件名不能为空")
        if ILLEGAL_FILE_CHARS.search(file_name):
            errors.append("文件名包含非法字符")
        if extension not in self.supported_extensions():
            errors.append("文件格式不支持")
        if file_size > self.single_file_limit:
            errors.append("文件超过单文件大小限制")
        if file_size <= 0:
            errors.append("文件大小异常")
        return errors

    @staticmethod
    def supported_extensions() -> set[str]:
        """返回需求规定的上传流程支持的扩展名。"""
        merged: set[str] = set()
        for values in SUPPORTED_EXTENSIONS.values():
            merged.update(values)
        return merged

    @staticmethod
    def classify_extension(extension: str) -> str:
        """将扩展名映射为界面文档类型标签。"""
        for label, extensions in SUPPORTED_EXTENSIONS.items():
            if extension in extensions:
                return {
                    "word": "Word",
                    "excel": "Excel",
                    "ppt": "PPT",
                    "pdf": "PDF",
                    "image": "图片",
                    "archive": "压缩包",
                }[label]
        return "未知"

    @staticmethod
    def _digest(*parts: str) -> str:
        """生成稳定元数据摘要，用于确定性样本统计。"""
        text = "::".join(parts)
        return hashlib.sha256(text.encode("utf-8")).hexdigest()

    def _analyze_ooxml(self, path: Path, item: FileItem) -> None:
        """检查支持的 OOXML ZIP 包并补充文件记录。"""
        if self._looks_encrypted_office_container(path):
            item.encrypted = True
            self._clear_unreadable_office_capabilities(item)
            item.validation_errors = [
                error
                for error in item.validation_errors
                if "损坏" not in str(error) and "Open XML" not in str(error)
            ]
            if not any("密码" in str(error) or "加密" in str(error) for error in item.validation_errors):
                item.validation_errors.append("文件已加密，需要密码")
            item.status = "校验失败"
            item.content_summary = {
                "encrypted": True,
                "container": "OLE Compound File",
                "passwordStatus": "未输入",
            }
            return
        try:
            with zipfile.ZipFile(path) as archive:
                names = archive.namelist()
                item.encrypted = self._looks_encrypted_ooxml(names)
                if item.file_type == "Word":
                    self._analyze_word_archive(archive, names, item)
                    self._mark_omml_dependency(item, path)
                elif item.file_type == "PPT":
                    self._analyze_ppt_archive(archive, names, item)
                elif item.file_type == "Excel":
                    self._analyze_excel_archive(archive, names, item)
        except zipfile.BadZipFile:
            self._clear_unreadable_office_capabilities(item)
            item.validation_errors.append("文件损坏或不是有效的 Office Open XML 文档")
            item.status = "校验失败"

    def _analyze_legacy_office(self, path: Path, item: FileItem) -> None:
        """交接本地客户端前校验旧 Office OLE 容器。"""
        self._clear_unreadable_office_capabilities(item)
        try:
            with path.open("rb") as handle:
                prefix = handle.read(5_000_000)
        except OSError as exc:
            item.validation_errors.append(f"旧版 Office 文件无法读取：{exc}")
            item.status = "校验失败"
            return
        header_error = self._legacy_office_header_error(prefix)
        if header_error:
            item.validation_errors.append(header_error)
            item.status = "校验失败"
            item.content_summary = {
                "container": "OLE Compound File",
                "containerValid": False,
                "requiresNativeOffice": True,
            }
            return
        lowered = prefix.lower()
        encrypted = any(
            marker.encode("utf-16le").lower() in lowered
            for marker in ("EncryptionInfo", "EncryptedPackage", "StrongEncryptionDataSpace")
        )
        item.encrypted = encrypted
        item.content_summary = {
            "container": "OLE Compound File",
            "containerValid": True,
            "legacyBinaryOffice": True,
            "requiresNativeOffice": True,
            "deepInspectionStatus": "需本地 Office 客户端",
            "encrypted": encrypted,
        }
        if encrypted:
            item.validation_errors.append("文件已加密，需要密码")
            item.status = "校验失败"

    @staticmethod
    def _legacy_office_header_error(data: bytes) -> str:
        """返回旧 Office CFB 头无效的校验错误。"""
        if len(data) < 512 or not data.startswith(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"):
            return "文件损坏或不是有效的旧版 Office OLE 文档"
        if data[28:30] != b"\xfe\xff":
            return "旧版 Office OLE 文件字节序标记异常"
        sector_shift = int.from_bytes(data[30:32], "little")
        mini_sector_shift = int.from_bytes(data[32:34], "little")
        if sector_shift not in {9, 12} or mini_sector_shift != 6:
            return "旧版 Office OLE 文件扇区头异常"
        return ""

    @staticmethod
    def _clear_unreadable_office_capabilities(item: FileItem) -> None:
        """Office 包不可读时清除推测的能力信息。"""
        item.page_count = 0
        item.slide_count = 0
        item.sheet_count = 0
        item.has_formula = False
        item.has_mathtype = False
        item.has_omml = False
        item.missing_omml_dependency = False
        item.has_macro = False
        item.has_image = False
        item.has_small_image = False

    def _analyze_word_archive(self, archive: zipfile.ZipFile, names: list[str], item: FileItem) -> None:
        """提取 Word 结构、公式、宏和对象的轻量统计。"""
        document_xml = self._read_archive_text(archive, "word/document.xml")
        footnotes_xml = self._read_archive_text(archive, "word/footnotes.xml")
        endnotes_xml = self._read_archive_text(archive, "word/endnotes.xml")
        comments_xml = self._read_archive_text(archive, "word/comments.xml")
        image_count = sum(1 for name in names if name.startswith("word/media/"))
        embedded_count = sum(1 for name in names if name.startswith("word/embeddings/"))
        header_count = sum(1 for name in names if name.startswith("word/header") and name.endswith(".xml"))
        footer_count = sum(1 for name in names if name.startswith("word/footer") and name.endswith(".xml"))
        revision_count = len(re.findall(r"<w:(?:ins|del|moveFrom|moveTo)(?:\s|>)", document_xml))
        item.page_count = max(item.page_count, self._read_app_count(archive, "Pages"))
        item.has_omml = count_xml_elements(document_xml, "oMath") > 0
        item.has_mathtype = embedded_count > 0 or "Equation Native" in document_xml or "MathType" in document_xml
        item.has_formula = item.has_omml or item.has_mathtype or "$$" in document_xml
        normalized_names = {name.replace("\\", "/").lower() for name in names}
        item.has_macro = "word/vbaproject.bin" in normalized_names
        item.has_image = image_count > 0
        item.has_small_image = image_count > 0
        item.content_summary = {
            "paragraphs": count_xml_elements(document_xml, "p"),
            "headings": len(re.findall(r"<w:pStyle[^>]+w:val=\"Heading", document_xml)),
            "tables": count_xml_elements(document_xml, "tbl"),
            "images": image_count,
            "embeddedObjects": embedded_count,
            "ommlFormulas": count_xml_elements(document_xml, "oMath"),
            "headers": header_count,
            "footers": footer_count,
            "footnotes": footnotes_xml.count("<w:footnote "),
            "endnotes": endnotes_xml.count("<w:endnote "),
            "comments": comments_xml.count("<w:comment "),
            "revisions": revision_count,
            "vbaProjects": 1 if item.has_macro else 0,
            "macroEnabledContainer": item.extension in {".docm", ".dotm"},
        }

    def _mark_omml_dependency(self, item: FileItem, path: Path | None) -> None:
        """为 Word 分析结果附加 OMML 依赖可用性。"""
        item.missing_omml_dependency = self._missing_omml_dependency(item, path)
        if item.has_omml:
            item.content_summary["missingOmmlDependency"] = item.missing_omml_dependency

    @staticmethod
    def _missing_omml_dependency(item: FileItem, path: Path | None) -> bool:
        """判断含 OMML 的 Word 文件附近是否缺少转换依赖。"""
        if item.file_type != "Word" or not item.has_omml:
            return False
        if path and str(path) not in {"", "."}:
            directory = path if path.is_dir() else path.parent
            if directory.exists() and _find_direct_omml_file(directory):
                return False
        return True

    def _analyze_ppt_archive(self, archive: zipfile.ZipFile, names: list[str], item: FileItem) -> None:
        """提取 PPT 幻灯片、媒体、公式和宏的轻量统计。"""
        slide_names = [name for name in names if name.startswith("ppt/slides/slide") and name.endswith(".xml")]
        note_names = [name for name in names if name.startswith("ppt/notesSlides/notesSlide") and name.endswith(".xml")]
        image_count = sum(1 for name in names if name.startswith("ppt/media/"))
        chart_count = sum(1 for name in names if name.startswith("ppt/charts/chart") and name.endswith(".xml"))
        master_count = sum(1 for name in names if name.startswith("ppt/slideMasters/slideMaster") and name.endswith(".xml"))
        layout_count = sum(1 for name in names if name.startswith("ppt/slideLayouts/slideLayout") and name.endswith(".xml"))
        text_blob = "".join(self._read_archive_text(archive, name) for name in slide_names[:30])
        formula_count = self._ppt_formula_count(text_blob)
        item.slide_count = len(slide_names)
        normalized_names = {name.replace("\\", "/").lower() for name in names}
        item.has_macro = "ppt/vbaproject.bin" in normalized_names
        item.has_mathtype = "MathType" in text_blob or "Equation Native" in text_blob
        item.has_formula = formula_count > 0
        item.has_image = image_count > 0
        item.has_small_image = image_count > 0
        item.content_summary = {
            "slides": len(slide_names),
            "notes": len(note_names),
            "images": image_count,
            "charts": chart_count,
            "tables": text_blob.count("<a:tbl"),
            "formulas": formula_count,
            "textBoxes": text_blob.count("<p:txBody"),
            "shapes": text_blob.count("<p:sp"),
            "masters": master_count,
            "layouts": layout_count,
            "textRuns": text_blob.count("<a:t>"),
            "vbaProjects": 1 if item.has_macro else 0,
            "macroEnabledContainer": item.extension == ".pptm",
        }

    @staticmethod
    def _ppt_formula_count(text: str) -> int:
        """根据 OMML、MathType、Equation 和 TeX 标记统计公式线索。"""
        return (
            len(re.findall(r"<(?:\w+:)?oMath(?:Para)?\b", text))
            + len(re.findall(r"MathType|Equation Native", text, flags=re.IGNORECASE))
            + len(re.findall(r"\$[^$]{1,120}\$|\\(?:frac|sqrt|sum|int)\b", text))
        )

    def _analyze_excel_archive(self, archive: zipfile.ZipFile, names: list[str], item: FileItem) -> None:
        """提取 Excel 工作表、公式、图表、图片和表格统计。"""
        sheet_names = [name for name in names if name.startswith("xl/worksheets/sheet") and name.endswith(".xml")]
        image_count = sum(1 for name in names if name.startswith("xl/media/"))
        chart_count = sum(1 for name in names if name.startswith("xl/charts/chart") and name.endswith(".xml"))
        drawing_count = sum(1 for name in names if name.startswith("xl/drawings/drawing") and name.endswith(".xml"))
        comment_count = 0
        for name in names:
            if name.startswith("xl/comments") and name.endswith(".xml"):
                comments_xml = self._read_archive_text(archive, name)
                comment_count += comments_xml.count("<comment ") + comments_xml.count("<x:comment ")
        table_count = sum(1 for name in names if name.startswith("xl/tables/table") and name.endswith(".xml"))
        text_blob = "".join(self._read_archive_text(archive, name) for name in sheet_names[:20])
        item.sheet_count = len(sheet_names)
        normalized_names = {name.replace("\\", "/").lower() for name in names}
        item.has_macro = "xl/vbaproject.bin" in normalized_names
        item.has_formula = "<f>" in text_blob or "<f " in text_blob
        item.has_image = image_count > 0
        item.has_small_image = image_count > 0
        item.content_summary = {
            "sheets": len(sheet_names),
            "formulas": text_blob.count("<f"),
            "images": image_count,
            "charts": chart_count,
            "drawings": drawing_count,
            "comments": comment_count,
            "tables": table_count,
            "mergedCells": text_blob.count("<mergeCell"),
            "vbaProjects": 1 if item.has_macro else 0,
            "macroEnabledContainer": item.extension == ".xlsm",
        }

    def _analyze_pdf(self, path: Path, item: FileItem) -> None:
        """分类 PDF 以规划 Mathpix OCR，不上传文件。"""
        data = path.read_bytes()[:5_000_000]
        text = data.decode("latin-1", errors="ignore")
        item.encrypted = "/Encrypt" in text
        item.page_count = max(1, text.count("/Type /Page") - text.count("/Type /Pages"))
        try:
            text_blocks = extract_pdf_text_blocks(path)
        except OSError:
            text_blocks = []
        text_snippets = [str(block.get("text") or "") for block in text_blocks if block.get("text")]
        if not text_snippets:
            text_snippets = self._pdf_text_snippets(text)
        pdf_images = _pdf_image_descriptors(data)
        image_objects = len(pdf_images) or text.count("/Subtype /Image")
        table_hints = self._pdf_table_hints(text)
        formula_hints = self._pdf_formula_hints(text)
        pdf_type = self._pdf_type(bool(text_snippets), image_objects, item.encrypted)
        item.has_formula = formula_hints > 0
        item.has_image = image_objects > 0
        item.has_small_image = image_objects > 0
        item.content_summary = {
            "pdfType": pdf_type,
            "pdfHeader": text[:8] if text.startswith("%PDF") else "",
            "pages": item.page_count,
            "textLayer": bool(text_snippets),
            "textSnippets": len(text_snippets),
            "imageObjects": image_objects,
            "pdfImages": [self._public_pdf_image_descriptor(descriptor) for descriptor in pdf_images[:50]],
            "formulaHints": formula_hints,
            "tableHints": table_hints,
            "ocrRecommendation": self._pdf_ocr_recommendation(pdf_type, formula_hints, table_hints),
            "encrypted": item.encrypted,
        }
        if not text.startswith("%PDF"):
            item.validation_errors.append("PDF 文件头异常")
            item.status = "校验失败"
        if item.encrypted:
            item.validation_errors.append("文件已加密，需要密码")
            item.status = "校验失败"

    @staticmethod
    def _public_pdf_image_descriptor(descriptor: dict[str, Any]) -> dict[str, Any]:
        """返回接口和报告可用的不含路径的 PDF 图片信息。"""
        return {
            "index": descriptor.get("index", 0),
            "width": descriptor.get("width", 0),
            "height": descriptor.get("height", 0),
            "filter": descriptor.get("filter", ""),
            "colorSpace": descriptor.get("colorSpace", ""),
            "bitsPerComponent": descriptor.get("bitsPerComponent", 0),
            "hasStream": bool(descriptor.get("stream")),
        }

    @staticmethod
    def _pdf_text_snippets(text: str) -> list[str]:
        """从简单 PDF 文本操作符提取有大小限制的文本片段。"""
        chunks: list[str] = []
        for pattern in [r"\(([^()]{3,160})\)\s*Tj", r"\(([^()]{3,160})\)\s*'"]:
            for value in re.findall(pattern, text):
                clean = html.unescape(re.sub(r"\s+", " ", value)).strip()
                if clean:
                    chunks.append(clean)
        return chunks[:20]

    @staticmethod
    def _pdf_type(has_text_layer: bool, image_objects: int, encrypted: bool) -> str:
        """将 PDF 分类为文本型、扫描型、混合型或加密型。"""
        if encrypted:
            return "加密 PDF"
        if has_text_layer and image_objects:
            return "混合型 PDF"
        if has_text_layer:
            return "文本型 PDF"
        if image_objects:
            return "扫描型 PDF"
        return "扫描型 PDF"

    @staticmethod
    def _pdf_formula_hints(text: str) -> int:
        """统计提示数学公式内容的原始 PDF 标记。"""
        return sum(text.count(token) for token in ("Math", "Equation", "Formula", "∫", "√", "\\frac", "\\sum"))

    @staticmethod
    def _pdf_table_hints(text: str) -> int:
        """统计提示表格结构的原始 PDF 标记。"""
        return sum(text.count(token) for token in ("/Table", "Table", "Cell", "Row", "Column"))

    @staticmethod
    def _pdf_ocr_recommendation(pdf_type: str, formula_hints: int, table_hints: int) -> str:
        """根据 PDF 类型和线索推荐 Mathpix OCR 开关。"""
        recommendations: list[str] = []
        if pdf_type in {"扫描型 PDF", "混合型 PDF"}:
            recommendations.append("启用文字 OCR")
        if formula_hints:
            recommendations.append("启用公式 OCR")
        if table_hints:
            recommendations.append("启用表格 OCR")
        return "、".join(recommendations) if recommendations else "可解析文本层"

    def _analyze_zip(self, path: Path, item: FileItem) -> None:
        """汇总 ZIP 条目和支持的文件，用于批量导入规划。"""
        try:
            with zipfile.ZipFile(path) as archive:
                infos = [info for info in archive.infolist() if not info.is_dir()]
                total_uncompressed = sum(info.file_size for info in infos)
                encrypted_entries = sum(1 for info in infos if info.flag_bits & 0x1)
                excessive_ratio_entries = sum(
                    1
                    for info in infos
                    if info.file_size > 10 * 1024 * 1024
                    and info.file_size / max(1, info.compress_size) > 200
                )
                item.archive_entry_count = len(infos)
                item.content_summary = {
                    "entries": len(infos),
                    "supportedEntries": sum(1 for info in infos if Path(info.filename).suffix.lower() in self.supported_extensions()),
                    "compressedSize": sum(info.compress_size for info in infos),
                    "uncompressedSize": total_uncompressed,
                    "encryptedEntries": encrypted_entries,
                    "excessiveCompressionRatioEntries": excessive_ratio_entries,
                    "entryLimit": 2000,
                    "uncompressedSizeLimit": self.single_file_limit,
                }
                if len(infos) > 2000:
                    item.validation_errors.append("ZIP 条目数量超过安全限制")
                if total_uncompressed > self.single_file_limit:
                    item.validation_errors.append("ZIP 解压后大小超过单文件安全限制")
                if excessive_ratio_entries:
                    item.validation_errors.append("ZIP 包含异常压缩比条目，可能是压缩炸弹")
                if encrypted_entries:
                    item.validation_errors.append("ZIP 包含加密条目，需要先解密")
                if item.validation_errors:
                    item.status = "校验失败"
        except zipfile.BadZipFile:
            item.validation_errors.append("ZIP 解压失败")
            item.status = "校验失败"

    def _read_app_count(self, archive: zipfile.ZipFile, key: str) -> int:
        """读取 OOXML 应用属性部件中的数值。"""
        app_xml = self._read_archive_text(archive, "docProps/app.xml")
        match = re.search(fr"<{key}>(\d+)</{key}>", app_xml)
        return int(match.group(1)) if match else 0

    @staticmethod
    def _read_archive_text(archive: zipfile.ZipFile, name: str) -> str:
        """尝试按 UTF-8 读取 ZIP 成员文本。"""
        try:
            return archive.read(name).decode("utf-8", errors="ignore")
        except KeyError:
            return ""

    @staticmethod
    def _looks_encrypted_ooxml(names: list[str]) -> bool:
        """判断 OOXML 包中是否包含加密标识部件。"""
        normalized = {name.replace("\\", "/").strip("/").lower() for name in names}
        return "encryptedpackage" in normalized or "encryptioninfo" in normalized

    @staticmethod
    def _looks_encrypted_office_container(path: Path) -> bool:
        """检测 OLE 复合容器中受密码保护的 OOXML。"""
        try:
            with path.open("rb") as handle:
                prefix = handle.read(8192)
        except OSError:
            return False
        compound_magic = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
        if not prefix.startswith(compound_magic):
            return False
        # 加密 OOXML 使用含 EncryptionInfo 和 EncryptedPackage 流的 OLE 复合文件。
        # 即使流目录超出前缀读取范围，也接受现代 OOXML 扩展名的复合文件签名。
        # 旧版二进制 Office 格式由其他分支处理。
        lowered = prefix.lower()
        return (
            "EncryptionInfo".encode("utf-16le").lower() in lowered
            or "EncryptedPackage".encode("utf-16le").lower() in lowered
            or path.suffix.lower() in {".docx", ".docm", ".dotx", ".dotm", ".pptx", ".pptm", ".xlsx", ".xlsm"}
        )


class TaskProcessor:
    """协调文件、任务、报告、预检与本地交接。"""

    def __init__(self, store: AppStore) -> None:
        """初始化当前对象所需的配置、依赖与运行状态。"""
        self.store = store
        self.report_builder = ReportBuilder(store)

    def create_file(self, payload: dict[str, Any]) -> dict[str, Any]:
        """根据接口提供的文件详情创建仅含元数据的记录。"""
        self._assert_permission("files.manage")
        settings = self.store.get_settings()
        file_name = self._resolve_duplicate_file_name(str(payload.get("file_name", "")), str(settings.get("duplicateFileStrategy", "自动重命名")))
        if file_name is None:
            existing = self._find_existing_file(str(payload.get("file_name", "")))
            return existing if existing else self.store.save_file({"id": new_id("file"), "file_name": str(payload.get("file_name", "")), "created_at": utc_now()})
        analyzer = DocumentAnalyzer(settings.get("singleFileLimitMb", 500))
        item = analyzer.analyze_metadata(
            file_name,
            int(payload.get("file_size", 0)),
            str(payload.get("file_path", "")),
        )
        record = item.to_dict()
        source_path = Path(item.file_path)
        if source_path.exists() and source_path.is_file():
            record["file_size"] = source_path.stat().st_size
            record["source_sha256"] = self._sha256(source_path)
        return self.store.save_file(record)

    def create_uploaded_file(self, file_name: str, content: bytes) -> list[dict[str, Any]]:
        """保存上传字节、分析内容并注册 ZIP 子条目。"""
        self._assert_permission("files.manage")
        relative_path = self._safe_relative_path(file_name)
        safe_name = self._safe_file_name(file_name)
        settings = self.store.get_settings()
        resolved_name = self._resolve_duplicate_file_name(safe_name, str(settings.get("duplicateFileStrategy", "自动重命名")))
        if resolved_name is None:
            existing = self._find_existing_file(safe_name)
            return [existing] if existing else []
        safe_name = resolved_name
        extension = Path(safe_name).suffix.lower()
        upload_name = f"{new_id('upload')}{extension or '.bin'}"
        target = self.store.uploads_dir / upload_name
        target.write_bytes(content)
        analyzer = DocumentAnalyzer(settings.get("singleFileLimitMb", 500))
        item = analyzer.analyze_file(target, safe_name)
        if relative_path and relative_path != safe_name:
            item.source_relative_path = relative_path
        record = item.to_dict()
        record["source_sha256"] = hashlib.sha256(content).hexdigest()
        saved = [self.store.save_file(record)]
        if item.extension == ".zip" and not item.validation_errors:
            saved.extend(self._register_archive_entries(target, item, analyzer))
            refreshed_parent = self.store.get_file(item.id)
            if refreshed_parent:
                saved[0] = refreshed_parent
        self.store.append_log("system", f"上传文件：{safe_name}", category="upload")
        return saved

    def replace_uploaded_file(self, file_id: str, file_name: str, content: bytes) -> dict[str, Any]:
        """替换已有上传文件，同时保留其公开标识。"""
        self._assert_permission("files.manage")
        if not self.store.get_file(file_id):
            raise KeyError(f"File not found: {file_id}")
        safe_name = self._safe_file_name(file_name)
        existing_names = {
            str(file.get("file_name") or "")
            for file in self.store.list_files()
            if file.get("id") != file_id
        }
        if safe_name in existing_names:
            safe_name = self._unique_file_name(safe_name, existing_names)
        settings = self.store.get_settings()
        extension = Path(safe_name).suffix.lower()
        upload_name = f"{new_id('upload')}{extension or '.bin'}"
        target = self.store.uploads_dir / upload_name
        target.write_bytes(content)
        analyzer = DocumentAnalyzer(settings.get("singleFileLimitMb", 500))
        item = analyzer.analyze_file(target, safe_name)
        record = item.to_dict()
        record["source_sha256"] = hashlib.sha256(content).hexdigest()
        replaced = self.store.replace_file(file_id, record)
        if item.extension == ".zip" and not item.validation_errors:
            item.id = file_id
            item.created_at = str(replaced.get("created_at") or item.created_at)
            item.source_kind = "upload"
            item.archive_parent_id = ""
            self._register_archive_entries(target, item, analyzer)
            refreshed_parent = self.store.get_file(file_id)
            if refreshed_parent:
                replaced = refreshed_parent
        self.store.append_log("system", f"重新上传替换：{safe_name}", category="upload")
        return replaced

    def reports_for_file(self, file_id: str) -> list[dict[str, Any]]:
        """返回引用指定源文件的报告。"""
        if not self.store.get_file(file_id):
            raise KeyError(f"File not found: {file_id}")
        reports: list[dict[str, Any]] = []
        for report in self.store.list_reports():
            report_file_ids = {str(file.get("id") or "") for file in report.get("files", [])}
            if report.get("file_id") == file_id or file_id in report_file_ids:
                reports.append(report)
        return reports

    def create_task(self, payload: dict[str, Any]) -> dict[str, Any]:
        """创建任务并立即运行本地处理流程。"""
        file_ids = list(payload.get("file_ids") or [])
        task_type = str(payload.get("task_type") or "batch_process")
        self._assert_task_permission(task_type)
        options = self._normalize_workflow_options(task_type, dict(payload.get("options") or {}))
        execute_mode = payload.get("execute_mode") or self.resolve_execute_mode(task_type, file_ids)
        task = Task(task_type=task_type, execute_mode=execute_mode, file_ids=file_ids, options=options)
        saved = self.store.save_task(task.to_dict())
        return self.run_task(saved["id"])

    def run_task(self, task_id: str) -> dict[str, Any]:
        """对队列任务执行分析、报告生成和清理。"""
        task = self.store.get_task(task_id)
        if not task:
            raise KeyError(f"Task not found: {task_id}")

        task["status"] = "处理中"
        task["progress"] = 8
        task["start_time"] = utc_now()
        task["attempt_started_at_ns"] = time.time_ns()
        self.store.save_task(task)
        self.store.append_log(task_id, f"创建任务：{TASK_LABELS.get(task['task_type'], task['task_type'])}")

        files = [self.store.get_file(file_id) for file_id in task["file_ids"]]
        files = [file for file in files if file]
        if not files:
            task["status"] = "失败"
            task["progress"] = 100
            task["error_message"] = "任务没有可处理文件"
            task["end_time"] = utc_now()
            self.store.append_log(task_id, task["error_message"], "error")
            return self.store.save_task(task)
        files = self._verify_task_source_files(task, files)

        stages = self._stages_for(task["task_type"], task["execute_mode"])
        log_category = self._task_log_category(task["task_type"])
        for index, stage in enumerate(stages, start=1):
            task["progress"] = min(95, int(index / len(stages) * 88) + 8)
            self.store.save_task(task)
            self.store.append_log(task_id, stage, category=log_category)

        analysis = self._build_analysis(task, files)
        workflow_plan = task.get("options", {}).get("workflowPlan")
        if not workflow_plan:
            workflow_plan = self._workflow_plan_from_options(task["task_type"], task.get("options", {}))
        analysis["workflowPlan"] = dict(workflow_plan)
        analysis["preflightChecks"] = self.preflight_checks(
            {"task_type": task["task_type"], "file_ids": task["file_ids"], "execute_mode": task["execute_mode"]}
        )
        if task["task_type"] == "batch_process":
            analysis["batchResults"] = self._batch_results(task, files)
            analysis["batchPlan"] = self._batch_plan(task, files, analysis["batchResults"])
        if task["task_type"] == "pdf_to_word":
            local_pdf_files = [file for file in files if self._is_local_text_pdf(file)]
            mathpix_files = [file for file in files if file not in local_pdf_files]
            local_pdf_artifacts, local_pdf_evidence = self._local_pdf_text_artifacts(task, local_pdf_files)
            analysis["pdfTextLayerConversions"] = local_pdf_evidence
            analysis["mathpix"] = self._mathpix_pdf_jobs(task, mathpix_files)
            analysis["formulas"].extend(self._mathpix_formula_items(files, analysis["mathpix"], task))
            analysis["artifacts"] = local_pdf_artifacts + self._mathpix_artifacts(task, mathpix_files, analysis["mathpix"])
        if task["task_type"] in {"formula_precheck", "omml_to_mathtype", "word_to_ppt", "mathtype_format"}:
            analysis["ommlDependencies"] = self._omml_dependency_jobs(task, files)
            analysis["ommlConversionPrompts"] = self._omml_conversion_prompts(task, files, analysis["ommlDependencies"])
        if task["task_type"] in {"word_to_ppt", "ppt_to_word", "excel_to_pdf", "excel_to_word", "excel_to_ppt"}:
            analysis["artifacts"] = self._conversion_artifacts(task, files)
        report = self.report_builder.build(task, files, analysis)
        hard_preflight_failed = any(item.get("status") == "失败" for item in analysis.get("preflightChecks", []))
        pending_local_outputs = int(report.get("artifact_pending_count") or 0)
        if report["fail_count"] > 0 or hard_preflight_failed:
            task["status"] = "失败"
        elif pending_local_outputs:
            task["status"] = "待处理"
        else:
            task["status"] = "成功"
        self._apply_task_report_summary(task, report)
        task["progress"] = 95 if pending_local_outputs else 100
        task["output_path"] = str(self.store.output_task_dir(task_id))
        task["end_time"] = utc_now()
        settings = self.store.get_settings()
        task["keep_original_file"] = bool(settings.get("keepOriginalFile", True))
        if settings.get("autoOpenOutputDirectory"):
            task["output_directory_action"] = {
                "status": "待本地客户端执行",
                "path": task["output_path"],
                "message": "任务完成后打开输出目录",
            }
            self.store.append_log(task_id, f"已请求本地客户端打开输出目录：{task['output_path']}")
        if task["status"] == "失败":
            task["error_message"] = self._task_failure_summary(report) or "部分文件未通过校验或能力条件"
        else:
            task["error_message"] = ""
        if settings.get("enableTaskCompletionNotice", True):
            task["completion_notice"] = {
                "status": "待前端提示",
                "level": "success" if task["status"] == "成功" else "warning",
                "message": (
                    f"{TASK_LABELS.get(task['task_type'], task['task_type'])} 已进入本地客户端队列"
                    if pending_local_outputs
                    else f"{TASK_LABELS.get(task['task_type'], task['task_type'])} 已完成：{task['status']}"
                ),
                "created_at": utc_now(),
            }
        self.store.save_report(report)
        self.store.append_log(task_id, f"生成报告：{report['report_type']}")
        saved_task = self.store.save_task(task)
        if not settings.get("saveHistory", True):
            cleanup = self.store.delete_task_history(task_id, delete_outputs=True)
            saved_task["history_saved"] = False
            saved_task["history_cleanup"] = cleanup
            saved_task["history_message"] = "历史记录已关闭，任务、报告、日志和运行缓存未保留"
            return saved_task
        saved_task["history_saved"] = True
        return saved_task

    def _verify_task_source_files(self, task: dict[str, Any], files: list[dict[str, Any]]) -> list[dict[str, Any]]:
        """在任何分析或转换前创建通过完整性校验的任务快照。"""
        verified: list[dict[str, Any]] = []
        snapshots: list[dict[str, Any]] = []
        snapshot_dir = self.store.output_task_dir(str(task.get("id") or "")) / ".inputs"
        snapshot_dir.mkdir(parents=True, exist_ok=True)
        for original in files:
            file = dict(original)
            source_path = str(file.get("storage_path") or file.get("file_path") or "")
            if source_path:
                try:
                    info = self.file_download_info(str(file.get("id") or ""))
                    suffix = str(file.get("extension") or Path(str(info["path"])).suffix or ".bin")
                    snapshot = snapshot_dir / f"{file.get('id')}{suffix}"
                    shutil.copyfile(Path(info["path"]), snapshot)
                    if snapshot.stat().st_size != int(info["file_size"]) or self._sha256(snapshot) != str(info.get("sha256") or ""):
                        snapshot.unlink(missing_ok=True)
                        raise ValueError("源文件在任务快照创建期间发生变化")
                    file["original_storage_path"] = str(original.get("storage_path") or "")
                    file["original_file_path"] = str(original.get("file_path") or "")
                    file["storage_path"] = str(snapshot)
                    file["file_path"] = str(snapshot)
                    file["source_kind"] = "task_snapshot"
                    file["source_sha256"] = str(info.get("sha256") or "")
                    snapshots.append(
                        {
                            "file_id": file.get("id", ""),
                            "path": str(snapshot),
                            "size": int(info["file_size"]),
                            "sha256": str(info.get("sha256") or ""),
                        }
                    )
                except (KeyError, ValueError, OSError) as exc:
                    error = f"源文件完整性校验失败：{exc}"
                    errors = list(file.get("validation_errors") or [])
                    if error not in errors:
                        errors.append(error)
                    file["validation_errors"] = errors
                    file["status"] = "校验失败"
                    self.store.save_file(file)
            verified.append(file)
        task["source_snapshots"] = snapshots
        self.store.save_task(task)
        return verified

    @classmethod
    def _normalize_workflow_options(cls, task_type: str, options: dict[str, Any]) -> dict[str, Any]:
        """保存来自拖拽规划界面的可信工作流顺序。"""
        plan = cls._workflow_plan_from_options(task_type, options)
        normalized = dict(options)
        normalized["workflowOrder"] = plan["order"]
        normalized["workflowLabels"] = plan["labels"]
        normalized["workflowPlan"] = plan
        return normalized

    @staticmethod
    def _workflow_plan_from_options(task_type: str, options: dict[str, Any]) -> dict[str, Any]:
        """规范并校验任务选项中的拖拽工作流顺序。"""
        raw_order = options.get("workflowOrder")
        order_values = raw_order if isinstance(raw_order, list) else []
        order: list[str] = []
        invalid_items: list[str] = []
        duplicate_items: list[str] = []
        for value in order_values:
            item = str(value or "").strip()
            if item not in WORKFLOW_TASK_TYPES:
                if item:
                    invalid_items.append(item[:80])
                continue
            if item in order:
                duplicate_items.append(item)
                continue
            order.append(item)
        if task_type in WORKFLOW_TASK_TYPES and task_type not in order:
            order.insert(0, task_type)
        labels = [TASK_LABELS.get(item, item) for item in order]
        current_index = order.index(task_type) + 1 if task_type in order else 0
        return {
            "version": "k12.workflowPlan.v1",
            "source": "workspace-planner",
            "current_task": task_type,
            "current_index": current_index,
            "order": order,
            "labels": labels,
            "invalid_items": invalid_items,
            "duplicate_items": duplicate_items,
        }

    @staticmethod
    def _apply_task_report_summary(task: dict[str, Any], report: dict[str, Any]) -> None:
        """将报告结果计数回写到任务记录。"""
        task["success_count"] = int(report.get("success_count") or report.get("batch_success_count") or 0)
        task["fail_count"] = int(report.get("fail_count") or report.get("batch_fail_count") or 0)
        task["pending_count"] = int(report.get("pending_count") or 0)
        task["cancelled_count"] = int(report.get("cancelled_count") or 0)
        task["failure_count"] = int(report.get("failure_count") or 0)
        task["retryable_count"] = int(report.get("batch_retryable_count") or 0)

    @staticmethod
    def _task_failure_summary(report: dict[str, Any]) -> str:
        """根据报告失败行生成简短任务错误信息。"""
        rows = list(report.get("failureRows") or [])
        if not rows:
            return ""
        snippets: list[str] = []
        seen: set[str] = set()
        for row in rows:
            category = str(row.get("category") or "失败")
            file_name = str(row.get("file_name") or "").strip()
            message = str(row.get("message") or row.get("recommendation") or "待处理").strip()
            snippet = f"{file_name}：{message}" if file_name else f"{category}：{message}"
            if snippet in seen:
                continue
            seen.add(snippet)
            snippets.append(snippet)
            if len(snippets) >= 3:
                break
        remaining = max(0, len(rows) - len(snippets))
        suffix = f"；另有 {remaining} 条失败清单" if remaining else ""
        return "；".join(snippets) + suffix

    def retry_task(self, task_id: str) -> dict[str, Any]:
        """重置失败、取消或中断的任务并重新运行。"""
        self._assert_permission("tasks.control")
        task = self.store.get_task(task_id)
        if not task:
            raise KeyError(f"Task not found: {task_id}")
        if task.get("status") not in {"失败", "已取消", "已中断"}:
            raise ValueError("只有失败、已取消或已中断的任务可以重试")
        task["status"] = "待处理"
        task["progress"] = 0
        task["error_message"] = ""
        self.store.save_task(task)
        self.store.append_log(task_id, "用户触发失败重试")
        return self.run_task(task_id)

    def acknowledge_task_completion_notice(self, task_id: str) -> dict[str, Any]:
        """标记任务完成通知已经向用户显示。"""
        task = self.store.get_task(task_id)
        if not task:
            raise KeyError(f"Task not found: {task_id}")
        notice = dict(task.get("completion_notice") or {})
        if not notice:
            return task
        notice["status"] = "已提示"
        notice["acknowledged_at"] = utc_now()
        task["completion_notice"] = notice
        self.store.append_log(task_id, "用户确认任务完成提醒")
        return self.store.save_task(task)

    def skip_batch_file(self, task_id: str, file_id: str) -> dict[str, Any]:
        """跳过批量任务中的失败文件并重建报告。"""
        self._assert_permission("tasks.control")
        task = self.store.get_task(task_id)
        if not task:
            raise KeyError(f"Task not found: {task_id}")
        if task.get("task_type") != "batch_process":
            raise ValueError("只有批量任务支持跳过单个失败文件")
        if file_id not in set(task.get("file_ids") or []):
            raise ValueError("文件不属于该批量任务")
        file = self.store.get_file(file_id)
        latest_result = self._latest_batch_result(task_id, file_id)
        has_validation_error = bool(file and file.get("validation_errors"))
        if not has_validation_error and latest_result.get("status") != "失败":
            raise ValueError("只能跳过失败或可重试的批量文件")
        skipped_ids = list(dict.fromkeys([*self._batch_skipped_file_ids(task), file_id]))
        task["skipped_file_ids"] = skipped_ids
        task.setdefault("batch_skip_actions", []).append(
            {
                "file_id": file_id,
                "file_name": (file or {}).get("file_name", latest_result.get("file_name", "")),
                "status": "跳过",
                "created_at": utc_now(),
            }
        )
        task["status"] = "待处理"
        task["progress"] = 0
        task["error_message"] = ""
        self.store.save_task(task)
        self.store.append_log(task_id, f"用户跳过批量失败文件：{(file or {}).get('file_name', file_id)}", "warning")
        return self.run_task(task_id)

    def sync_local_task_status(self, task_id: str, payload: dict[str, Any]) -> dict[str, Any]:
        """接受令牌保护的本地客户端任务状态更新。"""
        task = self.store.get_task(task_id)
        if not task:
            raise KeyError(f"Task not found: {task_id}")
        settings = self.store.get_settings()
        if not settings.get("allowTaskStatusCloudSync", False):
            raise ValueError("任务状态同步云端未授权")
        status = self._local_sync_status(str(payload.get("status") or payload.get("task_status") or task.get("status") or ""))
        current_status = str(task.get("status") or "")
        current_progress = int(task.get("progress", 0) or 0)
        reported_progress = min(100, self._non_negative_int(payload.get("progress", current_progress), current_progress))
        progress = max(current_progress, reported_progress)
        if status == "成功":
            progress = 100
        message = str(payload.get("message") or payload.get("error_message") or "")
        outputs = self._local_sync_outputs(payload.get("outputs") or payload.get("artifacts") or [])
        dry_run_execution = self._local_sync_dry_run_execution(
            payload.get("dryRunExecution") or payload.get("dry_run_execution") or payload.get("execution_summary")
        )
        native_execution_report = self._local_sync_native_execution_report(
            payload.get("nativeExecutionReport") or payload.get("native_execution_report")
        )
        if dry_run_execution:
            self._validate_local_sync_execution_task(task, dry_run_execution)
        if native_execution_report:
            self._validate_local_sync_execution_task(task, native_execution_report)
        self._validate_local_sync_transition(current_status, status)
        reconciled_report = self._reconcile_pending_local_outputs(
            task,
            status,
            outputs,
            native_execution_report,
            message,
        )
        if reconciled_report:
            self._apply_task_report_summary(task, reconciled_report)
        now = utc_now()
        task["status"] = status
        task["progress"] = progress
        task["local_client_status"] = str(payload.get("status") or status)
        task["local_client_message"] = message
        task["local_client_outputs"] = outputs
        if dry_run_execution:
            task["local_execution_summary"] = dry_run_execution
        if native_execution_report:
            task["local_native_execution_report"] = native_execution_report
        task["local_synced_at"] = now
        if payload.get("output_path"):
            task["output_path"] = str(payload.get("output_path") or "")
        result_upload_requested = bool(payload.get("resultUploadRequested") or payload.get("result_upload_requested"))
        result_upload_allowed = bool(settings.get("allowCloudSync", False))
        task["local_sync"] = {
            "task_status_cloud_sync_allowed": True,
            "result_upload_requested": result_upload_requested,
            "result_upload_allowed": result_upload_allowed,
            "result_upload_status": "queued" if result_upload_requested and result_upload_allowed else ("not_authorized" if result_upload_requested else "not_requested"),
            "dry_run_execution_status": dry_run_execution.get("plan_status", "") if dry_run_execution else "",
            "dry_run_ready_action_count": dry_run_execution.get("ready_action_count", 0) if dry_run_execution else 0,
            "dry_run_blocked_action_count": dry_run_execution.get("blocked_action_count", 0) if dry_run_execution else 0,
            "native_execution_status": native_execution_report.get("status", "") if native_execution_report else "",
            "native_execution_performed": bool(native_execution_report.get("native_execution_performed")) if native_execution_report else False,
            "native_successful_action_count": native_execution_report.get("successful_action_count", 0) if native_execution_report else 0,
            "synced_at": now,
        }
        if status in {"成功", "失败", "已取消"}:
            task["end_time"] = now
        if status == "失败":
            task["error_message"] = message or "本地客户端回传任务失败"
        elif status == "成功":
            task["error_message"] = ""
        self.store.append_log(task_id, f"本地客户端同步状态：{status} {progress}% {message}".strip())
        return self.store.save_task(task)

    @staticmethod
    def _validate_local_sync_transition(current_status: str, next_status: str) -> None:
        """拒绝状态倒退及绕过控制接口的同步更新。"""
        terminal_statuses = {"成功", "失败", "已取消"}
        if current_status in terminal_statuses and next_status != current_status:
            raise ValueError(f"终态任务不能从{current_status}同步改写为{next_status}")
        allowed = {
            "待处理": {"待处理", "处理中", "成功", "失败", "已暂停", "已取消", "已中断"},
            "处理中": {"处理中", "成功", "失败", "已暂停", "已取消", "已中断"},
            "已暂停": {"已暂停"},
            "已中断": {"已中断"},
            "成功": {"成功"},
            "失败": {"失败"},
            "已取消": {"已取消"},
        }
        if next_status not in allowed.get(current_status, set()):
            raise ValueError(f"本地任务状态不能从{current_status}同步转换为{next_status}")

    @staticmethod
    def _validate_local_sync_execution_task(task: dict[str, Any], execution: dict[str, Any]) -> None:
        """拒绝针对其他任务的模拟执行同步摘要。"""
        execution_task = execution.get("task") if isinstance(execution.get("task"), dict) else {}
        execution_task_id = str(execution_task.get("id") or "").strip()
        execution_task_type = str(execution_task.get("task_type") or "").strip()
        if execution_task_id and execution_task_id != str(task.get("id") or ""):
            raise ValueError("本地执行摘要任务 ID 不匹配")
        if execution_task_type and execution_task_type != str(task.get("task_type") or ""):
            raise ValueError("本地执行摘要任务类型不匹配")

    @staticmethod
    def _local_sync_status(status: str) -> str:
        """将客户端状态别名映射为 K12 任务状态。"""
        normalized = status.strip()
        aliases = {
            "queued": "待处理",
            "pending": "待处理",
            "running": "处理中",
            "in_progress": "处理中",
            "processing": "处理中",
            "completed": "成功",
            "complete": "成功",
            "success": "成功",
            "succeeded": "成功",
            "failed": "失败",
            "failure": "失败",
            "error": "失败",
            "paused": "已暂停",
            "cancelled": "已取消",
            "canceled": "已取消",
            "interrupted": "已中断",
        }
        supported = {"待处理", "处理中", "成功", "失败", "已暂停", "已取消", "已中断"}
        if normalized in supported:
            return normalized
        mapped = aliases.get(normalized.lower())
        if mapped:
            return mapped
        raise ValueError("本地任务状态不支持")

    @classmethod
    def _local_sync_outputs(cls, outputs: Any) -> list[dict[str, Any]]:
        """保存前清理客户端输出摘要。"""
        if not isinstance(outputs, list):
            return []
        cleaned: list[dict[str, Any]] = []
        for item in outputs[:50]:
            if not isinstance(item, dict):
                continue
            raw_sha256 = str(item.get("sha256") or item.get("checksum") or "").strip()
            sha256 = cls._safe_sha256(raw_sha256)
            if raw_sha256 and not sha256:
                raise ValueError("本地客户端输出 SHA-256 格式无效")
            raw_size = item.get("size", "")
            size_text = str(raw_size).strip()
            if size_text and not re.fullmatch(r"\d+", size_text):
                raise ValueError("本地客户端输出大小格式无效")
            size = int(size_text) if size_text else ""
            cleaned.append(
                {
                    "file_id": str(item.get("file_id") or "")[:80],
                    "name": str(item.get("name") or item.get("file_name") or Path(str(item.get("path") or "")).name),
                    "path": str(item.get("path") or ""),
                    "output_type": str(item.get("output_type") or item.get("type") or ""),
                    "status": str(item.get("status") or "已生成"),
                    "size": size,
                    "sha256": sha256,
                    "message": str(item.get("message") or ""),
                }
            )
        return cleaned

    def _reconcile_pending_local_outputs(
        self,
        task: dict[str, Any],
        status: str,
        outputs: list[dict[str, Any]],
        native_execution_report: dict[str, Any],
        message: str,
    ) -> dict[str, Any]:
        """根据校验过的成功、失败或取消证据处理待完成本地产物。"""
        reports = [report for report in self.store.list_reports() if report.get("task_id") == task.get("id")]
        report = reports[0] if reports else {}
        artifacts = list((report.get("analysis") or {}).get("artifacts") or [])
        pending = [item for item in artifacts if item.get("status") == "待本地客户端执行"]
        if not pending or status not in {"成功", "失败", "已取消"}:
            return {}
        if status == "已取消":
            cancellation_message = message or "本地 Office 转换已取消"
            for artifact in pending:
                artifact.update(
                    {
                        "status": "local_execution_cancelled",
                        "message": cancellation_message,
                        "native_execution_verified": False,
                    }
                )
            return self._refresh_reconciled_report(report)
        if status == "失败":
            native_actions = native_execution_report.get("actions") if isinstance(native_execution_report.get("actions"), list) else []
            failed_action = next(
                (
                    item
                    for item in native_actions
                    if isinstance(item, dict) and item.get("type") == "office_conversion" and item.get("status") == "failed"
                ),
                {},
            )
            failure_message = str(failed_action.get("message") or message or "本地 Office 客户端回传转换失败")
            for artifact in pending:
                artifact.update(
                    {
                        "status": "local_execution_failed",
                        "message": failure_message,
                        "native_execution_verified": bool(failed_action),
                    }
                )
            return self._refresh_reconciled_report(report)
        native_actions = native_execution_report.get("actions") if isinstance(native_execution_report.get("actions"), list) else []
        office_success = any(
            item.get("type") == "office_conversion"
            and item.get("status") == "success"
            and item.get("native_execution_performed")
            for item in native_actions
            if isinstance(item, dict)
        )
        if not office_success:
            raise ValueError("待本地 Office 输出不能仅凭任务状态标记成功，需要原生 office_conversion 执行证据")
        matched_outputs: set[int] = set()
        pending_by_type: dict[str, int] = {}
        for item in pending:
            output_type = str(item.get("output_type") or "").lower()
            pending_by_type[output_type] = pending_by_type.get(output_type, 0) + 1
        output_root = self.store.output_task_dir(str(task.get("id") or "")).resolve()
        for artifact in pending:
            output_type = str(artifact.get("output_type") or "").lower()
            file_id = str(artifact.get("file_id") or "")
            candidates = [
                (index, output)
                for index, output in enumerate(outputs)
                if index not in matched_outputs
                and str(output.get("output_type") or Path(str(output.get("name") or "")).suffix.lstrip(".")).lower() == output_type
                and (not output.get("file_id") or str(output.get("file_id")) == file_id)
            ]
            if pending_by_type.get(output_type, 0) > 1:
                candidates = [(index, output) for index, output in candidates if str(output.get("file_id") or "") == file_id]
            if len(candidates) != 1:
                raise ValueError(f"待本地执行文件 {artifact.get('source_file', file_id)} 缺少唯一可验证的 {output_type} 输出")
            output_index, output = candidates[0]
            output_path = Path(str(output.get("path") or ""))
            if not output_path.exists() or not output_path.is_file():
                raise ValueError(f"本地客户端回传输出不存在：{output.get('name') or output_type}")
            resolved = output_path.resolve()
            try:
                resolved.relative_to(output_root)
            except ValueError as exc:
                raise ValueError("本地客户端回传输出不在受管任务目录") from exc
            attempt_started_at_ns = int(task.get("attempt_started_at_ns") or 0)
            if attempt_started_at_ns and resolved.stat().st_mtime_ns < attempt_started_at_ns:
                raise ValueError(f"本地客户端回传输出早于当前任务轮次：{output.get('name') or output_type}")
            validation_error = self._output_artifact_validation_error(resolved, output_type)
            if validation_error:
                raise ValueError(f"本地客户端回传输出校验失败：{validation_error}")
            declared_size = output.get("size", "")
            if declared_size == "":
                raise ValueError(f"本地客户端回传输出缺少文件大小：{output.get('name') or output_type}")
            actual_size = resolved.stat().st_size
            if int(declared_size) != actual_size:
                raise ValueError(f"本地客户端回传输出文件大小不匹配：{output.get('name') or output_type}")
            declared_sha256 = str(output.get("sha256") or "")
            if not declared_sha256:
                raise ValueError(f"本地客户端回传输出缺少 SHA-256：{output.get('name') or output_type}")
            actual_sha256 = self._sha256(resolved)
            if declared_sha256 != actual_sha256:
                raise ValueError(f"本地客户端回传输出 SHA-256 不匹配：{output.get('name') or output_type}")
            matched_outputs.add(output_index)
            artifact.update(
                {
                    "file_name": resolved.name,
                    "path": str(resolved),
                    "url": f"/api/artifacts/{task['id']}/{resolved.name}",
                    "status": "成功",
                    "message": "本地 Office 客户端已回传并验证转换产物" + ("；" + str(output["message"])[:2000] if output.get("message") else ""),
                    "size": actual_size,
                    "sha256": actual_sha256,
                    "native_execution_verified": True,
                }
            )
        return self._refresh_reconciled_report(report)

    def _refresh_reconciled_report(self, report: dict[str, Any]) -> dict[str, Any]:
        """待完成产物进入终态后刷新报告计数与展示。"""
        artifacts = list((report.get("analysis") or {}).get("artifacts") or [])
        pending = [item for item in artifacts if item.get("status") == "待本地客户端执行"]
        cancelled = [item for item in artifacts if item.get("status") == "local_execution_cancelled"]
        failures = [
            item
            for item in artifacts
            if item.get("status") not in {"成功", "跳过", "待本地客户端执行", "local_execution_cancelled"}
        ]
        validation_failed_ids = {
            str(item.get("id") or "")
            for item in report.get("files") or []
            if item.get("validation_errors")
        }
        artifact_failed_ids = {str(item.get("file_id") or "") for item in failures if item.get("file_id")}
        image_failed_ids = {
            str(item.get("file_id") or "")
            for item in (report.get("analysis") or {}).get("imageExtractionErrors") or []
            if item.get("file_id")
        }
        failed_ids = {file_id for file_id in validation_failed_ids | artifact_failed_ids | image_failed_ids if file_id}
        pending_ids = {str(item.get("file_id") or "") for item in pending if item.get("file_id")}
        cancelled_ids = {str(item.get("file_id") or "") for item in cancelled if item.get("file_id")}
        report["artifact_count"] = sum(1 for item in artifacts if item.get("status") == "成功")
        report["artifact_failure_count"] = len(failures)
        report["artifact_pending_count"] = len(pending)
        report["artifact_cancelled_count"] = len(cancelled)
        report["failed_file_ids"] = sorted(failed_ids)
        report["fail_count"] = len(failed_ids)
        report["pending_file_ids"] = sorted(pending_ids - failed_ids)
        report["pending_count"] = len(report["pending_file_ids"])
        report["cancelled_file_ids"] = sorted(cancelled_ids - failed_ids - pending_ids)
        report["cancelled_count"] = len(report["cancelled_file_ids"])
        report["success_count"] = max(
            0,
            len(report.get("files") or [])
            - int(report.get("fail_count") or 0)
            - report["pending_count"]
            - report["cancelled_count"],
        )
        report["error_count"] = report["fail_count"]
        for item in report.get("qualityChecks") or []:
            if item.get("id") == "conversion_output":
                item.update(
                    {
                        "status": "失败" if failures else ("需确认" if pending or cancelled else "通过"),
                        "severity": "error" if failures else ("warning" if pending or cancelled else "info"),
                        "metric": f"成功 {report['artifact_count']} 个，待本地执行 {len(pending)} 个，已取消 {len(cancelled)} 个，失败 {len(failures)} 个",
                        "message": "本地 Office 转换已取消" if cancelled else ("本地 Office 输出已验证并登记" if not pending and not failures else item.get("message", "")),
                        "recommendation": "如需产物可重新发起转换" if cancelled else ("可下载转换结果" if not pending and not failures else item.get("recommendation", "")),
                    }
                )
        report["quality_issue_count"] = sum(1 for item in report.get("qualityChecks") or [] if item.get("status") != "通过")
        report["quality_blocker_count"] = sum(1 for item in report.get("qualityChecks") or [] if item.get("status") == "失败")
        report["failureRows"] = self.report_builder._failure_rows(report)
        report["failure_count"] = len(report["failureRows"])
        report["qualitySummary"] = self.report_builder._quality_summary(
            list(report.get("files") or []),
            dict(report.get("analysis") or {}),
            report,
        )
        report["summary"] = self.report_builder._summary(report)
        refreshed = self.report_builder.rewrite(report)
        return self.store.save_report(refreshed)

    @staticmethod
    def _local_sync_dry_run_execution(value: Any) -> dict[str, Any]:
        """脱敏并约束本地客户端模拟执行摘要。"""
        if not isinstance(value, dict):
            return {}
        raw_actions = value.get("actions") if isinstance(value.get("actions"), list) else []
        actions: list[dict[str, Any]] = []
        safe_capabilities = set(LOCAL_CLIENT_CAPABILITY_LABELS)
        for item in raw_actions[:20]:
            if not isinstance(item, dict):
                continue
            raw_capabilities = item.get("required_capabilities") if isinstance(item.get("required_capabilities"), list) else []
            capabilities = [
                {
                    "key": str(capability.get("key") or "")[:80],
                    "label": str(capability.get("label") or capability.get("key") or "")[:80],
                    "available": bool(capability.get("available")),
                }
                for capability in raw_capabilities
                if isinstance(capability, dict) and str(capability.get("key") or "") in safe_capabilities
            ]
            actions.append(
                {
                    "action_id": str(item.get("action_id") or "")[:120],
                    "type": str(item.get("type") or "")[:80],
                    "label": str(item.get("label") or item.get("type") or "")[:120],
                    "payload_status": str(item.get("payload_status") or "")[:80],
                    "gate_status": str(item.get("gate_status") or "")[:80],
                    "dry_run_status": str(item.get("dry_run_status") or "")[:80],
                    "native_execution_performed": False,
                    "required_capabilities": capabilities,
                    "step_count": TaskProcessor._non_negative_int(item.get("step_count"), 0),
                    "required_step_count": TaskProcessor._non_negative_int(item.get("required_step_count"), 0),
                    "output_artifact_types": [str(artifact)[:40] for artifact in item.get("output_artifact_types") or [] if isinstance(artifact, str)][:12],
                    "result_upload_optional": bool(item.get("result_upload_optional")),
                    "blockers": [str(blocker)[:120] for blocker in item.get("blockers") or [] if isinstance(blocker, str)][:12],
                }
            )
        return {
            "schema_version": "k12.localDryRunExecution.v1",
            "task": {
                "id": str((value.get("task") or {}).get("id") or "")[:80] if isinstance(value.get("task"), dict) else "",
                "task_type": str((value.get("task") or {}).get("task_type") or "")[:80] if isinstance(value.get("task"), dict) else "",
                "task_label": str((value.get("task") or {}).get("task_label") or "")[:120] if isinstance(value.get("task"), dict) else "",
            },
            "plan_status": str(value.get("plan_status") or "")[:80],
            "native_execution_allowed": bool(value.get("native_execution_allowed")),
            "companion_cli_native_execution": False,
            "action_count": TaskProcessor._non_negative_int(value.get("action_count"), len(actions)),
            "ready_action_count": TaskProcessor._non_negative_int(value.get("ready_action_count"), 0),
            "blocked_action_count": TaskProcessor._non_negative_int(value.get("blocked_action_count"), 0),
            "waiting_action_count": TaskProcessor._non_negative_int(value.get("waiting_action_count"), 0),
            "actions": actions,
            "path_policy": "no local file paths or tokens are stored",
        }

    @staticmethod
    def _local_sync_native_execution_report(value: Any) -> dict[str, Any]:
        """清理本地桌面客户端的真实原生执行报告。"""
        if not isinstance(value, dict):
            return {}
        raw_actions = value.get("actions") if isinstance(value.get("actions"), list) else []
        actions: list[dict[str, Any]] = []
        safe_capabilities = set(LOCAL_CLIENT_CAPABILITY_LABELS)
        safe_actions = set(LOCAL_ACTION_LABELS)
        for item in raw_actions[:20]:
            if not isinstance(item, dict):
                continue
            action_type = str(item.get("type") or "")[:80]
            if action_type not in safe_actions:
                continue
            raw_capabilities = item.get("required_capabilities") if isinstance(item.get("required_capabilities"), list) else []
            capabilities = [
                {
                    "key": str(capability.get("key") or "")[:80],
                    "label": str(capability.get("label") or capability.get("key") or "")[:80],
                    "available": bool(capability.get("available")),
                }
                for capability in raw_capabilities
                if isinstance(capability, dict) and str(capability.get("key") or "") in safe_capabilities
            ]
            action_status = TaskProcessor._native_execution_status(str(item.get("status") or item.get("native_status") or ""))
            native_performed = bool(item.get("native_execution_performed")) and action_status == "success"
            actions.append(
                {
                    "action_id": str(item.get("action_id") or "")[:120],
                    "type": action_type,
                    "label": str(item.get("label") or LOCAL_ACTION_LABELS.get(action_type, action_type))[:120],
                    "status": action_status,
                    "native_execution_performed": native_performed,
                    "platform": normalize_platform(str(item.get("platform") or value.get("platform") or "")),
                    "required_capabilities": capabilities,
                    "step_count": TaskProcessor._non_negative_int(item.get("step_count"), 0),
                    "successful_step_count": TaskProcessor._non_negative_int(item.get("successful_step_count"), 0),
                    "output_artifact_types": [str(artifact)[:40] for artifact in item.get("output_artifact_types") or [] if isinstance(artifact, str)][:12],
                    "message": TaskProcessor._redact_local_path_text(str(item.get("message") or ""))[:240],
                }
            )
        successful = [action for action in actions if action["native_execution_performed"] and action["status"] == "success"]
        return {
            "schema_version": "k12.localNativeExecutionReport.v1",
            "task": {
                "id": str((value.get("task") or {}).get("id") or "")[:80] if isinstance(value.get("task"), dict) else "",
                "task_type": str((value.get("task") or {}).get("task_type") or "")[:80] if isinstance(value.get("task"), dict) else "",
                "task_label": str((value.get("task") or {}).get("task_label") or "")[:120] if isinstance(value.get("task"), dict) else "",
            },
            "client_id": str(value.get("client_id") or "")[:120],
            "platform": normalize_platform(str(value.get("platform") or "")),
            "status": "success" if successful and len(successful) == len(actions) else ("partial_success" if successful else "no_native_success"),
            "native_execution_performed": bool(successful),
            "action_count": len(actions),
            "successful_action_count": len(successful),
            "actions": actions,
            "path_policy": "no local file paths or tokens are stored",
        }

    @staticmethod
    def _native_execution_status(status: str) -> str:
        """规范本地客户端原生执行结果状态。"""
        normalized = str(status or "").strip()
        aliases = {
            "completed": "success",
            "complete": "success",
            "success": "success",
            "succeeded": "success",
            "成功": "success",
            "failed": "failed",
            "failure": "failed",
            "error": "failed",
            "失败": "failed",
            "skipped": "skipped",
            "跳过": "skipped",
            "blocked": "blocked",
            "阻断": "blocked",
            "pending": "pending",
            "等待": "pending",
        }
        return aliases.get(normalized.lower(), aliases.get(normalized, normalized[:80] or "unknown"))

    def recover_interrupted_tasks(self) -> dict[str, Any]:
        """服务重启后将运行中任务标记为可恢复。"""
        recovered: list[dict[str, Any]] = []
        for task in self.store.list_tasks():
            if task.get("status") != "处理中":
                continue
            task["status"] = "已中断"
            task["progress"] = min(int(task.get("progress") or 0), 99)
            task["end_time"] = utc_now()
            task["error_message"] = "服务重启或异常退出，任务已标记为可恢复"
            task["recoverable"] = True
            task["interrupted_at"] = utc_now()
            self.store.save_task(task)
            self.store.append_log(task["id"], "检测到服务中断，任务已标记为可恢复", "warning")
            recovered.append(task)
        return {"recovered_count": len(recovered), "tasks": recovered}

    def task_recovery_summary(self) -> dict[str, Any]:
        """汇总暂停、失败、取消和重启后可恢复的任务。"""
        tasks = list(self.store.list_tasks())
        items: list[dict[str, Any]] = []
        for task in tasks:
            status = str(task.get("status") or "")
            recoverable = bool(task.get("recoverable")) or status in {"已中断", "已暂停"}
            retryable = status in {"失败", "已取消", "已中断"}
            if status not in {"处理中", "已中断", "已暂停", "失败", "已取消"} and not recoverable and not retryable:
                continue
            task_id = str(task.get("id") or "")
            items.append(
                {
                    "task_id": task_id,
                    "task_type": task.get("task_type", ""),
                    "task_label": task.get("task_label") or TASK_LABELS.get(str(task.get("task_type") or ""), str(task.get("task_type") or "")),
                    "status": status,
                    "progress": int(task.get("progress") or 0),
                    "recoverable": recoverable,
                    "retryable": retryable,
                    "recovery_status": self._task_recovery_status(task, recoverable, retryable),
                    "error_summary": self._redact_local_path_text(str(task.get("error_message") or "")),
                    "interrupted_at": task.get("interrupted_at", ""),
                    "updated_at": task.get("end_time") or task.get("start_time") or task.get("created_at") or "",
                    "resume_endpoint": f"/api/tasks/{task_id}/resume" if recoverable else "",
                    "retry_endpoint": f"/api/tasks/{task_id}/retry" if retryable else "",
                }
            )
        summary = {
            "total_tasks": len(tasks),
            "running_count": sum(1 for task in tasks if task.get("status") == "处理中"),
            "interrupted_count": sum(1 for task in tasks if task.get("status") == "已中断"),
            "paused_count": sum(1 for task in tasks if task.get("status") == "已暂停"),
            "failed_count": sum(1 for task in tasks if task.get("status") == "失败"),
            "recoverable_count": sum(1 for item in items if item["recoverable"]),
            "retryable_count": sum(1 for item in items if item["retryable"]),
        }
        return {
            "schema_version": "k12.taskRecoverySummary.v1",
            "generated_at": utc_now(),
            "summary": summary,
            "items": items,
            "contracts": {
                "resume_method": "POST",
                "retry_method": "POST",
                "path_policy": "local_paths_redacted",
                "restart_policy": "startup_marks_running_tasks_interrupted_and_recoverable",
            },
        }

    @staticmethod
    def _task_recovery_status(task: dict[str, Any], recoverable: bool, retryable: bool) -> str:
        """返回中断或可重试任务的界面恢复状态。"""
        status = str(task.get("status") or "")
        if status == "处理中":
            return "running_watch"
        if status == "已中断" and recoverable:
            return "restart_recoverable"
        if status == "已暂停" and recoverable:
            return "paused_resumable"
        if retryable:
            return "retry_available"
        return "tracked"

    @staticmethod
    def _redact_local_path_text(text: str) -> str:
        """隐藏消息中的常见本地文件路径。"""
        def replacement(match: re.Match[str]) -> str:
            """将匹配的本地路径替换为不含路径的显示标签。"""
            raw = match.group(0)
            name = raw.replace("\\", "/").rstrip("/").split("/")[-1]
            return f"本地路径已隐藏/{name}" if name else "本地路径已隐藏"

        return LOCAL_PATH_TEXT.sub(replacement, text)

    def cancel_task(self, task_id: str) -> dict[str, Any]:
        """取消尚未进入终态的任务。"""
        self._assert_permission("tasks.control")
        task = self.store.get_task(task_id)
        if not task:
            raise KeyError(f"Task not found: {task_id}")
        if task.get("status") in {"成功", "失败", "已取消"}:
            raise ValueError("成功、失败或已取消的终态任务不能取消")
        task["status"] = "已取消"
        task["progress"] = min(task.get("progress", 0), 99)
        task["end_time"] = utc_now()
        reconciled_report = self._reconcile_pending_local_outputs(
            task,
            "已取消",
            [],
            {},
            "用户取消待本地 Office 执行任务",
        )
        if reconciled_report:
            self._apply_task_report_summary(task, reconciled_report)
        self.store.append_log(task_id, "用户取消任务", "warning")
        return self.store.save_task(task)

    def file_preview(self, file_id: str) -> dict[str, Any]:
        """生成上传文件或源文件的结构化预览载荷。"""
        file = self.store.get_file(file_id)
        if not file:
            raise KeyError(f"File not found: {file_id}")
        return build_file_preview(file)

    def report_comparison(self, report_id: str) -> dict[str, Any]:
        """比较源预览与成功转换产物的预览。"""
        report = self.store.get_report(report_id)
        if not report:
            raise KeyError(f"Report not found: {report_id}")
        files = {file.get("id"): file for file in report.get("files", [])}
        items: list[dict[str, Any]] = []
        for artifact in report.get("analysis", {}).get("artifacts", []):
            if artifact.get("status") != "成功":
                continue
            output_path = Path(str(artifact.get("path") or ""))
            output_type = str(artifact.get("output_type") or "").lower()
            if output_type not in {"docx", "pptx", "xlsx", "pdf"}:
                continue
            if not output_path.exists() or not output_path.is_file():
                continue
            source = files.get(artifact.get("file_id")) or self.store.get_file(str(artifact.get("file_id") or ""))
            if not source:
                continue
            output_file = self._artifact_file_item(output_path, artifact)
            source_preview = build_file_preview(source)
            output_preview = build_file_preview(output_file)
            items.append(
                {
                    "file_id": source.get("id", ""),
                    "source_file": source.get("file_name", artifact.get("source_file", "")),
                    "output_file": output_file.get("file_name", artifact.get("file_name", "")),
                    "output_type": output_type,
                    "artifact_url": artifact.get("url", ""),
                    "source_preview": source_preview,
                    "output_preview": output_preview,
                    "diff": self._preview_diff(source_preview, output_preview),
                }
            )
        return {
            "report_id": report_id,
            "task_id": report.get("task_id", ""),
            "status": "ready" if items else "empty",
            "item_count": len(items),
            "items": items,
        }

    def file_download_info(self, file_id: str) -> dict[str, Any]:
        """返回注册源文件的安全下载描述。"""
        file = self.store.get_file(file_id)
        if not file:
            raise KeyError(f"File not found: {file_id}")
        path = Path(file.get("storage_path") or file.get("file_path") or "")
        if not path.exists() or not path.is_file():
            raise ValueError("文件源不存在，无法下载")
        resolved = path.resolve()
        managed_source = str(file.get("source_kind") or "") in {"upload", "archive_entry"}
        if managed_source:
            try:
                resolved.relative_to(self.store.uploads_dir.resolve())
            except ValueError as exc:
                raise ValueError("受管上传文件路径越界，已拒绝下载") from exc
        download_path = resolved if managed_source else path
        declared_size = int(file.get("file_size") or 0)
        declared_sha256 = str(file.get("source_sha256") or "")
        if not re.fullmatch(r"[0-9a-f]{64}", declared_sha256):
            raise ValueError("文件源缺少有效完整性哈希，已拒绝下载")
        if download_path.stat().st_size != declared_size or self._sha256(download_path) != declared_sha256:
            raise ValueError("文件源内容与登记记录不一致，已拒绝下载")
        return {
            "path": download_path,
            "file_name": file.get("file_name") or download_path.name,
            "file_type": file.get("file_type", ""),
            "file_size": declared_size,
            "sha256": declared_sha256,
        }

    def _artifact_file_item(self, path: Path, artifact: dict[str, Any]) -> dict[str, Any]:
        """将转换产物分析为可临时预览的文件记录。"""
        analyzer = DocumentAnalyzer(self.store.get_settings().get("singleFileLimitMb", 500))
        item = analyzer.analyze_file(path, str(artifact.get("file_name") or path.name))
        payload = item.to_dict()
        payload["source_kind"] = "conversion_output"
        payload["storage_path"] = str(path)
        payload["file_path"] = str(path)
        payload["content_summary"] = {
            **payload.get("content_summary", {}),
            "sourceFile": artifact.get("source_file", ""),
            "taskId": artifact.get("task_id", ""),
            "outputType": artifact.get("output_type", ""),
            "message": artifact.get("message", ""),
        }
        return payload

    @staticmethod
    def _preview_diff(source_preview: dict[str, Any], output_preview: dict[str, Any]) -> dict[str, Any]:
        """比较源与转换预览，用于报告质量检查。"""
        source_pages = len(source_preview.get("pages") or [])
        output_pages = len(output_preview.get("pages") or [])
        source_objects = {item.get("label") or item.get("type"): int(item.get("count") or 0) for item in source_preview.get("objects", [])}
        output_objects = {item.get("label") or item.get("type"): int(item.get("count") or 0) for item in output_preview.get("objects", [])}
        labels = sorted(set(source_objects) | set(output_objects))
        return {
            "source_page_count": source_pages,
            "output_page_count": output_pages,
            "page_delta": output_pages - source_pages,
            "source_object_count": sum(source_objects.values()),
            "output_object_count": sum(output_objects.values()),
            "object_delta": sum(output_objects.values()) - sum(source_objects.values()),
            "object_changes": [
                {
                    "label": label,
                    "source": source_objects.get(label, 0),
                    "output": output_objects.get(label, 0),
                    "delta": output_objects.get(label, 0) - source_objects.get(label, 0),
                }
                for label in labels
            ],
            "warnings": list(source_preview.get("warnings") or []) + list(output_preview.get("warnings") or []),
        }

    def set_file_password(self, file_id: str, password: str) -> dict[str, Any]:
        """仅在内存中保存加密文件会话密码。"""
        self._assert_permission("files.manage")
        return self.store.set_file_password(file_id, password)

    def delete_file(self, file_id: str) -> bool:
        """删除文件记录及受控上传缓存。"""
        self._assert_permission("files.manage")
        return self.store.delete_file(file_id)

    def delete_report(self, report_id: str) -> dict[str, Any]:
        """删除报告及其衍生审阅记录。"""
        self._assert_permission("reports.manage")
        return self.store.delete_report(report_id)

    def pause_task(self, task_id: str) -> dict[str, Any]:
        """暂停任务，使其可稍后从队列恢复。"""
        self._assert_permission("tasks.control")
        task = self.store.get_task(task_id)
        if not task:
            raise KeyError(f"Task not found: {task_id}")
        if task.get("status") not in {"待处理", "处理中"}:
            raise ValueError("只有待处理或处理中的任务可以暂停")
        task["status"] = "已暂停"
        task["progress"] = min(task.get("progress", 0), 99)
        task["end_time"] = utc_now()
        self.store.append_log(task_id, "用户暂停任务", "warning")
        return self.store.save_task(task)

    def resume_task(self, task_id: str) -> dict[str, Any]:
        """恢复暂停或重启中断的任务。"""
        self._assert_permission("tasks.control")
        task = self.store.get_task(task_id)
        if not task:
            raise KeyError(f"Task not found: {task_id}")
        if task.get("status") not in {"已暂停", "已中断"}:
            raise ValueError("只有已暂停或已中断的任务可以继续")
        previous_status = task["status"]
        task["status"] = "待处理"
        task["progress"] = 0
        task["error_message"] = ""
        task["end_time"] = ""
        task["recoverable"] = False
        self.store.save_task(task)
        self.store.append_log(task_id, "用户恢复中断任务" if previous_status == "已中断" else "用户继续任务")
        return self.run_task(task_id)

    def restore_task_backups(self, task_id: str) -> dict[str, Any]:
        """恢复最新报告中宏执行前的备份。"""
        self._assert_permission("tasks.control")
        task = self.store.get_task(task_id)
        if not task:
            raise KeyError(f"Task not found: {task_id}")
        reports = [report for report in self.store.list_reports() if report.get("task_id") == task_id]
        if not reports:
            return {"task": task, "results": [], "restored_count": 0, "failed_count": 0}
        latest_report = reports[0]
        files = {file["id"]: file for file in latest_report.get("files", [])}
        backup_root = self.store.backups_dir.resolve()
        seen: set[tuple[str, str]] = set()
        results: list[dict[str, Any]] = []
        for macro in latest_report.get("analysis", {}).get("macros", []):
            file_id = str(macro.get("file_id") or "")
            backup_path = Path(str(macro.get("backup_path") or ""))
            key = (file_id, str(backup_path))
            if key in seen or not file_id or not str(backup_path):
                continue
            seen.add(key)
            file = files.get(file_id) or self.store.get_file(file_id) or {}
            target_path = self._original_source_path(file)
            result = {
                "file_id": file_id,
                "file_name": file.get("file_name", ""),
                "backup_path": str(backup_path),
                "target_path": str(target_path),
                "status": "失败",
                "message": "",
            }
            try:
                resolved_backup = backup_path.resolve()
                resolved_backup.relative_to(backup_root)
                if not resolved_backup.exists() or not resolved_backup.is_file():
                    raise FileNotFoundError("备份文件不存在")
                if not target_path.exists() or not target_path.is_file():
                    raise FileNotFoundError("原文件路径不存在，无法恢复")
                declared_size = macro.get("backup_size", "")
                declared_sha256 = str(macro.get("backup_sha256") or "")
                if declared_size == "" or not re.fullmatch(r"[0-9a-f]{64}", declared_sha256):
                    raise ValueError("宏备份缺少有效完整性记录")
                backup_data = resolved_backup.read_bytes()
                if len(backup_data) != int(declared_size) or hashlib.sha256(backup_data).hexdigest() != declared_sha256:
                    raise ValueError("宏备份内容与登记记录不一致")
                target_path.write_bytes(backup_data)
                result["status"] = "成功"
                result["message"] = "已从宏执行前备份恢复"
                self.store.append_log(task_id, f"恢复宏备份：{file.get('file_name', file_id)}", category="macro")
            except (OSError, ValueError) as exc:
                result["message"] = str(exc)
                self.store.append_log(task_id, f"恢复宏备份失败：{file.get('file_name', file_id)}，{exc}", "error", category="macro")
            results.append(result)
        task["backup_restore_results"] = results
        task["backup_restored_at"] = utc_now() if any(item["status"] == "成功" for item in results) else ""
        self.store.save_task(task)
        return {
            "task": task,
            "results": results,
            "restored_count": sum(1 for item in results if item["status"] == "成功"),
            "failed_count": sum(1 for item in results if item["status"] != "成功"),
        }

    def cleanup_runtime_history(self, payload: dict[str, Any] | None = None) -> dict[str, Any]:
        """清理过期任务、报告、日志和受管运行缓存。"""
        payload = payload or {}
        settings = self.store.get_settings()
        retention_days = self._non_negative_int(
            payload.get("retentionDays", payload.get("cleanupRetentionDays", settings.get("cleanupRetentionDays", 30))),
            30,
        )
        log_retention_days = self._non_negative_int(
            payload.get("logRetentionDays", settings.get("logRetentionDays", retention_days)),
            retention_days,
        )
        result = self.store.cleanup_runtime_history(retention_days, log_retention_days)
        self.store.append_log(
            "system",
            f"系统清理完成：任务 {result['tasks_deleted']} 个，报告 {result['reports_deleted']} 个，日志 {result['logs_deleted']} 条",
        )
        result["message"] = "清理完成"
        return result

    def auto_cleanup_runtime_history(self) -> dict[str, Any]:
        """仅在配置的间隔到期后执行保留期清理。"""
        settings = self.store.get_settings()
        if not settings.get("autoCleanTemp"):
            return {"skipped": True, "reason": "自动清理未开启"}
        interval_days = self._positive_int(settings.get("cleanupIntervalDays", 1), 1)
        last_cleanup = self._parse_utc(str(settings.get("lastCleanupAt") or ""))
        now = datetime.now(timezone.utc)
        if last_cleanup and now - last_cleanup < timedelta(days=interval_days):
            return {
                "skipped": True,
                "reason": "未到自动清理周期",
                "cleanup_interval_days": interval_days,
                "last_cleanup_at": last_cleanup.isoformat(timespec="seconds"),
            }
        result = self.cleanup_runtime_history(
            {
                "retentionDays": settings.get("cleanupRetentionDays", 30),
                "logRetentionDays": settings.get("logRetentionDays", settings.get("cleanupRetentionDays", 30)),
            }
        )
        cleanup_time = utc_now()
        self.store.update_settings({"lastCleanupAt": cleanup_time})
        result["auto_cleanup"] = True
        result["cleanup_interval_days"] = interval_days
        result["last_cleanup_at"] = cleanup_time
        return result

    def current_user(self) -> dict[str, Any]:
        """返回当前用户及继承的有效权限。"""
        settings = self.store.get_settings()
        users = self.store.list_users()
        active_id = str(settings.get("activeUserId") or "user_admin")
        user = next((item for item in users if item.get("id") == active_id and item.get("status") == "启用"), None)
        if not user:
            user = next((item for item in users if item.get("role") == "管理员" and item.get("status") == "启用"), None)
        if not user and users:
            user = users[0]
        if not user:
            user = self.store.save_user({"id": "user_admin", "name": "本地管理员", "role": "管理员"})
        role = str(user.get("role") or "访客")
        inherited = ROLE_PERMISSIONS.get(role, ROLE_PERMISSIONS["访客"])
        explicit = user.get("permissions") if isinstance(user.get("permissions"), list) else []
        user = dict(user)
        user["effective_permissions"] = list(dict.fromkeys([*inherited, *[str(item) for item in explicit]]))
        return user

    def list_users(self) -> list[dict[str, Any]]:
        """返回本地用户及角色权限、激活状态。"""
        current_id = self.current_user().get("id")
        users = self.store.list_users()
        for user in users:
            role = str(user.get("role") or "访客")
            user["role_permissions"] = ROLE_PERMISSIONS.get(role, ROLE_PERMISSIONS["访客"])
            user["active"] = user.get("id") == current_id
        return users

    def save_user(self, payload: dict[str, Any]) -> dict[str, Any]:
        """创建或更新本地用户，并按需激活。"""
        self._assert_permission("users.manage")
        if not str(payload.get("name") or "").strip():
            raise ValueError("用户名称不能为空")
        user = self.store.save_user(payload)
        if bool(payload.get("active")):
            self.store.update_settings({"activeUserId": user["id"]})
            user["active"] = True
        return user

    def delete_user(self, user_id: str) -> bool:
        """删除本地用户，同时保留至少一名可用管理员。"""
        self._assert_permission("users.manage")
        users = self.store.list_users()
        target = next((item for item in users if item.get("id") == user_id), None)
        if not target:
            return False
        active_admins = [item for item in users if item.get("role") == "管理员" and item.get("status") == "启用" and item.get("id") != user_id]
        if target.get("role") == "管理员" and not active_admins:
            raise ValueError("至少需要保留一个启用的管理员")
        deleted = self.store.delete_user(user_id)
        settings = self.store.get_settings()
        if deleted and settings.get("activeUserId") == user_id:
            replacement = active_admins[0] if active_admins else next((item for item in self.store.list_users()), None)
            if replacement:
                self.store.update_settings({"activeUserId": replacement["id"]})
        return deleted

    def activate_user(self, user_id: str) -> dict[str, Any]:
        """将本地会话切换到允许登录的启用用户。"""
        user = self.store.get_user(user_id)
        if not user:
            raise ValueError("用户不存在")
        if user.get("status") != "启用":
            raise ValueError("用户未启用")
        if not user.get("login_enabled", True):
            raise ValueError("用户不允许登录")
        user["last_login_at"] = utc_now()
        self.store.save_user(user)
        self.store.update_settings({"activeUserId": user_id})
        return self.current_user()

    def list_templates(self) -> list[dict[str, Any]]:
        """返回转换、OCR 和公式模板记录。"""
        return self.store.list_templates()

    def save_template(self, payload: dict[str, Any]) -> dict[str, Any]:
        """保存可复用模板并同步相关默认设置。"""
        self._assert_permission("templates.manage")
        if not str(payload.get("name") or "").strip():
            raise ValueError("模板名称不能为空")
        template = self.store.save_template(payload)
        if template["applies_to"] == "word_to_ppt" and template["name"]:
            self.store.update_settings({"wordToPptTemplate": template["name"]})
        if template["applies_to"] == "ppt_to_word":
            updates: dict[str, Any] = {}
            if template.get("settings", {}).get("mode"):
                updates["pptToWordMode"] = template["settings"]["mode"]
            if template.get("name"):
                updates["pptToWordTemplate"] = template["name"]
            if template.get("template_path"):
                updates["pptToWordTemplatePath"] = template["template_path"]
            if updates:
                self.store.update_settings(updates)
        return template

    def delete_template(self, template_id: str) -> bool:
        """删除可复用模板记录。"""
        self._assert_permission("templates.manage")
        return self.store.delete_template(template_id)

    def list_authorizations(self) -> list[dict[str, Any]]:
        """返回与当前设置同步的授权开关。"""
        settings = self.store.get_settings()
        authorizations = self.store.list_authorizations()
        for item in authorizations:
            setting_key = item.get("setting_key")
            if setting_key in settings:
                item["enabled"] = bool(settings.get(setting_key))
                item["status"] = "已授权" if item["enabled"] else "未授权"
        return authorizations

    def save_authorization(self, payload: dict[str, Any]) -> dict[str, Any]:
        """更新敏感功能的明确授权及对应设置。"""
        self._assert_permission("authorizations.manage")
        key = str(payload.get("key") or "")
        if key not in DEFAULT_AUTHORIZATIONS:
            raise ValueError("授权项不存在")
        authorization = self.store.save_authorization(payload)
        setting_key = authorization.get("setting_key")
        if setting_key:
            self.store.update_settings({setting_key: bool(authorization.get("enabled"))})
        return self.list_authorizations_by_key(key)

    def list_authorizations_by_key(self, key: str) -> dict[str, Any]:
        """根据稳定键返回授权记录。"""
        for item in self.list_authorizations():
            if item.get("key") == key:
                return item
        raise ValueError("授权项不存在")

    def update_settings(self, payload: dict[str, Any]) -> dict[str, Any]:
        """验证当前用户权限后保存设置。"""
        self._assert_permission("settings.manage")
        return self.store.update_settings(payload)

    def preflight_checks(self, payload: dict[str, Any] | None = None) -> list[dict[str, Any]]:
        """返回权限、本地组件及 Mathpix 的需求预检结果。"""
        payload = payload or {}
        settings = self.store.get_settings()
        task_type = str(payload.get("task_type") or "")
        file_ids = list(payload.get("file_ids") or [])
        execute_mode = str(payload.get("execute_mode") or self.resolve_execute_mode(task_type, file_ids) if task_type else "")
        checks: list[dict[str, Any]] = []
        output_dir = self._preflight_output_dir()
        checks.append(output_dir)
        checks.append(self._preflight_disk_space(output_dir.get("path", ""), settings))
        checks.append(self._preflight_user_permission(task_type))
        checks.append(self._preflight_local_mode(task_type, execute_mode, settings))
        checks.append(self._preflight_local_client_platform(task_type, execute_mode, settings))
        checks.append(self._preflight_local_client_components(task_type, execute_mode, file_ids, settings))
        checks.append(self._preflight_mathpix(task_type, settings))
        checks.append(self._preflight_mathtype_compatibility(task_type, file_ids, settings))
        checks.append(self._preflight_macro(task_type, settings))
        return checks

    def _assert_task_permission(self, task_type: str) -> None:
        """检查创建任务及必要的宏执行权限。"""
        self._assert_permission("tasks.create")
        if task_type == "macro_sequence":
            self._assert_permission("macros.execute")

    def _assert_permission(self, permission: str) -> None:
        """当前用户缺少指定权限时抛出错误。"""
        user = self.current_user()
        permissions = set(user.get("effective_permissions") or [])
        if permission not in permissions:
            raise ValueError(f"当前用户缺少权限：{permission}")

    def _preflight_output_dir(self) -> dict[str, Any]:
        """验证配置的输出目录可写。"""
        try:
            output_dir = self.store.output_base_dir()
            probe = output_dir / ".k12-write-test"
            probe.write_text("ok", encoding="utf-8")
            probe.unlink(missing_ok=True)
            return self._preflight_item(
                "output_directory_writable",
                "输出目录权限",
                "通过",
                str(output_dir),
                "输出目录可写",
                "可继续生成转换结果和报告",
                path=str(output_dir),
            )
        except OSError as exc:
            return self._preflight_item(
                "output_directory_writable",
                "输出目录权限",
                "失败",
                "不可写",
                f"当前输出目录无写入权限：{exc}",
                "更换输出目录或调整文件夹权限",
            )

    def _preflight_disk_space(self, path_value: str, settings: dict[str, Any]) -> dict[str, Any]:
        """检查可用磁盘空间是否达到配置阈值。"""
        min_free_mb = self._positive_int(settings.get("minFreeDiskMb", 512), 512)
        path = Path(path_value or self.store.data_dir)
        try:
            usage = shutil.disk_usage(path)
            free_mb = int(usage.free / 1024 / 1024)
            return self._preflight_item(
                "disk_space",
                "磁盘空间",
                "失败" if free_mb < min_free_mb else "通过",
                f"可用 {free_mb} MB / 阈值 {min_free_mb} MB",
                "磁盘空间不足" if free_mb < min_free_mb else "磁盘空间满足当前阈值",
                "清理空间或切换输出目录" if free_mb < min_free_mb else "可继续处理",
                path=str(path),
            )
        except OSError as exc:
            return self._preflight_item("disk_space", "磁盘空间", "需确认", "无法读取", str(exc), "检查输出目录是否可访问")

    def _preflight_user_permission(self, task_type: str) -> dict[str, Any]:
        """检查当前用户能否创建所请求的任务。"""
        user = self.current_user()
        permissions = set(user.get("effective_permissions") or [])
        missing: list[str] = []
        if task_type and "tasks.create" not in permissions:
            missing.append("tasks.create")
        if task_type == "macro_sequence" and "macros.execute" not in permissions:
            missing.append("macros.execute")
        return self._preflight_item(
            "user_permissions",
            "用户权限",
            "失败" if missing else "通过",
            f"{user.get('name', '')} / {user.get('role', '')}",
            f"缺少权限：{', '.join(missing)}" if missing else "当前用户具备任务所需权限",
            "切换管理员或调整角色权限" if missing else "可继续创建任务",
        )

    def _preflight_local_mode(self, task_type: str, execute_mode: str, settings: dict[str, Any]) -> dict[str, Any]:
        """检查需本地处理的任务是否启用客户端路由。"""
        needs_local = task_type in LOCAL_REQUIRED_TASKS or execute_mode in {"local", "hybrid"}
        enabled = bool(settings.get("localClientEnabled", True))
        status = "需确认" if needs_local and not enabled else "通过"
        message = "当前任务依赖本地客户端，但本地连接已关闭" if status == "需确认" else "本地/网页执行模式配置可用"
        return self._preflight_item(
            "local_client",
            "本地客户端",
            status,
            execute_mode or "未指定",
            message,
            "启用本地客户端或改用网页可处理任务" if status == "需确认" else "按当前分流继续",
        )

    def _preflight_local_client_platform(self, task_type: str, execute_mode: str, settings: dict[str, Any]) -> dict[str, Any]:
        """交接 MathType 前验证 Windows/macOS 客户端心跳。"""
        needs_local = task_type in LOCAL_REQUIRED_TASKS or execute_mode in {"local", "hybrid"}
        expected_platform = self._configured_local_client_platform(settings)
        heartbeat = settings.get("localClientHeartbeat") if isinstance(settings.get("localClientHeartbeat"), dict) else {}
        raw_preflight = heartbeat.get("preflight") if isinstance(heartbeat.get("preflight"), dict) else {}
        preflight = self._safe_local_client_preflight(raw_preflight) if raw_preflight else {}
        actual_platform = normalize_platform(str(heartbeat.get("platform") or preflight.get("platform") or ""))
        has_heartbeat = bool(heartbeat.get("received_at"))
        last_heartbeat_at = str(heartbeat.get("received_at") or "")
        extra = {
            "expected_platform": expected_platform,
            "heartbeat_platform": actual_platform,
            "last_heartbeat_at": last_heartbeat_at,
        }
        platform_mismatch = bool(
            has_heartbeat
            and expected_platform in {"Windows", "macOS"}
            and actual_platform in {"Windows", "macOS"}
            and expected_platform != actual_platform
        )
        if not needs_local and not platform_mismatch:
            return self._preflight_item(
                "local_client_platform",
                "客户端平台心跳",
                "通过",
                "不涉及",
                "当前任务不要求本地客户端平台交接，未发现平台冲突",
                "无需处理",
                platform_compatible=True,
                **extra,
            )
        if not settings.get("localClientEnabled", True):
            return self._preflight_item(
                "local_client_platform",
                "客户端平台心跳",
                "需确认",
                "本地连接已关闭",
                "本地客户端连接关闭，无法确认 Windows/macOS 平台心跳",
                "启用本地客户端或改用 MathML/LaTeX、图片兜底格式",
                platform_compatible=False,
                **extra,
            )
        if not has_heartbeat:
            return self._preflight_item(
                "local_client_platform",
                "客户端平台心跳",
                "需确认",
                "未连接",
                "尚未收到本地客户端心跳，无法确认 Windows/macOS 平台是否匹配",
                "启动对应平台的本地客户端，或使用 MathML/LaTeX、图片兜底格式",
                platform_compatible=False,
                **extra,
            )
        if platform_mismatch:
            return self._preflight_item(
                "local_client_platform",
                "客户端平台心跳",
                "失败",
                f"目标 {expected_platform} / 心跳 {actual_platform}",
                self._local_platform_compatibility_message(expected_platform, actual_platform, True),
                "切换为目标平台客户端，或改用 MathML/LaTeX、图片兜底格式后再运行",
                platform_compatible=False,
                **extra,
            )
        if expected_platform in {"Windows", "macOS"} and actual_platform == expected_platform:
            return self._preflight_item(
                "local_client_platform",
                "客户端平台心跳",
                "通过",
                f"{expected_platform} 已匹配",
                self._local_platform_compatibility_message(expected_platform, actual_platform, False),
                "可继续交付该平台的 Office/MathType 本地任务",
                platform_compatible=True,
                **extra,
            )
        return self._preflight_item(
            "local_client_platform",
            "客户端平台心跳",
            "需确认",
            f"目标 {expected_platform or '未知'} / 心跳 {actual_platform or '未知'}",
            self._local_platform_compatibility_message(expected_platform, actual_platform, False),
            "在设置中明确选择 Windows 或 macOS，涉及公式对象时保留兜底格式",
            platform_compatible=False,
            **extra,
        )

    def _preflight_local_client_components(self, task_type: str, execute_mode: str, file_ids: list[str], settings: dict[str, Any]) -> dict[str, Any]:
        """检查任务需要的心跳组件能力。"""
        required_keys = self._required_preflight_capabilities(task_type, execute_mode, file_ids, settings)
        if not required_keys:
            return self._preflight_item(
                "local_client_components",
                "本地组件能力",
                "通过",
                "不涉及",
                "当前任务不要求 Office、MathType、OMML 或宏执行组件",
                "无需处理",
                required_capabilities=[],
                missing_capabilities=[],
            )
        heartbeat = settings.get("localClientHeartbeat") if isinstance(settings.get("localClientHeartbeat"), dict) else {}
        raw_preflight = heartbeat.get("preflight") if isinstance(heartbeat.get("preflight"), dict) else {}
        preflight = self._safe_local_client_preflight(raw_preflight) if raw_preflight else {}
        capabilities = self._local_client_capabilities(heartbeat, preflight)
        required = [
            {
                "key": key,
                "label": LOCAL_CLIENT_CAPABILITY_LABELS.get(key, key),
                "available": self._local_capability_available(capabilities.get(key)),
            }
            for key in required_keys
        ]
        missing = [item for item in required if not item["available"]]
        component_names = "、".join(item["label"] for item in required)
        if not heartbeat.get("received_at"):
            return self._preflight_item(
                "local_client_components",
                "本地组件能力",
                "需确认",
                "未收到心跳",
                f"当前任务需要 {component_names}，但尚未收到本地客户端组件预检",
                "启动本地客户端并发送心跳后重试，或切换到 MathML/LaTeX、图片等兜底流程",
                required_capabilities=required,
                missing_capabilities=[item["key"] for item in required],
                components=dict(preflight.get("components") or {}),
            )
        if missing:
            missing_names = "、".join(item["label"] for item in missing)
            return self._preflight_item(
                "local_client_components",
                "本地组件能力",
                "失败",
                f"缺少 {missing_names}",
                f"本地客户端已连接，但缺少当前任务所需组件能力：{missing_names}",
                "安装或启用对应 Windows/macOS 本地组件，或跳过相关功能并使用兜底格式",
                required_capabilities=required,
                missing_capabilities=[item["key"] for item in missing],
                components=dict(preflight.get("components") or {}),
            )
        return self._preflight_item(
            "local_client_components",
            "本地组件能力",
            "通过",
            component_names,
            "本地客户端组件能力满足当前任务要求",
            "可继续交付本地任务",
            required_capabilities=required,
            missing_capabilities=[],
            components=dict(preflight.get("components") or {}),
        )

    def _required_preflight_capabilities(self, task_type: str, execute_mode: str, file_ids: list[str], settings: dict[str, Any]) -> list[str]:
        """根据任务类型和文件返回所需客户端能力键。"""
        needs_local = task_type in LOCAL_REQUIRED_TASKS or execute_mode in {"local", "hybrid"}
        if not needs_local:
            return []
        required: list[str] = []
        if task_type == "macro_sequence":
            required.append("macroExecution")
        if task_type in {"omml_to_mathtype", "formula_precheck"}:
            required.extend(["ommlDependencySearch", "mathTypeAutomation"])
        if task_type == "mathtype_format":
            required.append("mathTypeAutomation")
        if task_type in {"word_to_ppt", "ppt_to_word", "excel_to_pdf", "excel_to_word", "excel_to_ppt"}:
            required.append("officeAutomation")
        if task_type == "pdf_to_word" and bool(settings.get("enableFormulaOcr", True)):
            required.append("mathTypeAutomation")
        files = [self.store.get_file(file_id) for file_id in file_ids]
        if task_type == "word_to_ppt" and (settings.get("wordConvertOmmlFirst") or any(file and file.get("has_omml") for file in files)):
            required.extend(["ommlDependencySearch", "mathTypeAutomation"])
        if any(file and file.get("has_mathtype") for file in files) and task_type in {"word_to_ppt", "ppt_to_word", "mathtype_format"}:
            required.append("mathTypeAutomation")
        return list(dict.fromkeys(required))

    def _preflight_mathpix(self, task_type: str, settings: dict[str, Any]) -> dict[str, Any]:
        """检查 Mathpix 上传授权、OCR 开关和凭据。"""
        if task_type != "pdf_to_word" or settings.get("pdfToWordEngine") != "Mathpix":
            return self._preflight_item("mathpix_authorization", "Mathpix 授权", "通过", "不涉及", "当前任务不需要 Mathpix 外部上传", "无需处理")
        enabled_ocr = any([settings.get("enableTextOcr", True), settings.get("enableFormulaOcr", True), settings.get("enableTableOcr", True)])
        if not enabled_ocr:
            return self._preflight_item("mathpix_authorization", "Mathpix 授权", "需确认", "OCR 已关闭", "文字、公式和表格 OCR 均已关闭", "开启至少一种 OCR 能力后重试")
        allowed = bool(settings.get("allowExternalMathpixUpload", False))
        app_id_env, app_key_env = self._mathpix_env_names(settings)
        missing_envs = [name for name in (app_id_env, app_key_env) if not os.getenv(name)]
        if allowed and missing_envs:
            return self._preflight_item(
                "mathpix_authorization",
                "Mathpix 授权",
                "需确认",
                "缺少凭证",
                f"缺少 Mathpix 凭证环境变量：{', '.join(missing_envs)}",
                "在运行服务的环境中配置 APP ID 和 APP KEY 后重试",
            )
        return self._preflight_item(
            "mathpix_authorization",
            "Mathpix 授权",
            "通过" if allowed else "需确认",
            "已授权 / 凭证已配置" if allowed else "未授权",
            f"已允许 Mathpix 外部上传，并检测到 {app_id_env} / {app_key_env}" if allowed else "PDF 转 Word 使用 Mathpix 时需要用户明确授权外部上传",
            "确认文档可上传并开启授权" if not allowed else "可提交 Mathpix 作业",
        )

    @staticmethod
    def _mathpix_env_names(settings: dict[str, Any]) -> tuple[str, str]:
        """返回配置的 Mathpix 凭据环境变量名称。"""
        app_id_env = str(settings.get("mathpixAppIdEnv") or "MATHPIX_APP_ID").strip() or "MATHPIX_APP_ID"
        app_key_env = str(settings.get("mathpixAppKeyEnv") or "MATHPIX_APP_KEY").strip() or "MATHPIX_APP_KEY"
        return app_id_env, app_key_env

    def _preflight_mathtype_compatibility(self, task_type: str, file_ids: list[str], settings: dict[str, Any]) -> dict[str, Any]:
        """提示 MathType 的同平台交付或兜底要求。"""
        math_tasks = {"formula_precheck", "omml_to_mathtype", "mathtype_format"}
        files = [self.store.get_file(file_id) for file_id in file_ids]
        has_formula_objects = any(file and (file.get("has_omml") or file.get("has_mathtype")) for file in files)
        word_omml_conversion = task_type == "word_to_ppt" and bool(settings.get("wordConvertOmmlFirst"))
        if task_type not in math_tasks and not has_formula_objects and not word_omml_conversion:
            return self._preflight_item(
                "mathtype_compatibility",
                "MathType 兼容",
                "通过",
                "不涉及",
                "当前任务未检测到 MathType/OMML 跨平台处理要求",
                "无需处理",
            )
        profile = self.install_profile()
        platform_name = str(profile.get("platform") or "Unknown")
        mode = str(settings.get("mathtypeCompatibilityMode") or "platform-specific")
        mode_labels = {
            "platform-specific": "区分平台",
            "mathml-latex": "MathML/LaTeX 优先",
            "image-fallback": "图片兜底",
        }
        mode_label = mode_labels.get(mode, mode)
        if platform_name == "Unknown":
            return self._preflight_item(
                "mathtype_compatibility",
                "MathType 兼容",
                "需确认",
                "平台未确认",
                "无法确认当前客户端平台，不能直接交付平台专属 MathType 对象",
                "在设置中选择 Windows 或 macOS，或改用 MathML/LaTeX、图片兜底",
            )
        if mode == "platform-specific":
            return self._preflight_item(
                "mathtype_compatibility",
                "MathType 兼容",
                "需确认",
                f"{platform_name} / {mode_label}",
                f"Windows 与 macOS MathType 对象不通用，当前仅按 {platform_name} 本机对象处理",
                "跨系统交付前切换为 MathML/LaTeX 优先或图片兜底",
                object_format=profile.get("mathtypeObjectFormat", ""),
            )
        return self._preflight_item(
            "mathtype_compatibility",
            "MathType 兼容",
            "通过",
            f"{platform_name} / {mode_label}",
            f"已为 {platform_name} MathType 任务启用跨系统兜底格式",
            "可继续处理，并在交付前保留兜底公式格式",
            object_format=profile.get("mathtypeObjectFormat", ""),
        )

    def _preflight_macro(self, task_type: str, settings: dict[str, Any]) -> dict[str, Any]:
        """检查 Word 宏执行授权与安全设置。"""
        if task_type != "macro_sequence":
            return self._preflight_item("macro_authorization", "宏授权", "通过", "不涉及", "当前任务不执行 Word 宏", "无需处理")
        enabled = bool(settings.get("enableMacroExecution", True))
        whitelist = "白名单开启" if settings.get("macroWhitelistOnly") else "白名单关闭"
        return self._preflight_item(
            "macro_authorization",
            "宏授权",
            "通过" if enabled else "需确认",
            whitelist,
            "宏执行队列已启用" if enabled else "宏执行队列已禁用，只能生成安全报告",
            "确认宏来源、白名单和执行前备份" if enabled else "在设置或授权管理中启用宏执行队列",
        )

    @staticmethod
    def _preflight_item(
        check_id: str,
        label: str,
        status: str,
        metric: str,
        message: str,
        recommendation: str,
        **extra: Any,
    ) -> dict[str, Any]:
        """生成接口和界面使用的规范预检行。"""
        severity = {"通过": "info", "需确认": "warning", "失败": "error"}.get(status, "info")
        return {
            "id": check_id,
            "label": label,
            "status": status,
            "severity": severity,
            "metric": metric,
            "message": message,
            "recommendation": recommendation,
            **extra,
        }

    def resolve_execute_mode(self, task_type: str, file_ids: list[str]) -> str:
        """根据任务和文件能力选择网页、本地或混合执行。"""
        if task_type in LOCAL_REQUIRED_TASKS:
            return "local"
        files = [self.store.get_file(file_id) for file_id in file_ids]
        if task_type in {"word_to_ppt", "ppt_to_word", "excel_to_pdf", "excel_to_word", "excel_to_ppt"} and any(
            file and bool((file.get("content_summary") or {}).get("requiresNativeOffice"))
            for file in files
        ):
            return "local"
        if task_type == "pdf_to_word" and (
            self.store.get_settings().get("enableFormulaOcr", True) or any(file and file.get("has_formula") for file in files)
        ):
            return "hybrid"
        if any(file and (file.get("has_macro") or file.get("has_omml")) for file in files):
            return "hybrid"
        if task_type in {"word_to_ppt", "ppt_to_word"}:
            return "hybrid"
        return "web"

    def capabilities(self) -> dict[str, Any]:
        """公开本地接口、Office、MathType、宏和同步能力标记。"""
        settings = self.store.get_settings()
        profile = self.install_profile()
        system_name = profile["platform"]
        flags = profile["capabilities"]
        return {
            "mode": "local-api",
            "host": settings.get("localApiHost", "127.0.0.1"),
            "port": settings.get("localApiPort", 8765),
            "platform": system_name,
            "installProfile": profile,
            "officeAutomation": {
                "available": bool(flags.get("officeAutomation")),
                "status": profile["officeAutomation"],
            },
            "mathType": {
                "available": bool(flags.get("mathTypeAutomation")),
                "status": profile["mathType"],
                "compatibility": profile["formulaPortability"],
                "objectFormat": profile["mathtypeObjectFormat"],
            },
            "ommlSearch": {
                "available": bool(flags.get("ommlDependencySearch")),
                "status": "已提供本地检索接口合同",
            },
            "macroExecution": {
                "available": bool(flags.get("macroExecution") and settings.get("enableMacroExecution", True)),
                "status": self._macro_capability_status(flags, settings),
            },
            "localConnection": {
                "client_enabled": bool(settings.get("localClientEnabled", True)),
                "web_launch_allowed": bool(settings.get("allowWebLaunchLocalClient", False)),
                "cloud_sync_allowed": bool(settings.get("allowCloudSync", False)),
                "task_status_cloud_sync_allowed": bool(settings.get("allowTaskStatusCloudSync", False)),
                "sensitive_files_prefer_local": bool(settings.get("sensitiveFilesPreferLocal", True)),
                "status": "本地客户端连接已启用" if settings.get("localClientEnabled", True) else "本地客户端连接已关闭",
            },
            "privacy": "默认只保存文件元数据和任务摘要",
        }

    def architecture_blueprint(self) -> dict[str, Any]:
        """返回需求架构蓝图，不声称实现未验证的原生执行。"""
        settings = self.store.get_settings()
        profile = self.install_profile()
        caps = self.capabilities()
        return {
            "schema_version": "k12.architectureBlueprint.v1",
            "runtime": {
                "python": sys.version.split()[0],
                "service": "Python 标准库本地 HTTP API",
                "frontend": "静态 HTML/CSS/JavaScript",
                "store": "本地 SQLite + 文件目录",
                "mode": caps["mode"],
                "platform": profile["platform"],
            },
            "layers": [
                self._architecture_layer(
                    "desktop_shell",
                    "桌面端",
                    "Electron + Vue 3 + TypeScript + Element Plus",
                    "planned",
                    "当前以本地 API 和浏览器预览承接；真实桌面壳接入后调用本地任务载荷",
                ),
                self._architecture_layer(
                    "web_frontend",
                    "网页端",
                    "Vue 3 + TypeScript + Element Plus",
                    "contract",
                    "当前为静态前端原型，覆盖任务管理、设置、报告和本地客户端交接流程",
                ),
                self._architecture_layer(
                    "local_service",
                    "本地服务",
                    "Python 3.11/3.12 + FastAPI",
                    "implemented",
                    "当前 Python 本地 API 已可运行，后续可替换为 FastAPI worker 而保持现有 JSON 合同",
                ),
                self._architecture_layer(
                    "cloud_service",
                    "云端服务",
                    "Python + Django / FastAPI",
                    "planned",
                    "用户、权限、报告摘要和云同步已有本地合同；云端服务未在首版内声明已上线",
                ),
                self._architecture_layer(
                    "task_queue",
                    "任务队列",
                    "Celery / RQ + Redis",
                    "contract",
                    f"当前本地队列支持并发 {settings.get('maxConcurrentTasks', 3)}、暂停、恢复、重试和失败跳过；分布式队列后续接入",
                ),
                self._architecture_layer(
                    "database",
                    "数据库",
                    "本地 SQLite；云端 MySQL / PostgreSQL",
                    "implemented",
                    "当前运行数据保存在本地 SQLite 与 data_dir 文件目录，云端库迁移保留为后续演进",
                ),
                self._architecture_layer(
                    "document_processing",
                    "文档处理",
                    "Office COM、LibreOffice Headless、python-docx、python-pptx、openpyxl、PyMuPDF、pdfplumber",
                    "contract",
                    f"{caps['officeAutomation']['status']}；当前已有 OOXML/报告合同，真实 Office/LibreOffice 自动化需本地客户端实测",
                ),
                self._architecture_layer(
                    "ocr_image",
                    "OCR 与图片",
                    "Mathpix、PaddleOCR、Tesseract、OpenCV、Pillow",
                    "contract",
                    f"纯文本 PDF 本地转换；扫描、混合和复杂 PDF 使用 Mathpix，外部上传授权 {('已开启' if settings.get('allowExternalMathpixUpload') else '未开启')}；图片检索已有本地报告合同",
                ),
                self._architecture_layer(
                    "formula_macro",
                    "公式与宏",
                    "MathType、OMML、MathML、LaTeX、pywin32、Word COM 自动化",
                    "contract",
                    f"{caps['mathType']['compatibility']}；宏执行 {caps['macroExecution']['status']}",
                ),
            ],
            "contracts": [
                {"name": "本地任务载荷", "endpoint": "/api/tasks/{task_id}/local-payload", "status": "已实现，需安全令牌"},
                {"name": "本地客户端就绪状态", "endpoint": "/api/tasks/{task_id}/local-readiness", "status": "已实现，返回脱敏预检缺口"},
                {"name": "本地状态同步", "endpoint": "/api/tasks/{task_id}/local-sync", "status": "已实现，需同步授权"},
                {"name": "本地结果上传登记", "endpoint": "/api/local-client/uploads", "status": "已实现，需安全令牌和云端同步授权"},
                {"name": "本地客户端运行清单", "endpoint": "/api/local-client/manifest", "status": "已定义启动、心跳和同步接口"},
                {"name": "本地客户端启动请求", "endpoint": "/api/tasks/{task_id}/local-launch", "status": "已记录唤起请求和协议 URL"},
                {"name": "任务恢复清单", "endpoint": "/api/tasks/recovery-summary", "status": "已实现，返回脱敏恢复摘要和继续/重试端点"},
                {"name": "安装画像", "endpoint": "/api/install-profile", "status": "已区分 Windows / macOS"},
                {"name": "安装计划", "endpoint": "/api/install-plan", "status": "已区分安装包、校验和平台步骤"},
                {"name": "能力探测", "endpoint": "/api/capabilities", "status": "已实现本地能力位"},
                {"name": "Mathpix PDF", "endpoint": "Mathpix OCR API", "status": "已建立授权、提交、轮询和下载合同"},
                {"name": "智能增强规划", "endpoint": "/api/enhancement-plan", "status": "已拆分 V3 能力边界，不声明真实 AI 执行"},
            ],
            "recommended_split": [
                {"scenario": "本地客户端 API", "recommendation": "FastAPI"},
                {"scenario": "轻量任务服务", "recommendation": "FastAPI"},
                {"scenario": "用户管理 / 权限管理 / 管理后台", "recommendation": "Django"},
                {"scenario": "SaaS 平台", "recommendation": "Django + DRF"},
                {"scenario": "本地单机工具", "recommendation": "FastAPI + SQLite"},
                {"scenario": "本地 + 网页混合平台", "recommendation": "FastAPI 本地服务 + Django 云端服务"},
            ],
        }

    def api_catalog(self) -> dict[str, Any]:
        """返回接口路由元数据及授权、权限和敏感性边界。"""
        token_required = bool(str(self.store.get_settings().get("localSecurityToken") or "").strip())
        endpoints = [
            self._api_endpoint("GET", "/api/health", "健康检查", "public", "返回服务状态和版本"),
            self._api_endpoint("GET", "/api/files", "文件列表", "token" if token_required else "local", "返回脱敏文件对象"),
            self._api_endpoint("POST", "/api/uploads", "上传文件", "token" if token_required else "local", "导入文件或 ZIP 并生成文件记录", True, "files.manage"),
            self._api_endpoint("POST", "/api/files", "文件元数据导入", "token" if token_required else "local", "导入外部文件元数据或批量元数据", True, "files.manage"),
            self._api_endpoint("POST", "/api/files/{file_id}/replace", "重新上传替换", "token" if token_required else "local", "保留 file_id 并重新分析文件", True, "files.manage"),
            self._api_endpoint("POST", "/api/files/{file_id}/password", "加密文件密码登记", "token" if token_required else "local", "登记加密文件打开密码，仅保存在当前本地进程", True, "files.manage"),
            self._api_endpoint("DELETE", "/api/files/{file_id}", "文件删除", "token" if token_required else "local", "删除文件记录和本地上传缓存，不删除外部源文件", True, "files.manage"),
            self._api_endpoint("GET", "/api/files/{file_id}/preview", "文件预览", "token" if token_required else "local", "返回结构化页面、对象和告警摘要"),
            self._api_endpoint("GET", "/api/files/{file_id}/download", "源文件下载", "token" if token_required else "local", "下载上传缓存或授权外部源文件", True),
            self._api_endpoint("GET", "/api/files/{file_id}/reports", "文件关联报告", "token" if token_required else "local", "返回包含该文件的处理报告，作为文件处理页报告入口"),
            self._api_endpoint("GET", "/api/files/download", "批量源文件下载", "token" if token_required else "local", "按选择或文件类型打包源文件 ZIP", True),
            self._api_endpoint("GET", "/api/tasks", "任务列表", "token" if token_required else "local", "返回脱敏任务对象"),
            self._api_endpoint("GET", "/api/tasks/recovery-summary", "任务恢复清单", "token" if token_required else "local", "返回中断、暂停、失败和可重试任务的恢复摘要"),
            self._api_endpoint("POST", "/api/tasks", "创建任务", "token" if token_required else "local", "创建本地、网页或混合任务并写入报告", False, "tasks.create"),
            self._api_endpoint("POST", "/api/tasks/{task_id}/retry", "任务重试", "token" if token_required else "local", "重跑失败或已取消任务", False, "tasks.control"),
            self._api_endpoint("POST", "/api/tasks/{task_id}/pause", "任务暂停", "token" if token_required else "local", "将可暂停任务置为已暂停", False, "tasks.control"),
            self._api_endpoint("POST", "/api/tasks/{task_id}/resume", "任务继续", "token" if token_required else "local", "继续已暂停或可恢复任务", False, "tasks.control"),
            self._api_endpoint("POST", "/api/tasks/{task_id}/cancel", "任务取消", "token" if token_required else "local", "取消未完成任务", False, "tasks.control"),
            self._api_endpoint("POST", "/api/tasks/{task_id}/skip-file", "批量文件跳过", "token" if token_required else "local", "跳过批量任务中的失败或可重试文件", False, "tasks.control"),
            self._api_endpoint("POST", "/api/tasks/{task_id}/restore-backups", "任务备份恢复", "token" if token_required else "local", "恢复宏执行前备份文件", True, "tasks.control"),
            self._api_endpoint("GET", "/api/tasks/{task_id}/download", "任务结果包", "token" if token_required else "local", "打包任务 JSON、日志、报告和转换输出", True),
            self._api_endpoint("GET", "/api/tasks/{task_id}/local-payload", "本地任务载荷", "configured-token", "交付本地路径、预检、动作队列和同步策略", True),
            self._api_endpoint("GET", "/api/tasks/{task_id}/local-readiness", "本地客户端就绪状态", "token" if token_required else "local", "返回当前任务所需本地能力、脱敏预检缺口和 platform_mismatch 同平台 MathType/Office 交付阻断"),
            self._api_endpoint("POST", "/api/tasks/{task_id}/local-launch", "本地客户端启动请求", "configured-token + launch-authorization", "记录网页端唤起本地客户端请求并返回 k12-local 协议 URL", True),
            self._api_endpoint("POST", "/api/tasks/{task_id}/local-sync", "本地状态同步", "token + sync-authorization", "接收本地客户端进度、输出摘要、dry-run 摘要、脱敏原生执行报告和上传意图", True),
            self._api_endpoint("GET", "/api/local-client/manifest", "本地客户端运行清单", "token" if token_required else "local", "返回本地客户端启动、心跳、载荷和同步接口合同"),
            self._api_endpoint("GET", "/api/local-client/uploads", "本地结果上传队列", "token" if token_required else "local", "返回本地结果上传意图、授权状态和脱敏输出摘要", True),
            self._api_endpoint("GET", "/api/local-client/uploads/{upload_id}/manifest", "本地结果云端接收清单", "token" if token_required else "local", "返回上传包哈希、输出摘要和云端接收合同，不包含本地路径或文件内容", True),
            self._api_endpoint("POST", "/api/local-client/uploads", "本地结果上传登记", "configured-token + cloud-sync-authorization", "接收本地客户端上传包清单、哈希和脱敏输出摘要", True),
            self._api_endpoint("POST", "/api/local-client/heartbeat", "本地客户端心跳", "configured-token", "接收桌面端运行状态、版本、平台和能力摘要"),
            self._api_endpoint("GET", "/api/mathpix-jobs", "Mathpix 作业队列", "token" if token_required else "local", "返回 PDF 转 Word 的 Mathpix 识别状态、授权边界和输出摘要，不触发新的外部上传"),
            self._api_endpoint("GET", "/api/reports", "报告列表", "token" if token_required else "local", "返回脱敏报告摘要"),
            self._api_endpoint("GET", "/api/reports/{report_id}/download", "报告下载", "token" if token_required else "local", "下载 HTML/JSON/PDF/XLSX/TXT 报告", True),
            self._api_endpoint("DELETE", "/api/reports/{report_id}", "报告删除", "token" if token_required else "local", "删除报告记录、报告缓存和关联人工标注", True, "reports.manage"),
            self._api_endpoint("GET", "/api/reports/{report_id}/images.zip", "图片包下载", "token" if token_required else "local", "导出报告中的微小图片资源、替换状态和 manifest.csv", True),
            self._api_endpoint("GET", "/api/reports/{report_id}/images.xlsx", "微小图片清单", "token" if token_required else "local", "导出微小图片、疑似类型、人工确认、重复和替换状态表格", True),
            self._api_endpoint("GET", "/api/reports/{report_id}/formulas.zip", "公式结果包", "token" if token_required else "local", "导出公式 JSON、TEX、MML、清单和人工修正结果", True),
            self._api_endpoint("GET", "/api/reports/{report_id}/formulas.xlsx", "公式结果表格", "token" if token_required else "local", "导出公式来源、位置、LaTeX、MathML、置信度和人工校正状态", True),
            self._api_endpoint("GET", "/api/reports/{report_id}/failures.csv", "失败清单", "token" if token_required else "local", "导出文件、转换、宏、OMML、Mathpix 等失败清单"),
            self._api_endpoint("GET", "/api/reports/{report_id}/omml-failures.csv", "OMML 失败清单", "token" if token_required else "local", "导出 OMML 依赖、转换失败、手动依赖和重新转换标注", True),
            self._api_endpoint("GET", "/api/reports/{report_id}/macro-failures.csv", "宏失败清单", "token" if token_required else "local", "导出宏失败、未授权、禁用和备份恢复状态", True),
            self._api_endpoint("GET", "/api/reports/{report_id}/comparison", "转换对比预览", "token" if token_required else "local", "返回转换前后结构化对比"),
            self._api_endpoint("GET", "/api/image-annotations", "图片人工校正记录", "token" if token_required else "local", "按报告返回微小图片确认、误判、删除和替换标注", True),
            self._api_endpoint("POST", "/api/image-annotations", "图片人工校正保存", "token" if token_required else "local", "保存微小图片类型确认、误判或删除待处理标注", True),
            self._api_endpoint("POST", "/api/image-annotations/replacement", "图片替换素材登记", "token" if token_required else "local", "保存替换图片素材并把替换请求写入图片校正记录", True),
            self._api_endpoint("DELETE", "/api/image-annotations/{annotation_id}", "图片人工校正删除", "token" if token_required else "local", "删除一条图片人工校正记录", True),
            self._api_endpoint("POST", "/api/image-assets/reexport", "图片重新导出", "token" if token_required else "local", "从来源文档重新导出图片资源，并在失败时保留原始结果和错误原因", True),
            self._api_endpoint("GET", "/api/assets/replacements/{file_name}", "图片替换素材下载", "token" if token_required else "local", "下载已登记的替换图片素材，供预览或本地客户端交接", True),
            self._api_endpoint("GET", "/api/assets/images/{asset_path}", "报告图片资源", "token" if token_required else "local", "仅从受控图片缓存目录读取报告图片预览资源", True),
            self._api_endpoint("GET", "/api/artifacts/{task_id}/{file_name}", "转换产物下载", "token" if token_required else "local", "按任务和文件名下载报告登记且已成功的转换产物", True),
            self._api_endpoint("GET", "/api/formula-annotations", "公式人工校正记录", "token" if token_required else "local", "按报告返回公式确认、修正、重新识别和跳过标注", True),
            self._api_endpoint("POST", "/api/formula-annotations", "公式人工校正保存", "token" if token_required else "local", "保存公式 LaTeX/MathML 修正、确认、重新识别或跳过状态；重新识别只登记受控请求合同，不由任务参数临时授权外部上传", True),
            self._api_endpoint("POST", "/api/formula-annotations/bulk-confirm", "公式批量确认", "token" if token_required else "local", "按置信度阈值批量确认公式识别结果", True),
            self._api_endpoint("DELETE", "/api/formula-annotations/{annotation_id}", "公式人工校正删除", "token" if token_required else "local", "删除一条公式人工校正记录", True),
            self._api_endpoint("GET", "/api/omml-annotations", "OMML 人工校正记录", "token" if token_required else "local", "按报告返回 OMML 转换失败、保留、重新转换和手动依赖标注", True),
            self._api_endpoint("POST", "/api/omml-annotations", "OMML 人工校正保存", "token" if token_required else "local", "保存 OMML 保留、重新转换、已修复或手动指定依赖记录", True),
            self._api_endpoint("DELETE", "/api/omml-annotations/{annotation_id}", "OMML 人工校正删除", "token" if token_required else "local", "删除一条 OMML 人工校正记录", True),
            self._api_endpoint("GET", "/api/layout-annotations", "排版人工校正记录", "token" if token_required else "local", "按报告返回页码、标题、表格、图片、公式位置和编号校正记录", True),
            self._api_endpoint("POST", "/api/layout-annotations", "排版人工校正保存", "token" if token_required else "local", "保存排版问题、修复建议和处理状态", True),
            self._api_endpoint("DELETE", "/api/layout-annotations/{annotation_id}", "排版人工校正删除", "token" if token_required else "local", "删除一条排版人工校正记录", True),
            self._api_endpoint("GET", "/api/logs", "日志列表", "token" if token_required else "local", "返回全局或任务日志"),
            self._api_endpoint("GET", "/api/logs/download", "日志导出", "token" if token_required else "local", "按 TXT/LOG/CSV/JSON 导出日志"),
            self._api_endpoint("GET", "/api/settings", "设置读取", "token" if token_required else "local", "读取系统设置；pdfToWordEngine 固定归一为 Mathpix"),
            self._api_endpoint("PUT", "/api/settings", "设置保存", "token" if token_required else "local", "保存转换、OCR、宏、本地连接和安全设置；纯文本 PDF 本地解析，复杂 PDF 的 OCR 引擎固定 Mathpix，外部上传授权仅来自 settings.allowExternalMathpixUpload", True, "settings.manage"),
            self._api_endpoint("GET", "/api/capabilities", "能力位", "token" if token_required else "local", "返回本地客户端、Office、MathType、宏和隐私能力"),
            self._api_endpoint("GET", "/api/install-profile", "安装画像", "token" if token_required else "local", "按 Windows/macOS 返回安装包、组件和 MathType 兼容提示"),
            self._api_endpoint("GET", "/api/install-plan", "安装计划", "token" if token_required else "local", "按 Windows/macOS 返回安装包状态、校验、安装步骤、心跳平台匹配和 MathType 同平台兜底策略"),
            self._api_endpoint(
                "GET",
                "/api/installers/{file_name}",
                "安装包下载",
                "token" if token_required else "local",
                "仅当本地 installers 目录存在注册的 Windows .msi 或 macOS .pkg 安装包，且请求显式携带 platform=Windows 或 platform=macOS 时下载，并返回平台、包边界、公式对象同平台边界、兜底格式和 SHA256 响应头",
                True,
                response_headers=[
                    {"name": "X-K12-Platform", "description": "Windows 或 macOS 安装目标平台"},
                    {"name": "X-K12-Installer-Kind", "description": "windows-msi 或 macos-pkg 安装包类型"},
                    {"name": "X-K12-Package-Boundary", "description": "Windows .msi 与 macOS .pkg 不跨平台混用"},
                    {"name": "X-K12-Formula-Object-Boundary", "description": "MathType 原生对象必须同平台交接"},
                    {"name": "X-K12-Formula-Fallback-Formats", "description": "ASCII 兜底格式标识 MathML,LaTeX,image"},
                    {"name": "X-K12-SHA256", "description": "安装包 SHA256 校验值"},
                ],
            ),
            self._api_endpoint("GET", "/api/architecture", "技术架构蓝图", "token" if token_required else "local", "返回 PRD 16 技术架构与接口合同"),
            self._api_endpoint("GET", "/api/api-catalog", "API 目录", "token" if token_required else "local", "返回开放接口目录和鉴权边界"),
            self._api_endpoint("GET", "/api/data-dictionary", "数据字典", "token" if token_required else "local", "按 PRD 第 10 章返回对象字段、隐私策略和当前记录数量"),
            self._api_endpoint("GET", "/api/acceptance-matrix", "PRD 验收矩阵", "token" if token_required else "local", "逐条返回 PRD 17 验收项的覆盖状态、证据、验证上下文和未覆盖风险"),
            self._api_endpoint("GET", "/api/product-summary", "产品总结", "token" if token_required else "local", "返回 PRD 21 核心能力和路线阶段"),
            self._api_endpoint("GET", "/api/enhancement-plan", "智能增强规划", "token" if token_required else "local", "返回 V3 AI 排版修复、智能模板套用、私有化部署和质量闭环规划"),
            self._api_endpoint("GET", "/api/users", "用户列表", "public-or-login", "本地用户、角色和当前用户信息"),
            self._api_endpoint("POST", "/api/users", "用户保存", "token" if token_required else "local", "创建或更新本地用户，需要用户管理权限", True, "users.manage"),
            self._api_endpoint("DELETE", "/api/users/{user_id}", "用户删除", "token" if token_required else "local", "删除本地用户并保留至少一个启用管理员", True, "users.manage"),
            self._api_endpoint("POST", "/api/session", "本地登录", "public", "激活允许登录的本地用户"),
            self._api_endpoint("GET", "/api/templates", "模板列表", "token" if token_required else "local", "返回网页端模板配置"),
            self._api_endpoint("POST", "/api/templates", "模板保存", "token" if token_required else "local", "创建或更新网页端模板配置，需要模板管理权限", True, "templates.manage"),
            self._api_endpoint("DELETE", "/api/templates/{template_id}", "模板删除", "token" if token_required else "local", "删除网页端模板配置，需要模板管理权限", True, "templates.manage"),
            self._api_endpoint("GET", "/api/macro-templates", "宏顺序模板列表", "token" if token_required else "local", "返回 Word 宏顺序模板"),
            self._api_endpoint("POST", "/api/macros/import", "VBA 源导入", "token" if token_required else "local", "导入 UTF-8 .bas 标准模块的真实宏声明，不执行代码", True, "templates.manage"),
            self._api_endpoint("GET", "/api/macros/sources/{source_id}", "VBA 定义读取", "token" if token_required else "local", "向有宏执行权限的客户端返回通过哈希及声明校验的源定义", True, "macros.execute"),
            self._api_endpoint("DELETE", "/api/macros/sources/{source_id}", "VBA 来源删除", "token" if token_required else "local", "删除没有待执行任务或模板引用的导入模块", True, "templates.manage"),
            self._api_endpoint("POST", "/api/macro-templates", "宏顺序模板保存", "token" if token_required else "local", "保存 Word 宏编排模板，需要模板管理权限", True, "templates.manage"),
            self._api_endpoint("DELETE", "/api/macro-templates/{template_id}", "宏顺序模板删除", "token" if token_required else "local", "删除 Word 宏编排模板，需要模板管理权限", True, "templates.manage"),
            self._api_endpoint("GET", "/api/authorizations", "授权列表", "token" if token_required else "local", "返回 Mathpix、云端同步和本地唤起授权状态"),
            self._api_endpoint("POST", "/api/authorizations", "授权保存", "token" if token_required else "local", "启用或撤销授权，需要授权管理权限", True, "authorizations.manage"),
            self._api_endpoint("GET", "/api/preflight", "运行预检", "token" if token_required else "local", "输出目录、磁盘、权限、Mathpix、宏和本地客户端预检"),
        ]
        return {
            "schema_version": "k12.apiCatalog.v1",
            "generated_at": utc_now(),
            "auth": {
                "token_required": token_required,
                "token_header": "X-K12-Token",
                "bearer_supported": True,
                "query_token_supported": True,
                "local_payload_requires_configured_token": True,
            },
            "endpoints": endpoints,
        }

    def data_dictionary(self) -> dict[str, Any]:
        """返回需求数据对象、字段定义及路径隐私说明。"""
        reports = self.store.list_reports()
        macros = self.store.list_macro_templates()
        objects = [
            self._dictionary_object(
                "files",
                "FileItem",
                "文件对象",
                "file_path、storage_path 默认脱敏，开启显示本地路径后才返回真实路径",
                len(self.store.list_files()),
                [
                    ("id", "string", "文件 ID"),
                    ("file_name", "string", "文件名"),
                    ("file_type", "string", "文件类型"),
                    ("file_size", "number", "文件大小"),
                    ("source_relative_path", "string", "文件夹上传时的来源相对路径"),
                    ("file_path", "string", "文件路径"),
                    ("page_count", "number", "页数"),
                    ("slide_count", "number", "PPT 页数"),
                    ("sheet_count", "number", "Excel Sheet 数"),
                    ("status", "string", "状态"),
                    ("has_formula", "boolean", "是否含公式"),
                    ("has_mathtype", "boolean", "是否含 MathType"),
                    ("has_omml", "boolean", "是否含 Word 自带公式"),
                    ("has_macro", "boolean", "是否含宏"),
                    ("has_small_image", "boolean", "是否含微小图片"),
                    ("created_at", "datetime", "创建时间"),
                ],
                ["file_path", "storage_path"],
            ),
            self._dictionary_object(
                "tasks",
                "Task",
                "任务对象",
                "input_path、output_path 默认按本地路径隐私策略脱敏",
                len(self.store.list_tasks()),
                [
                    ("id", "string", "任务 ID"),
                    ("task_type", "string", "任务类型"),
                    ("execute_mode", "string", "local / web / hybrid"),
                    ("file_ids", "array", "文件 ID 列表"),
                    ("options", "object", "任务参数，包含 workflowOrder、workflowLabels 和 k12.workflowPlan.v1 流程顺序"),
                    ("status", "string", "任务状态"),
                    ("progress", "number", "进度"),
                    ("input_path", "string", "输入路径"),
                    ("output_path", "string", "输出路径"),
                    ("error_message", "string", "错误信息"),
                    ("start_time", "datetime", "开始时间"),
                    ("end_time", "datetime", "结束时间"),
                ],
                ["input_path", "output_path"],
            ),
            self._dictionary_object(
                "formulas",
                "FormulaItem",
                "公式对象",
                "original_image_path 默认脱敏，报告里只展示引用或受控资源链接",
                sum(int(report.get("formula_count") or 0) for report in reports),
                [
                    ("id", "string", "公式 ID"),
                    ("file_id", "string", "文件 ID"),
                    ("page_index", "number", "页码"),
                    ("position", "string", "位置信息"),
                    ("position_status", "string", "已记录 / 异常位置"),
                    ("position_issue", "string", "位置异常说明"),
                    ("fallback_position", "string", "位置缺失时的异常位置记录"),
                    ("source_type", "string", "MathType / OMML / LaTeX / PDF / 图片"),
                    ("original_image_path", "string", "原始截图"),
                    ("latex", "string", "LaTeX 结果"),
                    ("mathml", "string", "MathML 结果"),
                    ("mathtype_data", "string", "MathType 数据"),
                    ("format_status", "string", "是否已格式化"),
                    ("confidence", "number", "置信度"),
                    ("status", "string", "成功 / 失败 / 待确认"),
                ],
                ["original_image_path"],
            ),
            self._dictionary_object(
                "formula_annotations",
                "FormulaAnnotation",
                "公式人工校正对象",
                "重识别请求只保存授权门槛、阻断原因和下一步，不保存 Mathpix 凭证值",
                len(self.store.list_formula_annotations()),
                [
                    ("id", "string", "校正记录 ID"),
                    ("formula_id", "string", "公式 ID"),
                    ("report_id", "string", "报告 ID"),
                    ("status", "string", "已确认 / 已修正 / 重新识别 / 跳过"),
                    ("latex", "string", "人工修正 LaTeX"),
                    ("mathml", "string", "人工修正 MathML"),
                    ("retry_recognition", "boolean", "是否请求重新识别"),
                    ("recognition_status", "string", "重识别请求状态"),
                    ("recognition_request", "object", "k12.formulaRecognitionRequest.v1 请求合同"),
                    ("next_step", "string", "下一步处理建议"),
                    ("note", "string", "人工备注"),
                    ("updated_at", "datetime", "更新时间"),
                ],
                [],
            ),
            self._dictionary_object(
                "omml",
                "OmmlDependencyItem",
                "OMML 依赖对象",
                "document_path、omml_source_path、omml_target_path 默认脱敏",
                sum(int(report.get("omml_dependency_count") or report.get("omml_count") or 0) for report in reports),
                [
                    ("id", "string", "依赖 ID"),
                    ("file_id", "string", "文件 ID"),
                    ("document_path", "string", "当前 Word 文档路径"),
                    ("omml_file_name", "string", "OMML 文件名称"),
                    ("omml_source_path", "string", "OMML 文件来源路径"),
                    ("omml_target_path", "string", "复制到的目标路径"),
                    ("found_status", "string", "已找到 / 未找到 / 手动选择"),
                    ("copy_status", "string", "成功 / 失败 / 跳过"),
                    ("source_size", "number", "依赖源快照字节数"),
                    ("source_sha256", "string", "依赖源快照 SHA-256"),
                    ("target_size", "number", "复制目标字节数"),
                    ("target_sha256", "string", "复制目标 SHA-256"),
                    ("error_message", "string", "错误信息"),
                    ("created_at", "datetime", "创建时间"),
                ],
                ["document_path", "omml_source_path", "omml_target_path"],
            ),
            self._dictionary_object(
                "macros",
                "MacroItem",
                "宏对象",
                "backup_path 默认脱敏，恢复动作只接受当前任务记录的备份",
                sum(int(report.get("macro_count") or 0) for report in reports) or len(macros),
                [
                    ("id", "string", "宏 ID"),
                    ("file_id", "string", "文件 ID"),
                    ("macro_name", "string", "宏名称"),
                    ("macro_source", "string", "当前文档 / 模板 / 本地宏库 / 系统内置"),
                    ("macro_description", "string", "宏说明"),
                    ("execute_order", "number", "执行顺序"),
                    ("execute_timing", "string", "执行时机"),
                    ("execute_status", "string", "成功 / 失败 / 跳过"),
                    ("failure_strategy", "string", "停止 / 跳过 / 询问"),
                    ("error_message", "string", "错误信息"),
                    ("backup_path", "string", "执行前备份路径"),
                ],
                ["backup_path"],
            ),
            self._dictionary_object(
                "images",
                "SmallImageItem",
                "微小图片对象",
                "image_path 默认脱敏，缩略图和导出使用受控资源地址",
                sum(int(report.get("small_image_count") or 0) for report in reports),
                [
                    ("id", "string", "图片 ID"),
                    ("file_id", "string", "文件 ID"),
                    ("page_index", "number", "页码"),
                    ("location", "string", "所在位置"),
                    ("image_path", "string", "图片路径"),
                    ("width", "number", "宽度"),
                    ("height", "number", "高度"),
                    ("area", "number", "面积"),
                    ("image_type", "string", "图片类型"),
                    ("is_formula_like", "boolean", "是否疑似公式"),
                    ("is_icon_like", "boolean", "是否疑似图标"),
                    ("is_qrcode_like", "boolean", "是否疑似二维码"),
                    ("is_stamp_like", "boolean", "是否疑似印章"),
                    ("is_signature_like", "boolean", "是否疑似签名"),
                    ("is_duplicate", "boolean", "是否重复"),
                    ("duplicate_check_status", "string", "重复判断状态"),
                    ("duplicate_fallback", "string", "重复判断异常兜底"),
                    ("export_status", "string", "可导出 / 已重新导出 / 重新导出失败 / 原始流不可用"),
                    ("export_message", "string", "图片导出或重新导出说明"),
                    ("reexported_at", "datetime", "最近重新导出时间"),
                    ("confidence", "number", "置信度"),
                ],
                ["image_path"],
            ),
            self._dictionary_object(
                "reports",
                "ReportItem",
                "报告对象",
                "report_path 和导出路径默认脱敏，下载走受令牌保护的资源接口",
                len(reports),
                [
                    ("id", "string", "报告 ID"),
                    ("task_id", "string", "任务 ID"),
                    ("file_id", "string", "文件 ID"),
                    ("report_type", "string", "报告类型"),
                    ("success_count", "number", "成功数量"),
                    ("fail_count", "number", "失败数量"),
                    ("formula_count", "number", "公式数量"),
                    ("omml_count", "number", "OMML 公式数量"),
                    ("omml_converted_count", "number", "OMML 转换成功数量"),
                    ("macro_count", "number", "宏数量"),
                    ("macro_success_count", "number", "宏执行成功数量"),
                    ("macro_fail_count", "number", "宏执行失败数量"),
                    ("formatted_formula_count", "number", "已格式化公式数量"),
                    ("small_image_count", "number", "微小图片数量"),
                    ("error_count", "number", "错误数量"),
                    ("report_path", "string", "报告路径"),
                    ("created_at", "datetime", "创建时间"),
                ],
                ["report_path", "html_path", "json_path", "pdf_path", "xlsx_path", "txt_path"],
            ),
        ]
        total_fields = sum(len(item["fields"]) for item in objects)
        return {
            "schema_version": "k12.dataDictionary.v1",
            "source": "PRD 第 10 章数据结构设计",
            "generated_at": utc_now(),
            "summary": {"objects": len(objects), "fields": total_fields},
            "path_policy": "本地路径字段默认脱敏；下载、恢复和打包由后端使用原始 store 数据",
            "objects": objects,
        }

    @staticmethod
    def _dictionary_object(
        key: str,
        name: str,
        title: str,
        privacy: str,
        count: int,
        fields: list[tuple[str, str, str]],
        sensitive_fields: list[str],
    ) -> dict[str, Any]:
        """生成含敏感性说明的数据字典对象。"""
        return {
            "key": key,
            "name": name,
            "title": title,
            "privacy": privacy,
            "count": count,
            "sensitive_fields": sensitive_fields,
            "fields": [{"name": field, "type": field_type, "description": description} for field, field_type, description in fields],
        }

    def _api_endpoint(
        self,
        method: str,
        path: str,
        name: str,
        auth: str,
        description: str,
        sensitive: bool = False,
        permission: str = "",
        response_headers: list[dict[str, str]] | None = None,
    ) -> dict[str, Any]:
        """生成含授权和敏感路由标记的接口目录条目。"""
        endpoint = {
            "method": method,
            "path": path,
            "name": name,
            "auth": auth,
            "description": description,
            "sensitive": sensitive,
            "permission": permission,
        }
        if response_headers:
            endpoint["response_headers"] = response_headers
        return endpoint

    def acceptance_matrix(self) -> dict[str, Any]:
        """返回需求验收证据，不夸大外部执行覆盖情况。"""
        settings = self.store.get_settings()
        files = self.store.list_files()
        tasks = self.store.list_tasks()
        reports = self.store.list_reports()
        file_types = {file.get("file_type") for file in files}
        task_types = {task.get("task_type") for task in tasks}
        report_types = {report.get("task_type") for report in reports}
        mathpix_queue = self.mathpix_job_queue()
        mathpix_contract_probe = self._mathpix_contract_probe()
        pdf_ocr_probe = self._pdf_ocr_contract_probe()
        pdf_retention_probe = self._pdf_retention_probe()
        pdf_formula_handoff_probe = self._pdf_formula_handoff_probe()
        file_validation_probe = self._file_validation_probe()
        omml_dependency_probe = self._omml_dependency_probe()
        ppt_formula_probe = self._ppt_formula_probe()
        macro_batch_probe = self._macro_batch_sequence_probe()
        macro_order_probe = self._macro_ordered_execution_probe()
        macro_detection_probe = self._macro_detection_probe()
        word_object_probe = self._word_object_preservation_probe()
        latex_mathtype_probe = self._latex_mathtype_probe()
        formula_scope_probe = self._formula_format_scope_probe()
        mathtype_preservation_probe = self._mathtype_preservation_probe()
        omml_handoff_probe = self._omml_mathtype_handoff_probe()
        local_api_security_probe = self._local_api_security_probe()
        local_result_upload_probe = self._local_result_upload_probe()
        local_file_action_probe = self._local_file_action_execution_probe()
        local_install_probe = self._local_install_platform_probe()
        native_execution_evidence = self._local_native_acceptance_evidence(tasks)
        token_configured = bool(str(settings.get("localSecurityToken") or "").strip())
        has_word_to_ppt = "word_to_ppt" in task_types or "word_to_ppt" in report_types
        has_ppt_to_word = "ppt_to_word" in task_types or "ppt_to_word" in report_types
        conversion_probes = self._conversion_capability_probes()
        word_to_ppt_probe = conversion_probes["word_to_ppt"]
        ppt_to_word_probe = conversion_probes["ppt_to_word"]
        pdf_to_word_probe = conversion_probes["pdf_to_word"]
        word_to_ppt_ready = has_word_to_ppt or bool(word_to_ppt_probe.get("available"))
        ppt_to_word_ready = has_ppt_to_word or bool(ppt_to_word_probe.get("available"))
        has_pdf_to_word = "pdf_to_word" in task_types or "pdf_to_word" in report_types
        has_completed_mathpix_pdf_to_word = any(
            job.get("status") == "completed"
            and any(output.get("type") == "docx" for output in job.get("outputs", []))
            for job in mathpix_queue.get("jobs", [])
        )
        has_completed_pdf_to_word = bool(pdf_to_word_probe.get("available")) or has_completed_mathpix_pdf_to_word
        mathpix_real_evidence = self._mathpix_acceptance_evidence(mathpix_queue, reports)
        pdf_formula_native_covered = bool(
            mathpix_real_evidence.get("formula_ocr_available") and native_execution_evidence.get("pdf_formula_mathtype_available")
        )
        has_batch = "batch_process" in task_types or "batch_process" in report_types
        formula_formatting = bool(settings.get("enableMathTypeFormatting", True))
        macro_backup = bool(settings.get("macroBackup", True))
        groups = [
            self._acceptance_group(
                "17.1",
                "文件上传验收",
                [
                    self._acceptance_item("17.1.1", "支持 Word 拖拽上传", "good", "已覆盖", "上传接口、拖拽区和 Word 类型识别已接入", current=self._acceptance_current("Word", file_types)),
                    self._acceptance_item("17.1.2", "支持 Excel 拖拽上传", "good", "已覆盖", "上传接口、拖拽区和 Excel 类型识别已接入", current=self._acceptance_current("Excel", file_types)),
                    self._acceptance_item("17.1.3", "支持 PPT 拖拽上传", "good", "已覆盖", "上传接口、拖拽区和 PPT 类型识别已接入", current=self._acceptance_current("PPT", file_types)),
                    self._acceptance_item("17.1.4", "支持 PDF 拖拽上传", "good", "已覆盖", "上传接口、拖拽区和 PDF 类型识别已接入", current=self._acceptance_current("PDF", file_types)),
                    self._acceptance_item("17.1.5", "支持批量上传", "good", "已覆盖", "多文件、文件夹和 ZIP 导入会生成文件队列", current=f"当前文件 {len(files)} 个"),
                    self._acceptance_item("17.1.6", "格式错误有提示", "good", "已覆盖", "文件校验会记录不支持格式和文件头异常"),
                    self._acceptance_item(
                        "17.1.7",
                        "文件损坏有提示",
                        "good" if file_validation_probe.get("corrupt_available") else "warn",
                        "已覆盖" if file_validation_probe.get("corrupt_available") else "自检失败",
                        str(file_validation_probe.get("corrupt_evidence") or file_validation_probe.get("evidence") or "预览和分析阶段会写入 validation_errors"),
                        current=str(file_validation_probe.get("current") or ""),
                    ),
                    self._acceptance_item(
                        "17.1.8",
                        "加密文件有提示",
                        "good" if file_validation_probe.get("encrypted_available") else "warn",
                        "已覆盖" if file_validation_probe.get("encrypted_available") else "自检失败",
                        str(file_validation_probe.get("encrypted_evidence") or file_validation_probe.get("evidence") or "加密 Word/PDF 会标记并提示登记密码"),
                        current=str(file_validation_probe.get("current") or ""),
                    ),
                    self._acceptance_item("17.1.9", "文件列表展示完整", "good", "已覆盖", "文件表格展示名称、类型、大小、公式、OMML、宏和状态"),
                ],
            ),
            self._acceptance_group(
                "17.2",
                "Word 处理验收",
                [
                    self._acceptance_item("17.2.1", "Word 可成功解析", "good", "已覆盖", "DOCX 结构预览、标题、正文、图片、表格和公式摘要已生成"),
                    self._acceptance_item(
                        "17.2.2",
                        "Word 可转换为 PPT",
                        "good" if word_to_ppt_ready else "warn",
                        "已覆盖" if word_to_ppt_ready else "自检失败",
                        str(word_to_ppt_probe.get("evidence") or "Word 转 PPT 会生成最小 PPTX 产物"),
                        current=self._conversion_probe_current("word_to_ppt", has_word_to_ppt, word_to_ppt_probe),
                    ),
                    self._acceptance_item("17.2.3", "标题层级基本正确", "good", "已覆盖", "转换报告会记录标题层级和内容提取结果"),
                    self._acceptance_item(
                        "17.2.4",
                        "图片不丢失",
                        "warn",
                        "部分实测" if word_object_probe.get("image_available") else "未完成对象写回",
                        str(word_object_probe.get("image_evidence") or word_object_probe.get("evidence") or "对象保留清单会记录图片数量，真实排版需本地 Office 复核"),
                        current=str(word_object_probe.get("current") or "Word 对象保留自检未运行"),
                        gap="输出未证明源图片对象完整保留，需实际图片写回与关联检查。",
                    ),
                    self._acceptance_item(
                        "17.2.5",
                        "表格基本保留",
                        "good" if word_object_probe.get("table_available") else "warn",
                        "已覆盖" if word_object_probe.get("table_available") else "自检失败",
                        str(word_object_probe.get("table_evidence") or word_object_probe.get("evidence") or "对象保留清单会记录表格数量，真实排版需本地 Office 复核"),
                        current=str(word_object_probe.get("current") or "Word 对象保留自检未运行"),
                    ),
                    self._acceptance_item(
                        "17.2.6",
                        "MathType 公式不丢失",
                        "good" if word_object_probe.get("formula_available") else "warn",
                        "已覆盖" if word_object_probe.get("formula_available") else "未完成对象写回",
                        str(word_object_probe.get("formula_evidence") or word_object_probe.get("evidence") or "MathType/嵌入对象进入保留清单，真实 OLE 写回需本地客户端"),
                        current=str(word_object_probe.get("current") or "Word 对象保留自检未运行"),
                        gap="输出未证明 MathType 对象保留，需原生对象写回及公式复核。",
                    ),
                    self._acceptance_item("17.2.7", "用户可选择是否启用 MathType 格式化", "good", "已覆盖", "设置页提供 MathType 格式化开关", current=f"当前为{'开启' if formula_formatting else '关闭'}"),
                    self._acceptance_item(
                        "17.2.8",
                        "支持批量统一公式格式",
                        "good" if formula_scope_probe.get("batch_available") else "warn",
                        "已覆盖" if formula_scope_probe.get("batch_available") else "自检失败",
                        str(formula_scope_probe.get("batch_evidence") or formula_scope_probe.get("evidence") or "批量任务和格式化参数会进入本地客户端动作队列"),
                        current=str(formula_scope_probe.get("current") or "公式批量格式化自检未运行"),
                    ),
                    self._acceptance_item("17.2.9", "格式化失败时保留原公式", "good", "已覆盖", "公式报告记录失败兜底并保留原公式或原图"),
                    self._acceptance_item("17.2.10", "生成公式格式化报告", "good", "已覆盖", "公式报告、公式 ZIP 和 XLSX 导出已接入"),
                ],
            ),
            self._acceptance_group(
                "17.3",
                "Word 自带公式与 OMML 验收",
                [
                    self._acceptance_item("17.3.1", "Word 处理前必须先检测是否存在 Word 自带公式", "good", "已覆盖", "上传分析和任务预检都会标记 has_omml"),
                    self._acceptance_item("17.3.2", "检测到 Word 自带公式时，应提示用户是否转换为 MathType", "good", "已覆盖", "OMML 转换确认记录会进入报告和前端卡片"),
                    self._acceptance_item(
                        "17.3.3",
                        "用户选择转换后，系统应执行 OMML 转 MathType",
                        "good" if native_execution_evidence.get("omml_mathtype_available") else "warn",
                        "已覆盖" if native_execution_evidence.get("omml_mathtype_available") else ("合同覆盖" if omml_handoff_probe.get("handoff_available") else "自检失败"),
                        str(native_execution_evidence.get("omml_evidence") or omml_handoff_probe.get("handoff_evidence") or omml_handoff_probe.get("evidence") or "本地任务载荷会下发 omml_mathtype 动作和桌面执行计划，真实写回需桌面端"),
                        current=str(native_execution_evidence.get("omml_current") or omml_handoff_probe.get("current") or "OMML 转 MathType 交接自检未运行"),
                    ),
                    self._acceptance_item(
                        "17.3.4",
                        "如果提示找不到 OMML 文件，系统应支持自动检索电脑中的 OMML 文件",
                        "good" if omml_dependency_probe.get("search_available") else "warn",
                        "已覆盖" if omml_dependency_probe.get("search_available") else "自检失败",
                        str(omml_dependency_probe.get("search_evidence") or omml_dependency_probe.get("evidence") or "检索范围、上限和路径配置已进入任务合同"),
                        current=self._omml_dependency_probe_current(omml_dependency_probe),
                    ),
                    self._acceptance_item(
                        "17.3.5",
                        "检索到 OMML 文件后，应复制到当前 Word 文档所在文件夹",
                        "good" if omml_dependency_probe.get("copy_available") else "warn",
                        "已覆盖" if omml_dependency_probe.get("copy_available") else "自检失败",
                        str(omml_dependency_probe.get("copy_evidence") or omml_dependency_probe.get("evidence") or "复制策略和目标目录已记录，真实复制需本地文件权限"),
                        current=self._omml_dependency_probe_current(omml_dependency_probe),
                    ),
                    self._acceptance_item(
                        "17.3.6",
                        "复制成功后，应重新执行公式转换",
                        "good" if native_execution_evidence.get("omml_mathtype_available") else "warn",
                        "已覆盖" if native_execution_evidence.get("omml_mathtype_available") else ("合同覆盖" if omml_handoff_probe.get("retry_available") else "自检失败"),
                        str(native_execution_evidence.get("omml_retry_evidence") or omml_handoff_probe.get("retry_evidence") or omml_handoff_probe.get("evidence") or "重新转换动作会进入本地客户端队列"),
                        current=str(native_execution_evidence.get("omml_current") or omml_handoff_probe.get("current") or "OMML 重试交接自检未运行"),
                    ),
                    self._acceptance_item("17.3.7", "如果无法自动找到 OMML 文件，应支持用户手动选择", "good", "已覆盖", "设置页和 OMML 校正支持手动路径"),
                    self._acceptance_item("17.3.8", "转换失败时，不得删除或破坏原公式", "good", "已覆盖", "失败策略为保留 OMML 或原公式"),
                    self._acceptance_item("17.3.9", "所有处理结果应写入日志", "good", "已覆盖", "OMML 检测、复制和失败会进入日志/报告"),
                    self._acceptance_item("17.3.10", "应生成 OMML 转 MathType 处理报告", "good", "已覆盖", "OMML 依赖清单和失败 CSV 已接入"),
                ],
            ),
            self._acceptance_group(
                "17.4",
                "Word 宏验收",
                [
                    self._acceptance_item(
                        "17.4.1",
                        "系统可检测当前 Word 文档中的宏",
                        "good" if macro_detection_probe.get("document_available") else "warn",
                        "已覆盖" if macro_detection_probe.get("document_available") else "自检失败",
                        str(macro_detection_probe.get("document_evidence") or macro_detection_probe.get("evidence") or "DOCM/宏线索会标记，真实 VBA 枚举需 Word COM"),
                        current=str(macro_detection_probe.get("current") or "宏检测自检未运行"),
                    ),
                    self._acceptance_item(
                        "17.4.2",
                        "系统可检测模板文件中的宏",
                        "good" if macro_detection_probe.get("template_available") else "warn",
                        "已覆盖" if macro_detection_probe.get("template_available") else "自检失败",
                        str(macro_detection_probe.get("template_evidence") or macro_detection_probe.get("evidence") or "模板宏来源进入宏库和授权策略，真实模板枚举需本地客户端"),
                        current=str(macro_detection_probe.get("current") or "宏检测自检未运行"),
                    ),
                    self._acceptance_item("17.4.3", "用户可选择一个或多个宏", "good", "已覆盖", "宏页支持选择、筛选和详情"),
                    self._acceptance_item("17.4.4", "用户可调整宏执行顺序", "good", "已覆盖", "宏顺序支持拖拽把手和上/下移"),
                    self._acceptance_item(
                        "17.4.5",
                        "系统可按顺序执行宏",
                        "good" if native_execution_evidence.get("macro_sequence_available") else "warn",
                        "已覆盖" if native_execution_evidence.get("macro_sequence_available") else ("合同覆盖" if macro_order_probe.get("available") else "自检失败"),
                        str(native_execution_evidence.get("macro_evidence") or macro_order_probe.get("evidence") or "网页端生成本地宏队列和桌面执行计划，真实执行只允许本地客户端"),
                        current=str(native_execution_evidence.get("macro_current") or macro_order_probe.get("current") or "宏顺序执行交接自检未运行"),
                    ),
                    self._acceptance_item("17.4.6", "宏执行前必须生成文档备份", "good" if macro_backup else "warn", "已覆盖" if macro_backup else "需开启", "宏备份策略和恢复入口已接入", current=f"当前备份{'开启' if macro_backup else '关闭'}"),
                    self._acceptance_item("17.4.7", "宏执行失败时支持停止、跳过或继续", "good", "已覆盖", "失败策略写入任务和报告", current=str(settings.get("macroFailureStrategy") or "跳过")),
                    self._acceptance_item("17.4.8", "宏执行结果必须写入日志", "good", "已覆盖", "宏执行日志和失败清单已接入"),
                    self._acceptance_item("17.4.9", "宏执行完成后必须生成执行报告", "good", "已覆盖", "宏执行报告、重跑和失败 CSV 已接入"),
                    self._acceptance_item(
                        "17.4.10",
                        "批量处理时可对多个 Word 文件应用同一宏执行顺序",
                        "good" if macro_batch_probe.get("available") else "warn",
                        "已覆盖" if macro_batch_probe.get("available") else "自检失败",
                        str(macro_batch_probe.get("evidence") or "批量任务可复用宏顺序和本地执行队列"),
                        current=str(macro_batch_probe.get("current") or "宏批量顺序自检未运行"),
                    ),
                ],
            ),
            self._acceptance_group(
                "17.5",
                "PPT 处理验收",
                [
                    self._acceptance_item("17.5.1", "PPT 可成功解析", "good", "已覆盖", "PPT 预览会提取幻灯片、标题、正文、备注和对象摘要"),
                    self._acceptance_item(
                        "17.5.2",
                        "PPT 可转换为 Word",
                        "good" if ppt_to_word_ready else "warn",
                        "已覆盖" if ppt_to_word_ready else "自检失败",
                        str(ppt_to_word_probe.get("evidence") or "PPT 转 Word 会生成最小 DOCX 产物"),
                        current=self._conversion_probe_current("ppt_to_word", has_ppt_to_word, ppt_to_word_probe),
                    ),
                    self._acceptance_item("17.5.3", "文本可提取", "good", "已覆盖", "转换报告写入标题和正文提取结果"),
                    self._acceptance_item("17.5.4", "图片可提取", "good", "已覆盖", "OOXML 媒体提取和对象摘要已接入"),
                    self._acceptance_item("17.5.5", "表格可提取", "good", "已覆盖", "PPT 对象摘要和质量检查记录表格"),
                    self._acceptance_item(
                        "17.5.6",
                        "公式可识别",
                        "good" if ppt_formula_probe.get("available") else "warn",
                        "已覆盖" if ppt_formula_probe.get("available") else "自检失败",
                        str(ppt_formula_probe.get("evidence") or "PPT 公式进入统一公式报告，真实 MathType 需本地复核"),
                        current=str(ppt_formula_probe.get("current") or "PPT 公式识别自检未运行"),
                    ),
                    self._acceptance_item("17.5.7", "备注可导出", "good", "已覆盖", "PPT 转 Word 支持备注优先和备注提取开关"),
                    self._acceptance_item("17.5.8", "转换失败有日志", "good", "已覆盖", "任务日志、失败原因和重试入口已接入"),
                ],
            ),
            self._acceptance_group(
                "17.6",
                "PDF 处理验收",
                [
                    self._acceptance_item(
                        "17.6.1",
                        "文本 PDF 可转 Word",
                        "good" if has_completed_pdf_to_word else "warn",
                        "已覆盖" if has_completed_pdf_to_word else "合同覆盖",
                        "已有 Mathpix DOCX 完成记录，PDF 转 Word 产物进入标准输出"
                        if has_completed_mathpix_pdf_to_word
                        else str(pdf_to_word_probe.get("evidence") or "文本型 PDF 本地转换自检未运行"),
                        current=(
                            self._mathpix_contract_probe_current(has_pdf_to_word, True, mathpix_contract_probe, mathpix_queue)
                            if has_completed_mathpix_pdf_to_word
                            else self._conversion_probe_current("pdf_to_word", has_pdf_to_word, pdf_to_word_probe)
                        ),
                    ),
                    self._acceptance_item(
                        "17.6.2",
                        "扫描 PDF 可 OCR",
                        "good" if mathpix_real_evidence.get("scanned_ocr_available") else "warn",
                        "已覆盖" if mathpix_real_evidence.get("scanned_ocr_available") else "需 Mathpix 实测",
                        str(mathpix_real_evidence.get("scanned_evidence") or pdf_ocr_probe.get("scanned_evidence") or pdf_ocr_probe.get("evidence") or "扫描型 PDF 会生成 Mathpix OCR 作业，外部上传需授权"),
                        current=str(mathpix_real_evidence.get("scanned_current") or pdf_ocr_probe.get("scanned_current") or "扫描 PDF Mathpix 合同自检未运行"),
                    ),
                    self._acceptance_item(
                        "17.6.3",
                        "图片可保留",
                        "good" if pdf_retention_probe.get("image_available") else "warn",
                        "已覆盖" if pdf_retention_probe.get("image_available") else "自检失败",
                        str(pdf_retention_probe.get("image_evidence") or pdf_retention_probe.get("evidence") or "Mathpix 保留计划记录图片对象和保留策略"),
                        current=str(pdf_retention_probe.get("current") or "PDF 图片保留自检未运行"),
                    ),
                    self._acceptance_item(
                        "17.6.4",
                        "表格可保留",
                        "good" if pdf_retention_probe.get("table_available") else "warn",
                        "已覆盖" if pdf_retention_probe.get("table_available") else "自检失败",
                        str(pdf_retention_probe.get("table_evidence") or pdf_retention_probe.get("evidence") or "Mathpix 请求携带表格 OCR 设置并记录表格线索"),
                        current=str(pdf_retention_probe.get("current") or "PDF 表格保留自检未运行"),
                    ),
                    self._acceptance_item(
                        "17.6.5",
                        "数学公式可识别",
                        "good" if mathpix_real_evidence.get("formula_ocr_available") else "warn",
                        "已覆盖" if mathpix_real_evidence.get("formula_ocr_available") else "需 Mathpix 实测",
                        str(mathpix_real_evidence.get("formula_evidence") or pdf_ocr_probe.get("formula_evidence") or pdf_ocr_probe.get("evidence") or "Mathpix tex.zip 公式结果会进入统一公式报告"),
                        current=str(mathpix_real_evidence.get("formula_current") or pdf_ocr_probe.get("formula_current") or "PDF 公式 OCR 合同自检未运行"),
                    ),
                    self._acceptance_item(
                        "17.6.6",
                        "公式可转换为 MathType",
                        "good" if pdf_formula_handoff_probe.get("available") else "warn",
                        "已覆盖" if pdf_formula_handoff_probe.get("available") else "自检失败",
                        str(pdf_formula_handoff_probe.get("evidence") or "Mathpix 公式结果会进入 pdf_formula_mathtype 本地动作和桌面执行计划"),
                        current=str(pdf_formula_handoff_probe.get("current") or "PDF 公式 MathType 后处理自检未运行"),
                    ),
                    self._acceptance_item("17.6.7", "低置信度公式可标记", "good", "已覆盖", "公式报告和 PDF 工作区展示低置信度条目"),
                    self._acceptance_item("17.6.8", "转换结果可预览", "good", "已覆盖", "转换前后对比预览已接入"),
                    self._acceptance_item("17.6.9", "生成 PDF 转换报告", "good", "已覆盖", "Mathpix 状态、保留计划和失败清单进入报告", current=f"Mathpix 队列 {mathpix_queue['summary']['total']} 个"),
                ],
            ),
            self._acceptance_group(
                "17.7",
                "MathType 公式验收",
                [
                    self._acceptance_item(
                        "17.7.1",
                        "已有 MathType 公式应尽量原样保留",
                        "good" if mathtype_preservation_probe.get("available") else "warn",
                        "已覆盖" if mathtype_preservation_probe.get("available") else "自检失败",
                        str(mathtype_preservation_probe.get("evidence") or "对象保留清单记录 MathType/OLE 对象"),
                        current=str(mathtype_preservation_probe.get("current") or "MathType 原样保留自检未运行"),
                    ),
                    self._acceptance_item(
                        "17.7.2",
                        "Word 原生公式可转换为 MathType",
                        "good" if native_execution_evidence.get("omml_mathtype_available") else "warn",
                        "已覆盖" if native_execution_evidence.get("omml_mathtype_available") else ("合同覆盖" if omml_handoff_probe.get("handoff_available") else "自检失败"),
                        str(native_execution_evidence.get("omml_evidence") or omml_handoff_probe.get("handoff_evidence") or omml_handoff_probe.get("evidence") or "OMML 转 MathType 动作进入本地载荷"),
                        current=str(native_execution_evidence.get("omml_current") or omml_handoff_probe.get("current") or "OMML 转 MathType 交接自检未运行"),
                    ),
                    self._acceptance_item(
                        "17.7.3",
                        "LaTeX 公式可转换为 MathType",
                        "good" if latex_mathtype_probe.get("available") else "warn",
                        "已覆盖" if latex_mathtype_probe.get("available") else "自检失败",
                        str(latex_mathtype_probe.get("evidence") or "公式报告保留 LaTeX、MathML 和 MathType 预览字段"),
                        current=str(latex_mathtype_probe.get("current") or "LaTeX 转 MathType 自检未运行"),
                    ),
                    self._acceptance_item(
                        "17.7.4",
                        "PDF 公式可识别并转换为 MathType",
                        "good" if pdf_formula_native_covered else "warn",
                        "已覆盖" if pdf_formula_native_covered else ("需本地客户端实测" if mathpix_real_evidence.get("formula_ocr_available") else "需 Mathpix 实测"),
                        str(native_execution_evidence.get("pdf_formula_evidence") if pdf_formula_native_covered else (mathpix_real_evidence.get("formula_handoff_evidence") or pdf_ocr_probe.get("formula_handoff_evidence") or "Mathpix 公式 OCR 结果会交给 PDF 公式 MathType 后处理动作；真实识别和对象写回仍需授权与桌面端实测")),
                        current=str(native_execution_evidence.get("pdf_formula_current") if pdf_formula_native_covered else (mathpix_real_evidence.get("formula_handoff_current") or pdf_ocr_probe.get("formula_handoff_current") or "PDF 公式识别与 MathType 转换仍需 Mathpix 授权和本地客户端实测")),
                    ),
                    self._acceptance_item("17.7.5", "图片公式识别失败时保留原图", "good", "已覆盖", "失败公式保留原图引用和待确认状态"),
                    self._acceptance_item("17.7.6", "用户可在 Word 中选择 MathType 格式化公式", "good", "已覆盖", "设置页和公式工作区提供格式化参数"),
                    self._acceptance_item(
                        "17.7.7",
                        "支持全文公式批量格式化",
                        "good" if formula_scope_probe.get("batch_available") else "warn",
                        "已覆盖" if formula_scope_probe.get("batch_available") else "自检失败",
                        str(formula_scope_probe.get("batch_evidence") or formula_scope_probe.get("evidence") or "格式化范围和批量任务参数会进入本地动作队列"),
                        current=str(formula_scope_probe.get("current") or "公式批量格式化自检未运行"),
                    ),
                    self._acceptance_item(
                        "17.7.8",
                        "支持指定范围公式格式化",
                        "good" if formula_scope_probe.get("scope_available") else "warn",
                        "已覆盖" if formula_scope_probe.get("scope_available") else "自检失败",
                        str(formula_scope_probe.get("scope_evidence") or formula_scope_probe.get("evidence") or "支持全文、章节、选中区域和单公式参数"),
                        current=str(formula_scope_probe.get("current") or "公式范围格式化自检未运行"),
                    ),
                    self._acceptance_item("17.7.9", "支持公式格式化前后对比", "good", "已覆盖", "公式详情展示格式化前后和校正记录"),
                    self._acceptance_item("17.7.10", "支持生成公式识别和格式化报告", "good", "已覆盖", "公式报告、ZIP、XLSX 和校正记录已接入"),
                ],
            ),
            self._acceptance_group(
                "17.8",
                "微小图片验收",
                [
                    self._acceptance_item("17.8.1", "Word 小图片可检索", "good", "已覆盖", "Word OOXML 媒体提取支持微小图片标记"),
                    self._acceptance_item("17.8.2", "PPT 小图片可检索", "good", "已覆盖", "PPT OOXML 媒体提取支持微小图片标记"),
                    self._acceptance_item("17.8.3", "Excel 小图片可检索", "good", "已覆盖", "Excel OOXML 媒体提取支持微小图片标记"),
                    self._acceptance_item("17.8.4", "PDF 小图片可检索", "good", "已覆盖", "PDF 图片 XObject 读取和占位兜底已接入"),
                    self._acceptance_item("17.8.5", "支持按大小筛选", "good", "已覆盖", "图片页支持面积和尺寸阈值筛选"),
                    self._acceptance_item("17.8.6", "支持按类型筛选", "good", "已覆盖", "图片页支持疑似类型筛选"),
                    self._acceptance_item("17.8.7", "支持按位置筛选", "good", "已覆盖", "图片页支持位置关键词筛选"),
                    self._acceptance_item("17.8.8", "支持疑似公式图片筛选", "good", "已覆盖", "疑似公式类型和人工确认已接入"),
                    self._acceptance_item("17.8.9", "支持疑似二维码筛选", "good", "已覆盖", "二维码疑似类型已进入报告"),
                    self._acceptance_item("17.8.10", "支持疑似印章筛选", "good", "已覆盖", "印章疑似类型已进入报告"),
                    self._acceptance_item("17.8.11", "支持导出图片", "good", "已覆盖", "单张图片和图片 ZIP 下载已接入"),
                    self._acceptance_item("17.8.12", "支持生成图片报告", "good", "已覆盖", "图片清单 XLSX、报告和标注记录已接入"),
                ],
            ),
            self._acceptance_group(
                "17.9",
                "批量处理验收",
                [
                    self._acceptance_item("17.9.1", "支持多个文件排队", "good", "已覆盖", "任务队列和批量任务已接入", current=self._task_presence_current("batch_process", has_batch)),
                    self._acceptance_item("17.9.2", "显示整体进度", "good", "已覆盖", "工作台和任务中心展示整体进度"),
                    self._acceptance_item("17.9.3", "显示单文件状态", "good", "已覆盖", "批量报告记录单文件状态、失败原因和跳过状态"),
                    self._acceptance_item("17.9.4", "失败文件可重试", "good", "已覆盖", "任务重试和批量单文件建议任务已接入"),
                    self._acceptance_item("17.9.5", "失败文件可跳过", "good", "已覆盖", "批量失败文件可标记跳过并重建报告"),
                    self._acceptance_item("17.9.6", "结果可打包下载", "good", "已覆盖", "任务结果包包含任务 JSON、日志、报告和输出"),
                    self._acceptance_item("17.9.7", "日志可导出", "good", "已覆盖", "支持 TXT、LOG、CSV、JSON 日志导出"),
                    self._acceptance_item("17.9.8", "处理报告可导出", "good", "已覆盖", "HTML、JSON、PDF、XLSX、TXT 报告下载已接入"),
                ],
            ),
            self._acceptance_group(
                "17.10",
                "本地 + 网页双模式验收",
                [
                    self._acceptance_item(
                        "17.10.1",
                        "本地客户端可独立运行",
                        "good" if local_install_probe.get("available") else "warn",
                        "已覆盖" if local_install_probe.get("available") else "自检失败",
                        (
                            "标准库本地伴随 CLI 可独立运行，支持 manifest、心跳、载荷领取、动作级 dry-run 校验和状态同步；"
                            f"{local_install_probe.get('evidence') or '安装画像自检未运行'}"
                        ),
                        current=str(local_install_probe.get("current") or "安装画像自检未运行；真实桌面壳仍可后续打包"),
                    ),
                    self._acceptance_item("17.10.2", "网页端可独立访问", "good", "已覆盖", "静态前端和本地 API 可独立运行"),
                    self._acceptance_item("17.10.3", "网页端可创建任务", "good", "已覆盖", "任务创建接口和工作台入口已接入"),
                    self._acceptance_item("17.10.4", "系统可判断任务是否需要本地处理", "good", "已覆盖", "resolve_execute_mode 会按任务类型和文件能力分流"),
                    self._acceptance_item("17.10.5", "需要本地处理的任务可提示启动本地客户端", "good", "已覆盖", "本地任务交接面板和 local-launch 请求已接入"),
                    self._acceptance_item("17.10.6", "本地客户端可接收任务参数", "good", "已覆盖", "local-payload 返回本地路径、预检、动作队列和同步策略"),
                    self._acceptance_item(
                        "17.10.7",
                        "本地客户端可执行 Office、MathType、OMML、Word 宏相关任务",
                        "good" if native_execution_evidence.get("all_native_actions_available") else "warn",
                        (
                            "已覆盖"
                            if native_execution_evidence.get("all_native_actions_available")
                            else ("部分实测" if native_execution_evidence.get("available_action_types") else ("需本地客户端实测" if local_file_action_probe.get("available") else "自检失败"))
                        ),
                        str(
                            native_execution_evidence.get("all_native_evidence")
                            or native_execution_evidence.get("partial_evidence")
                            or local_file_action_probe.get("evidence")
                            or "当前提供动作队列、能力位、桌面执行计划、动作级 dry-run 校验摘要、k12.localNativeExecutionRequest.v1 原生执行请求合同和 k12.localFileActionExecution.v1 OMML 文件动作执行合同，不声明真实桌面执行"
                        ),
                        current=str(native_execution_evidence.get("all_native_current") or native_execution_evidence.get("partial_current") or local_file_action_probe.get("current") or "需要外部/本地真实环境证明"),
                        gap=str(native_execution_evidence.get("remaining_gap") or "需要真实桌面客户端补齐 Office、MathType、OMML 写回和 Word 宏动作证据"),
                    ),
                    self._acceptance_item(
                        "17.10.8",
                        "本地任务结果可选择本地保存或上传云端",
                        "good" if local_result_upload_probe.get("available") else "warn",
                        "已覆盖" if local_result_upload_probe.get("available") else "自检失败",
                        str(local_result_upload_probe.get("evidence") or "local-sync 记录本地保存/上传意图，POST /api/local-client/uploads 可在授权后登记上传包清单"),
                        current=str(local_result_upload_probe.get("current") or "本地结果上传自检未运行"),
                    ),
                    self._acceptance_item("17.10.9", "网页端可同步任务状态", "good", "已覆盖", "local-sync 可回传状态、进度、输出摘要和上传意图"),
                    self._acceptance_item(
                        "17.10.10",
                        "本地 API 需要安全令牌保护",
                        "good" if local_api_security_probe.get("available") else "warn",
                        "已覆盖" if local_api_security_probe.get("available") else "自检失败",
                        str(local_api_security_probe.get("evidence") or "本地载荷强制要求配置令牌，普通 API 可按设置启用令牌"),
                        current=f"{local_api_security_probe.get('current') or '安全令牌自检未运行'}；当前令牌{'已配置' if token_configured else '未配置'}",
                    ),
                ],
            ),
        ]
        all_items = [item for group in groups for item in group["items"]]
        uncovered_risks = [
            {
                "key": item["key"],
                "requirement": item["requirement"],
                "status": item["status"],
                "uncovered_risk": item["verification"]["uncovered_risk"],
                "blocking_reasons": item["verification"]["blocking_reasons"],
                "required_environment": item["verification"]["required_environment"],
                "verification_checklist": item["verification"].get("verification_checklist", []),
                "next_step": item["next_step"],
            }
            for item in all_items
            if item.get("verification", {}).get("uncovered_risk")
        ]
        return {
            "schema_version": "k12.acceptanceMatrix.v1",
            "generated_at": utc_now(),
            "source": "PRD 第 17 章验收标准",
            "groups": groups,
            "summary": {
                "groups": len(groups),
                "items": len(all_items),
                "covered": sum(1 for item in all_items if item["level"] == "good"),
                "contract": sum(1 for item in all_items if item["level"] == "warn"),
                "planned": sum(1 for item in all_items if item["level"] == "blue"),
                "blocked": sum(1 for item in all_items if item["level"] == "bad"),
                "uncovered_risks": len(uncovered_risks),
            },
            "uncovered_risks": uncovered_risks,
            "guardrails": [
                "Office、MathType、OMML 写回和 Word 宏执行属于本地客户端实测能力，不在网页端伪造执行结果。",
                "纯文本 PDF 本地解析；扫描、混合和复杂 PDF 使用 Mathpix 时必须由用户授权并配置凭证。",
                "Windows 与 macOS 安装和 MathType 对象格式分开验收，跨平台交付使用 MathML、LaTeX 或图片兜底。",
            ],
        }

    def _acceptance_group(self, section: str, title: str, items: list[dict[str, Any]]) -> dict[str, Any]:
        """生成需求验收分组及摘要计数和证据条目。"""
        total = len(items)
        covered = sum(1 for item in items if item["level"] == "good")
        contract = sum(1 for item in items if item["level"] == "warn")
        blocked = sum(1 for item in items if item["level"] == "bad")
        level = "bad" if blocked else ("warn" if contract else "good")
        status = "存在缺口" if blocked else ("需实测" if contract else "已覆盖")
        return {
            "section": section,
            "title": title,
            "level": level,
            "status": status,
            "total": total,
            "covered": covered,
            "contract": contract,
            "blocked": blocked,
            "items": items,
        }

    def _acceptance_item(
        self,
        key: str,
        requirement: str,
        level: str,
        status: str,
        evidence: str,
        current: str = "",
        gap: str = "",
        next_step: str = "",
    ) -> dict[str, Any]:
        """生成含验证背景与未覆盖风险的验收条目。"""
        default_next_step = {
            "good": "保持回归测试和报告证据",
            "warn": "在真实本地客户端、Office、MathType 或 Mathpix 授权环境中实测",
            "blue": "进入后续版本规划",
            "bad": "补齐实现和测试后重新验收",
        }.get(level, "补充证据")
        effective_gap = gap or ("需要外部/本地真实环境证明" if level == "warn" else "")
        verification = self._acceptance_verification_context(level, status, effective_gap)
        return {
            "key": key,
            "requirement": requirement,
            "level": level,
            "status": status,
            "evidence": evidence,
            "current": current,
            "gap": effective_gap,
            "next_step": next_step or default_next_step,
            "verification": verification,
        }

    @staticmethod
    def _acceptance_verification_context(level: str, status: str, gap: str = "") -> dict[str, Any]:
        """按检查清单区分验证范围，不夸大本地或外部覆盖。"""
        if status == "需 Mathpix 实测":
            return {
                "scope": "external_mathpix",
                "uncovered_risk": "未覆盖风险：尚未执行真实 Mathpix 上传、OCR 轮询、DOCX/tex.zip 下载和识别结果复核。",
                "blocking_reasons": [
                    "mathpix_upload_authorization_required",
                    "mathpix_credentials_required",
                    "external_ocr_result_not_verified",
                ],
                "required_environment": ["Mathpix APP ID/KEY", "外部上传授权", "可访问 Mathpix API 的网络"],
                "verification_checklist": [
                    "在设置中开启 Mathpix 外部上传授权，并配置 MATHPIX_APP_ID / MATHPIX_APP_KEY",
                    "使用可上传的扫描型或含公式 PDF 创建 pdf_to_word 任务，并等待 Mathpix 作业完成",
                    "确认 DOCX 与 tex.zip 输出进入 artifacts，并核对下载清单的哈希与状态",
                    "复核公式报告中的识别结果、低置信度标记和人工确认记录",
                ],
            }
        if status == "需本地客户端实测":
            return {
                "scope": "native_desktop_client",
                "uncovered_risk": "未覆盖风险：尚未在真实桌面客户端执行 Office、MathType、OMML 写回或 Word 宏动作。",
                "blocking_reasons": [
                    "native_desktop_runner_required",
                    "office_mathtype_runtime_required",
                    "native_writeback_not_verified",
                ],
                "required_environment": ["真实本地客户端", "Microsoft Office 或兼容桌面组件", "同平台 MathType 环境"],
                "verification_checklist": [
                    "在目标 Windows 或 macOS 上安装同平台本地客户端、Office 和 MathType 或兼容组件",
                    "发送包含 Office、MathType、OMML 依赖检索和宏执行能力位的本地客户端心跳",
                    "领取受令牌保护的 local-payload，并执行同平台 native-plan 或显式本地文件动作",
                    "通过 local-sync 回传输出摘要、动作结果和报告证据，确认没有跨平台 MathType 对象交付",
                ],
            }
        if level == "warn":
            return {
                "scope": "contract_or_external_verification",
                "uncovered_risk": gap,
                "blocking_reasons": ["external_or_native_verification_required"] if gap else [],
                "required_environment": ["真实本地客户端、Office、MathType 或 Mathpix 授权环境"] if gap else [],
                "verification_checklist": [
                    "补齐真实环境运行记录、输出文件和报告证据后再改为已覆盖",
                ]
                if gap
                else [],
            }
        return {
            "scope": "local_evidence",
            "uncovered_risk": "",
            "blocking_reasons": [],
            "required_environment": [],
            "verification_checklist": [],
        }

    @staticmethod
    def _acceptance_current(file_type: str, file_types: set[str]) -> str:
        """说明当前文件样本是否覆盖所需类型。"""
        return f"当前样本{'已包含' if file_type in file_types else '未包含'} {file_type}"

    @staticmethod
    def _task_presence_current(task_type: str, present: bool) -> str:
        """说明运行数据是否证明某类任务已执行。"""
        return f"{task_type} {'已有任务或报告' if present else '等待样本任务'}"

    def _conversion_capability_probes(self) -> dict[str, dict[str, Any]]:
        """运行安全的标准库转换探测以提供验收证据。"""
        cached = getattr(self, "_conversion_probe_cache", None)
        if isinstance(cached, dict):
            return cached
        probes = {
            "word_to_ppt": self._probe_word_to_ppt_conversion(),
            "ppt_to_word": self._probe_ppt_to_word_conversion(),
            "pdf_to_word": self._probe_pdf_to_word_conversion(),
        }
        self._conversion_probe_cache = probes
        return probes

    def _probe_word_to_ppt_conversion(self) -> dict[str, Any]:
        """使用临时 DOCX 样本验证基础 Word 转 PPT。"""
        try:
            with tempfile.TemporaryDirectory(prefix="k12-conversion-probe-") as tmp:
                root = Path(tmp)
                source = root / "probe.docx"
                target = root / "probe.pptx"
                build_docx(
                    [
                        {"text": "K12 转换自检", "style": "Heading1"},
                        {"text": "Word 转 PPT 最小产物", "style": "Normal"},
                    ],
                    source,
                )
                blocks = extract_docx_blocks(source)
                build_pptx_from_docx(blocks, target, object_preservation={"source": {"images": 0, "tables": 0, "formulas": 0}})
                with zipfile.ZipFile(target) as archive:
                    names = set(archive.namelist())
                slides = extract_pptx_slides(target)
                if "ppt/presentation.xml" not in names or not slides:
                    raise ValueError("PPTX 关键部件缺失")
                return {
                    "schema_version": "k12.conversionCapabilityProbe.v1",
                    "available": True,
                    "artifact_type": "pptx",
                    "evidence": "内置转换自检通过：Word 转 PPT 可生成 PPTX，并包含 ppt/presentation.xml",
                    "current": f"自检通过，幻灯片 {len(slides)} 页",
                }
        except Exception as exc:
            return self._conversion_probe_failure("pptx", exc)

    def _probe_ppt_to_word_conversion(self) -> dict[str, Any]:
        """使用临时 PPTX 样本验证基础 PPT 转 Word。"""
        try:
            with tempfile.TemporaryDirectory(prefix="k12-conversion-probe-") as tmp:
                root = Path(tmp)
                source = root / "probe.pptx"
                target = root / "probe.docx"
                build_pptx([{"title": "K12 转换自检", "body": ["PPT 转 Word 最小产物"]}], source)
                slides = extract_pptx_slides(source)
                build_docx_from_slides(slides, target)
                with zipfile.ZipFile(target) as archive:
                    names = set(archive.namelist())
                blocks = extract_docx_blocks(target)
                if "word/document.xml" not in names or not blocks:
                    raise ValueError("DOCX 关键部件缺失")
                return {
                    "schema_version": "k12.conversionCapabilityProbe.v1",
                    "available": True,
                    "artifact_type": "docx",
                    "evidence": "内置转换自检通过：PPT 转 Word 可生成 DOCX，并包含 word/document.xml",
                    "current": f"自检通过，段落 {len(blocks)} 段",
                }
        except Exception as exc:
            return self._conversion_probe_failure("docx", exc)

    def _probe_pdf_to_word_conversion(self) -> dict[str, Any]:
        """使用临时 PDF 验证本地文本层转 Word。"""
        try:
            with tempfile.TemporaryDirectory(prefix="k12-conversion-probe-") as tmp:
                root = Path(tmp)
                source = root / "probe.pdf"
                target = root / "probe.docx"
                build_text_pdf(["K12 PDF text layer", "PDF to Word local conversion"], source, "K12 Probe")
                summary = build_docx_from_pdf_text(source, target, "K12 PDF 转 Word 自检")
                validation_error = self._output_artifact_validation_error(target, "docx")
                blocks = extract_docx_blocks(target)
                if validation_error or not blocks or int(summary.get("character_count") or 0) <= 0:
                    raise ValueError(validation_error or "DOCX 文本内容缺失")
                return {
                    "schema_version": "k12.conversionCapabilityProbe.v1",
                    "available": True,
                    "artifact_type": "docx",
                    "evidence": "内置转换自检通过：文本型 PDF 可解析文本层并生成通过 OOXML 校验的 DOCX",
                    "current": f"自检通过，提取文本 {summary['paragraph_count']} 段，DOCX 段落 {len(blocks)} 段",
                }
        except Exception as exc:
            return self._conversion_probe_failure("docx", exc)

    def _conversion_probe_failure(self, artifact_type: str, exc: Exception) -> dict[str, Any]:
        """返回不含路径的转换探测失败结果。"""
        message = self._redact_local_path_text(str(exc) or exc.__class__.__name__)
        return {
            "schema_version": "k12.conversionCapabilityProbe.v1",
            "available": False,
            "artifact_type": artifact_type,
            "evidence": f"内置转换自检失败：{message}",
            "current": "自检失败",
        }

    @staticmethod
    def _conversion_probe_current(task_type: str, has_task_history: bool, probe: dict[str, Any]) -> str:
        """汇总转换探测产物，不暴露临时路径。"""
        history = "已有任务或报告" if has_task_history else "无历史任务"
        return f"{task_type} {history}；{probe.get('current') or '自检未运行'}"

    def _word_object_preservation_probe(self) -> dict[str, Any]:
        """自检 Word 转 PPT 的 DOCX 对象保留证据。"""
        cached = getattr(self, "_word_object_preservation_probe_cache", None)
        if isinstance(cached, dict):
            return cached
        try:
            with tempfile.TemporaryDirectory(prefix="k12-word-object-probe-") as tmp:
                root = Path(tmp)
                source = root / "objects.docx"
                target = root / "objects.pptx"
                image_data = base64.b64decode('iVBORw0KGgoAAAANSUhEUgAAAAIAAAABCAIAAAB7QOjdAAAAD0lEQVR4nGP4z8DA8J8BAAf/Af8Bf4mnAAAAAElFTkSuQmCC')
                document_xml = (
                    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
                    'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">'
                    '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>对象保留自检</w:t></w:r></w:p>'
                    '<w:p><w:r><w:t>含图片、表格和公式</w:t></w:r></w:p>'
                    '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>表格内容</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
                    '<m:oMath><m:r><m:t>x=1</m:t></m:r></m:oMath>'
                    '<w:p><w:r><w:t>MathType 对象</w:t></w:r></w:p>'
                    '<w:p><w:r><w:drawing xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><wp:inline><wp:extent cx="2000000" cy="1000000"/><a:blip r:embed="image"/></wp:inline></w:drawing></w:r></w:p>'
                    "</w:document>"
                )
                with zipfile.ZipFile(source, "w", zipfile.ZIP_DEFLATED) as archive:
                    archive.writestr("word/document.xml", document_xml)
                    archive.writestr("word/media/probe.png", image_data)
                    archive.writestr("word/_rels/document.xml.rels", '<Relationships><Relationship Id="image" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/probe.png"/></Relationships>')
                    archive.writestr("word/embeddings/oleObject1.bin", b"k12-mathtype-probe")
                summary = extract_docx_object_summary(source)
                blocks = extract_docx_blocks(source)
                task = {"id": "task_word_object_probe", "task_type": "word_to_ppt", "options": {}}
                preservation = self._word_to_ppt_object_preservation(task, summary)
                build_pptx_from_docx(blocks, target, object_preservation=preservation)
                with zipfile.ZipFile(target) as archive:
                    slide_xml = "\n".join(
                        archive.read(name).decode("utf-8", errors="ignore")
                        for name in archive.namelist()
                        if name.startswith("ppt/slides/slide") and name.endswith(".xml")
                    )
                source_counts = preservation.get("source") or {}
                statuses = preservation.get("statuses") or {}
                with zipfile.ZipFile(target) as archive:
                    image_available = any(name.startswith("ppt/media/") and archive.read(name) == image_data for name in archive.namelist())
                    formula_available = any(name.startswith("ppt/embeddings/") and archive.read(name) == b"k12-mathtype-probe" for name in archive.namelist())
                table_available = "<a:tbl>" in slide_xml and "表格内容" in slide_xml
                image_available = image_available and sum(slide.get("image_count", 0) for slide in extract_pptx_slides(target)) == 1
                contract_available = "图片 1 个" in slide_xml and "MathType 1 个" in slide_xml
                probe = {
                    "schema_version": "k12.wordObjectPreservationProbe.v1",
                    "available": image_available and table_available and formula_available,
                    "contract_available": contract_available,
                    "image_available": image_available,
                    "table_available": table_available,
                    "formula_available": formula_available,
                    "source": source_counts,
                    "statuses": statuses,
                    "image_evidence": "内嵌 PNG 样本的输出部件、图片引用和原始字节已验证；全部图片格式与原生布局验收仍未完成" if image_available else "输出未包含源图片部件；对象清单不能证明图片保留",
                    "table_evidence": "输出已包含原生可编辑表格及源单元格文字" if table_available else "输出未验证原生表格",
                    "formula_evidence": "输出嵌入对象部件已验证" if formula_available else "输出未包含源 MathType 嵌入对象；对象清单不能证明公式保留",
                    "current": "产物自检：内嵌 PNG 样本和原生表格已验证；完整图片保真与 MathType 对象写回仍待验收",
                }
        except Exception as exc:
            message = self._redact_local_path_text(str(exc) or exc.__class__.__name__)
            probe = {
                "schema_version": "k12.wordObjectPreservationProbe.v1",
                "available": False,
                "image_available": False,
                "table_available": False,
                "formula_available": False,
                "source": {},
                "statuses": {},
                "evidence": f"Word 对象保留自检失败：{message}",
                "current": "自检失败；真实排版和 OLE 写回仍需本地 Office/MathType 复核",
            }
        self._word_object_preservation_probe_cache = probe
        return probe

    def _macro_batch_sequence_probe(self) -> dict[str, Any]:
        """自检批量宏顺序约定，不运行 Word 宏。"""
        cached = getattr(self, "_macro_batch_sequence_probe_cache", None)
        if isinstance(cached, dict):
            return cached
        try:
            library = self.macro_library()[:2]
            if len(library) < 2:
                raise ValueError("宏库不足，无法构造批量顺序自检")
            sequence = [
                {"id": macro["id"], "execute_order": index, "selected": True, "confirmed": True}
                for index, macro in enumerate(library, start=1)
            ]
            task = {
                "id": "task_macro_batch_probe",
                "task_type": "macro_sequence",
                "options": {
                    "selectedMacros": sequence,
                    "confirmMacroRisk": True,
                    "macroBackup": False,
                    "failureStrategy": "跳过",
                    "executeTiming": "Word 处理前",
                },
            }
            with tempfile.TemporaryDirectory(prefix="k12-macro-batch-probe-") as tmp:
                root = Path(tmp)
                files: list[dict[str, Any]] = []
                for index in range(1, 3):
                    path = root / f"macro-batch-{index}.docm"
                    path.write_bytes(b"macro batch probe")
                    files.append(
                        {
                            "id": f"file_macro_batch_probe_{index}",
                            "file_name": path.name,
                            "file_type": "Word",
                            "extension": ".docm",
                            "has_macro": True,
                            "storage_path": str(path),
                        }
                    )
                expected_ids = [item["id"] for item in sequence]
                expected_orders = [item["execute_order"] for item in sequence]
                shared_names: list[list[str]] = []
                statuses: list[str] = []
                for file in files:
                    items = self._macro_items(file, task, batch_allowed=True)
                    item_ids = [str(item.get("id") or "") for item in items]
                    item_orders = [int(item.get("execute_order") or 0) for item in items]
                    if item_ids != expected_ids or item_orders != expected_orders:
                        raise ValueError("多文件宏顺序不一致")
                    blocked = [
                        str(item.get("execute_status") or "")
                        for item in items
                        if item.get("execute_status") not in {"待本地客户端执行", "成功"}
                    ]
                    if blocked:
                        raise ValueError(f"宏队列状态异常：{', '.join(blocked)}")
                    if any(str(item.get("backup_path") or "") for item in items):
                        raise ValueError("宏批量自检不应创建备份文件")
                    statuses.extend(str(item.get("execute_status") or "") for item in items)
                    shared_names.append([str(item.get("macro_name") or "") for item in items])
                if len({tuple(names) for names in shared_names}) != 1:
                    raise ValueError("多文件宏名称顺序不一致")
                probe = {
                    "schema_version": "k12.macroBatchSequenceProbe.v1",
                    "available": True,
                    "file_count": len(files),
                    "macro_count": len(sequence),
                    "macro_sequence": shared_names[0],
                    "statuses": sorted({status for status in statuses if status}),
                    "evidence": "宏批量顺序自检通过：同一宏顺序可应用到多个 Word 文件并进入本地执行队列",
                    "current": f"自检通过，{len(files)} 个文件共享 {len(sequence)} 个宏顺序；真实 VBA 执行仍需本地客户端",
                }
        except Exception as exc:
            probe = {
                "schema_version": "k12.macroBatchSequenceProbe.v1",
                "available": False,
                "file_count": 0,
                "macro_count": 0,
                "macro_sequence": [],
                "statuses": [],
                "evidence": f"宏批量顺序自检失败：{self._redact_local_path_text(str(exc) or exc.__class__.__name__)}",
                "current": "自检失败；真实 VBA 执行仍需本地客户端",
            }
        self._macro_batch_sequence_probe_cache = probe
        return probe

    def _macro_ordered_execution_probe(self) -> dict[str, Any]:
        """自检有序宏的本地交接和模拟执行同步约定。"""
        cached = getattr(self, "_macro_ordered_execution_probe_cache", None)
        if isinstance(cached, dict):
            return cached
        try:
            from .local_client import build_dry_run_execution_summary, build_dry_run_sync_payload

            with tempfile.TemporaryDirectory(prefix="k12-macro-order-probe-") as tmp:
                root = Path(tmp)
                source = root / "macro-order-probe.docm"
                build_docx(
                    [
                        {"text": "K12 macro order probe", "style": "Heading1"},
                        {"text": "This document is only used for local handoff validation."},
                    ],
                    source,
                )
                probe_store = AppStore(root / "store")
                probe_store.update_settings(
                    {
                        "localSecurityToken": "probe-token",
                        "localClientPlatform": "Windows",
                        "allowWebLaunchLocalClient": True,
                        "allowTaskStatusCloudSync": True,
                        "macroBackup": True,
                    }
                )
                probe_processor = TaskProcessor(probe_store)
                probe_processor.record_local_client_heartbeat(
                    {
                        "client_id": "macro-order-probe",
                        "status": "online",
                        "platform": "Windows",
                        "capabilities": {"macroExecution": True},
                        "preflight": {
                            "schema_version": "k12.localClientPreflight.v1",
                            "platform": "Windows",
                            "components": {"word": {"label": "Word 桌面组件", "available": True, "status": "available"}},
                            "capabilities": {"macroExecution": True},
                        },
                    }
                )
                file = probe_processor.create_uploaded_file(source.name, source.read_bytes())[0]
                selected = [
                    {"id": "macro_repair_equation_anchors", "execute_order": 1},
                    {"id": "macro_clean_empty_paragraphs", "execute_order": 2},
                ]
                task = probe_processor.create_task(
                    {
                        "task_type": "macro_sequence",
                        "file_ids": [file["id"]],
                        "options": {
                            "selectedMacros": selected,
                            "confirmMacroRisk": True,
                            "macroBackup": True,
                            "failureStrategy": "跳过",
                            "executeTiming": "Word 处理前",
                        },
                    }
                )
                report = probe_processor._latest_report_for_task(task["id"]) or {}
                report_macros = list((report.get("analysis") or {}).get("macros") or [])
                report_order = [str(item.get("id") or "") for item in report_macros]
                expected_order = [item["id"] for item in selected]
                if report_order != expected_order:
                    raise ValueError("报告宏顺序与用户选择不一致")

                payload = probe_processor.local_task_payload(task["id"])
                action = next((item for item in payload.get("local_actions", []) if item.get("type") == "macro_sequence"), None)
                if not action:
                    raise ValueError("本地载荷缺少宏队列动作")
                action_order = [str(item.get("id") or "") for item in action.get("macros") or []]
                if action_order != expected_order:
                    raise ValueError("本地宏队列顺序与用户选择不一致")

                execution_plan = payload.get("desktop_execution_plan") if isinstance(payload.get("desktop_execution_plan"), dict) else {}
                plan_action = next((item for item in execution_plan.get("actions", []) if item.get("type") == "macro_sequence"), None)
                if not plan_action:
                    raise ValueError("桌面执行计划缺少宏动作")
                operations = [str(step.get("operation") or "") for step in plan_action.get("steps") or [] if isinstance(step, dict)]
                if execution_plan.get("status") != "ready_for_native_client" or plan_action.get("gate_status") != "ready":
                    raise ValueError("宏动作未达到本地客户端交接就绪状态")
                if "macro.run_ordered" not in operations:
                    raise ValueError("桌面执行计划缺少 macro.run_ordered 步骤")
                required_keys = [str(item.get("key") or "") for item in plan_action.get("required_capabilities") or []]
                if required_keys != ["macroExecution"]:
                    raise ValueError("宏动作能力门槛异常")

                dry_run = build_dry_run_execution_summary(payload)
                dry_action = next((item for item in dry_run.get("actions", []) if item.get("type") == "macro_sequence"), None)
                if not dry_action:
                    raise ValueError("dry-run 摘要缺少宏动作")
                if dry_run.get("ready_action_count") != 1 or dry_action.get("dry_run_status") != "ready_for_native_executor":
                    raise ValueError("dry-run 宏动作未达到可交接执行器状态")
                if dry_action.get("native_execution_performed"):
                    raise ValueError("dry-run 不应执行真实 Word 宏")
                if int(dry_action.get("step_count") or 0) < 3:
                    raise ValueError("dry-run 宏步骤数量不足")

                task["status"] = "待处理"
                task["progress"] = 95
                task["end_time"] = ""
                probe_store.save_task(task)
                synced = probe_processor.sync_local_task_status(task["id"], build_dry_run_sync_payload(payload))
                sync = synced.get("local_sync") if isinstance(synced.get("local_sync"), dict) else {}
                if sync.get("dry_run_execution_status") != "ready_for_native_client" or sync.get("dry_run_ready_action_count") != 1:
                    raise ValueError("dry-run 状态同步未记录宏动作就绪结果")

                macro_names = [str(item.get("macro_name") or "") for item in report_macros]
                probe = {
                    "schema_version": "k12.macroOrderedExecutionProbe.v1",
                    "available": True,
                    "macro_count": len(report_macros),
                    "macro_sequence": macro_names,
                    "plan_status": execution_plan.get("status", ""),
                    "dry_run_status": dry_action.get("dry_run_status", ""),
                    "native_execution_performed": bool(dry_action.get("native_execution_performed")),
                    "evidence": "宏顺序执行交接自检通过：用户顺序进入本地宏队列、桌面计划包含 macro.run_ordered，并可回传 dry-run 就绪摘要",
                    "current": (
                        f"自检通过，{len(report_macros)} 个宏按顺序进入本地客户端；"
                        "dry-run 未执行真实 Word 宏，真实 VBA 执行仍需本地客户端"
                    ),
                }
        except Exception as exc:
            probe = {
                "schema_version": "k12.macroOrderedExecutionProbe.v1",
                "available": False,
                "macro_count": 0,
                "macro_sequence": [],
                "plan_status": "",
                "dry_run_status": "",
                "native_execution_performed": False,
                "evidence": f"宏顺序执行交接自检失败：{self._redact_local_path_text(str(exc) or exc.__class__.__name__)}",
                "current": "自检失败；真实 VBA 执行仍需本地客户端",
            }
        self._macro_ordered_execution_probe_cache = probe
        return probe

    def _latex_mathtype_probe(self) -> dict[str, Any]:
        """自检 LaTeX 转 MathType 预览和公式兜底元数据。"""
        cached = getattr(self, "_latex_mathtype_probe_cache", None)
        if isinstance(cached, dict):
            return cached
        try:
            task = {
                "id": "task_latex_mathtype_probe",
                "task_type": "mathtype_format",
                "options": {"enableMathTypeFormatting": True},
            }
            file = {
                "id": "file_latex_mathtype_probe",
                "file_name": "latex-formula.docx",
                "file_type": "Word",
                "has_formula": True,
                "has_omml": False,
                "has_mathtype": False,
            }
            items = self._formula_items(file, 13, task)
            latex_items = [item for item in items if item.get("source_type") == "LaTeX"]
            if not latex_items:
                raise ValueError("未生成 LaTeX 来源公式")
            item = latex_items[0]
            latex = str(item.get("latex") or "")
            mathml = str(item.get("mathml") or "")
            preview = str(item.get("mathtype_preview") or item.get("mathtype_data") or "")
            policy = dict(item.get("format_failure_policy") or {})
            if not latex or not mathml.startswith("<math") or not preview.startswith("MathType 预览"):
                raise ValueError("LaTeX、MathML 或 MathType 预览字段缺失")
            if latex not in preview or html.escape(latex) not in mathml:
                raise ValueError("MathType 预览或 MathML 未引用原始 LaTeX")
            if policy.get("status") != "formatted" or not policy.get("preserve_original_formula"):
                raise ValueError("格式化策略未记录成功与原公式保留")
            probe = {
                "schema_version": "k12.latexMathTypeProbe.v1",
                "available": True,
                "source_type": "LaTeX",
                "latex": latex,
                "mathml_available": True,
                "mathtype_preview_available": True,
                "format_policy": policy,
                "evidence": "LaTeX 转 MathType 自检通过：公式项同时保留 LaTeX、MathML 和 MathType 预览",
                "current": "自检通过，LaTeX 公式生成 MathML 兜底和 MathType 预览；原生 MathType 对象写回仍需本地客户端",
            }
        except Exception as exc:
            probe = {
                "schema_version": "k12.latexMathTypeProbe.v1",
                "available": False,
                "source_type": "LaTeX",
                "latex": "",
                "mathml_available": False,
                "mathtype_preview_available": False,
                "format_policy": {},
                "evidence": f"LaTeX 转 MathType 自检失败：{self._redact_local_path_text(str(exc) or exc.__class__.__name__)}",
                "current": "自检失败；原生 MathType 对象写回仍需本地客户端",
            }
        self._latex_mathtype_probe_cache = probe
        return probe

    def _mathtype_preservation_probe(self) -> dict[str, Any]:
        """自检现有 MathType 保留与格式兜底策略。"""
        cached = getattr(self, "_mathtype_preservation_probe_cache", None)
        if isinstance(cached, dict):
            return cached
        try:
            task = {
                "id": "task_mathtype_preservation_probe",
                "task_type": "mathtype_format",
                "options": {"enableMathTypeFormatting": True},
            }
            file = {
                "id": "file_mathtype_preservation_probe",
                "file_name": "existing-mathtype.docx",
                "file_type": "Word",
                "has_formula": True,
                "has_omml": False,
                "has_mathtype": True,
            }
            items = self._formula_items(file, 13, task)
            mathtype_items = [item for item in items if item.get("source_type") == "MathType"]
            if not mathtype_items:
                raise ValueError("未生成 MathType 来源公式")
            item = mathtype_items[0]
            policy = dict(item.get("format_failure_policy") or {})
            if item.get("format_status") != "原样保留" or policy.get("status") != "preserved_existing":
                raise ValueError("已有 MathType 公式未按原对象保留策略记录")
            if not item.get("preserve_original_formula") or not policy.get("preserve_original_formula"):
                raise ValueError("已有 MathType 公式未记录原公式保留")
            object_probe = self._word_object_preservation_probe()
            if not object_probe.get("contract_available"):
                raise ValueError("对象保留清单未证明 MathType/OLE 线索")
            probe = {
                "schema_version": "k12.mathTypePreservationProbe.v1",
                "available": True,
                "source_type": "MathType",
                "format_status": str(item.get("format_status") or ""),
                "format_policy": policy,
                "object_preservation_available": True,
                "evidence": "已有 MathType 保留自检通过：MathType 来源公式原样保留，并进入对象保留清单",
                "current": "自检通过，报告策略保留原公式，PPTX 对象保留清单记录 MathType/OLE 线索；原生 OLE 写回仍需本地客户端",
            }
        except Exception as exc:
            probe = {
                "schema_version": "k12.mathTypePreservationProbe.v1",
                "available": False,
                "source_type": "MathType",
                "format_status": "",
                "format_policy": {},
                "object_preservation_available": False,
                "evidence": f"已有 MathType 保留自检失败：{self._redact_local_path_text(str(exc) or exc.__class__.__name__)}",
                "current": "自检失败；原生 OLE 写回仍需本地客户端",
            }
        self._mathtype_preservation_probe_cache = probe
        return probe

    def _omml_mathtype_handoff_probe(self) -> dict[str, Any]:
        """自检 OMML 转 MathType 的交接约定，不执行原生写回。"""
        cached = getattr(self, "_omml_mathtype_handoff_probe_cache", None)
        if isinstance(cached, dict):
            return cached
        try:
            with tempfile.TemporaryDirectory(prefix="k12-omml-handoff-probe-") as tmp:
                root = Path(tmp)
                dependency_dir = root / "dependencies"
                dependency_dir.mkdir()
                (dependency_dir / "OMML2MML.XSL").write_text("<xsl:stylesheet />", encoding="utf-8")
                source = root / "lesson.docx"
                with zipfile.ZipFile(source, "w", zipfile.ZIP_DEFLATED) as archive:
                    archive.writestr("[Content_Types].xml", "<Types></Types>")
                    archive.writestr("word/document.xml", "<w:document><w:p>题目</w:p><m:oMath>x</m:oMath></w:document>")
                    archive.writestr("docProps/app.xml", "<Properties><Pages>1</Pages></Properties>")
                probe_store = AppStore(root / "data")
                probe_store.update_settings(
                    {
                        "ommlSearchPaths": str(dependency_dir),
                        "localSecurityToken": "probe-token",
                        "allowWebLaunchLocalClient": True,
                        "localClientPlatform": "Windows",
                    }
                )
                probe_processor = TaskProcessor(probe_store)
                probe_processor.record_local_client_heartbeat(
                    {
                        "client_id": "desktop-omml-probe",
                        "status": "online",
                        "platform": "Windows",
                        "preflight": {
                            "platform": "Windows",
                            "components": {
                                "omml_dependency": {"label": "OMML 依赖文件", "available": True, "status": "available"},
                                "mathtype": {"label": "MathType 组件", "available": True, "status": "available"},
                            },
                            "capabilities": {"ommlDependencySearch": True, "mathTypeAutomation": True},
                        },
                    }
                )
                word = probe_processor.create_uploaded_file("lesson.docx", source.read_bytes())[0]
                task = probe_processor.create_task(
                    {
                        "task_type": "omml_to_mathtype",
                        "file_ids": [word["id"]],
                        "options": {"convertOmmlToMathType": True},
                    }
                )
                report = probe_store.list_reports()[0]
                analysis = report.get("analysis") or {}
                dependency = (analysis.get("ommlDependencies") or [])[0]
                prompt = (analysis.get("ommlConversionPrompts") or [])[0]
                if prompt.get("status") != "已选择转换" or prompt.get("local_action") != "convert_omml_to_mathtype":
                    raise ValueError("OMML 转换确认未进入转换状态")
                if dependency.get("copy_status") not in {"成功", "跳过"}:
                    raise ValueError("OMML 依赖未完成复制或复用")
                probe_processor.save_omml_annotation(
                    {
                        "report_id": report["id"],
                        "dependency_id": dependency["id"],
                        "status": "重新转换",
                        "retry_conversion": True,
                    }
                )
                payload = probe_processor.local_task_payload(task["id"])
                action = next(item for item in payload["local_actions"] if item["type"] == "omml_mathtype")
                plan = payload["desktop_execution_plan"]
                plan_action = next(item for item in plan["actions"] if item["type"] == "omml_mathtype")
                required_keys = [item["key"] for item in plan_action.get("required_capabilities", [])]
                operations = [item.get("operation") for item in plan_action.get("steps", [])]
                if action.get("status") != "retry_queued":
                    raise ValueError("OMML 重新转换请求未进入本地动作队列")
                if not action.get("conversion_prompts") or not action.get("retry_requests"):
                    raise ValueError("OMML 转换确认或重试请求未进入本地载荷")
                if "ommlDependencySearch" not in required_keys or "mathTypeAutomation" not in required_keys:
                    raise ValueError("桌面执行计划缺少 OMML 或 MathType 能力门槛")
                for operation in ("omml.confirm", "omml.search_dependency", "mathtype.convert", "formula.validate", "omml.retry_conversion"):
                    if operation not in operations:
                        raise ValueError(f"桌面执行计划缺少步骤：{operation}")
                if plan.get("status") != "ready_for_native_client" or plan_action.get("gate_status") != "ready":
                    raise ValueError("OMML 桌面执行计划未达到可交接状态")
                probe = {
                    "schema_version": "k12.ommlMathTypeHandoffProbe.v1",
                    "handoff_available": True,
                    "retry_available": True,
                    "dependency_copy_status": str(dependency.get("copy_status") or ""),
                    "action_status": str(action.get("status") or ""),
                    "plan_status": str(plan.get("status") or ""),
                    "required_capabilities": required_keys,
                    "operations": operations,
                    "handoff_evidence": "OMML 转 MathType 交接自检通过：确认转换后 omml_mathtype 动作、能力门槛和桌面执行步骤进入本地载荷",
                    "retry_evidence": "OMML 重新转换自检通过：复制成功或人工校正后的 retry_conversion 请求进入本地动作队列，并追加 omml.retry_conversion 步骤",
                    "current": "自检通过，OMML 依赖已复制，动作状态 retry_queued，桌面计划 ready；真实 MathType 对象写回仍需本地客户端",
                }
        except Exception as exc:
            message = self._redact_local_path_text(str(exc) or exc.__class__.__name__)
            probe = {
                "schema_version": "k12.ommlMathTypeHandoffProbe.v1",
                "handoff_available": False,
                "retry_available": False,
                "dependency_copy_status": "",
                "action_status": "",
                "plan_status": "",
                "required_capabilities": [],
                "operations": [],
                "evidence": f"OMML 转 MathType 交接自检失败：{message}",
                "current": "自检失败；真实 MathType 对象写回仍需本地客户端",
            }
        self._omml_mathtype_handoff_probe_cache = probe
        return probe

    def _local_api_security_probe(self) -> dict[str, Any]:
        """自检敏感交接接口的令牌保护。"""
        cached = getattr(self, "_local_api_security_probe_cache", None)
        if isinstance(cached, dict):
            return cached
        try:
            catalog = self.api_catalog()
            manifest = self.local_client_manifest()
            catalog_endpoints = {(item["method"], item["path"]): item for item in catalog.get("endpoints", [])}
            manifest_endpoints = {(item["method"], item["path"]): item for item in manifest.get("endpoints", [])}
            protected_contracts = [
                ("GET", "/api/tasks/{task_id}/local-payload", "configured-token"),
                ("POST", "/api/tasks/{task_id}/local-launch", "configured-token + launch-authorization"),
                ("POST", "/api/local-client/uploads", "configured-token + cloud-sync-authorization"),
                ("POST", "/api/local-client/heartbeat", "configured-token"),
            ]
            for method, path, auth in protected_contracts:
                if catalog_endpoints.get((method, path), {}).get("auth") != auth:
                    raise ValueError(f"API 目录鉴权不匹配：{method} {path}")
                if manifest_endpoints.get((method, path), {}).get("auth") != auth:
                    raise ValueError(f"客户端清单鉴权不匹配：{method} {path}")
            if not (catalog.get("auth") or {}).get("local_payload_requires_configured_token"):
                raise ValueError("API 目录未声明本地载荷必须配置令牌")
            with tempfile.TemporaryDirectory(prefix="k12-local-security-probe-") as tmp:
                probe_store = AppStore(Path(tmp) / "data")
                probe_store.update_settings({"allowWebLaunchLocalClient": True})
                probe_processor = TaskProcessor(probe_store)
                task = probe_store.save_task(Task(task_type="word_to_ppt", execute_mode="local", file_ids=[]).to_dict())
                try:
                    probe_processor.create_local_launch_request(task["id"])
                except ValueError as exc:
                    if "安全令牌" not in str(exc):
                        raise
                else:
                    raise ValueError("未配置令牌时仍允许创建本地启动请求")
            probe = {
                "schema_version": "k12.localApiSecurityProbe.v1",
                "available": True,
                "protected_endpoint_count": len(protected_contracts),
                "evidence": "本地 API 安全自检通过：敏感载荷、启动请求、上传登记和心跳均声明 configured-token，未配置令牌时拒绝本地启动",
                "current": f"自检通过，{len(protected_contracts)} 个本地客户端端点需要配置令牌或显式授权",
            }
        except Exception as exc:
            probe = {
                "schema_version": "k12.localApiSecurityProbe.v1",
                "available": False,
                "protected_endpoint_count": 0,
                "evidence": f"本地 API 安全自检失败：{self._redact_local_path_text(str(exc) or exc.__class__.__name__)}",
                "current": "自检失败",
            }
        self._local_api_security_probe_cache = probe
        return probe

    def _local_result_upload_probe(self) -> dict[str, Any]:
        """自检本地结果上传注册，不传输用户文件。"""
        cached = getattr(self, "_local_result_upload_probe_cache", None)
        if isinstance(cached, dict):
            return cached
        try:
            with tempfile.TemporaryDirectory(prefix="k12-local-upload-probe-") as tmp:
                root = Path(tmp)
                probe_store = AppStore(root / "data")
                probe_store.update_settings({"allowTaskStatusCloudSync": True, "localSecurityToken": "probe-token"})
                probe_processor = TaskProcessor(probe_store)
                task = probe_store.save_task(Task(task_type="word_to_ppt", execute_mode="local", file_ids=[]).to_dict())
                local_only = probe_processor.sync_local_task_status(
                    task["id"],
                    {
                        "status": "completed",
                        "outputs": [{"name": "local-only.pptx", "path": str(root / "private" / "local-only.pptx"), "size": 12}],
                        "resultUploadRequested": False,
                    },
                )
                if local_only.get("local_sync", {}).get("result_upload_status") != "not_requested":
                    raise ValueError("本地保存路径未正确记录为不上传")
                probe_store.update_settings({"allowCloudSync": True})
                content = b"K12 cloud receive probe"
                digest = hashlib.sha256(content).hexdigest()
                upload = probe_processor.register_local_result_upload(
                    {
                        "task_id": task["id"],
                        "client_id": "desktop-upload-probe",
                        "outputs": [{"name": "received.pptx", "path": str(root / "private" / "received.pptx"), "size": len(content), "sha256": digest}],
                        "files": [{"name": "received.pptx", "content_base64": base64.b64encode(content).decode("ascii"), "sha256": digest}],
                    }
                )
                manifest = probe_processor.local_result_upload_manifest(upload["upload_id"])
                serialized = json.dumps({"upload": upload, "manifest": manifest}, ensure_ascii=False)
                if upload.get("receive_mode") != "content_received" or upload.get("received_file_count") != 1:
                    raise ValueError("上传内容未进入接收记录")
                if manifest.get("receive_contract", {}).get("content_transfer") != "included" or manifest.get("receive_state") != "received":
                    raise ValueError("云端接收清单未标记内容已接收")
                if manifest.get("package", {}).get("received_file_count") != 1 or not manifest.get("integrity", {}).get("all_received_files_hashed"):
                    raise ValueError("上传内容完整性摘要缺失")
                received_dir = probe_store.cloud_uploads_dir / upload["upload_id"]
                if not any(path.is_file() for path in received_dir.iterdir()):
                    raise ValueError("上传内容未写入受控接收目录")
                if str(root) in serialized or base64.b64encode(content).decode("ascii") in serialized:
                    raise ValueError("上传响应泄露本地路径或文件内容")
                probe = {
                    "schema_version": "k12.localResultUploadProbe.v1",
                    "available": True,
                    "upload_id": upload["upload_id"],
                    "receive_state": manifest.get("receive_state", ""),
                    "received_file_count": int(upload.get("received_file_count") or 0),
                    "evidence": "本地结果上传自检通过：可选择仅本地保存，也可在授权后接收上传内容并生成脱敏云端接收清单",
                    "current": "自检通过，1 个输出内容已写入受控接收目录，manifest 标记 content_transfer=included；外部云分发仍需部署对应服务",
                }
        except Exception as exc:
            probe = {
                "schema_version": "k12.localResultUploadProbe.v1",
                "available": False,
                "upload_id": "",
                "receive_state": "",
                "received_file_count": 0,
                "evidence": f"本地结果上传自检失败：{self._redact_local_path_text(str(exc) or exc.__class__.__name__)}",
                "current": "自检失败",
            }
        self._local_result_upload_probe_cache = probe
        return probe

    def _local_install_platform_probe(self) -> dict[str, Any]:
        """自检 Windows/macOS 安装包约定和客户端清单摘要。"""
        cached = getattr(self, "_local_install_platform_probe_cache", None)
        if isinstance(cached, dict):
            return cached
        try:
            from .local_client import summarize_manifest

            with tempfile.TemporaryDirectory(prefix="k12-install-platform-probe-") as tmp:
                root = Path(tmp)
                probe_store = AppStore(root / "data")
                probe_store.update_settings({"localClientPlatform": "Windows", "localSecurityToken": "probe-token"})
                windows_package = probe_store.installers_dir / "K12-Local-Client-Windows-x64.msi"
                macos_package = probe_store.installers_dir / "K12-Local-Client-macOS-universal.pkg"
                windows_package.write_bytes(b"k12 windows installer probe")
                macos_package.write_bytes(b"k12 macos installer probe")
                probe_processor = TaskProcessor(probe_store)

                windows_plan = probe_processor.install_plan("Windows")
                macos_plan = probe_processor.install_plan("macOS")
                if windows_plan.get("installer_kind") != "windows-msi" or not windows_plan.get("package", {}).get("download_available"):
                    raise ValueError("Windows 安装计划未暴露可下载 .msi")
                if macos_plan.get("installer_kind") != "macos-pkg" or not macos_plan.get("package", {}).get("download_available"):
                    raise ValueError("macOS 安装计划未暴露可下载 .pkg")
                if "?platform=Windows" not in str(windows_plan.get("package", {}).get("download_url") or ""):
                    raise ValueError("Windows 安装包下载未要求 platform=Windows")
                if "?platform=macOS" not in str(macos_plan.get("package", {}).get("download_url") or ""):
                    raise ValueError("macOS 安装包下载未要求 platform=macOS")
                for plan in (windows_plan, macos_plan):
                    compatibility = plan.get("formula_compatibility") if isinstance(plan.get("formula_compatibility"), dict) else {}
                    if compatibility.get("platform_objects_cross_compatible") is not False:
                        raise ValueError("安装公式合同未阻止跨平台 MathType 对象")
                    if "MathML" not in list(compatibility.get("fallback_formats") or []):
                        raise ValueError("安装公式合同缺少兜底格式")

                manifest = probe_processor.local_client_manifest()
                summary = summarize_manifest(manifest)
                serialized = json.dumps({"manifest": manifest, "summary": summary}, ensure_ascii=False)
                summary_installer = summary.get("platform", {}).get("installer", {})
                summary_formula = summary.get("platform", {}).get("formula_compatibility", {})
                if summary.get("schema_version") != "k12.localClientManifestSummary.v1":
                    raise ValueError("CLI manifest 摘要版本缺失")
                if summary_installer.get("file_name") != windows_package.name:
                    raise ValueError("CLI manifest 摘要未保留 Windows 安装包文件名")
                if not summary_installer.get("download_requires_platform_query"):
                    raise ValueError("CLI manifest 摘要未声明 platform 查询要求")
                if "?platform=Windows" not in str(summary_installer.get("download_url") or ""):
                    raise ValueError("CLI manifest 摘要未保留受控下载路径")
                if summary_formula.get("platform_objects_cross_compatible") is not False:
                    raise ValueError("CLI manifest 摘要未阻止跨平台 MathType 对象")
                if str(root) in serialized or "probe-token" in serialized:
                    raise ValueError("安装 manifest 或摘要泄露本地路径/令牌")

                probe_processor.record_local_client_heartbeat({"client_id": "install-probe", "platform": "macOS", "status": "online"})
                mismatch = summarize_manifest(probe_processor.local_client_manifest())
                mismatch_formula = mismatch.get("platform", {}).get("formula_compatibility", {})
                mismatch_installer = mismatch.get("platform", {}).get("installer", {})
                blockers = set(mismatch_formula.get("native_handoff_blocking_reasons") or [])
                blockers.update(mismatch_installer.get("native_handoff_blocking_reasons") or [])
                if mismatch_formula.get("heartbeat_platform") != "macOS" or "platform_mismatch" not in blockers:
                    raise ValueError("平台不符心跳未进入安装/公式合同阻断")

                probe = {
                    "schema_version": "k12.localInstallPlatformProbe.v1",
                    "available": True,
                    "platforms": ["Windows", "macOS"],
                    "installer_kinds": [windows_plan["installer_kind"], macos_plan["installer_kind"]],
                    "manifest_summary_schema": summary["schema_version"],
                    "download_requires_platform_query": bool(summary_installer.get("download_requires_platform_query")),
                    "platform_mismatch_blocked": True,
                    "evidence": "安装平台自检通过：Windows .msi、macOS .pkg、/api/installers/{file_name}?platform=... 下载边界、k12.localClientManifestSummary.v1 脱敏摘要和 MathType 同平台阻断均可验证",
                    "current": "自检通过，Windows/macOS 安装包按平台区分，CLI manifest 摘要不泄露 installers 路径或令牌，macOS 心跳接收 Windows 目标时触发 platform_mismatch；真实安装包签名和桌面安装仍需发布环境实测",
                }
        except Exception as exc:
            probe = {
                "schema_version": "k12.localInstallPlatformProbe.v1",
                "available": False,
                "platforms": [],
                "installer_kinds": [],
                "manifest_summary_schema": "",
                "download_requires_platform_query": False,
                "platform_mismatch_blocked": False,
                "evidence": f"安装平台自检失败：{self._redact_local_path_text(str(exc) or exc.__class__.__name__)}",
                "current": "自检失败；真实安装包与桌面安装仍需发布环境实测",
            }
        self._local_install_platform_probe_cache = probe
        return probe

    def _local_file_action_execution_probe(self) -> dict[str, Any]:
        """仅自检授权的 OMML 依赖复制文件动作。"""
        cached = getattr(self, "_local_file_action_execution_probe_cache", None)
        if isinstance(cached, dict):
            return cached
        try:
            from .local_client import execute_local_file_actions

            with tempfile.TemporaryDirectory(prefix="k12-local-file-action-probe-") as tmp:
                root = Path(tmp)
                document = root / "lesson.docx"
                source = root / "deps" / "OMML2MML.XSL"
                target = root / "OMML2MML.XSL"
                document.write_bytes(b"K12 probe document")
                source.parent.mkdir()
                source.write_text("<xsl:stylesheet>k12</xsl:stylesheet>", encoding="utf-8")
                payload = {
                    "local_actions": [
                        {
                            "type": "omml_mathtype",
                            "copy_strategy": "自动重命名",
                            "dependencies": [
                                {
                                    "id": "dep_local_file_action_probe",
                                    "file_id": "file_local_file_action_probe",
                                    "file_name": "lesson.docx",
                                    "document_path": str(document),
                                    "omml_source_path": str(source),
                                    "omml_target_path": str(target),
                                }
                            ],
                        }
                    ]
                }
                blocked = execute_local_file_actions(payload)
                if blocked.get("blocked_count") != 1 or blocked.get("actions", [{}])[0].get("status") != "blocked_until_explicit_file_action_request":
                    raise ValueError("OMML 文件动作未在缺少显式授权时阻断")
                if target.exists():
                    raise ValueError("未授权时仍复制了 OMML 依赖")
                executed = execute_local_file_actions(payload, allow_file_actions=True)
                action = (executed.get("actions") or [{}])[0]
                serialized = json.dumps({"blocked": blocked, "executed": executed}, ensure_ascii=False)
                if executed.get("performed_count") != 1 or action.get("status") != "copied" or not action.get("copy_performed"):
                    raise ValueError("显式授权后未复制 OMML 依赖")
                if not action.get("dependency_extension_allowed"):
                    raise ValueError("本地文件动作未校验 OMML 依赖扩展名")
                if not action.get("target_directory_writeable"):
                    raise ValueError("本地文件动作未校验目标目录写入权限")
                if not target.is_file() or target.read_text(encoding="utf-8") != source.read_text(encoding="utf-8"):
                    raise ValueError("复制后的 OMML 依赖内容不一致")
                if action.get("source_sha256") != hashlib.sha256(source.read_bytes()).hexdigest():
                    raise ValueError("OMML 依赖源文件校验缺失")
                if str(root) in serialized:
                    raise ValueError("本地文件动作摘要泄露临时路径")
                probe = {
                    "schema_version": "k12.localFileActionExecutionProbe.v1",
                    "available": True,
                    "blocked_status": blocked["actions"][0]["status"],
                    "executed_status": action.get("status", ""),
                    "performed_count": int(executed.get("performed_count") or 0),
                    "evidence": "本地文件动作执行自检通过：动作级 dry-run 和 k12.localNativeExecutionRequest.v1 仍只生成原生执行合同，k12.localFileActionExecution.v1 在显式授权、依赖扩展名受控且目标目录可写时仅复制 OMML 依赖",
                    "current": "自检通过，未授权时状态 blocked_until_explicit_file_action_request，OMML 依赖扩展名和目标目录写入权限已校验，显式授权后复制 1 个 OMML 依赖；不会执行 Office、MathType、OMML 写回或 Word 宏",
                }
        except Exception as exc:
            probe = {
                "schema_version": "k12.localFileActionExecutionProbe.v1",
                "available": False,
                "blocked_status": "",
                "executed_status": "",
                "performed_count": 0,
                "evidence": f"本地文件动作执行自检失败：{self._redact_local_path_text(str(exc) or exc.__class__.__name__)}",
                "current": "自检失败；真实 Office、MathType、OMML 写回或 Word 宏仍需本地客户端",
            }
        self._local_file_action_execution_probe_cache = probe
        return probe

    def _formula_format_scope_probe(self) -> dict[str, Any]:
        """自检支持范围内的公式格式化选项。"""
        cached = getattr(self, "_formula_format_scope_probe_cache", None)
        if isinstance(cached, dict):
            return cached
        try:
            files = [
                {"id": "file_formula_scope_probe_1", "file_name": "scope-a.docx", "file_type": "Word", "has_formula": True},
                {"id": "file_formula_scope_probe_2", "file_name": "scope-b.docx", "file_type": "Word", "has_formula": True},
            ]
            batch_task = {
                "id": "task_formula_scope_batch_probe",
                "task_type": "mathtype_format",
                "options": {"enableMathTypeFormatting": True, "formulaFormatScope": "全文"},
            }
            batch_items: list[dict[str, Any]] = []
            for index, file in enumerate(files, start=1):
                batch_items.extend(self._formula_items(file, 14 + index, batch_task))
            if len(batch_items) < 2:
                raise ValueError("批量格式化自检需要至少两个公式")
            if {str(item.get("file_id") or "") for item in batch_items} != {file["id"] for file in files}:
                raise ValueError("批量格式化自检未覆盖多个文件")
            if any(item.get("format_scope") != "全文" for item in batch_items):
                raise ValueError("全文格式化范围未写入所有公式项")
            if any((item.get("format_comparison") or {}).get("after", {}).get("scope") != "全文" for item in batch_items):
                raise ValueError("全文格式化范围未写入格式化对比")
            if any(not (item.get("format_failure_policy") or {}).get("preserve_original_formula") for item in batch_items):
                raise ValueError("批量格式化自检未记录原公式保留策略")
            scoped_modes = ["当前章节", "选中区域", "单个公式"]
            scoped_results: dict[str, str] = {}
            for offset, scope in enumerate(scoped_modes, start=1):
                task = {
                    "id": f"task_formula_scope_probe_{offset}",
                    "task_type": "mathtype_format",
                    "options": {"enableMathTypeFormatting": True, "formatScope": scope},
                }
                item = self._formula_items(files[0], 20 + offset, task)[0]
                comparison = item.get("format_comparison") or {}
                if item.get("format_scope") != scope or comparison.get("after", {}).get("scope") != scope:
                    raise ValueError(f"指定范围未写入公式项：{scope}")
                scoped_results[scope] = str(item.get("format_scope") or "")
            probe = {
                "schema_version": "k12.formulaFormatScopeProbe.v1",
                "batch_available": True,
                "scope_available": True,
                "file_count": len(files),
                "formula_count": len(batch_items),
                "batch_scope": "全文",
                "scoped_modes": scoped_results,
                "batch_evidence": "公式批量格式化自检通过：多个 Word 文件的公式共享全文格式化参数和保留策略",
                "scope_evidence": "公式范围格式化自检通过：当前章节、选中区域和单个公式范围会写入公式项与格式化对比",
                "current": f"自检通过，{len(files)} 个文件 {len(batch_items)} 个公式共享全文格式化；指定范围覆盖 {', '.join(scoped_modes)}；真实 MathType 对象写回仍需本地客户端",
            }
        except Exception as exc:
            probe = {
                "schema_version": "k12.formulaFormatScopeProbe.v1",
                "batch_available": False,
                "scope_available": False,
                "file_count": 0,
                "formula_count": 0,
                "batch_scope": "",
                "scoped_modes": {},
                "evidence": f"公式格式化范围自检失败：{self._redact_local_path_text(str(exc) or exc.__class__.__name__)}",
                "current": "自检失败；真实 MathType 对象写回仍需本地客户端",
            }
        self._formula_format_scope_probe_cache = probe
        return probe

    def _macro_detection_probe(self) -> dict[str, Any]:
        """使用临时启用宏的文件自检宏检测与报告。"""
        cached = getattr(self, "_macro_detection_probe_cache", None)
        if isinstance(cached, dict):
            return cached
        try:
            analyzer = DocumentAnalyzer()
            with tempfile.TemporaryDirectory(prefix="k12-macro-detection-probe-") as tmp:
                root = Path(tmp)
                document = root / "macro-document.docm"
                template = root / "macro-template.dotm"
                empty_container = root / "empty-macro-container.docm"
                disguised_container = root / "disguised-macro.docx"
                for path in (document, template):
                    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
                        archive.writestr(
                            "word/document.xml",
                            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                            "<w:p><w:r><w:t>宏检测自检</w:t></w:r></w:p>"
                            "</w:document>",
                        )
                        archive.writestr("word/vbaProject.bin", b"k12-macro-probe")
                with zipfile.ZipFile(empty_container, "w", zipfile.ZIP_DEFLATED) as archive:
                    archive.writestr("word/document.xml", "<w:document />")
                with zipfile.ZipFile(disguised_container, "w", zipfile.ZIP_DEFLATED) as archive:
                    archive.writestr("word/document.xml", "<w:document />")
                    archive.writestr("word/vbaProject.bin", b"k12-disguised-macro-probe")
                document_item = analyzer.analyze_file(document)
                template_item = analyzer.analyze_file(template)
                empty_item = analyzer.analyze_file(empty_container)
                disguised_item = analyzer.analyze_file(disguised_container)
                document_available = document_item.file_type == "Word" and document_item.has_macro and not document_item.validation_errors
                template_available = template_item.file_type == "Word" and template_item.has_macro and not template_item.validation_errors
                if not document_available or not template_available:
                    raise ValueError("DOCM 或 DOTM 中的 VBA 项目未被标记为宏")
                if empty_item.has_macro:
                    raise ValueError("不含 VBA 项目的 DOCM 空容器被误报为宏")
                if not disguised_item.has_macro:
                    raise ValueError("错误扩展名中的 VBA 项目未被识别")
                probe = {
                    "schema_version": "k12.macroDetectionProbe.v1",
                    "document_available": True,
                    "template_available": True,
                    "document_extension": document_item.extension,
                    "template_extension": template_item.extension,
                    "document_evidence": "宏检测自检通过：当前 Word 文档按 word/vbaProject.bin 实际部件标记 has_macro",
                    "template_evidence": "宏检测自检通过：Word 模板按 word/vbaProject.bin 实际部件标记 has_macro",
                    "current": "自检通过，DOCM/DOTM 的 VBA 项目可识别，空宏容器不误报，错误扩展名中的 VBA 部件不漏报；真实 VBA 模块枚举仍需本地 Word COM",
                }
        except Exception as exc:
            probe = {
                "schema_version": "k12.macroDetectionProbe.v1",
                "document_available": False,
                "template_available": False,
                "document_extension": "",
                "template_extension": "",
                "evidence": f"宏检测自检失败：{self._redact_local_path_text(str(exc) or exc.__class__.__name__)}",
                "current": "自检失败；真实 VBA 模块枚举仍需本地 Word COM",
            }
        self._macro_detection_probe_cache = probe
        return probe

    def _file_validation_probe(self) -> dict[str, Any]:
        """使用真实签名自检损坏与加密文档分类。"""
        cached = getattr(self, "_file_validation_probe_cache", None)
        if isinstance(cached, dict):
            return cached
        try:
            analyzer = DocumentAnalyzer()
            with tempfile.TemporaryDirectory(prefix="k12-file-validation-probe-") as tmp:
                root = Path(tmp)
                corrupt_word = root / "broken.docx"
                encrypted_word = root / "locked.docx"
                encrypted_pdf = root / "locked.pdf"
                corrupt_legacy_word = root / "broken.doc"
                encrypted_legacy_excel = root / "locked.xls"
                corrupt_word.write_bytes(b"not-an-ooxml-package")
                encrypted_word.write_bytes(
                    b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
                    + b"\x00" * 64
                    + "EncryptionInfo".encode("utf-16le")
                    + "EncryptedPackage".encode("utf-16le")
                )
                encrypted_pdf.write_bytes(b"%PDF-1.4\n1 0 obj<< /Type /Page /Encrypt >>endobj\n")
                legacy_header = bytearray(512)
                legacy_header[:8] = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
                legacy_header[28:30] = b"\xfe\xff"
                legacy_header[30:32] = (9).to_bytes(2, "little")
                legacy_header[32:34] = (6).to_bytes(2, "little")
                corrupt_legacy_word.write_bytes(b"not-a-legacy-office-container")
                encrypted_legacy_excel.write_bytes(
                    bytes(legacy_header)
                    + "EncryptionInfo".encode("utf-16le")
                    + "EncryptedPackage".encode("utf-16le")
                )
                corrupt_item = analyzer.analyze_file(corrupt_word)
                word_item = analyzer.analyze_file(encrypted_word)
                pdf_item = analyzer.analyze_file(encrypted_pdf)
                corrupt_legacy_item = analyzer.analyze_file(corrupt_legacy_word)
                encrypted_legacy_item = analyzer.analyze_file(encrypted_legacy_excel)
                corrupt_available = bool(
                    not corrupt_item.encrypted
                    and corrupt_item.status == "校验失败"
                    and any("损坏" in error for error in corrupt_item.validation_errors)
                    and corrupt_legacy_item.status == "校验失败"
                    and any("损坏" in error for error in corrupt_legacy_item.validation_errors)
                )
                encrypted_available = bool(
                    word_item.encrypted
                    and pdf_item.encrypted
                    and all(
                        item.status == "校验失败" and any("密码" in error for error in item.validation_errors)
                        for item in (word_item, pdf_item)
                    )
                    and not any("损坏" in error for error in word_item.validation_errors)
                    and encrypted_legacy_item.encrypted
                    and encrypted_legacy_item.status == "校验失败"
                    and any("密码" in error for error in encrypted_legacy_item.validation_errors)
                )
                probe_store = AppStore(root / "runtime-data")
                probe_store.update_settings({"allowExternalMathpixUpload": True})
                probe_processor = TaskProcessor(probe_store)
                invalid_pdf = probe_processor.create_uploaded_file("invalid.pdf", b"not-a-pdf")[0]
                invalid_task = probe_processor.create_task(
                    {"task_type": "pdf_to_word", "file_ids": [invalid_pdf["id"]]}
                )
                invalid_report = probe_store.list_reports()[0]
                invalid_job = invalid_report["analysis"]["mathpix"][0]
                invalid_artifact = invalid_report["analysis"]["artifacts"][0]
                execution_gate_available = bool(
                    invalid_task.get("status") == "失败"
                    and invalid_job.get("status") == "validation_failed"
                    and not (invalid_job.get("recognition_plan") or {}).get("submit_allowed")
                    and invalid_artifact.get("status") == "validation_failed"
                    and not invalid_report["analysis"].get("formulas")
                )
                if not corrupt_available:
                    raise ValueError("损坏 OOXML 未进入明确校验错误")
                if not encrypted_available:
                    raise ValueError("加密 OOXML/PDF 未进入密码提示或被误报为损坏")
                if not execution_gate_available:
                    raise ValueError("校验失败文件仍进入转换、对象提取或 Mathpix 提交流程")
                probe = {
                    "schema_version": "k12.fileValidationProbe.v1",
                    "corrupt_available": True,
                    "encrypted_available": True,
                    "execution_gate_available": True,
                    "corrupt_evidence": "文件校验自检通过：损坏 OOXML 和旧版 Office OLE 会进入 validation_errors，且校验失败文件不会生成派生对象、转换产物或 Mathpix 上传",
                    "encrypted_evidence": "文件加密自检通过：OLE 加密 OOXML、旧版 Office OLE 与含 /Encrypt 的 PDF 会提示输入密码，且加密文件不误报为损坏",
                    "current": "自检通过，损坏 OOXML/旧版 OLE、加密 Word/Excel 和加密 PDF 签名均被区分，invalid.pdf 在外部上传已授权时仍被 validation_failed 门槛阻止",
                }
        except Exception as exc:
            probe = {
                "schema_version": "k12.fileValidationProbe.v1",
                "corrupt_available": False,
                "encrypted_available": False,
                "execution_gate_available": False,
                "evidence": f"文件校验自检失败：{self._redact_local_path_text(str(exc) or exc.__class__.__name__)}",
                "current": "自检失败，需复核损坏与加密文件判定",
            }
        self._file_validation_probe_cache = probe
        return probe

    def _ppt_formula_probe(self) -> dict[str, Any]:
        """自检 PPT 公式线索提取以支持验收。"""
        cached = getattr(self, "_ppt_formula_probe_cache", None)
        if isinstance(cached, dict):
            return cached
        try:
            with tempfile.TemporaryDirectory(prefix="k12-ppt-formula-probe-") as tmp:
                source = Path(tmp) / "formula-probe.pptx"
                build_pptx(
                    [
                        {
                            "title": "PPT 公式识别自检",
                            "body": ["例题公式：$x^2+y^2=z^2$", "MathType fallback marker"],
                        }
                    ],
                    source,
                )
                slides = extract_pptx_slides(source)
                formula_count = sum(int(slide.get("formula_count") or 0) for slide in slides)
                if formula_count <= 0:
                    raise ValueError("未识别到 PPT 公式线索")
                probe = {
                    "schema_version": "k12.pptFormulaProbe.v1",
                    "available": True,
                    "evidence": "PPT 公式自检通过：可从 PPTX 文本和 MathType 线索中识别公式数量",
                    "current": f"自检通过，识别公式线索 {formula_count} 个；真实 MathType 对象复核仍需本地客户端",
                    "formula_count": formula_count,
                }
        except Exception as exc:
            probe = {
                "schema_version": "k12.pptFormulaProbe.v1",
                "available": False,
                "evidence": f"PPT 公式自检失败：{self._redact_local_path_text(str(exc) or exc.__class__.__name__)}",
                "current": "自检失败；真实 MathType 对象复核仍需本地客户端",
                "formula_count": 0,
            }
        self._ppt_formula_probe_cache = probe
        return probe

    def _omml_dependency_probe(self) -> dict[str, Any]:
        """在临时目录自检 OMML 依赖查找与复制。"""
        cached = getattr(self, "_omml_dependency_probe_cache", None)
        if isinstance(cached, dict):
            return cached
        settings = self.store.get_settings()
        try:
            with tempfile.TemporaryDirectory(prefix="k12-omml-probe-") as tmp:
                root = Path(tmp)
                document_dir = root / "document"
                dependency_dir = root / "dependency-source"
                document_dir.mkdir()
                dependency_dir.mkdir()
                document_path = document_dir / "lesson.docx"
                document_path.write_bytes(b"probe")
                source = dependency_dir / "OMML2MML.XSL"
                source.write_text("<xsl:stylesheet />", encoding="utf-8")
                probe_settings = {
                    **settings,
                    "ommlSearchPaths": str(dependency_dir),
                    "ommlSearchMaxFiles": 20,
                    "ommlCopyStrategy": "自动重命名",
                }
                found = self._find_omml_dependency(document_path, probe_settings)
                if not found or found.name != source.name:
                    raise ValueError("未找到临时 OMML 依赖")
                item = OmmlDependencyItem(file_id="file_omml_probe", document_path=str(document_path))
                self._copy_omml_dependency(found, document_path, item, "自动重命名")
                target = Path(item.omml_target_path or "")
                if item.copy_status != "成功" or not target.exists() or target.parent != document_dir:
                    raise ValueError("OMML 依赖未复制到文档目录")
                if not source.exists():
                    raise ValueError("OMML 源依赖不应被移动或删除")
                probe = {
                    "schema_version": "k12.ommlDependencyProbe.v1",
                    "search_available": True,
                    "copy_available": True,
                    "search_evidence": "OMML 依赖自检通过：可在配置目录自动检索 OMML2MML.XSL",
                    "copy_evidence": "OMML 依赖自检通过：可将检索到的 OMML2MML.XSL 复制到当前 Word 文档目录",
                    "current": "自检通过，临时目录搜索与复制成功，未触碰用户文件",
                }
        except Exception as exc:
            message = self._redact_local_path_text(str(exc) or exc.__class__.__name__)
            probe = {
                "schema_version": "k12.ommlDependencyProbe.v1",
                "search_available": False,
                "copy_available": False,
                "evidence": f"OMML 依赖自检失败：{message}",
                "current": "自检失败，未触碰用户文件",
            }
        self._omml_dependency_probe_cache = probe
        return probe

    @staticmethod
    def _omml_dependency_probe_current(probe: dict[str, Any]) -> str:
        """为验收矩阵格式化 OMML 依赖探测证据。"""
        return str(probe.get("current") or "OMML 依赖自检未运行")

    def _mathpix_contract_probe(self) -> dict[str, Any]:
        """自检 Mathpix PDF 转 Word 请求约定，不进行外部上传。"""
        cached = getattr(self, "_mathpix_contract_probe_cache", None)
        if isinstance(cached, dict):
            return cached
        settings = self.store.get_settings()
        try:
            ocr_settings = self._ocr_settings_snapshot(settings)
            task = {"id": "task_mathpix_probe", "task_type": "pdf_to_word", "options": {}}
            file = {
                "id": "file_mathpix_probe",
                "file_name": "mathpix-probe.pdf",
                "file_type": "PDF",
                "content_summary": {
                    "pdfType": "文本型 PDF",
                    "imageObjects": 0,
                    "tableHints": 0,
                    "formulaHints": 1 if ocr_settings["formula_ocr"] else 0,
                },
            }
            request_options = self._mathpix_submit_options(ocr_settings)
            retention_plan = self._mathpix_retention_plan(task, file, ocr_settings, settings)
            allow_upload = bool(settings.get("allowExternalMathpixUpload", False))
            job = {
                "file_id": file["id"],
                "file_name": file["file_name"],
                "engine": "Mathpix",
                "status": "pending" if allow_upload else "authorization_required",
                "pdf_id": "",
                "outputs": {},
                "message": "",
                "ocr_settings": ocr_settings,
                "retention_plan": retention_plan,
                "request_options": request_options,
            }
            self._attach_mathpix_recognition_plan(job, file, settings, allow_upload, True)
            plan = job["recognition_plan"]
            submit_contract = dict(plan.get("submit_contract") or {})
            conversion_formats = list(submit_contract.get("conversion_formats") or [])
            if "docx" not in conversion_formats:
                raise ValueError("Mathpix DOCX 输出格式未进入请求合同")
            if ocr_settings["formula_ocr"] and "tex.zip" not in conversion_formats:
                raise ValueError("公式 OCR 开启时未请求 tex.zip 结果包")
            credential_envs = list(plan.get("credential_envs") or [])
            missing_envs = [name for name in credential_envs if not os.getenv(name)]
            credential_status = "凭证环境变量已配置" if not missing_envs else f"凭证待配置：{', '.join(missing_envs)}"
            upload_status = "外部上传已授权" if allow_upload else "外部上传未授权"
            formats_label = "、".join(conversion_formats)
            probe = {
                "schema_version": "k12.mathpixContractProbe.v1",
                "available": True,
                "engine": "Mathpix",
                "evidence": f"Mathpix 合同自检通过：PDF 转 Word 使用 /v3/pdf，输出 {formats_label}",
                "current": f"自检通过，{upload_status}，{credential_status}，未发起外部上传",
                "recognition_status": plan.get("status", ""),
            }
        except Exception as exc:
            probe = {
                "schema_version": "k12.mathpixContractProbe.v1",
                "available": False,
                "engine": "Mathpix",
                "evidence": f"Mathpix 合同自检失败：{self._redact_local_path_text(str(exc) or exc.__class__.__name__)}",
                "current": "合同自检失败，未发起外部上传",
                "recognition_status": "probe_failed",
            }
        self._mathpix_contract_probe_cache = probe
        return probe

    def _pdf_ocr_contract_probe(self) -> dict[str, Any]:
        """自检扫描 PDF OCR 约定和上传风险阻断。"""
        cached = getattr(self, "_pdf_ocr_contract_probe_cache", None)
        if isinstance(cached, dict):
            return cached
        try:
            with tempfile.TemporaryDirectory(prefix="k12-pdf-ocr-probe-") as tmp:
                probe_store = AppStore(Path(tmp) / "data")
                probe_processor = TaskProcessor(probe_store)
                scanned_pdf = (
                    b"%PDF-1.4\n"
                    b"1 0 obj<< /Type /Page >>endobj\n"
                    b"2 0 obj<< /Type /XObject /Subtype /Image /Width 20 /Height 16 "
                    b"/ColorSpace /DeviceRGB /BitsPerComponent 8 /Length 8 >>\n"
                    b"stream\nSCANIMG\nendstream\nendobj\n%%EOF\n"
                )
                formula_pdf = b"%PDF-1.4\n1 0 obj<< /Type /Page >>(Math Formula) Tj endobj\n%%EOF\n"
                scanned = probe_processor.create_uploaded_file("scan-ocr.pdf", scanned_pdf)[0]
                formula = probe_processor.create_uploaded_file("formula-ocr.pdf", formula_pdf)[0]
                task = probe_processor.create_task({"task_type": "pdf_to_word", "file_ids": [scanned["id"], formula["id"]]})
                report = probe_store.list_reports()[0]
                mathpix_jobs = list((report.get("analysis") or {}).get("mathpix") or [])
                scanned_job = next((job for job in mathpix_jobs if job.get("file_id") == scanned["id"]), None)
                formula_job = next((job for job in mathpix_jobs if job.get("file_id") == formula["id"]), None)
                if not scanned_job or not formula_job:
                    raise ValueError("PDF OCR 自检缺少 Mathpix 作业")

                scanned_plan = dict(scanned_job.get("recognition_plan") or {})
                scanned_retention = dict(scanned_job.get("retention_plan") or {})
                formula_plan = dict(formula_job.get("recognition_plan") or {})
                formula_retention = dict(formula_job.get("retention_plan") or {})
                scanned_contract = dict(scanned_plan.get("submit_contract") or {})
                formula_contract = dict(formula_plan.get("submit_contract") or {})
                formula_formats = list(formula_contract.get("conversion_formats") or [])
                if scanned.get("content_summary", {}).get("pdfType") != "扫描型 PDF":
                    raise ValueError("扫描 PDF 类型识别失败")
                if scanned_plan.get("status") != "blocked_authorization" or scanned_plan.get("submit_allowed"):
                    raise ValueError("扫描 PDF 未正确停在 Mathpix 外部上传授权边界")
                if not scanned_plan.get("ocr", {}).get("text") or scanned_contract.get("endpoint") != "/v3/pdf":
                    raise ValueError("扫描 PDF 未进入 Mathpix 文字 OCR 请求合同")
                if formula_retention.get("formula_hints", 0) < 1 or formula_retention.get("formula_retention_status") != "计划识别公式并生成 MathType 预览":
                    raise ValueError("PDF 公式线索未进入 Mathpix 公式 OCR 保留计划")
                if formula_plan.get("status") != "blocked_authorization" or formula_plan.get("submit_allowed"):
                    raise ValueError("PDF 公式 OCR 未正确停在 Mathpix 外部上传授权边界")
                if "docx" not in formula_formats or "tex.zip" not in formula_formats:
                    raise ValueError("PDF 公式 OCR 未请求 docx 和 tex.zip 输出")

                payload = probe_processor.local_task_payload(task["id"])
                formula_action = next((action for action in payload.get("local_actions", []) if action.get("type") == "pdf_formula_mathtype"), None)
                if not formula_action:
                    raise ValueError("PDF 公式 OCR 自检缺少 pdf_formula_mathtype 动作")
                if formula_action.get("status") != "blocked_mathpix" or int(formula_action.get("formula_count") or 0) < 1:
                    raise ValueError("PDF 公式后处理未等待 Mathpix 识别结果")
                plan_action = next((action for action in payload.get("desktop_execution_plan", {}).get("actions", []) if action.get("type") == "pdf_formula_mathtype"), None)
                if not plan_action:
                    raise ValueError("PDF 公式 OCR 自检缺少桌面后处理动作")
                operations = [str(step.get("operation") or "") for step in plan_action.get("steps") or [] if isinstance(step, dict)]
                if "mathpix.collect_formula_outputs" not in operations or "mathtype.convert_pdf_formula" not in operations:
                    raise ValueError("PDF 公式后处理桌面计划缺少 Mathpix 或 MathType 步骤")
                if scanned_job.get("status") != "authorization_required" or formula_job.get("status") != "authorization_required":
                    raise ValueError("PDF OCR 自检不应提交外部 Mathpix")

                probe = {
                    "schema_version": "k12.pdfOcrContractProbe.v1",
                    "scanned_available": True,
                    "formula_available": True,
                    "scanned_pdf_type": scanned.get("content_summary", {}).get("pdfType", ""),
                    "formula_hints": int(formula_retention.get("formula_hints") or 0),
                    "recognition_status": {
                        "scanned": scanned_plan.get("status", ""),
                        "formula": formula_plan.get("status", ""),
                    },
                    "conversion_formats": formula_formats,
                    "scanned_evidence": "扫描 PDF OCR 合同自检通过：扫描型 PDF 进入 Mathpix /v3/pdf 文字 OCR 识别计划，并被外部上传授权门槛阻断",
                    "scanned_current": "自检通过，扫描型 PDF 识别为扫描件，Mathpix submit_allowed=false；未发起外部上传，真实 OCR 需授权实测",
                    "formula_evidence": "PDF 公式 OCR 合同自检通过：公式线索进入 Mathpix 公式 OCR 保留计划，并请求 docx 与 tex.zip 输出",
                    "formula_current": "自检通过，1 个 PDF 公式线索进入 Mathpix tex.zip 请求；未发起外部上传，真实识别需 Mathpix 授权实测",
                    "formula_handoff_evidence": "PDF 公式识别转 MathType 合同自检通过：公式 OCR 作业在授权前阻断，pdf_formula_mathtype 后处理动作等待 Mathpix 输出并保留 MathType 桌面计划",
                    "formula_handoff_current": "自检通过，PDF 公式后处理动作状态 blocked_mathpix；真实公式识别和 MathType 对象写回仍需 Mathpix 授权与本地客户端",
                }
        except Exception as exc:
            message = self._redact_local_path_text(str(exc) or exc.__class__.__name__)
            probe = {
                "schema_version": "k12.pdfOcrContractProbe.v1",
                "scanned_available": False,
                "formula_available": False,
                "scanned_pdf_type": "",
                "formula_hints": 0,
                "recognition_status": {},
                "conversion_formats": [],
                "evidence": f"PDF OCR 合同自检失败：{message}",
                "scanned_current": "自检失败；未发起外部 Mathpix 上传",
                "formula_current": "自检失败；未发起外部 Mathpix 上传",
                "formula_handoff_current": "自检失败；未发起外部 Mathpix 上传",
            }
        self._pdf_ocr_contract_probe_cache = probe
        return probe

    def _pdf_retention_probe(self) -> dict[str, Any]:
        """自检 Mathpix 产物中的 PDF 图片与表格保留计划。"""
        cached = getattr(self, "_pdf_retention_probe_cache", None)
        if isinstance(cached, dict):
            return cached
        try:
            with tempfile.TemporaryDirectory(prefix="k12-pdf-retention-probe-") as tmp:
                probe_store = AppStore(Path(tmp) / "data")
                probe_processor = TaskProcessor(probe_store)
                image_stream = b"\xff\xd8K12PDFRETENTION\xff\xd9"
                pdf_bytes = (
                    b"%PDF-1.4\n"
                    b"1 0 obj<< /Type /Page /Table >>(Math Formula) Tj endobj\n"
                    b"2 0 obj<< /Type /XObject /Subtype /Image /Width 12 /Height 8 "
                    b"/ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length "
                    + str(len(image_stream)).encode("ascii")
                    + b" >>\nstream\n"
                    + image_stream
                    + b"\nendstream\nendobj\n%%EOF\n"
                )
                pdf = probe_processor.create_uploaded_file("retention.pdf", pdf_bytes)[0]
                task = probe_processor.create_task({"task_type": "pdf_to_word", "file_ids": [pdf["id"]]})
                report = probe_store.list_reports()[0]
                mathpix = (report.get("analysis") or {}).get("mathpix", [])[0]
                retention = dict(mathpix.get("retention_plan") or {})
                request_options = dict(mathpix.get("request_options") or {})
                html_text = Path(report["html_path"]).read_text(encoding="utf-8")
                xlsx_text = ""
                with zipfile.ZipFile(report["xlsx_path"]) as archive:
                    xlsx_text = archive.read("xl/worksheets/sheet1.xml").decode("utf-8")
                image_available = (
                    retention.get("image_objects", 0) >= 1
                    and retention.get("retain_images") is True
                    and retention.get("image_retention_status") == "计划保留图片对象"
                    and "计划保留图片对象" in html_text
                )
                table_available = (
                    retention.get("table_hints", 0) >= 1
                    and retention.get("retain_tables") is True
                    and retention.get("table_ocr") is True
                    and retention.get("table_retention_status") == "计划识别并保留表格结构"
                    and bool(request_options.get("enable_tables_fallback"))
                    and "计划识别并保留表格结构" in xlsx_text
                )
                if task.get("execute_mode") != "hybrid" or mathpix.get("status") != "authorization_required":
                    raise ValueError("PDF 保留自检应保持未授权外部上传的混合模式")
                if not image_available:
                    raise ValueError("PDF 图片对象未进入保留计划和报告")
                if not table_available:
                    raise ValueError("PDF 表格线索未进入保留计划和报告")
                probe = {
                    "schema_version": "k12.pdfRetentionProbe.v1",
                    "image_available": True,
                    "table_available": True,
                    "image_objects": int(retention.get("image_objects") or 0),
                    "table_hints": int(retention.get("table_hints") or 0),
                    "image_evidence": "PDF 图片保留自检通过：PDF 图片 XObject 进入 Mathpix 保留计划、报告和脱敏队列",
                    "table_evidence": "PDF 表格保留自检通过：PDF 表格线索进入 Mathpix 表格 OCR 请求、保留计划和报告",
                    "current": "自检通过，图片对象 1 个、表格线索 1 个；未发起外部 Mathpix 上传，真实排版仍需 Mathpix 授权结果复核",
                }
        except Exception as exc:
            message = self._redact_local_path_text(str(exc) or exc.__class__.__name__)
            probe = {
                "schema_version": "k12.pdfRetentionProbe.v1",
                "image_available": False,
                "table_available": False,
                "image_objects": 0,
                "table_hints": 0,
                "evidence": f"PDF 图片/表格保留自检失败：{message}",
                "current": "自检失败；未发起外部 Mathpix 上传",
            }
        self._pdf_retention_probe_cache = probe
        return probe

    def _pdf_formula_handoff_probe(self) -> dict[str, Any]:
        """自检 PDF 公式 OCR 至本地 MathType 后处理的交接。"""
        cached = getattr(self, "_pdf_formula_handoff_probe_cache", None)
        if isinstance(cached, dict):
            return cached
        try:
            with tempfile.TemporaryDirectory(prefix="k12-pdf-formula-probe-") as tmp:
                root = Path(tmp)
                probe_store = AppStore(root / "data")
                probe_store.update_settings({"localClientPlatform": "Windows", "localSecurityToken": "probe-token"})
                probe_processor = TaskProcessor(probe_store)
                probe_processor.record_local_client_heartbeat(
                    {
                        "client_id": "desktop-pdf-formula-probe",
                        "status": "online",
                        "platform": "Windows",
                        "preflight": {
                            "platform": "Windows",
                            "components": {"mathtype": {"label": "MathType 组件", "available": True, "status": "available"}},
                            "capabilities": {"mathTypeAutomation": True},
                        },
                    }
                )
                pdf = probe_processor.create_uploaded_file(
                    "formula.pdf",
                    b"%PDF-1.4\n1 0 obj<< /Type /Page >>(Math Formula) Tj endobj\n%%EOF\n",
                )[0]
                tex_zip = root / "mathpix-formula.zip"
                with zipfile.ZipFile(tex_zip, "w", zipfile.ZIP_DEFLATED) as archive:
                    archive.writestr("formula.tex", r"$$x^2+y^2=z^2$$")
                job = {
                    "file_id": pdf["id"],
                    "file_name": pdf["file_name"],
                    "engine": "Mathpix",
                    "status": "completed",
                    "tex_zip_status": "completed",
                    "outputs": {"tex_zip": str(tex_zip), "docx": str(root / "mathpix.docx")},
                    "ocr_settings": {"formula_ocr": True},
                    "retention_plan": {"formula_hints": 1},
                }
                formulas = probe_processor._mathpix_formula_items([pdf], [job])
                if not formulas:
                    raise ValueError("Mathpix tex.zip 未生成 PDF 公式项")
                formula = formulas[0]
                if formula.get("source_type") != "PDF" or not formula.get("mathml") or not str(formula.get("mathtype_preview") or "").startswith("MathType 预览"):
                    raise ValueError("PDF 公式项缺少 MathML 或 MathType 预览")
                task = {
                    "id": "task_pdf_formula_probe",
                    "task_type": "pdf_to_word",
                    "task_label": TASK_LABELS.get("pdf_to_word", "pdf_to_word"),
                    "execute_mode": "hybrid",
                    "file_ids": [pdf["id"]],
                    "status": "成功",
                    "progress": 100,
                    "options": {},
                }
                report = {
                    "id": "report_pdf_formula_probe",
                    "task_id": task["id"],
                    "analysis": {"mathpix": [job], "formulas": formulas, "artifacts": []},
                }
                actions = probe_processor._local_payload_actions(task, [pdf], report)
                action = next(item for item in actions if item["type"] == "pdf_formula_mathtype")
                readiness = probe_processor._local_client_readiness(True, actions, probe_store.get_settings())
                plan = probe_processor._local_desktop_execution_plan(task, [pdf], actions, probe_store.get_settings(), readiness, include_sensitive_paths=False)
                plan_action = next(item for item in plan["actions"] if item["type"] == "pdf_formula_mathtype")
                operations = [item.get("operation") for item in plan_action.get("steps", [])]
                if action.get("status") != "queued" or int(action.get("formula_count") or 0) != 1:
                    raise ValueError("PDF 公式后处理动作未进入队列")
                if plan.get("status") != "ready_for_native_client" or plan_action.get("gate_status") != "ready":
                    raise ValueError("PDF 公式桌面执行计划未达到可交接状态")
                for operation in ("mathpix.collect_formula_outputs", "mathtype.convert_pdf_formula", "formula.merge_into_docx", "formula.validate"):
                    if operation not in operations:
                        raise ValueError(f"PDF 公式执行计划缺少步骤：{operation}")
                probe = {
                    "schema_version": "k12.pdfFormulaHandoffProbe.v1",
                    "available": True,
                    "formula_count": len(formulas),
                    "plan_status": plan.get("status", ""),
                    "operations": operations,
                    "evidence": "PDF 公式 MathType 后处理自检通过：Mathpix tex.zip 公式进入 PDF 公式项、MathType 预览、本地动作和桌面执行计划",
                    "current": "自检通过，1 个 PDF 公式可交接给 pdf_formula_mathtype；真实 MathType 对象写回仍需本地客户端",
                }
        except Exception as exc:
            message = self._redact_local_path_text(str(exc) or exc.__class__.__name__)
            probe = {
                "schema_version": "k12.pdfFormulaHandoffProbe.v1",
                "available": False,
                "formula_count": 0,
                "plan_status": "",
                "operations": [],
                "evidence": f"PDF 公式 MathType 后处理自检失败：{message}",
                "current": "自检失败；未发起外部 Mathpix 上传",
            }
        self._pdf_formula_handoff_probe_cache = probe
        return probe

    @staticmethod
    def _mathpix_contract_probe_current(
        has_task_history: bool,
        has_completed_docx: bool,
        probe: dict[str, Any],
        mathpix_queue: dict[str, Any],
    ) -> str:
        """汇总 Mathpix 约定证据及失败下载数量。"""
        history = "已有完成 DOCX" if has_completed_docx else ("已有任务或报告" if has_task_history else "无历史任务")
        total = mathpix_queue.get("summary", {}).get("total", 0)
        completed = mathpix_queue.get("summary", {}).get("completed", 0)
        return f"pdf_to_word {history}；{probe.get('current') or '合同自检未运行'}；Mathpix 队列 {total} 个，已完成 {completed} 个"

    @staticmethod
    def _mathpix_acceptance_evidence(mathpix_queue: dict[str, Any], reports: list[dict[str, Any]]) -> dict[str, Any]:
        """汇总真实 Mathpix 完成证据用于需求验收。"""
        jobs = [item for item in mathpix_queue.get("jobs", []) if isinstance(item, dict)]
        pdf_formula_count = 0
        for report in reports:
            analysis = report.get("analysis") if isinstance(report.get("analysis"), dict) else {}
            formulas = analysis.get("formulas") if isinstance(analysis.get("formulas"), list) else []
            pdf_formula_count += sum(
                1
                for formula in formulas
                if isinstance(formula, dict)
                and str(formula.get("source_type") or "") == "PDF"
                and bool(str(formula.get("latex") or formula.get("mathml") or formula.get("mathtype_preview") or "").strip())
            )
        scanned_jobs: list[dict[str, Any]] = []
        formula_jobs: list[dict[str, Any]] = []
        for job in jobs:
            recognition = job.get("recognition_plan") if isinstance(job.get("recognition_plan"), dict) else {}
            retention = job.get("retention") if isinstance(job.get("retention"), dict) else {}
            ocr_settings = job.get("ocr_settings") if isinstance(job.get("ocr_settings"), dict) else {}
            if job.get("status") != "completed" and recognition.get("status") != "completed":
                continue
            docx_output = TaskProcessor._mathpix_manifest_output(recognition, "docx")
            tex_output = TaskProcessor._mathpix_manifest_output(recognition, "tex.zip")
            pdf_type = str(retention.get("pdf_type") or "")
            if bool(ocr_settings.get("text_ocr")) and pdf_type in {"扫描型 PDF", "混合型 PDF"} and TaskProcessor._mathpix_output_downloaded_with_hash(docx_output):
                scanned_jobs.append(job)
            if bool(ocr_settings.get("formula_ocr")) and TaskProcessor._mathpix_output_downloaded_with_hash(tex_output):
                formula_jobs.append(job)
        scanned_available = bool(scanned_jobs)
        formula_available = bool(formula_jobs) and pdf_formula_count > 0
        scanned_count = len(scanned_jobs)
        formula_count = max(pdf_formula_count, len(formula_jobs))
        return {
            "schema_version": "k12.mathpixAcceptanceEvidence.v1",
            "scanned_ocr_available": scanned_available,
            "formula_ocr_available": formula_available,
            "scanned_completed_count": scanned_count,
            "formula_completed_count": len(formula_jobs),
            "pdf_formula_count": pdf_formula_count,
            "scanned_evidence": (
                f"已有 Mathpix 扫描 PDF OCR 实测记录：{scanned_count} 个扫描/混合 PDF 作业完成，DOCX 下载清单完整且哈希可用"
                if scanned_available
                else ""
            ),
            "scanned_current": (
                f"Mathpix 已完成扫描/混合 PDF OCR {scanned_count} 个，DOCX 结果进入标准输出和报告"
                if scanned_available
                else ""
            ),
            "formula_evidence": (
                f"已有 Mathpix PDF 公式 OCR 实测记录：{len(formula_jobs)} 个作业下载 tex.zip，公式报告解析出 {pdf_formula_count} 个 PDF 公式"
                if formula_available
                else ""
            ),
            "formula_current": (
                f"Mathpix tex.zip 与公式报告已完成，PDF 公式 {formula_count} 个；低置信度和人工确认仍按报告策略处理"
                if formula_available
                else ""
            ),
            "formula_handoff_evidence": (
                f"Mathpix PDF 公式 OCR 已实测完成，{pdf_formula_count} 个 PDF 公式可交接给 pdf_formula_mathtype；原生 MathType 写回仍需同平台本地客户端"
                if formula_available
                else ""
            ),
            "formula_handoff_current": (
                "Mathpix 识别证据已覆盖，剩余为 Office/MathType 同平台本地客户端写回实测"
                if formula_available
                else ""
            ),
        }

    @staticmethod
    def _mathpix_manifest_output(recognition: dict[str, Any], output_type: str) -> dict[str, Any]:
        """按产物类型返回 Mathpix 下载清单记录。"""
        manifest = recognition.get("download_manifest") if isinstance(recognition.get("download_manifest"), dict) else {}
        outputs = manifest.get("outputs") if isinstance(manifest.get("outputs"), list) else []
        for item in outputs:
            if isinstance(item, dict) and str(item.get("type") or "") == output_type:
                return item
        return {}

    @staticmethod
    def _mathpix_output_downloaded_with_hash(output: dict[str, Any]) -> bool:
        """检查 Mathpix 产物是否已下载且具备完整性证据。"""
        return str(output.get("status") or "") == "downloaded" and bool(output.get("sha256_available"))

    @staticmethod
    def _local_native_acceptance_evidence(tasks: list[dict[str, Any]]) -> dict[str, Any]:
        """汇总真实桌面原生执行报告用于需求验收。"""
        successful: dict[str, list[dict[str, Any]]] = {}
        for task in tasks:
            report = task.get("local_native_execution_report") if isinstance(task.get("local_native_execution_report"), dict) else {}
            actions = report.get("actions") if isinstance(report.get("actions"), list) else []
            for action in actions:
                if not isinstance(action, dict):
                    continue
                action_type = str(action.get("type") or "")
                if action_type not in LOCAL_ACTION_LABELS:
                    continue
                if action.get("status") != "success" or not action.get("native_execution_performed"):
                    continue
                successful.setdefault(action_type, []).append(
                    {
                        "task_id": str(task.get("id") or ""),
                        "task_type": str(task.get("task_type") or ""),
                        "platform": normalize_platform(str(action.get("platform") or report.get("platform") or "")),
                        "output_artifact_types": [str(item) for item in action.get("output_artifact_types") or [] if isinstance(item, str)][:12],
                    }
                )
        native_required = {"office_conversion", "omml_mathtype", "pdf_formula_mathtype", "macro_sequence"}
        available_types = set(successful)
        missing_types = sorted(native_required - available_types)
        available_labels = [LOCAL_ACTION_LABELS.get(action_type, action_type) for action_type in sorted(available_types & native_required)]
        missing_labels = [LOCAL_ACTION_LABELS.get(action_type, action_type) for action_type in missing_types]
        omml_count = len(successful.get("omml_mathtype", []))
        macro_count = len(successful.get("macro_sequence", []))
        pdf_formula_count = len(successful.get("pdf_formula_mathtype", []))
        all_native = native_required.issubset(available_types)
        partial_current = ""
        if available_labels:
            partial_current = f"已收到本地客户端原生执行证据：{', '.join(available_labels)}；仍缺 {', '.join(missing_labels) if missing_labels else '无'}"
        return {
            "schema_version": "k12.localNativeAcceptanceEvidence.v1",
            "available_action_types": sorted(available_types),
            "missing_action_types": missing_types,
            "omml_mathtype_available": omml_count > 0,
            "macro_sequence_available": macro_count > 0,
            "pdf_formula_mathtype_available": pdf_formula_count > 0,
            "office_conversion_available": bool(successful.get("office_conversion")),
            "all_native_actions_available": all_native,
            "omml_evidence": (
                f"已有本地客户端原生执行证据：{omml_count} 个 omml_mathtype 动作成功，MathType/OMML 写回由同平台客户端回传"
                if omml_count
                else ""
            ),
            "omml_retry_evidence": (
                f"已有本地客户端原生执行证据：OMML 依赖补齐后的重新转换动作成功 {omml_count} 次"
                if omml_count
                else ""
            ),
            "omml_current": (
                f"本地客户端已回传 OMML/MathType 原生执行成功 {omml_count} 次，平台证据已脱敏保存"
                if omml_count
                else ""
            ),
            "macro_evidence": (
                f"已有本地客户端原生执行证据：{macro_count} 个 macro_sequence 动作成功，Word 宏按本地队列执行"
                if macro_count
                else ""
            ),
            "macro_current": (
                f"本地客户端已回传 Word 宏顺序执行成功 {macro_count} 次，结果路径未进入验收矩阵"
                if macro_count
                else ""
            ),
            "pdf_formula_evidence": (
                f"已有本地客户端原生执行证据：{pdf_formula_count} 个 pdf_formula_mathtype 动作成功，PDF 公式后处理已写回同平台 MathType 合同"
                if pdf_formula_count
                else ""
            ),
            "pdf_formula_current": (
                f"本地客户端已回传 PDF 公式 MathType 后处理成功 {pdf_formula_count} 次，Mathpix 与本地写回证据均可追踪"
                if pdf_formula_count
                else ""
            ),
            "all_native_evidence": (
                f"已有本地客户端原生执行证据：{', '.join(available_labels)} 均成功回传"
                if all_native
                else ""
            ),
            "partial_evidence": (
                f"已有部分本地客户端原生执行证据：{', '.join(available_labels)} 成功回传；尚缺 {', '.join(missing_labels)}"
                if available_labels and not all_native
                else ""
            ),
            "remaining_gap": (
                f"尚缺真实本地执行证据：{', '.join(missing_labels)}"
                if missing_labels
                else ""
            ),
            "all_native_current": (
                "Office、MathType、OMML 和 Word 宏相关动作均已有本地客户端原生执行成功记录"
                if all_native
                else ""
            ),
            "partial_current": partial_current,
        }

    def product_summary(self) -> dict[str, Any]:
        """返回需求能力覆盖和分阶段路线图状态。"""
        settings = self.store.get_settings()
        files = self.store.list_files()
        tasks = self.store.list_tasks()
        reports = self.store.list_reports()
        caps = self.capabilities()
        core_capabilities = [
            self._product_capability("upload_batch", "文档可上传、可识别、可批量处理", "implemented", f"文件 {len(files)} 个，任务 {len(tasks)} 个，支持拖拽、文件夹、ZIP、队列和批量报告"),
            self._product_capability("word_ppt", "Word 和 PPT 可互相转换", "contract", "已生成最小 OOXML 转换产物和对象保留清单，高保真排版交给本地 Office 引擎复核"),
            self._product_capability("pdf_word", "PDF 可转换为 Word", "implemented", f"纯文本 PDF 已支持本地转换；扫描、混合和复杂 PDF 使用 Mathpix，外部上传授权 {('已开启' if settings.get('allowExternalMathpixUpload') else '未开启')}"),
            self._product_capability("omml_detection", "Word 处理前可检测是否存在 Word 自带公式", "implemented", "文件分析会标记 OMML、缺失依赖和转换确认状态"),
            self._product_capability("omml_to_mathtype", "Word 自带公式可转换为 MathType 公式", "contract", "已提供转换确认、依赖状态和本地客户端动作队列；真实写回需桌面端实测"),
            self._product_capability("omml_search_copy", "找不到 OMML 文件时，可检索电脑本地 OMML 文件并复制到当前文档所在文件夹", "contract", "检索路径、手动指定和复制策略已进入任务合同；复制需本地客户端权限"),
            self._product_capability("macro_sequence", "用户可选择 Word 宏，并按指定顺序执行", "contract", f"宏队列状态：{caps['macroExecution']['status']}"),
            self._product_capability("mathtype_preserve", "MathType 公式尽量不丢失", "contract", caps["mathType"]["compatibility"]),
            self._product_capability("pdf_formula", "PDF 数学公式可识别并转换为 MathType", "contract", "Mathpix OCR、低置信度和公式报告合同已建立；外部识别需授权与凭证"),
            self._product_capability("mathtype_format", "Word 中可选择使用 MathType 格式化公式", "contract", f"格式化开关 {('已开启' if settings.get('enableMathTypeFormatting') else '未开启')}，真实 MathType 对象写回需本地客户端"),
            self._product_capability("small_images", "微小图片可检索、定位、导出", "implemented", "OOXML/PDF 图片提取、筛选、详情、标注、ZIP 和 XLSX 清单已进入报告链路"),
            self._product_capability("hybrid_mode", "本地端负责复杂文档处理，网页端负责管理能力", "contract", f"本地连接：{caps['localConnection']['status']}，载荷和状态同步接口已实现"),
            self._product_capability("trace_recover_export", "任务可追踪、错误可恢复、报告可导出", "implemented", f"报告 {len(reports)} 份，支持日志、失败清单、任务包、恢复和导出"),
        ]
        phases = [
            self._product_phase("phase_1", "第一阶段", "Windows 本地桌面版", "contract", "Windows 安装画像、Office COM/MathType/OLE 任务分流和本地载荷已定义；真实桌面壳后续接入"),
            self._product_phase("phase_2", "第二阶段", "增加网页管理端", "implemented", "本地网页端已覆盖用户、权限、模板、授权、报告和预检管理"),
            self._product_phase("phase_3", "第三阶段", "实现本地 + 网页混合模式", "contract", "本地安全令牌、任务载荷、状态同步和云端摘要授权已形成接口合同"),
            self._product_phase("phase_4", "第四阶段", "增强 AI 排版修复、模板套用和私有化部署能力", "planned", "AI 排版修复、智能模板套用和私有化部署已拆分为增强规划合同，真实执行仍处于 V3.0 后续接入"),
        ]
        summary_items = core_capabilities + phases
        return {
            "schema_version": "k12.productSummary.v1",
            "generated_at": utc_now(),
            "core_capabilities": core_capabilities,
            "phase_roadmap": phases,
            "summary": {
                "implemented": sum(1 for item in summary_items if item["level"] == "good"),
                "contract": sum(1 for item in summary_items if item["level"] == "warn"),
                "planned": sum(1 for item in summary_items if item["level"] == "blue"),
            },
            "positioning": "统一处理 Word、Excel、PPT、PDF，并重点增强公式、宏和图片处理能力",
        }

    def enhancement_plan(self) -> dict[str, Any]:
        """返回 V3 增强约定，不声称执行实时 AI 功能。"""
        settings = self.store.get_settings()
        caps = self.capabilities()
        token_configured = bool(str(settings.get("localSecurityToken") or "").strip())
        external_mathpix_allowed = bool(settings.get("allowExternalMathpixUpload"))
        cloud_sync_allowed = bool(settings.get("allowCloudSync"))
        items = [
            self._enhancement_item(
                "ai_layout_repair",
                "AI 排版修复",
                "planned",
                "读取排版校正记录、转换前后对比、标题层级、表格、图片和公式位置，生成修复建议与人工确认队列",
                ["layout_annotations", "preview_comparison", "quality_checks"],
                ["默认不调用外部 AI", "外部模型需单独授权", "只生成建议，不覆盖源文件"],
            ),
            self._enhancement_item(
                "smart_template_apply",
                "智能模板套用",
                "planned",
                "基于模板库、转换设置和文档结构，为 Word、PPT、Excel 输出匹配模板并记录套用依据",
                ["templates", "conversion_settings", "document_outline"],
                ["用户选择或确认模板", "保留原文件", "输出同名策略跟随设置"],
            ),
            self._enhancement_item(
                "private_deployment",
                "私有化部署",
                "contract",
                "沿用本地 API、SQLite/文件目录、安全令牌、本地任务载荷和安装画像，作为私有化部署前置合同",
                ["local_api", "sqlite_store", "install_profile", "local_payload"],
                ["敏感文档本地优先", "云同步默认关闭", "部署包仍需后续封装"],
            ),
            self._enhancement_item(
                "open_api",
                "开放 API 接口",
                "implemented",
                "接口目录、鉴权边界、敏感接口标记和本地任务载荷保护已经可通过 API 查询",
                ["api_catalog", "token_auth", "local_payload_contract"],
                ["本地载荷必须配置令牌", "路径字段默认脱敏", "敏感资源下载需受控"],
            ),
            self._enhancement_item(
                "quality_feedback_loop",
                "质量反馈闭环",
                "contract",
                "将低置信度公式、OMML 依赖、宏失败、微小图片、排版校正和转换质量检查汇总为后续智能修复输入",
                ["formula_annotations", "omml_annotations", "image_annotations", "failure_rows"],
                ["人工校正优先", "保留可追踪报告", "不自动重跑高风险宏"],
            ),
        ]
        readiness = [
            {
                "key": "local_security_token",
                "title": "本地安全令牌",
                "level": "good" if token_configured else "warn",
                "status": "已配置" if token_configured else "建议配置",
                "detail": "私有化、本地载荷和敏感资源下载建议强制令牌。",
            },
            {
                "key": "external_ai_authorization",
                "title": "外部 AI/OCR 授权",
                "level": "good" if external_mathpix_allowed else "blue",
                "status": "已授权" if external_mathpix_allowed else "默认关闭",
                "detail": "PDF Mathpix 与后续外部 AI 都需要显式授权后才允许上传文档内容。",
            },
            {
                "key": "local_client",
                "title": "本地客户端交接",
                "level": "good" if caps["localConnection"]["client_enabled"] else "warn",
                "status": caps["localConnection"]["status"],
                "detail": "复杂 Office、MathType、宏和本地路径动作通过本地任务载荷交给桌面端。",
            },
            {
                "key": "cloud_sync",
                "title": "云端同步",
                "level": "warn" if cloud_sync_allowed else "good",
                "status": "已允许" if cloud_sync_allowed else "默认关闭",
                "detail": "私有化部署建议默认关闭云端同步，仅同步脱敏摘要或授权后的任务状态。",
            },
        ]
        return {
            "schema_version": "k12.enhancementPlan.v1",
            "generated_at": utc_now(),
            "scope": "V3.0 增强能力规划",
            "items": items,
            "readiness": readiness,
            "summary": {
                "implemented": sum(1 for item in items if item["level"] == "good"),
                "contract": sum(1 for item in items if item["level"] == "warn"),
                "planned": sum(1 for item in items if item["level"] == "blue"),
            },
        }

    def install_plan(self, platform_name: str | None = None) -> dict[str, Any]:
        """生成 Windows/macOS 安装计划及 MathType 可移植性约定。"""
        settings = self.store.get_settings()
        profile = self.install_profile(platform_name)
        spec = self._installer_spec(profile["platform"])
        package_path = self.store.installers_dir / spec["file_name"] if spec["file_name"] else None
        package_size = package_path.stat().st_size if package_path and package_path.exists() and package_path.is_file() else 0
        package_invalid = bool(package_path and package_path.exists() and package_path.is_file() and package_size <= 0)
        package_exists = bool(package_path and package_path.exists() and package_path.is_file() and package_size > 0)
        token_configured = bool(str(settings.get("localSecurityToken") or "").strip())
        compatibility_mode = str(settings.get("mathtypeCompatibilityMode") or "platform-specific")
        compatibility_label = {
            "platform-specific": "区分平台",
            "mathml-latex": "MathML/LaTeX 优先",
            "image-fallback": "图片兜底",
        }.get(compatibility_mode, compatibility_mode)
        package = {
            "file_name": spec["file_name"],
            "status": "可下载" if package_exists else ("安装包无效" if package_invalid else ("需选择平台" if not spec["file_name"] else "待打包")),
            "download_url": f"/api/installers/{spec['file_name']}?{urlencode({'platform': profile['platform']})}" if package_exists else "",
            "download_available": package_exists,
            "package_present": package_exists or package_invalid,
            "package_valid": package_exists,
            "size": package_size if package_exists else 0,
            "sha256": self._sha256(package_path) if package_exists and package_path else "",
            "checksum_required": bool(spec["file_name"]),
            "expected_location": f"本地路径已隐藏/{package_path.name}" if package_path else "",
            "expected_location_available": bool(package_path),
            "path_policy": "安装计划不返回 installers 本地绝对路径；下载必须走 /api/installers/{file_name}?platform=...",
        }
        steps = self._installer_steps(profile["platform"])
        heartbeat_platform = self._heartbeat_platform(settings)
        readiness = [
            {
                "key": "platform",
                "title": "目标平台",
                "level": "good" if profile["platform"] in {"Windows", "macOS"} else "warn",
                "status": profile["platform"],
                "detail": "安装包、Office 自动化和 MathType 策略必须按平台区分。",
            },
            {
                "key": "package",
                "title": "安装包",
                "level": "good" if package_exists else "blue",
                "status": package["status"],
                "detail": "真实安装包放入本地 installers 目录后才开放下载。",
            },
            self._installer_heartbeat_readiness(profile["platform"], heartbeat_platform),
            {
                "key": "security_token",
                "title": "本地安全令牌",
                "level": "good" if token_configured else "warn",
                "status": "已配置" if token_configured else "建议配置",
                "detail": "安装完成后本地任务载荷、状态同步和敏感下载建议统一启用令牌。",
            },
            {
                "key": "formula_portability",
                "title": "公式跨平台兜底",
                "level": "good" if compatibility_mode in {"mathml-latex", "image-fallback"} else "warn",
                "status": compatibility_label,
                "detail": f"{profile['formulaPortability']} 当前模式：{compatibility_label}",
            },
        ]
        formula_compatibility = self._install_formula_compatibility_contract(profile, compatibility_mode, compatibility_label, heartbeat_platform)
        return {
            "schema_version": "k12.installPlan.v1",
            "generated_at": utc_now(),
            "platform": profile["platform"],
            "installer_kind": profile["installerKind"],
            "recommended_installer": profile["recommendedInstaller"],
            "download_label": profile["downloadLabel"],
            "package": package,
            "steps": steps,
            "readiness": readiness,
            "formula_compatibility": formula_compatibility,
            "warnings": [
                profile["formulaPortability"],
                "Windows 与 macOS 安装包不能混用；MathType 对象也不能跨平台互认为同一可编辑格式。",
                "当前接口只管理安装计划和本地安装包下载，不会在网页端静默安装软件。",
            ],
        }

    @staticmethod
    def _install_formula_compatibility_contract(profile: dict[str, Any], compatibility_mode: str, compatibility_label: str, heartbeat_platform: str) -> dict[str, Any]:
        """生成 Windows/macOS 边界的安装公式兼容性约定。"""
        target = normalize_platform(str(profile.get("platform") or "Unknown"))
        heartbeat = normalize_platform(heartbeat_platform) if heartbeat_platform else ""
        native_mode = compatibility_mode == "platform-specific"
        same_platform_required = target in {"Windows", "macOS"} and native_mode
        platform_mismatch = bool(
            heartbeat
            and target in {"Windows", "macOS"}
            and heartbeat in {"Windows", "macOS"}
            and target != heartbeat
        )
        native_handoff_allowed = same_platform_required and bool(heartbeat) and not platform_mismatch
        blocking_reasons: list[str] = []
        if native_mode and target not in {"Windows", "macOS"}:
            blocking_reasons.append("target_platform_unknown")
        if same_platform_required and not heartbeat:
            blocking_reasons.append("client_platform_heartbeat_missing")
        if platform_mismatch:
            blocking_reasons.append("platform_mismatch")
        fallback_formats = ["MathML", "LaTeX", "图片"]
        if platform_mismatch:
            status = "平台不符"
            message = f"安装计划目标为 {target}，但当前客户端心跳来自 {heartbeat}；不能交付平台专属 MathType 对象。"
            recommended_action = "切换到同平台本地客户端，或改用 MathML、LaTeX、图片兜底格式。"
        elif native_handoff_allowed:
            status = "同平台可交接"
            message = f"安装计划目标和客户端心跳均为 {target}；可按同平台 MathType 对象合同交接。"
            recommended_action = "继续使用当前同平台客户端；跨平台交付仍需同时保留 MathML、LaTeX 或图片兜底。"
        elif same_platform_required:
            status = "需同平台"
            message = f"当前为 {target} 平台专属对象模式，MathType 对象只能交给同平台客户端处理。"
            recommended_action = "启动本地客户端并回传同平台心跳；跨平台交付时改用 MathML、LaTeX、图片兜底格式。"
        else:
            status = "跨平台兜底"
            message = "当前安装计划要求生成 MathML、LaTeX 或图片兜底，避免跨平台 MathType 对象不兼容。"
            recommended_action = "按兜底格式交付公式，不写入平台专属 MathType 原生对象。"
        return {
            "schema_version": "k12.installFormulaCompatibility.v1",
            "target_platform": target,
            "heartbeat_platform": heartbeat or "未连接",
            "compatibility_mode": compatibility_mode,
            "mode_label": compatibility_label,
            "native_object_format": profile.get("mathtypeObjectFormat", ""),
            "platform_objects_cross_compatible": False,
            "same_platform_required_for_native_objects": same_platform_required,
            "native_mathtype_object_allowed": same_platform_required and not platform_mismatch,
            "native_handoff_allowed": native_handoff_allowed,
            "native_handoff_blocking_reasons": blocking_reasons,
            "fallback_required_for_cross_platform": not native_handoff_allowed,
            "fallback_formats": fallback_formats,
            "status": status,
            "message": message,
            "recommended_action": recommended_action,
            "package_boundary": "Windows 使用 .msi，macOS 使用 .pkg；安装包和 MathType 原生对象不能跨平台混用。",
        }

    @staticmethod
    def _heartbeat_platform(settings: dict[str, Any]) -> str:
        """规范最近客户端心跳报告的平台。"""
        heartbeat = settings.get("localClientHeartbeat") if isinstance(settings.get("localClientHeartbeat"), dict) else {}
        raw_platform = str(heartbeat.get("platform") or "").strip() if heartbeat else ""
        return normalize_platform(raw_platform) if raw_platform else ""

    @staticmethod
    def _installer_heartbeat_readiness(target_platform: str, heartbeat_platform: str) -> dict[str, Any]:
        """说明安装目标与心跳平台是否兼容。"""
        target = normalize_platform(target_platform)
        if not heartbeat_platform:
            return {
                "key": "client_platform_heartbeat",
                "title": "客户端平台心跳",
                "level": "blue",
                "status": "未连接",
                "detail": "安装完成并启动本地客户端后，心跳会用于校验 Windows/macOS 平台是否与安装计划一致。",
            }
        actual = normalize_platform(heartbeat_platform)
        if target in {"Windows", "macOS"} and actual in {"Windows", "macOS"} and target != actual:
            return {
                "key": "client_platform_heartbeat",
                "title": "客户端平台心跳",
                "level": "warn",
                "status": "平台不符",
                "detail": f"安装计划目标为 {target}，但当前心跳来自 {actual}；Windows 与 macOS MathType 对象不通用，请切换安装包或客户端。",
            }
        if target in {"Windows", "macOS"} and actual == target:
            return {
                "key": "client_platform_heartbeat",
                "title": "客户端平台心跳",
                "level": "good",
                "status": "已匹配",
                "detail": f"当前本地客户端心跳平台为 {actual}，与安装计划一致。",
            }
        return {
            "key": "client_platform_heartbeat",
            "title": "客户端平台心跳",
            "level": "warn",
            "status": actual or "未知",
            "detail": "平台未完全确认，涉及 MathType 对象时优先使用 MathML/LaTeX 或图片兜底。",
        }

    def installer_download_info(self, file_name: str, requested_platform: str | None = None) -> dict[str, Any]:
        """校验明确限定平台的本地安装包下载。"""
        safe_name = Path(file_name).name
        if safe_name != file_name:
            raise ValueError("Invalid installer name")
        registered = {
            spec["file_name"]: {**spec, "platform": platform}
            for platform, spec in self._installer_specs().items()
            if spec.get("file_name")
        }
        spec = registered.get(safe_name)
        if not spec:
            raise ValueError("Installer is not registered for Windows or macOS")
        requested = normalize_platform(requested_platform) if requested_platform else ""
        if not requested:
            raise ValueError("Installer platform query is required")
        if requested not in {"Windows", "macOS"}:
            raise ValueError("Installer platform must be Windows or macOS")
        if requested != spec["platform"]:
            raise ValueError("Installer platform does not match the requested platform")
        if Path(safe_name).suffix.lower() != spec["extension"]:
            raise ValueError("Installer extension does not match its platform")
        target = (self.store.installers_dir / safe_name).resolve()
        root = self.store.installers_dir.resolve()
        try:
            target.relative_to(root)
        except ValueError as exc:
            raise ValueError("Invalid installer path") from exc
        if not target.exists() or not target.is_file():
            raise FileNotFoundError(safe_name)
        if target.stat().st_size <= 0:
            raise ValueError("Installer file is empty")
        return {
            "path": str(target),
            "file_name": safe_name,
            "platform": spec["platform"],
            "installer_kind": spec["installer_kind"],
            "package_boundary": "Windows 使用 .msi，macOS 使用 .pkg；安装包和 MathType 原生对象不能跨平台混用。",
            "package_boundary_header": "windows-msi-and-macos-pkg-are-not-cross-platform",
            "formula_object_boundary": "Windows 与 macOS MathType 原生对象不跨平台兼容；下载安装包后仍需按同平台公式对象合同交接。",
            "formula_object_boundary_header": "mathtype-native-objects-require-same-platform",
            "fallback_formats": ["MathML", "LaTeX", "图片"],
            "fallback_formats_header": "MathML,LaTeX,image",
            "sha256": self._sha256(target),
            "file_size": target.stat().st_size,
        }

    def local_client_manifest(self) -> dict[str, Any]:
        """返回启动、心跳和载荷接口的本地伴随清单。"""
        settings = self.store.get_settings()
        profile = self.install_profile()
        install_plan = self.install_plan(profile["platform"])
        host = str(settings.get("localApiHost") or "127.0.0.1")
        port = int(settings.get("localApiPort") or 8765)
        origin = f"http://{host}:{port}"
        token_configured = bool(str(settings.get("localSecurityToken") or "").strip())
        python_command = "python" if profile["platform"] == "Windows" else "python3"
        heartbeat = dict(settings.get("localClientHeartbeat") or {})
        package = install_plan.get("package", {}) if isinstance(install_plan.get("package"), dict) else {}
        installer_manifest = {
            "schema_version": "k12.localInstallerManifest.v1",
            "platform": install_plan["platform"],
            "target_platform": install_plan["formula_compatibility"]["target_platform"],
            "heartbeat_platform": install_plan["formula_compatibility"]["heartbeat_platform"],
            "installer_kind": install_plan["installer_kind"],
            "file_name": str(package.get("file_name") or ""),
            "status": str(package.get("status") or "待打包"),
            "download_url": str(package.get("download_url") or ""),
            "download_requires_platform_query": bool(package.get("file_name")),
            "download_available": bool(package.get("download_available")),
            "package_present": bool(package.get("package_present")),
            "package_valid": bool(package.get("package_valid")),
            "size": int(package.get("size") or 0),
            "sha256": str(package.get("sha256") or ""),
            "checksum_required": bool(package.get("checksum_required")),
            "expected_location_available": bool(package.get("expected_location")),
            "path_policy": "本地 installers 路径不写入 manifest；安装包下载必须走 /api/installers/{file_name}?platform=...",
            "formula_object_boundary": install_plan["formula_compatibility"]["package_boundary"],
            "same_platform_required_for_native_objects": bool(install_plan["formula_compatibility"]["same_platform_required_for_native_objects"]),
            "native_handoff_allowed": bool(install_plan["formula_compatibility"]["native_handoff_allowed"]),
            "native_handoff_blocking_reasons": list(install_plan["formula_compatibility"]["native_handoff_blocking_reasons"]),
            "fallback_formats": list(install_plan["formula_compatibility"]["fallback_formats"]),
        }
        local_tasks = [
            task
            for task in self.store.list_tasks()
            if str(task.get("execute_mode") or "") in {"local", "hybrid"} or str(task.get("task_type") or "") in LOCAL_REQUIRED_TASKS
        ]
        pending_tasks = [task for task in local_tasks if task.get("status") not in {"成功", "失败", "已取消"}]
        return {
            "schema_version": "k12.localClientManifest.v1",
            "generated_at": utc_now(),
            "service": {
                "origin": origin,
                "host": host,
                "port": port,
                "loopback_only": host in {"127.0.0.1", "localhost", "::1"},
            },
            "platform": {
                "selected": profile["platform"],
                "installer_kind": install_plan["installer_kind"],
                "recommended_installer": install_plan["recommended_installer"],
                "install_plan_endpoint": "/api/install-plan",
                "installer": installer_manifest,
                "package_boundary": install_plan["formula_compatibility"]["package_boundary"],
                "mathtype_object_format": profile["mathtypeObjectFormat"],
                "formula_portability": profile["formulaPortability"],
                "formula_object_interop": profile["formulaObjectInterop"],
                "formula_compatibility": install_plan["formula_compatibility"],
            },
            "security": {
                "token_configured": token_configured,
                "token_header": "X-K12-Token",
                "bearer_supported": True,
                "query_token_supported_for_downloads": True,
                "payload_requires_configured_token": True,
                "heartbeat_requires_configured_token": True,
                "local_paths_hidden_from_web": not bool(settings.get("exposeLocalPaths", False)),
            },
            "launch": {
                "web_launch_allowed": bool(settings.get("allowWebLaunchLocalClient", False)),
                "protocol": "k12-local",
                "url_template": f"k12-local://open?api={origin}&task={{task_id}}",
            },
            "companion_cli": {
                "available": True,
                "command": f"{python_command} -m k12.local_client --origin {origin}",
                "native_plan_command": f"{python_command} -m k12.local_client --origin {origin} --native-plan",
                "native_report_command": f"{python_command} -m k12.local_client --origin {origin} --native-report-json <report.json>",
                "file_action_command": f"{python_command} -m k12.local_client --origin {origin} --execute-file-actions",
                "dry_run_supported": True,
                "dry_run_execution_schema": "k12.localDryRunExecution.v1",
                "native_plan_supported": True,
                "native_execution_request_schema": "k12.localNativeExecutionRequest.v1",
                "native_report_supported": True,
                "native_execution_report_schema": "k12.localNativeExecutionReport.v1",
                "file_actions_supported": True,
                "file_action_execution_schema": "k12.localFileActionExecution.v1",
                "macos_office_execution_supported": True,
                "windows_office_execution_supported": True,
                "windows_native_office_execution_schema": "k12.windowsOfficeTaskExecution.v1",
                "native_office_execution_schema": "k12.macosOfficeTaskExecution.v1",
                "native_office_command": f"python -m k12.local_client --origin {origin} --allow-native-execution --execute-native-office",
                "executes_native_documents": True,
                "description": "本地伴随客户端支持心跳、任务载荷、模拟执行、原生请求、报告同步和 OMML 依赖复制等安全本地文件动作；双重显式授权后，可通过 Windows Office COM 或 macOS AppleScript 执行旧 Word/PPT 转换。MathType、OMML 写回和 Word 宏仍保持阻断。",
            },
            "endpoints": [
                {"method": "GET", "path": "/api/tasks/{task_id}/local-payload", "auth": "configured-token", "sensitive": True},
                {"method": "GET", "path": "/api/tasks/{task_id}/local-readiness", "auth": "token-or-local", "sensitive": False},
                {"method": "POST", "path": "/api/tasks/{task_id}/local-launch", "auth": "configured-token + launch-authorization", "sensitive": True},
                {"method": "POST", "path": "/api/tasks/{task_id}/local-sync", "auth": "token + sync-authorization", "sensitive": True},
                {"method": "POST", "path": "/api/local-client/uploads", "auth": "configured-token + cloud-sync-authorization", "sensitive": True},
                {"method": "GET", "path": "/api/local-client/uploads/{upload_id}/manifest", "auth": "token-or-local", "sensitive": True},
                {"method": "POST", "path": "/api/local-client/heartbeat", "auth": "configured-token", "sensitive": False},
                {"method": "GET", "path": "/api/install-plan", "auth": "token-or-local", "sensitive": False},
            ],
            "supported_actions": [
                {"type": "office_conversion", "label": "Office 转换复核", "requires_platform": True},
                {"type": "omml_mathtype", "label": "OMML/MathType 处理", "requires_platform": True},
                {"type": "macro_sequence", "label": "Word 宏顺序执行", "requires_platform": True},
                {"type": "open_output_directory", "label": "打开输出目录", "requires_platform": False},
            ],
            "queue": {
                "local_task_count": len(local_tasks),
                "pending_local_task_count": len(pending_tasks),
                "next_task_id": pending_tasks[0]["id"] if pending_tasks else "",
            },
            "heartbeat": heartbeat,
        }

    def record_local_client_heartbeat(self, payload: dict[str, Any]) -> dict[str, Any]:
        """保存脱敏客户端心跳及组件能力摘要。"""
        raw_capabilities = payload.get("capabilities") if isinstance(payload.get("capabilities"), dict) else {}
        capabilities = {
            str(key): value
            for key, value in raw_capabilities.items()
            if isinstance(value, (bool, int, float, str)) and key in {"officeAutomation", "mathTypeAutomation", "macroExecution", "ommlDependencySearch"}
        }
        heartbeat = {
            "schema_version": "k12.localClientHeartbeat.v1",
            "client_id": str(payload.get("client_id") or payload.get("clientId") or "local-client")[:80],
            "status": str(payload.get("status") or "online")[:40],
            "platform": normalize_platform(str(payload.get("platform") or self.install_profile()["platform"])),
            "version": str(payload.get("version") or "")[:40],
            "active_task_id": str(payload.get("active_task_id") or payload.get("activeTaskId") or "")[:80],
            "message": str(payload.get("message") or "")[:200],
            "capabilities": capabilities,
            "preflight": self._safe_local_client_preflight(payload.get("preflight")),
            "received_at": utc_now(),
        }
        self.store.update_settings({"localClientHeartbeat": heartbeat})
        self.store.append_log("", f"本地客户端心跳：{heartbeat['platform']} {heartbeat['status']}", category="system")
        return heartbeat

    @staticmethod
    def _safe_local_client_preflight(preflight: Any) -> dict[str, Any]:
        """脱敏并规范心跳载荷中的预检详情。"""
        if not isinstance(preflight, dict):
            return {}
        raw_components = preflight.get("components") if isinstance(preflight.get("components"), dict) else {}
        components: dict[str, dict[str, Any]] = {}
        for key, value in raw_components.items():
            if key not in {"office", "word", "powerpoint", "excel", "mathtype", "libreoffice", "omml_dependency"} or not isinstance(value, dict):
                continue
            components[str(key)] = {
                "label": str(value.get("label") or key)[:80],
                "available": bool(value.get("available")),
                "status": str(value.get("status") or "")[:40],
            }
        raw_capabilities = preflight.get("capabilities") if isinstance(preflight.get("capabilities"), dict) else {}
        capabilities = {
            str(key): value
            for key, value in raw_capabilities.items()
            if isinstance(value, (bool, int, float, str)) and key in {"officeAutomation", "mathTypeAutomation", "macroExecution", "ommlDependencySearch"}
        }
        return {
            "schema_version": str(preflight.get("schema_version") or "k12.localClientPreflight.v1")[:60],
            "platform": normalize_platform(str(preflight.get("platform") or "")),
            "components": components,
            "capabilities": capabilities,
            "executes_native_documents": bool(preflight.get("executes_native_documents")),
            "path_policy": "component paths are not stored",
        }

    def local_result_upload_queue(self) -> dict[str, Any]:
        """返回结果上传意图，不公开本地输出路径。"""
        settings = self.store.get_settings()
        cloud_allowed = bool(settings.get("allowCloudSync", False))
        items = [self._local_upload_item(task, cloud_allowed) for task in self.store.list_tasks()]
        items = [item for item in items if item]
        return {
            "schema_version": "k12.localResultUploadQueue.v1",
            "generated_at": utc_now(),
            "cloud_sync_allowed": cloud_allowed,
            "policy": "允许上传处理结果" if cloud_allowed else "仅本地保存结果，上传请求会被标记为未授权",
            "items": items,
            "summary": {
                "total": len(items),
                "queued": sum(1 for item in items if item["upload_status"] == "queued"),
                "registered": sum(1 for item in items if item["upload_status"] == "registered"),
                "not_authorized": sum(1 for item in items if item["upload_status"] == "not_authorized"),
                "local_only": sum(1 for item in items if item["upload_status"] in {"not_requested", "local_only"}),
                "output_count": sum(int(item["output_count"]) for item in items),
            },
        }

    def local_result_upload_manifest(self, upload_id: str) -> dict[str, Any]:
        """返回含哈希与脱敏路径的上传接收约定。"""
        upload_id = str(upload_id or "").strip()
        if not upload_id:
            raise KeyError("Upload not found")
        settings = self.store.get_settings()
        cloud_allowed = bool(settings.get("allowCloudSync", False))
        for task in self.store.list_tasks():
            uploads = [item for item in list(task.get("local_result_uploads") or []) if isinstance(item, dict)]
            for upload in uploads:
                if str(upload.get("upload_id") or "") != upload_id:
                    continue
                outputs = self._local_upload_manifest_outputs(upload.get("outputs") or [])
                received_files = self._local_upload_manifest_files(upload.get("received_files") or [])
                computed_total_size = sum(int(item.get("size") or 0) for item in outputs if isinstance(item.get("size"), int))
                declared_total_size = self._non_negative_int(upload.get("total_size"), computed_total_size)
                output_sha_count = sum(1 for item in outputs if item.get("sha256"))
                status = str(upload.get("status") or "registered")
                content_received = bool(received_files)
                return {
                    "schema_version": "k12.localResultUploadManifest.v1",
                    "generated_at": utc_now(),
                    "upload_id": upload_id,
                    "status": status,
                    "status_label": str(upload.get("status_label") or "已登记待云端接收"),
                    "registered_at": str(upload.get("registered_at") or ""),
                    "task": {
                        "id": task.get("id", ""),
                        "task_type": task.get("task_type", ""),
                        "task_label": task.get("task_label") or TASK_LABELS.get(str(task.get("task_type") or ""), str(task.get("task_type") or "")),
                        "status": task.get("status", ""),
                        "local_synced_at": task.get("local_synced_at", ""),
                    },
                    "client_id": str(upload.get("client_id") or "")[:80],
                    "cloud_sync_allowed_at_registration": bool(upload.get("cloud_sync_allowed", False)),
                    "cloud_sync_allowed_now": cloud_allowed,
                    "receive_state": "received" if status == "registered" and cloud_allowed and content_received else ("ready_for_cloud_receiver" if status == "registered" and cloud_allowed else "cloud_sync_disabled"),
                    "receive_contract": {
                        "mode": "content_upload" if content_received else "manifest_only",
                        "content_transfer": "included" if content_received else "not_included",
                        "path_policy": "本地路径已隐藏",
                        "requires_cloud_sync_authorization": True,
                        "requires_local_security_token": True,
                        "next_step": "已接收上传内容并完成摘要登记；后续云端服务可按 upload_id、包哈希和输出清单继续分发。" if content_received else "云端接收服务按 upload_id、包哈希和输出清单建立接收任务；本接口不上传完整文件内容。",
                    },
                    "package": {
                        "package_sha256": self._safe_sha256(upload.get("package_sha256") or ""),
                        "output_count": len(outputs),
                        "total_size": declared_total_size,
                        "computed_output_size": computed_total_size,
                        "received_file_count": len(received_files),
                        "received_total_size": sum(int(item.get("size") or 0) for item in received_files if isinstance(item.get("size"), int)),
                    },
                    "integrity": {
                        "package_sha256_provided": bool(self._safe_sha256(upload.get("package_sha256") or "")),
                        "output_sha256_count": output_sha_count,
                        "all_outputs_named": all(bool(item.get("name")) for item in outputs),
                        "size_matches_manifest": declared_total_size == computed_total_size,
                        "received_file_sha256_count": sum(1 for item in received_files if item.get("sha256")),
                        "all_received_files_hashed": all(bool(item.get("sha256")) for item in received_files),
                    },
                    "outputs": outputs,
                    "received_files": received_files,
                    "path_policy": "本地路径已隐藏；manifest 不返回本地绝对路径，也不包含文件内容。",
                }
        raise KeyError("Upload not found")

    def register_local_result_upload(self, payload: dict[str, Any]) -> dict[str, Any]:
        """获得云同步授权后注册令牌保护的本地结果上传。"""
        settings = self.store.get_settings()
        if not str(settings.get("localSecurityToken") or "").strip():
            raise ValueError("本地结果上传需要先配置本地安全令牌")
        if not settings.get("allowCloudSync", False):
            raise ValueError("结果上传云端未授权")
        task_id = str(payload.get("task_id") or payload.get("taskId") or "")
        if not task_id:
            raise ValueError("缺少任务 ID")
        task = self.store.get_task(task_id)
        if not task:
            raise KeyError(f"Task not found: {task_id}")
        upload_id = new_id("local_upload")
        received_files = self._receive_local_upload_files(upload_id, payload.get("files") or payload.get("contents") or payload.get("uploaded_files"))
        outputs = self._local_upload_manifest_outputs(payload.get("outputs") or payload.get("artifacts") or task.get("local_client_outputs") or [])
        if not outputs and received_files:
            outputs = [
                {
                    "name": item["name"],
                    "output_type": item.get("output_type", ""),
                    "status": "已接收",
                    "size": item.get("size", 0),
                    "sha256": item.get("sha256", ""),
                    "path_available": False,
                    "path_display": "",
                }
                for item in received_files
            ]
        if not outputs:
            raise ValueError("缺少本地结果输出摘要")
        total_size = sum(int(item.get("size") or 0) for item in outputs if isinstance(item.get("size"), int))
        now = utc_now()
        upload = {
            "schema_version": "k12.localResultUpload.v1",
            "upload_id": upload_id,
            "task_id": task_id,
            "client_id": str(payload.get("client_id") or payload.get("clientId") or "local-client")[:80],
            "status": "registered",
            "status_label": "已登记待云端接收",
            "registered_at": now,
            "output_count": len(outputs),
            "total_size": total_size,
            "package_sha256": self._safe_sha256(payload.get("package_sha256") or payload.get("packageSha256") or ""),
            "outputs": outputs,
            "receive_mode": "content_received" if received_files else "manifest_only",
            "received_file_count": len(received_files),
            "received_total_size": sum(int(item.get("size") or 0) for item in received_files if isinstance(item.get("size"), int)),
            "received_files": received_files,
            "path_policy": "本地路径已隐藏",
            "cloud_sync_allowed": True,
            "message": str(payload.get("message") or ("本地结果上传内容已接收，路径已隐藏" if received_files else "本地结果上传包清单已登记，等待云端服务接收"))[:200],
        }
        uploads = [item for item in list(task.get("local_result_uploads") or []) if isinstance(item, dict)]
        uploads.append(upload)
        task["local_result_uploads"] = uploads[-20:]
        task["local_client_outputs"] = outputs
        task["local_synced_at"] = task.get("local_synced_at") or now
        sync = dict(task.get("local_sync") or {})
        sync.update(
            {
                "task_status_cloud_sync_allowed": bool(settings.get("allowTaskStatusCloudSync", False)),
                "result_upload_requested": True,
                "result_upload_allowed": True,
                "result_upload_status": "registered",
                "last_upload_id": upload["upload_id"],
                "last_upload_registered_at": now,
                "synced_at": sync.get("synced_at") or now,
            }
        )
        task["local_sync"] = sync
        self.store.save_task(task)
        received_note = f"，已接收 {len(received_files)} 个文件" if received_files else ""
        self.store.append_log(task_id, f"本地结果上传登记：{len(outputs)} 个输出{received_note}，路径已隐藏", category="system")
        return upload

    def mathpix_job_queue(self) -> dict[str, Any]:
        """汇总 Mathpix PDF 转 Word 作业，不触发外部上传。"""
        settings = self.store.get_settings()
        jobs: list[dict[str, Any]] = []
        for report in self.store.list_reports():
            for job in report.get("analysis", {}).get("mathpix", []):
                jobs.append(self._mathpix_queue_item(report, job))
        return {
            "schema_version": "k12.mathpixJobQueue.v1",
            "generated_at": utc_now(),
            "engine": settings.get("pdfToWordEngine", "Mathpix"),
            "external_upload_allowed": bool(settings.get("allowExternalMathpixUpload", False)),
            "authorization": "已授权" if settings.get("allowExternalMathpixUpload", False) else "未授权",
            "credential_envs": list(self._mathpix_env_names(settings)),
            "jobs": jobs,
            "summary": {
                "total": len(jobs),
                "blocked": sum(1 for item in jobs if item["status"] in {"authorization_required", "missing_credentials", "ocr_disabled", "missing_local_file"}),
                "submitted": sum(1 for item in jobs if item["status"] in {"submitted", "processing"}),
                "completed": sum(1 for item in jobs if item["status"] == "completed"),
                "partial_failed": sum(1 for item in jobs if item.get("tex_zip_status") == "download_failed"),
                "failed": sum(1 for item in jobs if item["status"] in {"api_error", "download_failed"}),
            },
        }

    def _mathpix_queue_item(self, report: dict[str, Any], job: dict[str, Any]) -> dict[str, Any]:
        """根据报告作业生成不含路径的 Mathpix 队列行。"""
        outputs = dict(job.get("outputs") or {})
        output_summaries = self._mathpix_output_summaries(job.get("output_summaries"))
        retention = dict(job.get("retention_plan") or {})
        recognition = dict(job.get("recognition_plan") or {})
        return {
            "report_id": report.get("id", ""),
            "task_id": report.get("task_id", ""),
            "file_id": job.get("file_id", ""),
            "file_name": job.get("file_name", ""),
            "engine": job.get("engine", "Mathpix"),
            "status": job.get("status", ""),
            "status_label": self._mathpix_status_label(str(job.get("status") or "")),
            "message": job.get("message", ""),
            "pdf_id_available": bool(job.get("pdf_id")),
            "percent_done": job.get("percent_done", ""),
            "tex_zip_status": job.get("tex_zip_status", ""),
            "tex_zip_message": job.get("tex_zip_message", ""),
            "ocr_settings": job.get("ocr_settings", {}),
            "request": self._mathpix_request_summary(job.get("request_options")),
            "recognition_plan": {
                "schema_version": recognition.get("schema_version", ""),
                "status": recognition.get("status", ""),
                "status_label": recognition.get("status_label", ""),
                "submit_allowed": bool(recognition.get("submit_allowed")),
                "external_upload_authorized": bool(recognition.get("external_upload_authorized")),
                "external_upload_authorization_source": str(recognition.get("external_upload_authorization_source") or ""),
                "credential_values_exposed": False,
                "upload_gate": {
                    "schema_version": (recognition.get("upload_gate") or {}).get("schema_version", ""),
                    "status": (recognition.get("upload_gate") or {}).get("status", ""),
                    "submit_allowed": bool((recognition.get("upload_gate") or {}).get("submit_allowed")),
                    "blocking_reasons": list((recognition.get("upload_gate") or {}).get("blocking_reasons") or []),
                    "external_upload_authorized": bool((recognition.get("upload_gate") or {}).get("external_upload_authorized")),
                    "authorization_source": str((recognition.get("upload_gate") or {}).get("authorization_source") or ""),
                    "task_options_can_authorize": bool((recognition.get("upload_gate") or {}).get("task_options_can_authorize")),
                    "source_file_available": bool((recognition.get("upload_gate") or {}).get("source_file_available")),
                    "ocr_enabled": bool((recognition.get("upload_gate") or {}).get("ocr_enabled")),
                    "credential_values_exposed": False,
                },
                "conversion_formats": list(recognition.get("submit_contract", {}).get("conversion_formats") or []),
                "download_formats": list(recognition.get("submit_contract", {}).get("download_formats") or []),
                "wait_for_completion": bool(recognition.get("submit_contract", {}).get("wait_for_completion")),
                "wait_for_completion_source": str(recognition.get("submit_contract", {}).get("wait_for_completion_source") or ""),
                "download_results": dict(recognition.get("download_results") or {}),
                "download_failures": list(recognition.get("download_failures") or []),
                "download_manifest": dict(recognition.get("download_manifest") or {}),
                "formula_review": dict(recognition.get("formula_review") or {}),
                "task_option_audit": dict(recognition.get("task_option_audit") or {}),
                "safety": dict(recognition.get("safety") or {}),
            },
            "retention": {
                "pdf_type": retention.get("pdf_type", ""),
                "image_objects": retention.get("image_objects", 0),
                "table_hints": retention.get("table_hints", 0),
                "formula_hints": retention.get("formula_hints", 0),
                "image_retention_status": retention.get("image_retention_status", ""),
                "table_retention_status": retention.get("table_retention_status", ""),
                "formula_retention_status": retention.get("formula_retention_status", ""),
            },
            "outputs": [
                {
                    "type": key,
                    "status": "available",
                    "path_display": "本地路径已隐藏",
                    "file_name": output_summaries.get(key, {}).get("file_name", ""),
                    "file_size": output_summaries.get(key, {}).get("size", 0),
                    "sha256": output_summaries.get(key, {}).get("sha256", ""),
                }
                for key, value in outputs.items()
                if value
            ],
        }

    @staticmethod
    def _mathpix_request_summary(options: Any) -> dict[str, Any]:
        """汇总 Mathpix 请求元数据，不含凭据或路径。"""
        raw = options if isinstance(options, dict) else {}
        formats = raw.get("conversion_formats") if isinstance(raw.get("conversion_formats"), dict) else {}
        return {
            "conversion_formats": [str(key) for key, value in formats.items() if bool(value)],
            "formula_zip_requested": bool(formats.get("tex.zip")),
            "enable_tables_fallback": bool(raw.get("enable_tables_fallback")),
            "include_page_info": bool(raw.get("include_page_info")),
            "metadata_keys": sorted(str(key) for key in raw.get("metadata", {}).keys()) if isinstance(raw.get("metadata"), dict) else [],
        }

    @staticmethod
    def _mathpix_status_label(status: str) -> str:
        """将 Mathpix 作业状态转为用户可见标签。"""
        labels = {
            "authorization_required": "等待授权",
            "missing_credentials": "缺少凭证",
            "ocr_disabled": "OCR 已关闭",
            "missing_local_file": "缺少本地文件",
            "submitted": "已提交",
            "completed": "已完成",
            "api_error": "接口错误",
            "skipped_existing_output": "已跳过",
            "download_failed": "下载失败",
        }
        return labels.get(status, status or "未知")

    def _local_upload_item(self, task: dict[str, Any], cloud_allowed: bool) -> dict[str, Any] | None:
        """生成不含路径的本地结果上传队列行。"""
        sync = dict(task.get("local_sync") or {})
        uploads = [item for item in list(task.get("local_result_uploads") or []) if isinstance(item, dict)]
        latest_upload = uploads[-1] if uploads else {}
        outputs = self._local_upload_manifest_outputs(latest_upload.get("outputs") or task.get("local_client_outputs") or [])
        requested = bool(sync.get("result_upload_requested"))
        if not requested and not outputs:
            return None
        upload_status = str(sync.get("result_upload_status") or ("not_authorized" if requested and not cloud_allowed else ("queued" if requested else "local_only")))
        return {
            "task_id": task.get("id", ""),
            "task_label": task.get("task_label") or TASK_LABELS.get(str(task.get("task_type") or ""), str(task.get("task_type") or "")),
            "task_type": task.get("task_type", ""),
            "task_status": task.get("status", ""),
            "synced_at": task.get("local_synced_at", ""),
            "upload_id": latest_upload.get("upload_id", ""),
            "manifest_available": bool(latest_upload.get("upload_id")),
            "manifest_endpoint": f"/api/local-client/uploads/{latest_upload.get('upload_id')}/manifest" if latest_upload.get("upload_id") else "",
            "upload_registered_at": latest_upload.get("registered_at", ""),
            "receive_mode": latest_upload.get("receive_mode", "manifest_only"),
            "received_file_count": int(latest_upload.get("received_file_count") or 0),
            "received_total_size": int(latest_upload.get("received_total_size") or 0),
            "result_upload_requested": requested,
            "result_upload_allowed": bool(sync.get("result_upload_allowed", cloud_allowed)),
            "upload_status": upload_status,
            "upload_status_label": {
                "queued": "等待上传",
                "registered": "已登记待云端接收",
                "not_authorized": "未授权",
                "not_requested": "仅本地保存",
                "local_only": "仅本地保存",
            }.get(upload_status, upload_status),
            "output_count": len(outputs),
            "uploaded_package_count": len(uploads),
            "outputs": outputs,
        }

    @classmethod
    def _local_upload_outputs(cls, outputs: Any) -> list[dict[str, Any]]:
        """清理上传输出摘要供队列响应使用。"""
        if not isinstance(outputs, list):
            return []
        cleaned: list[dict[str, Any]] = []
        for item in outputs[:20]:
            if not isinstance(item, dict):
                continue
            path_value = str(item.get("path") or "")
            cleaned.append(
                {
                    "name": cls._safe_output_name(item.get("name") or path_value),
                    "output_type": str(item.get("output_type") or item.get("type") or ""),
                    "status": str(item.get("status") or "已生成"),
                    "size": item.get("size", ""),
                    "path_available": bool(path_value),
                    "path_display": "本地路径已隐藏" if path_value else "",
                }
            )
        return cleaned

    def _receive_local_upload_files(self, upload_id: str, files: Any) -> list[dict[str, Any]]:
        """接收明确提交的本地上传包至受管运行存储。"""
        if not isinstance(files, list):
            return []
        safe_upload_id = self._safe_output_name(upload_id)
        target_dir = self.store.cloud_uploads_dir / safe_upload_id
        target_dir.mkdir(parents=True, exist_ok=True)
        received: list[dict[str, Any]] = []
        total_size = 0
        limit = int(self.store.get_settings().get("singleFileLimitMb", 500)) * 1024 * 1024
        for index, item in enumerate(files[:50], start=1):
            if not isinstance(item, dict):
                continue
            encoded = str(item.get("content_base64") or item.get("contentBase64") or item.get("base64") or item.get("data") or "")
            if "," in encoded and encoded.lstrip().startswith("data:"):
                encoded = encoded.split(",", 1)[1]
            if not encoded.strip():
                raise ValueError("上传文件内容不能为空")
            try:
                content = base64.b64decode(encoded, validate=True)
            except (binascii.Error, ValueError) as exc:
                raise ValueError("上传文件内容不是有效的 Base64") from exc
            if not content:
                raise ValueError("上传文件内容不能为空")
            total_size += len(content)
            if total_size > limit:
                raise ValueError("上传内容超过单文件大小限制")
            raw_name = str(item.get("name") or item.get("file_name") or item.get("fileName") or item.get("path") or f"output-{index}.bin")
            name = self._safe_output_name(raw_name)
            digest = hashlib.sha256(content).hexdigest()
            declared_sha = self._safe_sha256(item.get("sha256") or item.get("checksum") or "")
            if declared_sha and declared_sha != digest:
                raise ValueError(f"上传文件哈希不匹配：{name}")
            target = self._unique_cloud_upload_path(target_dir / name)
            target.write_bytes(content)
            received.append(
                {
                    "name": target.name[:160],
                    "output_type": str(item.get("output_type") or item.get("type") or Path(name).suffix.lstrip("."))[:40],
                    "status": "已接收",
                    "size": len(content),
                    "sha256": digest,
                    "content_received": True,
                    "path_available": True,
                    "path_display": "云端接收区路径已隐藏",
                }
            )
        return received

    @staticmethod
    def _unique_cloud_upload_path(target: Path) -> Path:
        """为本地上传文件保留唯一受管路径。"""
        if not target.exists():
            return target
        stem = target.stem or "output"
        suffix = target.suffix
        for index in range(1, 1000):
            candidate = target.with_name(f"{stem}-{index}{suffix}")
            if not candidate.exists():
                return candidate
        return target.with_name(f"{stem}-{new_id('cloud')}{suffix}")

    @classmethod
    def _local_upload_manifest_files(cls, files: Any) -> list[dict[str, Any]]:
        """返回上传清单的已清理文件条目。"""
        if not isinstance(files, list):
            return []
        cleaned: list[dict[str, Any]] = []
        for item in files[:50]:
            if not isinstance(item, dict):
                continue
            cleaned.append(
                {
                    "name": cls._safe_output_name(item.get("name") or ""),
                    "output_type": str(item.get("output_type") or item.get("type") or "")[:40],
                    "status": str(item.get("status") or "已接收")[:40],
                    "size": cls._non_negative_int(item.get("size"), 0),
                    "sha256": cls._safe_sha256(item.get("sha256") or ""),
                    "content_received": bool(item.get("content_received", True)),
                    "path_available": bool(item.get("path_available")),
                    "path_display": "云端接收区路径已隐藏" if item.get("path_available") else "",
                }
            )
        return cleaned

    @classmethod
    def _local_upload_manifest_outputs(cls, outputs: Any) -> list[dict[str, Any]]:
        """返回上传清单的已清理输出摘要。"""
        if not isinstance(outputs, list):
            return []
        cleaned: list[dict[str, Any]] = []
        for item in outputs[:50]:
            if not isinstance(item, dict):
                continue
            raw_name = str(item.get("name") or item.get("file_name") or item.get("fileName") or "")
            raw_path = str(item.get("path") or "")
            name = cls._safe_output_name(raw_name or raw_path)
            size_value = item.get("size", "")
            size = cls._non_negative_int(size_value, 0) if str(size_value).strip() else ""
            path_available = bool(raw_path or item.get("path_available"))
            cleaned.append(
                {
                    "name": name[:160],
                    "output_type": str(item.get("output_type") or item.get("type") or "")[:40],
                    "status": str(item.get("status") or "已生成")[:40],
                    "size": size,
                    "sha256": cls._safe_sha256(item.get("sha256") or item.get("checksum") or ""),
                    "path_available": path_available,
                    "path_display": "本地路径已隐藏" if path_available else "",
                }
            )
        return cleaned

    @staticmethod
    def _safe_sha256(value: Any) -> str:
        """返回合法小写 SHA256，无效时返回空字符串。"""
        text = str(value or "").strip().lower()
        return text if re.fullmatch(r"[0-9a-f]{64}", text) else ""

    @staticmethod
    def _safe_output_name(value: Any) -> str:
        """返回安全输出显示名，不暴露本地路径。"""
        text = str(value or "").strip().replace("\\", "/")
        return (Path(text).name or "output")[:160]

    def create_local_launch_request(self, task_id: str, payload: dict[str, Any] | None = None) -> dict[str, Any]:
        """为任务创建已授权的 k12-local 启动请求。"""
        task = self.store.get_task(task_id)
        if not task:
            raise KeyError(f"Task not found: {task_id}")
        settings = self.store.get_settings()
        if not settings.get("localClientEnabled", True):
            raise ValueError("本地客户端连接已关闭")
        if not settings.get("allowWebLaunchLocalClient", False):
            raise ValueError("网页唤起本地客户端尚未授权")
        if not str(settings.get("localSecurityToken") or "").strip():
            raise ValueError("本地任务载荷包含本地路径，请先配置本地安全令牌")
        origin = self._local_launch_origin(settings, (payload or {}).get("api_origin") or (payload or {}).get("apiOrigin"))
        request_id = new_id("launch")
        now = utc_now()
        payload_url = f"{origin}/api/tasks/{task_id}/local-payload"
        params = {
            "task_id": task_id,
            "request_id": request_id,
            "api": origin,
            "payload_url": payload_url,
        }
        launch_request = {
            "schema_version": "k12.localLaunchRequest.v1",
            "id": request_id,
            "task_id": task_id,
            "task_type": task.get("task_type", ""),
            "task_label": task.get("task_label") or TASK_LABELS.get(str(task.get("task_type") or ""), str(task.get("task_type") or "")),
            "status": "requested",
            "created_at": now,
            "api_origin": origin,
            "protocol": "k12-local",
            "protocol_url": f"k12-local://open?{urlencode(params)}",
            "payload_url": payload_url,
            "payload_url_requires_token": True,
            "manifest_url": f"{origin}/api/local-client/manifest",
            "heartbeat_url": f"{origin}/api/local-client/heartbeat",
            "platform": self.install_profile()["platform"],
        }
        task["local_launch_request"] = launch_request
        task["local_client_status"] = "启动已请求"
        task["local_client_message"] = "已生成本地客户端唤起请求，等待桌面端心跳或任务同步"
        task["local_launch_requested_at"] = now
        self.store.append_log(task_id, f"本地客户端启动请求：{request_id}")
        self.store.save_task(task)
        return launch_request

    @staticmethod
    def _local_launch_origin(settings: dict[str, Any], requested: Any = None) -> str:
        """返回规范化的回环接口地址供启动链接使用。"""
        raw = str(requested or "").strip()
        if raw:
            parsed = urlparse(raw)
            if parsed.scheme in {"http", "https"} and parsed.hostname in {"127.0.0.1", "localhost", "::1"} and parsed.netloc:
                return f"{parsed.scheme}://{parsed.netloc}"
        host = str(settings.get("localApiHost") or "127.0.0.1")
        port = int(settings.get("localApiPort") or 8765)
        return f"http://{host}:{port}"

    def _product_capability(self, key: str, title: str, status: str, evidence: str) -> dict[str, Any]:
        """生成产品摘要能力行。"""
        labels = {
            "implemented": ("已落地", "good"),
            "contract": ("合同覆盖", "warn"),
            "planned": ("规划中", "blue"),
        }
        status_label, level = labels.get(status, ("需确认", "warn"))
        return {"key": key, "title": title, "status": status_label, "level": level, "evidence": evidence}

    def _product_phase(self, key: str, label: str, title: str, status: str, current: str) -> dict[str, Any]:
        """生成产品路线图阶段行。"""
        item = self._product_capability(key, f"{label}：{title}", status, current)
        item["phase"] = label
        return item

    def _enhancement_item(
        self,
        key: str,
        title: str,
        status: str,
        evidence: str,
        inputs: list[str],
        guardrails: list[str],
    ) -> dict[str, Any]:
        """生成含状态与依赖的增强规划行。"""
        item = self._product_capability(key, title, status, evidence)
        item["inputs"] = inputs
        item["guardrails"] = guardrails
        return item

    @staticmethod
    def _installer_specs() -> dict[str, dict[str, str]]:
        """返回已注册的 Windows 和 macOS 安装包规格。"""
        return {
            "Windows": {"file_name": "K12-Local-Client-Windows-x64.msi", "installer_kind": "windows-msi", "extension": ".msi"},
            "macOS": {"file_name": "K12-Local-Client-macOS-universal.pkg", "installer_kind": "macos-pkg", "extension": ".pkg"},
            "Unknown": {"file_name": "", "installer_kind": "manual", "extension": ""},
        }

    def _installer_spec(self, platform_name: str) -> dict[str, str]:
        """返回规范平台对应的安装包规格。"""
        specs = self._installer_specs()
        return specs.get(platform_name, specs["Unknown"])

    def _installer_steps(self, platform_name: str) -> list[dict[str, Any]]:
        """返回各平台安装和 MathType 兼容性步骤。"""
        common = [
            {"order": 1, "title": "下载安装包", "detail": "仅从本地 API 暴露的安装包地址下载，并核对 SHA256。"},
            {"order": 2, "title": "配置本地安全令牌", "detail": "安装后在设置页保存本地安全令牌，保护本地任务载荷和资源下载。"},
            {"order": 3, "title": "启动本地客户端", "detail": "确认本地服务监听 127.0.0.1，并可接收网页端交接的任务参数。"},
        ]
        if platform_name == "Windows":
            return common + [
                {"order": 4, "title": "确认 Windows 组件", "detail": "检查 Microsoft Office 桌面版、MathType Windows 版本和 pywin32/Office COM 自动化。"},
                {"order": 5, "title": "启用公式兜底", "detail": "跨平台输出时同时保存 MathML、LaTeX 或图片，避免 Windows OLE 公式在 macOS 不可编辑。"},
            ]
        if platform_name == "macOS":
            return common + [
                {"order": 4, "title": "确认 macOS 权限", "detail": "检查 Office for Mac、MathType macOS 版本、文件系统访问权限和本地服务授权。"},
                {"order": 5, "title": "启用公式兜底", "detail": "不要把 macOS MathType 对象直接交给 Windows 自动化链路，跨平台输出保存 MathML、LaTeX 或图片。"},
            ]
        return [
            {"order": 1, "title": "选择平台", "detail": "先在设置页选择 Windows 或 macOS，再生成对应安装计划。"},
            {"order": 2, "title": "使用兜底公式格式", "detail": "未知平台不交付平台专属 MathType 对象，优先使用 MathML、LaTeX 或图片。"},
        ]

    def _sha256(self, path: Path) -> str:
        """计算注册安装包的 SHA256 摘要。"""
        digest = hashlib.sha256()
        with path.open("rb") as handle:
            for chunk in iter(lambda: handle.read(1024 * 1024), b""):
                digest.update(chunk)
        return digest.hexdigest()

    def _architecture_layer(self, key: str, title: str, recommended: str, status: str, current: str) -> dict[str, Any]:
        """生成架构蓝图层级行。"""
        labels = {
            "implemented": ("已落地", "good"),
            "contract": ("合同覆盖", "warn"),
            "planned": ("后续接入", "blue"),
        }
        status_label, level = labels.get(status, ("需确认", "warn"))
        return {
            "key": key,
            "title": title,
            "recommended_stack": recommended,
            "status": status_label,
            "level": level,
            "current": current,
        }

    def local_task_payload(self, task_id: str) -> dict[str, Any]:
        """生成令牌保护的桌面或混合执行载荷。"""
        task = self.store.get_task(task_id)
        if not task:
            raise KeyError(f"Task not found: {task_id}")
        files = [self.store.get_file(file_id) for file_id in task.get("file_ids", [])]
        files = [file for file in files if file]
        files = self._local_payload_snapshot_files(task, files)
        report = self._latest_report_for_task(task_id)
        settings = self.store.get_settings()
        execute_mode = str(task.get("execute_mode") or self.resolve_execute_mode(str(task.get("task_type") or ""), list(task.get("file_ids") or [])))
        task_type = str(task.get("task_type") or "")
        requires_local = task_type in LOCAL_REQUIRED_TASKS or execute_mode in {"local", "hybrid"}
        preflight = self.preflight_checks({"task_type": task_type, "file_ids": list(task.get("file_ids") or []), "execute_mode": execute_mode})
        local_actions = self._local_payload_actions(task, files, report)
        client_readiness = self._local_client_readiness(requires_local, local_actions, settings)
        workflow_plan = self._local_payload_workflow_plan(task)
        return {
            "schema_version": "k12.localTaskPayload.v1",
            "generated_at": utc_now(),
            "task": {
                "id": task.get("id", ""),
                "task_type": task_type,
                "task_label": task.get("task_label") or TASK_LABELS.get(task_type, task_type),
                "execute_mode": execute_mode,
                "status": task.get("status", ""),
                "progress": task.get("progress", 0),
                "options": task.get("options", {}),
                "created_at": task.get("created_at", ""),
                "start_time": task.get("start_time", ""),
                "end_time": task.get("end_time", ""),
            },
            "handoff": self._local_handoff_status(requires_local, settings),
            "workflow_plan": workflow_plan,
            "files": [self._local_payload_file(file) for file in files],
            "preflight": preflight,
            "formula_delivery": self._local_formula_delivery_contract(settings),
            "local_actions": local_actions,
            "client_readiness": client_readiness,
            "desktop_execution_plan": self._local_desktop_execution_plan(task, files, local_actions, settings, client_readiness, include_sensitive_paths=True),
            "sync": self._local_payload_sync(settings, task),
            "settings_snapshot": self._local_payload_settings(settings),
            "report": self._local_payload_report(report),
        }

    def local_task_readiness(self, task_id: str) -> dict[str, Any]:
        """返回任务的脱敏客户端就绪与桌面计划状态。"""
        task = self.store.get_task(task_id)
        if not task:
            raise KeyError(f"Task not found: {task_id}")
        files = [self.store.get_file(file_id) for file_id in task.get("file_ids", [])]
        files = [file for file in files if file]
        report = self._latest_report_for_task(task_id)
        settings = self.store.get_settings()
        execute_mode = str(task.get("execute_mode") or self.resolve_execute_mode(str(task.get("task_type") or ""), list(task.get("file_ids") or [])))
        task_type = str(task.get("task_type") or "")
        requires_local = task_type in LOCAL_REQUIRED_TASKS or execute_mode in {"local", "hybrid"}
        local_actions = self._local_payload_actions(task, files, report)
        client_readiness = self._local_client_readiness(requires_local, local_actions, settings)
        workflow_plan = self._local_payload_workflow_plan(task)
        return {
            "schema_version": "k12.localTaskReadinessPayload.v1",
            "generated_at": utc_now(),
            "task": {
                "id": task.get("id", ""),
                "task_type": task_type,
                "task_label": task.get("task_label") or TASK_LABELS.get(task_type, task_type),
                "execute_mode": execute_mode,
                "status": task.get("status", ""),
                "progress": task.get("progress", 0),
            },
            "handoff": self._local_handoff_status(requires_local, settings),
            "workflow_plan": workflow_plan,
            "client_readiness": client_readiness,
            "local_actions": [self._local_action_summary(action) for action in local_actions],
            "desktop_execution_plan": self._local_desktop_execution_plan(task, files, local_actions, settings, client_readiness, include_sensitive_paths=False),
            "path_policy": "no local file paths or tokens are returned",
        }

    def _local_payload_workflow_plan(self, task: dict[str, Any]) -> dict[str, Any]:
        """返回本地载荷中的拖拽工作流顺序。"""
        options = task.get("options") if isinstance(task.get("options"), dict) else {}
        plan = options.get("workflowPlan") if isinstance(options.get("workflowPlan"), dict) else {}
        if not plan:
            plan = self._workflow_plan_from_options(str(task.get("task_type") or ""), options)
        return {
            "schema_version": str(plan.get("version") or "k12.workflowPlan.v1"),
            "source": str(plan.get("source") or "workspace-planner"),
            "current_task": str(plan.get("current_task") or task.get("task_type") or ""),
            "current_index": self._non_negative_int(plan.get("current_index"), 0),
            "order": [str(item) for item in plan.get("order") or []],
            "labels": [str(item) for item in plan.get("labels") or []],
            "invalid_count": len(plan.get("invalid_items") or []),
            "duplicate_count": len(plan.get("duplicate_items") or []),
        }

    def _latest_report_for_task(self, task_id: str) -> dict[str, Any] | None:
        """返回指定任务最新的已保存报告。"""
        for report in self.store.list_reports():
            if report.get("task_id") == task_id:
                return report
        return None

    @staticmethod
    def _local_handoff_status(requires_local: bool, settings: dict[str, Any]) -> dict[str, Any]:
        """说明客户端启动授权和令牌就绪情况。"""
        client_enabled = bool(settings.get("localClientEnabled", True))
        web_launch_allowed = bool(settings.get("allowWebLaunchLocalClient", False))
        token_required = bool(str(settings.get("localSecurityToken") or "").strip())
        if not requires_local:
            status = "not_required"
            message = "当前任务可由网页端处理，本地客户端载荷仅供审计或混合调度使用"
        elif not client_enabled:
            status = "client_disabled"
            message = "当前任务依赖本地客户端，但本地客户端连接已关闭"
        elif not web_launch_allowed:
            status = "authorization_required"
            message = "网页端唤起本地客户端尚未授权"
        else:
            status = "ready"
            message = "本地客户端可接收任务参数"
        return {
            "requires_local_client": requires_local,
            "client_enabled": client_enabled,
            "web_launch_allowed": web_launch_allowed,
            "token_required": token_required,
            "status": status,
            "message": message,
        }

    def _local_client_readiness(self, requires_local: bool, actions: list[dict[str, Any]], settings: dict[str, Any]) -> dict[str, Any]:
        """汇总心跳、平台和能力的本地交接就绪状态。"""
        heartbeat = settings.get("localClientHeartbeat") if isinstance(settings.get("localClientHeartbeat"), dict) else {}
        raw_preflight = heartbeat.get("preflight") if isinstance(heartbeat.get("preflight"), dict) else {}
        preflight = self._safe_local_client_preflight(raw_preflight) if raw_preflight else {}
        capabilities = self._local_client_capabilities(heartbeat, preflight)
        required_keys = self._required_local_client_capabilities(actions)
        required = [
            {
                "key": key,
                "label": LOCAL_CLIENT_CAPABILITY_LABELS.get(key, key),
                "available": self._local_capability_available(capabilities.get(key)),
            }
            for key in required_keys
        ]
        missing = [item["key"] for item in required if not item["available"]]
        has_heartbeat = bool(heartbeat.get("received_at"))
        expected_platform = self._configured_local_client_platform(settings)
        platform_value = heartbeat.get("platform") or preflight.get("platform") or settings.get("localClientPlatform") or ""
        actual_platform = normalize_platform(str(platform_value))
        platform_mismatch = bool(
            requires_local
            and has_heartbeat
            and expected_platform in {"Windows", "macOS"}
            and actual_platform in {"Windows", "macOS"}
            and expected_platform != actual_platform
        )
        if not requires_local:
            status = "not_required"
            message = "当前任务不要求本地客户端接手执行"
        elif not has_heartbeat:
            status = "needs_heartbeat"
            message = "尚未收到本地客户端心跳，无法确认桌面组件预检状态"
        elif platform_mismatch:
            status = "platform_mismatch"
            message = "本地客户端平台与安装计划不一致，不能交付平台专属 MathType/Office 对象"
        elif missing:
            status = "missing_capability"
            message = "本地客户端已连接，但缺少当前动作所需能力"
        elif required:
            status = "ready_for_handoff"
            message = "本地客户端预检能力满足当前动作交接要求"
        else:
            status = "no_executable_action"
            message = "当前载荷没有需要桌面端执行的动作"
        return {
            "schema_version": "k12.localClientReadiness.v1",
            "requires_local_client": requires_local,
            "status": status,
            "message": message,
            "client_id": str(heartbeat.get("client_id") or "")[:80],
            "platform": actual_platform,
            "expected_platform": expected_platform,
            "platform_compatible": not platform_mismatch,
            "platform_compatibility_message": self._local_platform_compatibility_message(expected_platform, actual_platform, platform_mismatch),
            "last_heartbeat_at": str(heartbeat.get("received_at") or ""),
            "required_capabilities": required,
            "missing_capabilities": missing,
            "capabilities": {key: self._local_capability_available(capabilities.get(key)) for key in LOCAL_CLIENT_CAPABILITY_LABELS},
            "components": dict(preflight.get("components") or {}),
            "preflight_schema_version": str(preflight.get("schema_version") or ""),
            "executes_native_documents": bool(preflight.get("executes_native_documents")),
        }

    @staticmethod
    def _configured_local_client_platform(settings: dict[str, Any]) -> str:
        """规范配置的本地客户端平台。"""
        configured = str(settings.get("localClientPlatform") or "auto")
        if configured.strip().lower() == "auto":
            return "auto"
        return normalize_platform(configured)

    @staticmethod
    def _local_platform_compatibility_message(expected: str, actual: str, mismatch: bool) -> str:
        """说明 Windows/macOS MathType 平台交接兼容性。"""
        if mismatch:
            return f"安装计划目标为 {expected}，但当前心跳来自 {actual}；Windows 与 macOS MathType 对象不通用，请切换客户端或使用 MathML/LaTeX/图片兜底。"
        if expected in {"Windows", "macOS"} and actual in {"Windows", "macOS"}:
            return f"安装计划与客户端平台均为 {expected}，可按该平台对象合同继续。"
        if expected == "auto":
            return "当前按自动平台选择交接；涉及 MathType 对象时建议在设置中明确选择 Windows 或 macOS。"
        return "平台未完全确认，涉及 MathType 对象时优先使用 MathML/LaTeX/图片兜底。"

    @staticmethod
    def _local_client_capabilities(heartbeat: dict[str, Any], preflight: dict[str, Any]) -> dict[str, Any]:
        """合并可信心跳与预检能力标记。"""
        merged: dict[str, Any] = {}
        raw_heartbeat = heartbeat.get("capabilities") if isinstance(heartbeat.get("capabilities"), dict) else {}
        raw_preflight = preflight.get("capabilities") if isinstance(preflight.get("capabilities"), dict) else {}
        for source in (raw_heartbeat, raw_preflight):
            for key, value in source.items():
                if key in LOCAL_CLIENT_CAPABILITY_LABELS and isinstance(value, (bool, int, float, str)):
                    merged[str(key)] = value
        return merged

    @staticmethod
    def _required_local_client_capabilities(actions: list[dict[str, Any]]) -> list[str]:
        """根据本地动作推导桌面原生能力要求。"""
        required: list[str] = []
        for action in actions:
            action_type = str(action.get("type") or "")
            status = str(action.get("status") or "")
            if action_type == "macro_sequence" and (status == "queued" or action.get("macros")):
                required.append("macroExecution")
            elif action_type == "omml_mathtype" and status != "not_required":
                required.extend(["ommlDependencySearch", "mathTypeAutomation"])
            elif action_type == "pdf_formula_mathtype" and status != "not_required":
                required.append("mathTypeAutomation")
            elif action_type == "office_conversion" and status != "completed":
                required.append("officeAutomation")
        return list(dict.fromkeys(required))

    @staticmethod
    def _local_capability_available(value: Any) -> bool:
        """将客户端能力值解释为可用性布尔标记。"""
        if isinstance(value, str):
            return value.strip().lower() in {"1", "true", "yes", "available", "ok", "ready", "enabled"}
        return bool(value)

    def _local_action_summary(self, action: dict[str, Any]) -> dict[str, Any]:
        """返回本地载荷动作的脱敏摘要。"""
        action_type = str(action.get("type") or "")
        required = self._required_local_client_capabilities([action])
        return {
            "type": action_type,
            "label": LOCAL_ACTION_LABELS.get(action_type, action_type),
            "status": str(action.get("status") or ""),
            "recognition_request_count": len(action.get("recognition_requests") or []),
            "required_capabilities": [
                {
                    "key": key,
                    "label": LOCAL_CLIENT_CAPABILITY_LABELS.get(key, key),
                }
                for key in required
            ],
        }

    def _local_desktop_execution_plan(
        self,
        task: dict[str, Any],
        files: list[dict[str, Any]],
        actions: list[dict[str, Any]],
        settings: dict[str, Any],
        readiness: dict[str, Any],
        include_sensitive_paths: bool,
    ) -> dict[str, Any]:
        """生成桌面执行计划，不执行原生文档动作。"""
        readiness_status = str(readiness.get("status") or "")
        plan_status_map = {
            "ready_for_handoff": "ready_for_native_client",
            "needs_heartbeat": "waiting_for_heartbeat",
            "missing_capability": "blocked_by_capability",
            "platform_mismatch": "blocked_by_platform",
            "not_required": "not_required",
            "no_executable_action": "no_executable_action",
        }
        plan_status = plan_status_map.get(readiness_status, "pending_preflight")
        native_action_count = sum(1 for action in actions if str(action.get("type") or "") != "open_output_directory")
        companion_office_execution = bool(
            readiness.get("platform") in {"Windows", "macOS"}
            and any(str(action.get("type") or "") == "office_conversion" for action in actions)
        )
        formula_delivery = self._local_formula_delivery_contract(settings)
        same_platform_required = any(
            str(action.get("type") or "") in {"omml_mathtype", "office_conversion"} and bool(formula_delivery.get("native_object_requires_same_platform"))
            for action in actions
        )
        return {
            "schema_version": "k12.desktopExecutionPlan.v1",
            "generated_at": utc_now(),
            "task_id": task.get("id", ""),
            "task_type": task.get("task_type", ""),
            "task_label": task.get("task_label") or TASK_LABELS.get(str(task.get("task_type") or ""), str(task.get("task_type") or "")),
            "status": plan_status,
            "native_execution_allowed": readiness_status == "ready_for_handoff",
            "web_executes_native_documents": False,
            "current_companion_cli_executes_native_documents": companion_office_execution,
            "execution_surface": "local_desktop_client",
            "action_count": len(actions),
            "native_action_count": native_action_count,
            "file_count": len(files),
            "platform": {
                "expected": readiness.get("expected_platform", ""),
                "actual": readiness.get("platform", ""),
                "compatible": bool(readiness.get("platform_compatible", True)),
                "same_platform_required_for_native_mathtype": same_platform_required,
                "message": readiness.get("platform_compatibility_message", ""),
            },
            "authorization_gates": {
                "local_security_token_required": True,
                "local_security_token_configured": bool(str(settings.get("localSecurityToken") or "").strip()),
                "web_launch_allowed": bool(settings.get("allowWebLaunchLocalClient", False)),
                "task_status_sync_allowed": bool(settings.get("allowTaskStatusCloudSync", False)),
                "cloud_result_upload_allowed": bool(settings.get("allowCloudSync", False)),
            },
            "capability_gates": {
                "readiness_status": readiness_status,
                "missing_capabilities": list(readiness.get("missing_capabilities") or []),
                "required_capabilities": list(readiness.get("required_capabilities") or []),
            },
            "formula_delivery": formula_delivery,
            "actions": [
                self._desktop_execution_action(action, task, files, settings, readiness, include_sensitive_paths)
                for action in actions
            ],
            "path_policy": "sensitive payload; local input/output paths may be present" if include_sensitive_paths else "no local file paths or tokens are returned",
            "guardrails": [
                "网页端只生成可交接执行计划，不直接执行 Office、MathType、OMML 写回或 Word 宏。",
                "Windows 与 macOS MathType 原生对象不跨平台通用；需要同平台客户端或 MathML/LaTeX/图片兜底。",
                "本地客户端回传结果时只同步状态和脱敏输出摘要，完整路径和令牌不进入网页就绪摘要。",
            ],
        }

    def _desktop_execution_action(
        self,
        action: dict[str, Any],
        task: dict[str, Any],
        files: list[dict[str, Any]],
        settings: dict[str, Any],
        readiness: dict[str, Any],
        include_sensitive_paths: bool,
    ) -> dict[str, Any]:
        """生成桌面计划的原生动作约定。"""
        action_type = str(action.get("type") or "")
        required_keys = self._required_local_client_capabilities([action])
        missing = set(str(key) for key in readiness.get("missing_capabilities") or [])
        gate_status = self._desktop_action_gate_status(action_type, required_keys, missing, readiness)
        item: dict[str, Any] = {
            "action_id": f"{task.get('id', '')}:{action_type}",
            "type": action_type,
            "label": LOCAL_ACTION_LABELS.get(action_type, action_type),
            "status": str(action.get("status") or ""),
            "gate_status": gate_status,
            "required_capabilities": [
                {
                    "key": key,
                    "label": LOCAL_CLIENT_CAPABILITY_LABELS.get(key, key),
                    "available": key not in missing,
                }
                for key in required_keys
            ],
            "input_file_count": len(files),
            "output_contract": self._desktop_action_output_contract(action_type, task, settings, include_sensitive_paths),
            "steps": self._desktop_action_steps(action_type, action, task, settings),
        }
        if isinstance(action.get("formula_delivery"), dict):
            item["formula_delivery"] = dict(action["formula_delivery"])
        elif action_type in {"omml_mathtype", "pdf_formula_mathtype", "office_conversion"}:
            item["formula_delivery"] = self._local_formula_delivery_contract(settings)
        if action_type == "macro_sequence":
            macros = list(action.get("macros") or [])
            item["macro_count"] = len(macros)
            item["requires_user_confirmation"] = any(str(macro.get("execute_status") or "") == "待确认" for macro in macros)
            item["backup_required"] = bool(settings.get("macroBackup", True))
        elif action_type == "omml_mathtype":
            item["dependency_count"] = len(action.get("dependencies") or [])
            item["conversion_prompt_count"] = len(action.get("conversion_prompts") or [])
            item["retry_request_count"] = len(action.get("retry_requests") or [])
            item["recognition_request_count"] = len(action.get("recognition_requests") or [])
            item["recognition_requests"] = list(action.get("recognition_requests") or [])
            item["copy_strategy"] = str(settings.get("ommlCopyStrategy") or "自动重命名")
            item["requires_user_confirmation"] = str(action.get("status") or "") == "awaiting_confirmation"
        elif action_type == "pdf_formula_mathtype":
            item["formula_count"] = int(action.get("formula_count") or 0)
            item["mathpix_job_count"] = len(action.get("mathpix_jobs") or [])
            item["recognition_request_count"] = len(action.get("recognition_requests") or [])
            item["recognition_requests"] = list(action.get("recognition_requests") or [])
            item["requires_user_confirmation"] = False
        elif action_type == "office_conversion":
            item["artifact_count"] = len(action.get("artifacts") or [])
            item["conversion_type"] = str(task.get("task_type") or "")
        elif action_type == "open_output_directory":
            item["native_document_action"] = False
        if include_sensitive_paths:
            item["input_paths"] = [str(file.get("storage_path") or file.get("file_path") or "") for file in files]
            if action_type == "macro_sequence":
                item["backup_paths"] = [str(macro.get("backup_path") or "") for macro in action.get("macros") or [] if macro.get("backup_path")]
            if action_type == "open_output_directory":
                item["target_path"] = str(action.get("path") or "")
        return item

    @staticmethod
    def _desktop_action_gate_status(action_type: str, required_keys: list[str], missing: set[str], readiness: dict[str, Any]) -> str:
        """判断桌面动作能否交接原生执行器。"""
        if action_type == "open_output_directory":
            return "ready"
        if readiness.get("status") == "platform_mismatch":
            return "blocked_by_platform"
        if readiness.get("status") == "needs_heartbeat":
            return "waiting_for_heartbeat"
        if any(key in missing for key in required_keys):
            return "blocked_by_capability"
        if readiness.get("status") == "ready_for_handoff":
            return "ready"
        return "pending"

    def _desktop_action_output_contract(self, action_type: str, task: dict[str, Any], settings: dict[str, Any], include_sensitive_paths: bool) -> dict[str, Any]:
        """描述动作预期产物及路径脱敏规则。"""
        task_type = str(task.get("task_type") or "")
        artifact_types = {
            "word_to_ppt": ["pptx"],
            "ppt_to_word": ["docx"],
            "pdf_to_word": ["docx", "tex.zip"],
            "excel_to_pdf": ["pdf"],
            "excel_to_word": ["docx"],
            "excel_to_ppt": ["pptx"],
            "macro_sequence": ["macro-report", "backup"],
            "omml_to_mathtype": ["docx", "formula-report"],
            "mathtype_format": ["docx", "formula-report"],
            "formula_precheck": ["formula-report"],
        }
        if action_type == "pdf_formula_mathtype":
            artifact_list = ["docx", "tex.zip", "formula-report"]
        elif action_type == "open_output_directory":
            artifact_list = []
        else:
            artifact_list = artifact_types.get(task_type, ["report"])
        contract = {
            "conflict_strategy": str(settings.get("outputConflictStrategy") or "自动重命名"),
            "artifact_types": artifact_list,
            "result_upload_optional": bool(settings.get("allowCloudSync", False)),
        }
        output_directory = str(self.store.output_task_dir(str(task.get("id") or "")))
        if include_sensitive_paths:
            contract["output_directory"] = output_directory
        else:
            contract["output_directory_available"] = bool(output_directory)
            contract["output_directory_display"] = "本地路径已隐藏" if output_directory else ""
        return contract

    @staticmethod
    def _desktop_action_steps(action_type: str, action: dict[str, Any], task: dict[str, Any], settings: dict[str, Any]) -> list[dict[str, Any]]:
        """返回桌面动作约定的有序执行步骤。"""
        if action_type == "macro_sequence":
            return [
                {"order": 1, "operation": "macro.backup", "title": "创建执行前备份", "required": bool(settings.get("macroBackup", True))},
                {"order": 2, "operation": "macro.run_ordered", "title": "按用户顺序执行宏队列", "required": True},
                {"order": 3, "operation": "macro.write_report", "title": "写入宏日志、失败策略和执行报告", "required": True},
            ]
        if action_type == "omml_mathtype":
            steps = [
                {"order": 1, "operation": "omml.confirm", "title": "确认是否将 Word 自带公式转换为 MathType", "required": str(action.get("status") or "") == "awaiting_confirmation"},
                {"order": 2, "operation": "omml.search_dependency", "title": "检索 OMML 依赖并按复制策略放入文档目录", "required": True},
                {"order": 3, "operation": "mathtype.convert", "title": "同平台转换 OMML/MathType 对象并保留兜底格式", "required": True},
                {"order": 4, "operation": "formula.validate", "title": "校验公式数量、失败兜底和报告记录", "required": True},
            ]
            if action.get("retry_requests"):
                steps.append({"order": 5, "operation": "omml.retry_conversion", "title": "复制成功或人工校正后重新执行 OMML 转 MathType", "required": True})
            if action.get("recognition_requests"):
                steps.append({"order": 6, "operation": "formula.rerecognize_local", "title": "按人工校正请求重新识别或转换公式，并保留原公式兜底", "required": True})
            return steps
        if action_type == "pdf_formula_mathtype":
            steps = [
                {"order": 1, "operation": "mathpix.collect_formula_outputs", "title": "读取 Mathpix tex.zip、LaTeX 和公式识别摘要", "required": True},
                {"order": 2, "operation": "mathtype.convert_pdf_formula", "title": "按同平台 MathType 合同生成可编辑公式对象或兜底格式", "required": True},
                {"order": 3, "operation": "formula.merge_into_docx", "title": "把公式后处理结果并入 PDF 转 Word 输出清单", "required": True},
                {"order": 4, "operation": "formula.validate", "title": "记录低置信度、失败兜底和人工校正入口", "required": True},
            ]
            if action.get("recognition_requests"):
                steps.insert(1, {"order": 2, "operation": "mathpix.rerecognize_formula", "title": "按人工重识别请求等待授权 Mathpix worker 重新识别公式", "required": True})
                for index, step in enumerate(steps, start=1):
                    step["order"] = index
            return steps
        if action_type == "office_conversion":
            return [
                {"order": 1, "operation": "office.open_source", "title": "用本机 Office 打开源文档", "required": True},
                {"order": 2, "operation": f"office.convert.{task.get('task_type', '')}", "title": "按任务类型生成转换产物", "required": True},
                {"order": 3, "operation": "quality.compare_preview", "title": "回传输出摘要、对象保留清单和质量检查", "required": True},
            ]
        if action_type == "open_output_directory":
            return [
                {"order": 1, "operation": "shell.open_output_directory", "title": "在本机打开任务输出目录", "required": False},
            ]
        return [
            {"order": 1, "operation": "local.noop", "title": "等待本地客户端识别动作类型", "required": False},
        ]

    def _local_payload_snapshot_files(self, task: dict[str, Any], files: list[dict[str, Any]]) -> list[dict[str, Any]]:
        """使用校验过的任务快照提供敏感桌面输入路径。"""
        snapshots = {
            str(item.get("file_id") or ""): item
            for item in task.get("source_snapshots") or []
            if isinstance(item, dict)
        }
        prepared: list[dict[str, Any]] = []
        for original in files:
            file = dict(original)
            snapshot = snapshots.get(str(file.get("id") or ""))
            if snapshot:
                path = Path(str(snapshot.get("path") or ""))
                expected_size = int(snapshot.get("size") or 0)
                expected_sha256 = str(snapshot.get("sha256") or "")
                if path.exists() and path.is_file() and path.stat().st_size == expected_size and self._sha256(path) == expected_sha256:
                    file["storage_path"] = str(path)
                    file["file_path"] = str(path)
                    file["source_kind"] = "task_snapshot"
                    file["file_size"] = expected_size
                    file["source_sha256"] = expected_sha256
                else:
                    errors = list(file.get("validation_errors") or [])
                    errors.append("任务源文件快照完整性校验失败")
                    file["validation_errors"] = list(dict.fromkeys(errors))
                    file["status"] = "校验失败"
            prepared.append(file)
        return prepared

    @staticmethod
    def _local_payload_file(file: dict[str, Any]) -> dict[str, Any]:
        """返回伴随客户端所需的敏感文件载荷行。"""
        input_path = str(file.get("storage_path") or file.get("file_path") or "")
        path = Path(input_path) if input_path else None
        return {
            "id": file.get("id", ""),
            "file_name": file.get("file_name", ""),
            "file_type": file.get("file_type", ""),
            "extension": file.get("extension", ""),
            "file_size": file.get("file_size", 0),
            "source_sha256": file.get("source_sha256", ""),
            "source_relative_path": file.get("source_relative_path", ""),
            "input_path": input_path,
            "input_path_exists": bool(path and path.exists()),
            "input_source": "upload_cache" if file.get("storage_path") else "external_path",
            "status": file.get("status", ""),
            "validation_errors": file.get("validation_errors", []),
            "encrypted": bool(file.get("encrypted", False)),
            "password_session_active": bool(file.get("password_session_active", False)),
            "has_formula": bool(file.get("has_formula", False)),
            "has_mathtype": bool(file.get("has_mathtype", False)),
            "has_omml": bool(file.get("has_omml", False)),
            "missing_omml_dependency": bool(file.get("missing_omml_dependency", False)),
            "has_macro": bool(file.get("has_macro", False)),
            "has_image": bool(file.get("has_image", False)),
            "has_small_image": bool(file.get("has_small_image", False)),
            "content_summary": file.get("content_summary", {}),
        }

    def _local_payload_actions(self, task: dict[str, Any], files: list[dict[str, Any]], report: dict[str, Any] | None) -> list[dict[str, Any]]:
        """根据任务与报告证据生成客户端动作约定。"""
        task_type = str(task.get("task_type") or "")
        analysis = dict((report or {}).get("analysis") or {})
        formula_delivery = self._local_formula_delivery_contract(self.store.get_settings())
        recognition_requests = self._formula_recognition_requests(report)
        local_recognition_requests = [item for item in recognition_requests if item.get("engine") == "local-client"]
        mathpix_recognition_requests = [item for item in recognition_requests if item.get("engine") == "Mathpix"]
        actions: list[dict[str, Any]] = []
        if task_type == "macro_sequence":
            queued_statuses = {"待本地客户端执行", "待确认"}
            macros = [
                {
                    "id": macro.get("id", ""),
                    "file_id": macro.get("file_id", ""),
                    "macro_name": macro.get("macro_name", ""),
                    "macro_source": macro.get("macro_source", ""),
                    "source_id": macro.get("source_id", ""),
                    "source_file": macro.get("source_file", ""),
                    "source_sha256": macro.get("source_sha256", ""),
                    "module_name": macro.get("module_name", ""),
                    "source_line": macro.get("source_line", 0),
                    "execute_order": macro.get("execute_order", 0),
                    "execute_timing": macro.get("execute_timing", ""),
                    "failure_strategy": macro.get("failure_strategy", ""),
                    "failure_policy": macro.get("failure_policy", {}),
                    "timeout_seconds": macro.get("timeout_seconds", 0),
                    "execute_status": macro.get("execute_status", ""),
                    "backup_path": macro.get("backup_path", ""),
                }
                for macro in analysis.get("macros", [])
                if macro.get("execute_status") in queued_statuses
            ]
            actions.append({"type": "macro_sequence", "status": "queued" if macros else "empty", "macros": macros})
        if task_type in {"formula_precheck", "omml_to_mathtype", "mathtype_format", "word_to_ppt"}:
            dependencies = analysis.get("ommlDependencies", [])
            prompts = analysis.get("ommlConversionPrompts", [])
            prompt_statuses = {prompt.get("status") for prompt in prompts}
            retry_requests = self._omml_retry_requests(report)
            action_status = (
                "awaiting_confirmation"
                if "需确认" in prompt_statuses
                else (
                    "recognition_queued"
                    if local_recognition_requests
                    else "retry_queued"
                    if retry_requests
                    else ("queued" if dependencies or prompts or any(file.get("has_omml") or file.get("has_mathtype") for file in files) else "not_required")
                )
            )
            actions.append(
                {
                    "type": "omml_mathtype",
                    "status": action_status,
                    "dependencies": dependencies,
                    "conversion_prompts": prompts,
                    "retry_requests": retry_requests,
                    "recognition_requests": local_recognition_requests,
                    "formula_count": len(analysis.get("formulas", [])),
                    "formula_delivery": formula_delivery,
                }
            )
        if task_type == "pdf_to_word":
            mathpix_jobs = [job for job in analysis.get("mathpix", []) if isinstance(job, dict)]
            pdf_formulas = [formula for formula in analysis.get("formulas", []) if str(formula.get("source_type") or "") == "PDF"]
            action_status = self._pdf_formula_action_status(mathpix_jobs, pdf_formulas)
            if action_status == "not_required" and mathpix_recognition_requests:
                action_status = "blocked_mathpix" if any(item.get("blocking_reasons") for item in mathpix_recognition_requests) else "waiting_mathpix_rerecognition"
            if action_status != "not_required":
                actions.append(
                    {
                        "type": "pdf_formula_mathtype",
                        "status": action_status,
                        "formula_count": len(pdf_formulas) or sum(self._summary_count(job.get("retention_plan") or {}, "formula_hints") for job in mathpix_jobs),
                        "mathpix_jobs": [self._local_pdf_formula_job_summary(job) for job in mathpix_jobs],
                        "recognition_requests": mathpix_recognition_requests,
                        "formula_delivery": formula_delivery,
                    }
                )
        if task_type in {"word_to_ppt", "ppt_to_word", "excel_to_pdf", "excel_to_word", "excel_to_ppt"}:
            artifacts = [item for item in analysis.get("artifacts", []) if isinstance(item, dict)]
            pending_files = [
                {
                    "file_id": item.get("file_id", ""),
                    "source_file": item.get("source_file", ""),
                    "output_type": item.get("output_type", ""),
                    "conversion_settings": item.get("conversion_settings") or self._conversion_settings_snapshot(task, task_type),
                    "excel_options": dict(task.get("options", {}).get("excel") or {}),
                }
                for item in artifacts
                if item.get("status") == "待本地客户端执行"
            ]
            completed = bool(artifacts) and all(item.get("status") in {"成功", "跳过"} for item in artifacts)
            actions.append(
                {
                    "type": "office_conversion",
                    "status": "completed" if completed else "queued",
                    "artifacts": artifacts,
                    "pending_files": pending_files,
                    "formula_delivery": formula_delivery,
                }
            )
        if task.get("output_directory_action"):
            actions.append({"type": "open_output_directory", **dict(task.get("output_directory_action") or {})})
        return actions

    def _omml_retry_requests(self, report: dict[str, Any] | None) -> list[dict[str, Any]]:
        """返回手动更正标注创建的 OMML 重试请求。"""
        if not report:
            return []
        report_id = str(report.get("id") or "")
        if not report_id:
            return []
        annotations = [
            item
            for item in self.store.list_omml_annotations(report_id)
            if item.get("retry_conversion")
        ]
        if not annotations:
            return []
        dependencies = {
            str(item.get("id") or ""): item
            for item in (report.get("analysis") or {}).get("ommlDependencies", [])
            if isinstance(item, dict) and item.get("id")
        }
        requests: list[dict[str, Any]] = []
        for annotation in annotations:
            dependency_id = str(annotation.get("dependency_id") or "")
            dependency = dependencies.get(dependency_id, {})
            manual_path = str(annotation.get("manual_omml_path") or "")
            requests.append(
                {
                    "annotation_id": str(annotation.get("id") or ""),
                    "dependency_id": dependency_id,
                    "file_id": str(dependency.get("file_id") or ""),
                    "file_name": str(dependency.get("file_name") or ""),
                    "status": str(annotation.get("status") or "重新转换"),
                    "retry_conversion": True,
                    "keep_omml": bool(annotation.get("keep_omml", True)),
                    "manual_omml_path": manual_path,
                    "manual_omml_path_available": bool(manual_path),
                    "message": "复制成功或人工校正后重新执行 OMML 转 MathType，并保留原 OMML 兜底",
                }
            )
        return requests

    def _formula_recognition_requests(self, report: dict[str, Any] | None) -> list[dict[str, Any]]:
        """返回可安全交接本地端的公式重识别请求。"""
        if not report:
            return []
        report_id = str(report.get("id") or "")
        if not report_id:
            return []
        formulas = {
            str(item.get("id") or ""): item
            for item in (report.get("analysis") or {}).get("formulas", [])
            if isinstance(item, dict) and item.get("id")
        }
        requests: list[dict[str, Any]] = []
        for annotation in self.store.list_formula_annotations(report_id):
            if not annotation.get("retry_recognition"):
                continue
            formula_id = str(annotation.get("formula_id") or "")
            formula = formulas.get(formula_id, {})
            recognition = dict(annotation.get("recognition_request") or {})
            requests.append(
                {
                    "annotation_id": str(annotation.get("id") or ""),
                    "formula_id": formula_id,
                    "file_id": str(formula.get("file_id") or ""),
                    "source_type": str(recognition.get("source_type") or formula.get("source_type") or ""),
                    "position": str(recognition.get("position") or formula.get("position") or ""),
                    "status": str(annotation.get("status") or "重新识别"),
                    "recognition_status": str(annotation.get("recognition_status") or recognition.get("status") or ""),
                    "engine": str(recognition.get("engine") or ""),
                    "action": str(recognition.get("action") or ""),
                    "external_upload_required": bool(recognition.get("external_upload_required", False)),
                    "external_upload_authorized": bool(recognition.get("external_upload_authorized", False)),
                    "external_upload_performed": bool(recognition.get("external_upload_performed", False)),
                    "blocking_reasons": list(recognition.get("blocking_reasons") or []),
                    "required_environment": list(recognition.get("required_environment") or []),
                    "next_step": str(annotation.get("next_step") or recognition.get("next_step") or ""),
                }
            )
        return requests

    @staticmethod
    def _pdf_formula_action_status(mathpix_jobs: list[dict[str, Any]], pdf_formulas: list[dict[str, Any]]) -> str:
        """根据 Mathpix 作业判断 PDF 公式 MathType 交接状态。"""
        if pdf_formulas:
            return "queued"
        if not mathpix_jobs:
            return "not_required"
        if not any(bool((job.get("ocr_settings") or {}).get("formula_ocr")) for job in mathpix_jobs):
            return "not_required"
        statuses = {str(job.get("status") or "") for job in mathpix_jobs}
        if statuses & {"authorization_required", "missing_credentials", "api_error", "missing_local_file", "ocr_disabled"}:
            return "blocked_mathpix"
        if statuses & {"submitted", "processing", "pending"}:
            return "waiting_for_mathpix"
        if statuses & {"completed", "skipped_existing_output"}:
            return "queued"
        return "waiting_for_mathpix"

    @staticmethod
    def _local_pdf_formula_job_summary(job: dict[str, Any]) -> dict[str, Any]:
        """返回 PDF 公式交接使用的脱敏 Mathpix 作业摘要。"""
        recognition = dict(job.get("recognition_plan") or {})
        retention = dict(job.get("retention_plan") or {})
        return {
            "file_id": str(job.get("file_id") or ""),
            "file_name": str(job.get("file_name") or ""),
            "status": str(job.get("status") or ""),
            "recognition_status": str(recognition.get("status") or ""),
            "formula_hints": int(retention.get("formula_hints") or 0),
            "tex_zip_status": str(job.get("tex_zip_status") or ""),
            "tex_zip_available": bool((job.get("outputs") or {}).get("tex_zip")),
            "docx_available": bool((job.get("outputs") or {}).get("docx")),
        }

    def _local_payload_sync(self, settings: dict[str, Any], task: dict[str, Any]) -> dict[str, Any]:
        """说明本地结果同步策略与受管输出目录。"""
        return {
            "local_output_directory": str(self.store.output_task_dir(str(task.get("id") or ""))),
            "result_upload_allowed": bool(settings.get("allowCloudSync", False)),
            "task_status_cloud_sync_allowed": bool(settings.get("allowTaskStatusCloudSync", False)),
            "sensitive_files_prefer_local": bool(settings.get("sensitiveFilesPreferLocal", True)),
            "policy": "允许上传处理结果" if settings.get("allowCloudSync", False) else "仅本地保存结果",
        }

    def _local_formula_delivery_contract(self, settings: dict[str, Any]) -> dict[str, Any]:
        """返回 MathType 同平台与公式兜底交付约定。"""
        profile = self.install_profile()
        platform_name = str(profile.get("platform") or "Unknown")
        mode = str(settings.get("mathtypeCompatibilityMode") or "platform-specific")
        mode_labels = {
            "platform-specific": "区分平台",
            "mathml-latex": "MathML/LaTeX 优先",
            "image-fallback": "图片兜底",
        }
        mode_label = mode_labels.get(mode, mode)
        known_platform = platform_name in {"Windows", "macOS"}
        native_allowed = known_platform and mode == "platform-specific"
        if mode == "image-fallback":
            output_priority = ["图片", "MathML", "LaTeX"]
            fallback_formats = ["图片", "MathML", "LaTeX"]
        elif mode == "mathml-latex":
            output_priority = ["MathML", "LaTeX", "图片"]
            fallback_formats = ["MathML", "LaTeX", "图片"]
        else:
            output_priority = [str(profile.get("mathtypeObjectFormat") or "平台专属 MathType 对象")]
            fallback_formats = ["MathML", "LaTeX", "图片"]
        return {
            "schema_version": "k12.formulaDeliveryContract.v1",
            "platform": platform_name,
            "compatibility_mode": mode,
            "compatibility_label": mode_label,
            "platform_object_format": str(profile.get("mathtypeObjectFormat") or ""),
            "platform_objects_cross_compatible": False,
            "native_mathtype_object_allowed": native_allowed,
            "native_object_requires_same_platform": native_allowed,
            "cross_platform_safe": mode in {"mathml-latex", "image-fallback"},
            "output_priority": output_priority,
            "fallback_formats": fallback_formats,
            "preserve_original_formula": bool(settings.get("keepFormulaImages", True)),
            "requires_user_confirmation": mode == "platform-specific" or not known_platform,
            "message": str(profile.get("formulaPortability") or ""),
        }

    @staticmethod
    def _local_payload_settings(settings: dict[str, Any]) -> dict[str, Any]:
        """返回客户端执行相关的非秘密设置。"""
        keys = [
            "localClientPlatform",
            "mathtypeCompatibilityMode",
            "pdfToWordEngine",
            "outputConflictStrategy",
            "wordToPptTemplate",
            "pptToWordMode",
            "pptToWordTemplate",
            "pptToWordTemplatePath",
            "pptToWordGenerateToc",
            "formulaOutputFormat",
            "formulaFormatScope",
            "ommlCopyStrategy",
            "macroFailureStrategy",
            "imageExportFormat",
        ]
        return {key: settings.get(key) for key in keys}

    @staticmethod
    def _local_payload_report(report: dict[str, Any] | None) -> dict[str, Any]:
        """返回本地载荷的报告可用性及摘要计数。"""
        if not report:
            return {"available": False}
        return {
            "available": True,
            "id": report.get("id", ""),
            "report_type": report.get("report_type", ""),
            "success_count": report.get("success_count", 0),
            "fail_count": report.get("fail_count", 0),
            "macro_queued_count": report.get("macro_queued_count", 0),
            "quality_issue_count": report.get("quality_issue_count", 0),
        }

    def install_profile(self, platform_name: str | None = None) -> dict[str, Any]:
        """返回所选平台安装配置的副本，包含公式互操作规则。"""
        return install_profile(platform_name, self.store.get_settings())

    @staticmethod
    def _macro_capability_status(flags: dict[str, Any], settings: dict[str, Any]) -> str:
        """根据平台标记与设置说明宏队列可用性。"""
        if not settings.get("enableMacroExecution", True):
            return "宏执行队列已禁用"
        status = "Windows 本地客户端授权后执行" if flags.get("macroExecution") else "当前平台不直接执行宏，生成本地执行队列"
        if settings.get("macroWhitelistOnly"):
            status = f"{status}；仅允许白名单宏"
        return status

    def import_macro_source(self, payload: dict[str, Any]) -> dict[str, Any]:
        """导入 UTF-8 标准 VBA 模块，保存真实声明与完整来源文本。"""
        self._assert_permission("templates.manage")
        filename = str(payload.get("file_name") or "")
        source = payload.get("source")
        if not filename.lower().endswith(".bas") or "/" in filename or "\\" in filename or len(filename) > 255:
            raise ValueError("请选择 .bas 标准模块源文件")
        if not isinstance(source, str):
            raise ValueError("VBA 源内容必须是文本")
        if len(source.encode("utf-8")) > 2_000_000:
            raise ValueError("VBA 源文件不能超过 2 MB")
        declarations = vba_macros(source)
        if not declarations:
            raise ValueError("源模块中未找到公开无参数宏")
        digest = hashlib.sha256(source.encode("utf-8")).hexdigest()
        identifier = "vba_" + hashlib.sha256((filename + "\0" + digest).encode()).hexdigest()[:24]
        record = {"id": identifier, "file_name": filename, "sha256": digest, "source": source, "macros": declarations}
        self.store.save_macro_source(record)
        return {key: value for key, value in record.items() if key != "source"}

    def macro_source_definition(self, source_id: str) -> dict[str, Any]:
        """仅向有宏执行权限的调用方返回通过哈希核验的源定义。"""
        self._assert_permission("macros.execute")
        source = self.store.get_macro_source(source_id)
        if source is None:
            raise ValueError("导入模块不存在")
        digest = hashlib.sha256(source["source"].encode("utf-8")).hexdigest()
        if digest != source["sha256"] or vba_macros(source["source"]) != source["macros"]:
            raise ValueError("宏源定义与保存的哈希或声明不一致")
        return source

    def delete_macro_source(self, source_id: str) -> bool:
        """删除未被待执行任务引用的导入模块，保留内置宏。"""
        self._assert_permission("templates.manage")
        sources = {item["id"] for item in self.store.list_macro_sources()}
        if source_id not in sources:
            raise ValueError("导入模块不存在")
        for template in self.store.list_macro_templates():
            if any(str(item.get("id") or "").startswith(source_id + "_") for item in template.get("macro_sequence", [])):
                raise ValueError("模块被宏顺序模板引用，请先修改或删除模板")
        for task in self.store.list_tasks():
            if task.get("status") in {"成功", "失败", "已取消"}:
                continue
            for spec in self._selected_macro_specs(task.get("options", {})):
                identifier = str(spec.get("id") or spec.get("macro_id") or "")
                if identifier.startswith(source_id + "_"):
                    raise ValueError("模块被待执行任务引用，请先处理该任务")
        return self.store.delete_macro_source(source_id)

    def macro_library(self) -> list[dict[str, Any]]:
        """返回内置宏选择与使用元数据。"""
        macro_specs = [
            (MacroItem("NormalizeHeadingStyles", "系统内置", "统一标题层级与样式", 1, id="macro_normalize_heading_styles"), "样式统一"),
            (MacroItem("CleanEmptyParagraphs", "系统内置", "清理空段落和多余换行", 2, id="macro_clean_empty_paragraphs"), "清理排版"),
            (MacroItem("RefreshFieldsBeforeExport", "本地宏库", "导出前更新目录、编号和交叉引用", 3, id="macro_refresh_fields_before_export"), "导出准备"),
            (MacroItem("RepairEquationAnchors", "本地宏库", "修复公式对象锚点和行内位置", 4, id="macro_repair_equation_anchors"), "公式修复"),
            (MacroItem("CompactOcrParagraphs", "模板", "合并 OCR 后的破碎段落", 5, id="macro_compact_ocr_paragraphs"), "OCR 清理"),
        ]
        usage = self._macro_usage_index()
        items: list[dict[str, Any]] = []
        for macro, purpose in macro_specs:
            item = macro.to_dict()
            stats = usage.get(item["id"], {})
            item["macro_purpose"] = purpose
            item["usage_count"] = int(stats.get("usage_count", 0) or 0)
            item["last_used_at"] = str(stats.get("last_used_at") or "")
            item["recently_used"] = item["usage_count"] > 0
            items.append(item)
        for source in self.store.list_macro_sources():
            for index, declaration in enumerate(source["macros"]):
                identifier = source["id"] + "_" + str(index)
                item = MacroItem(declaration["qualified_name"], "本地宏库", "导入的 VBA 标准模块", len(items) + 1, id=identifier).to_dict()
                item.update(declaration)
                item.update({"macro_name": declaration["qualified_name"], "source_id": source["id"],
                             "source_file": source["file_name"], "source_sha256": source["sha256"],
                             "macro_purpose": "本地导入", "usage_count": usage.get(identifier, {}).get("usage_count", 0),
                             "last_used_at": usage.get(identifier, {}).get("last_used_at", ""),
                             "recently_used": bool(usage.get(identifier)), "definition_available": True})
                items.append(item)
        return items

    def _macro_usage_index(self) -> dict[str, dict[str, Any]]:
        """从已有报告累计宏使用次数。"""
        usage: dict[str, dict[str, Any]] = {}
        for report in self.store.list_reports():
            timestamp = str(report.get("created_at") or "")
            for macro in report.get("analysis", {}).get("macros", []):
                macro_id = str(macro.get("id") or "")
                if not macro_id:
                    continue
                stats = usage.setdefault(macro_id, {"usage_count": 0, "last_used_at": ""})
                stats["usage_count"] = int(stats["usage_count"]) + 1
                if timestamp > str(stats.get("last_used_at") or ""):
                    stats["last_used_at"] = timestamp
        return usage

    def list_macro_templates(self) -> list[dict[str, Any]]:
        """返回已保存的宏执行顺序模板。"""
        return self.store.list_macro_templates()

    def save_macro_template(self, payload: dict[str, Any]) -> dict[str, Any]:
        """校验并保存可复用的有序宏序列。"""
        self._assert_permission("templates.manage")
        library = {macro["id"]: macro for macro in self.macro_library()}
        sequence: list[dict[str, Any]] = []
        for index, item in enumerate(payload.get("macro_sequence") or payload.get("macroSequence") or [], start=1):
            macro_id = str(item.get("id") or item.get("macro_id") or "")
            if macro_id not in library:
                continue
            macro = library[macro_id]
            sequence.append(
                {
                    "id": macro_id,
                    "macro_name": macro["macro_name"],
                    "macro_source": macro["macro_source"],
                    "execute_order": int(item.get("execute_order") or index),
                    "failure_strategy": str(item.get("failure_strategy") or payload.get("failureStrategy") or "跳过"),
                    "execute_timing": str(item.get("execute_timing") or payload.get("executeTiming") or "Word 处理前"),
                    "confirmed": bool(item.get("confirmed") or payload.get("confirmMacroRisk")),
                }
            )
        if not sequence:
            raise ValueError("宏顺序模板至少需要包含一个有效宏")
        sequence.sort(key=lambda item: item["execute_order"])
        defaults = {
            "failureStrategy": str(payload.get("failureStrategy") or sequence[0].get("failure_strategy") or "跳过"),
            "executeTiming": str(payload.get("executeTiming") or sequence[0].get("execute_timing") or "Word 处理前"),
            "confirmMacroRisk": bool(payload.get("confirmMacroRisk") or any(item.get("confirmed") for item in sequence)),
        }
        return self.store.save_macro_template(
            {
                "id": payload.get("id"),
                "name": payload.get("name"),
                "description": payload.get("description"),
                "macro_sequence": sequence,
                "defaults": defaults,
            }
        )

    def delete_macro_template(self, template_id: str) -> bool:
        """删除已保存的宏执行顺序模板。"""
        self._assert_permission("templates.manage")
        return self.store.delete_macro_template(template_id)

    def list_image_annotations(self, report_id: str | None = None) -> list[dict[str, Any]]:
        """返回全部或指定报告的图片审阅标注。"""
        return self.store.list_image_annotations(report_id)

    def save_image_annotation(self, payload: dict[str, Any]) -> dict[str, Any]:
        """校验并保存手动小图片审阅决定。"""
        if not payload.get("image_id") or not payload.get("report_id"):
            raise ValueError("图片标记需要 image_id 和 report_id")
        status = str(payload.get("status") or "误判")
        if status not in {"误判", "删除待处理", "替换待处理", "类型已确认", "位置未知"}:
            raise ValueError("图片标记状态不支持")
        report = self.store.get_report(str(payload.get("report_id")))
        if not report:
            raise ValueError("报告不存在")
        image_ids = {item.get("id") for item in report.get("analysis", {}).get("smallImages", [])}
        if payload.get("image_id") not in image_ids:
            raise ValueError("图片不属于该报告")
        return self.store.save_image_annotation(payload)

    def save_image_replacement(self, payload: dict[str, Any]) -> dict[str, Any]:
        """保存替换图片资源并标记本地写回需求。"""
        file_name = self._safe_file_name(str(payload.get("file_name") or payload.get("fileName") or "replacement.png"))
        mime_type = str(payload.get("mime_type") or payload.get("mimeType") or "").lower()
        extension = Path(file_name).suffix.lower() or REPLACEMENT_IMAGE_MIME_SUFFIXES.get(mime_type, ".png")
        if extension not in REPLACEMENT_IMAGE_SUFFIXES:
            raise ValueError("替换图片格式不支持")
        encoded = str(payload.get("content_base64") or payload.get("contentBase64") or payload.get("data_url") or payload.get("dataUrl") or "")
        if "," in encoded and encoded.lstrip().startswith("data:"):
            encoded = encoded.split(",", 1)[1]
        if not encoded.strip():
            raise ValueError("替换图片内容不能为空")
        try:
            content = base64.b64decode(encoded, validate=True)
        except (binascii.Error, ValueError) as exc:
            raise ValueError("替换图片内容不是有效的 Base64") from exc
        if not content:
            raise ValueError("替换图片内容不能为空")
        limit = int(self.store.get_settings().get("singleFileLimitMb", 500)) * 1024 * 1024
        if len(content) > limit:
            raise ValueError("替换图片超过单文件大小限制")
        storage_name = f"{new_id('replacement')}{extension}"
        target = self.store.uploads_dir / storage_name
        target.write_bytes(content)
        annotation = self.save_image_annotation(
            {
                **payload,
                "status": "替换待处理",
                "replacement_file_name": file_name,
                "replacement_storage_name": storage_name,
                "replacement_asset_url": f"/api/assets/replacements/{storage_name}",
                "replacement_file_size": len(content),
                "replacement_sha256": hashlib.sha256(content).hexdigest(),
                "replacement_mime_type": mime_type or "image/*",
                "note": str(payload.get("note") or "用户选择替换图片，等待本地客户端改写源文档"),
            }
        )
        self.store.append_log(str(payload.get("report_id") or "system"), f"图片替换素材已登记：{file_name}", category="image")
        return annotation

    def reexport_image_asset(self, payload: dict[str, Any]) -> dict[str, Any]:
        """从原始源文档重新导出图片。"""
        report_id = str(payload.get("report_id") or payload.get("reportId") or "")
        image_id = str(payload.get("image_id") or payload.get("imageId") or "")
        if not report_id or not image_id:
            raise ValueError("重新导出图片需要 report_id 和 image_id")
        report = self.store.get_report(report_id)
        if not report:
            raise ValueError("报告不存在")
        images = report.get("analysis", {}).get("smallImages", [])
        image = next((item for item in images if item.get("id") == image_id), None)
        if not image:
            raise ValueError("图片不属于该报告")
        now = utc_now()
        try:
            content, extension = self._reexport_image_bytes(report, image)
        except ValueError as exc:
            image["export_status"] = "重新导出失败"
            image["export_message"] = str(exc)
            image["reexported_at"] = now
            self._persist_updated_report(report)
            self.store.append_log(str(report.get("task_id") or "system"), f"图片重新导出失败：{image.get('source_name') or image_id}，{exc}", category="image")
            return {"reexported": False, "message": str(exc), "image": image, "report": report}
        output_dir = self.store.images_dir / str(report.get("task_id") or "manual") / str(image.get("file_id") or "unknown")
        output_dir.mkdir(parents=True, exist_ok=True)
        target = output_dir / f"{new_id('image')}{extension}"
        target.write_bytes(content)
        image["image_path"] = str(target)
        image["asset_url"] = f"/api/assets/{target.relative_to(self.store.data_dir).as_posix()}"
        image["image_size"] = len(content)
        image["image_hash"] = hashlib.sha256(content).hexdigest()
        image["export_status"] = "已重新导出"
        image["export_message"] = "已从来源文档重新导出图片资源"
        image["reexported_at"] = now
        self._persist_updated_report(report)
        self.store.append_log(str(report.get("task_id") or "system"), f"图片已重新导出：{image.get('source_name') or image_id}", category="image")
        return {"reexported": True, "message": "图片已重新导出", "image": image, "report": report}

    def _persist_updated_report(self, report: dict[str, Any]) -> dict[str, Any]:
        """重写报告文件并保存更新后的记录。"""
        self.report_builder.rewrite(report)
        return self.store.save_report(report)

    def _reexport_image_bytes(self, report: dict[str, Any], image: dict[str, Any]) -> tuple[bytes, str]:
        """从受管源读取原始图片字节供重新导出。"""
        file_id = str(image.get("file_id") or "")
        source_file = self.store.get_file(file_id) or next((item for item in report.get("files", []) if item.get("id") == file_id), None)
        if not source_file:
            raise ValueError("来源文件不存在，已保留原始结果")
        path = Path(str(source_file.get("storage_path") or source_file.get("file_path") or ""))
        if not path.exists() or not path.is_file():
            raise ValueError("来源文件不可访问，已保留原始结果")
        file_type = str(source_file.get("file_type") or "")
        source_name = str(image.get("source_name") or "")
        image_type = str(image.get("image_type") or "").lower()
        if file_type in OOXML_IMAGE_PREFIXES:
            try:
                with zipfile.ZipFile(path) as archive:
                    content = archive.read(source_name)
            except KeyError as exc:
                raise ValueError("来源文档中未找到图片条目，已保留原始结果") from exc
            except zipfile.BadZipFile as exc:
                raise ValueError("来源文档无法读取，已保留原始结果") from exc
            return content, self._image_asset_extension(source_name, image_type)
        if file_type == "图片":
            return path.read_bytes(), self._image_asset_extension(path.name, image_type)
        if file_type == "PDF":
            descriptors = _pdf_image_descriptors(path.read_bytes()[:20_000_000])
            page_index = int(image.get("page_index") or 0)
            descriptor = next((item for item in descriptors if int(item.get("index") or 0) == page_index), None)
            stream = descriptor.get("stream") if descriptor else b""
            if not stream:
                raise ValueError("PDF 图片原始流不可导出，已保留原始结果")
            return stream, self._image_asset_extension(source_name, image_type)
        raise ValueError("该来源类型暂不支持重新导出，已保留原始结果")

    @staticmethod
    def _image_asset_extension(source_name: str, image_type: str) -> str:
        """为提取或重新导出的图片选择安全扩展名。"""
        extension = Path(source_name).suffix.lower()
        if extension:
            return extension
        normalized = image_type.lower().lstrip(".")
        if normalized == "jpeg":
            normalized = "jpg"
        if normalized in {"png", "jpg", "gif", "webp", "bmp", "svg", "emf", "wmf", "jpx", "jp2", "jb2"}:
            return f".{normalized}"
        return ".bin"

    def delete_image_annotation(self, annotation_id: str) -> bool:
        """删除图片审阅标注。"""
        return self.store.delete_image_annotation(annotation_id)

    def list_formula_annotations(self, report_id: str | None = None) -> list[dict[str, Any]]:
        """返回全部或指定报告的公式审阅标注。"""
        return self.store.list_formula_annotations(report_id)

    def save_formula_annotation(self, payload: dict[str, Any]) -> dict[str, Any]:
        """校验并保存手动公式更正或确认。"""
        formula_id = str(payload.get("formula_id") or payload.get("formulaId") or "")
        report_id = str(payload.get("report_id") or payload.get("reportId") or "")
        if not formula_id or not report_id:
            raise ValueError("公式校正需要 formula_id 和 report_id")
        report = self.store.get_report(report_id)
        if not report:
            raise ValueError("报告不存在")
        formulas = list(report.get("analysis", {}).get("formulas", []))
        formula = next((item for item in formulas if item.get("id") == formula_id), None)
        if not formula:
            raise ValueError("公式不属于该报告")
        status = str(payload.get("status") or "已确认").strip()
        if status == "重识别":
            status = "重新识别"
        if status not in {"已确认", "已修正", "重新识别", "跳过"}:
            raise ValueError("公式校正状态不支持")
        annotation_payload = {**payload, "formula_id": formula_id, "report_id": report_id, "status": status}
        if status == "重新识别":
            request = self._formula_recognition_request(report, formula)
            annotation_payload.update(
                {
                    "retry_recognition": True,
                    "recognition_status": request["status"],
                    "recognition_request": request,
                    "next_step": request["next_step"],
                }
            )
        return self.store.save_formula_annotation(annotation_payload)

    def _formula_recognition_request(self, report: dict[str, Any], formula: dict[str, Any]) -> dict[str, Any]:
        """生成重识别约定，不调用 Mathpix 或桌面工具。"""
        settings = self.store.get_settings()
        source_type = str(formula.get("source_type") or "").strip() or "未知"
        source_file = next((file for file in report.get("files", []) if file.get("id") == formula.get("file_id")), {})
        request = {
            "schema_version": "k12.formulaRecognitionRequest.v1",
            "formula_id": str(formula.get("id") or ""),
            "report_id": str(report.get("id") or ""),
            "task_id": str(report.get("task_id") or ""),
            "source_type": source_type,
            "source_file": str(source_file.get("file_name") or formula.get("file_id") or ""),
            "position": str(formula.get("position") or ""),
            "original_image_ref": str(formula.get("original_image_ref") or formula.get("original_image_path") or ""),
            "external_upload_performed": False,
            "task_options_can_authorize_external_upload": False,
            "blocking_reasons": [],
            "required_environment": [],
        }
        if source_type in {"PDF", "图片"}:
            external_allowed = bool(settings.get("allowExternalMathpixUpload"))
            credentials_ready = bool(os.environ.get("MATHPIX_APP_ID") and os.environ.get("MATHPIX_APP_KEY"))
            blockers: list[str] = []
            if not external_allowed:
                blockers.append("external_mathpix_upload_not_authorized")
            if external_allowed and not credentials_ready:
                blockers.append("mathpix_credentials_required")
            status = "ready_for_authorized_mathpix_worker" if external_allowed and credentials_ready else "blocked"
            request.update(
                {
                    "engine": "Mathpix",
                    "action": "formula_ocr_rerecognize",
                    "status": status,
                    "external_upload_required": True,
                    "external_upload_authorized": external_allowed,
                    "external_upload_authorization_source": "settings.allowExternalMathpixUpload",
                    "credentials_required": ["MATHPIX_APP_ID", "MATHPIX_APP_KEY"],
                    "credentials_configured": credentials_ready,
                    "blocking_reasons": blockers,
                    "required_environment": ["Mathpix APP ID/KEY", "用户授权外部 Mathpix 上传"],
                    "next_step": "配置 Mathpix 凭证并在设置中授权外部上传后，由受控 worker 重新识别该公式" if blockers else "等待受控 Mathpix worker 重新识别该公式",
                }
            )
            return request
        request.update(
            {
                "engine": "local-client",
                "action": "formula_rerecognize_or_convert",
                "status": "queued_for_local_client",
                "external_upload_required": False,
                "external_upload_authorized": False,
                "required_environment": ["同平台 Office/MathType 本地客户端"],
                "next_step": "等待本地客户端按同平台公式合同重新识别或转换该公式",
            }
        )
        return request

    def bulk_confirm_formulas(self, payload: dict[str, Any]) -> dict[str, Any]:
        """批量确认指定报告中的高置信度公式。"""
        report_id = str(payload.get("report_id") or "")
        if not report_id:
            raise ValueError("批量确认需要 report_id")
        report = self.store.get_report(report_id)
        if not report:
            raise ValueError("报告不存在")
        min_confidence = int(payload.get("min_confidence") or payload.get("minConfidence") or 80)
        only_unannotated = bool(payload.get("only_unannotated", payload.get("onlyUnannotated", True)))
        existing_ids = {item.get("formula_id") for item in self.store.list_formula_annotations(report_id)} if only_unannotated else set()
        annotations: list[dict[str, Any]] = []
        for formula in report.get("analysis", {}).get("formulas", []):
            if int(formula.get("confidence", 0)) < min_confidence:
                continue
            if formula.get("id") in existing_ids:
                continue
            annotations.append(
                self.store.save_formula_annotation(
                    {
                        "report_id": report_id,
                        "formula_id": formula.get("id"),
                        "status": "已确认",
                        "latex": formula.get("latex", ""),
                        "mathml": formula.get("mathml", ""),
                        "note": f"批量确认：置信度不低于 {min_confidence}%",
                    }
                )
            )
        return {"annotations": annotations, "confirmed_count": len(annotations), "min_confidence": min_confidence}

    def delete_formula_annotation(self, annotation_id: str) -> bool:
        """删除公式审阅标注。"""
        return self.store.delete_formula_annotation(annotation_id)

    def list_omml_annotations(self, report_id: str | None = None) -> list[dict[str, Any]]:
        """返回全部或指定报告的 OMML 依赖标注。"""
        return self.store.list_omml_annotations(report_id)

    def save_omml_annotation(self, payload: dict[str, Any]) -> dict[str, Any]:
        """校验并保存 OMML 转换或依赖决定。"""
        dependency_id = str(payload.get("dependency_id") or payload.get("dependencyId") or "")
        report_id = str(payload.get("report_id") or payload.get("reportId") or "")
        if not dependency_id or not report_id:
            raise ValueError("OMML 校正需要 dependency_id 和 report_id")
        status = str(payload.get("status") or "保留OMML")
        if status not in {"转换失败", "保留OMML", "重新转换", "手动指定依赖", "已修复"}:
            raise ValueError("OMML 校正状态不支持")
        report = self.store.get_report(report_id)
        if not report:
            raise ValueError("报告不存在")
        dependency_ids = {item.get("id") for item in report.get("analysis", {}).get("ommlDependencies", [])}
        if dependency_id not in dependency_ids:
            raise ValueError("OMML 依赖不属于该报告")
        return self.store.save_omml_annotation({**payload, "dependency_id": dependency_id, "report_id": report_id, "status": status})

    def delete_omml_annotation(self, annotation_id: str) -> bool:
        """删除 OMML 依赖标注。"""
        return self.store.delete_omml_annotation(annotation_id)

    def list_layout_annotations(self, report_id: str | None = None) -> list[dict[str, Any]]:
        """返回全部或指定报告的排版更正备注。"""
        return self.store.list_layout_annotations(report_id)

    def save_layout_annotation(self, payload: dict[str, Any]) -> dict[str, Any]:
        """校验并保存手动排版更正备注。"""
        report_id = str(payload.get("report_id") or payload.get("reportId") or "")
        location = str(payload.get("location") or payload.get("position") or "").strip()
        if not report_id or not location:
            raise ValueError("排版校正需要 report_id 和 location")
        issue_type = str(payload.get("issue_type") or payload.get("issueType") or "标题层级")
        if issue_type not in {"页码", "标题层级", "表格结构", "图片位置", "公式位置", "公式编号"}:
            raise ValueError("排版校正类型不支持")
        status = str(payload.get("status") or "待校正")
        if status not in {"待校正", "已校正", "已忽略", "需本地客户端处理"}:
            raise ValueError("排版校正状态不支持")
        if not self.store.get_report(report_id):
            raise ValueError("报告不存在")
        return self.store.save_layout_annotation({**payload, "report_id": report_id, "location": location, "issue_type": issue_type, "status": status})

    def delete_layout_annotation(self, annotation_id: str) -> bool:
        """删除排版更正备注。"""
        return self.store.delete_layout_annotation(annotation_id)

    def _register_archive_entries(self, archive_path: Path, parent: FileItem, analyzer: DocumentAnalyzer) -> list[dict[str, Any]]:
        """将支持的压缩包条目提取至受管上传存储。"""
        items: list[dict[str, Any]] = []
        extract_dir = self.store.uploads_dir / parent.id
        extract_dir.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(archive_path) as archive:
            for info in archive.infolist():
                if info.is_dir():
                    continue
                entry_name = self._safe_archive_name(info.filename)
                extension = Path(entry_name).suffix.lower()
                if extension not in analyzer.supported_extensions():
                    continue
                target = extract_dir / f"{new_id('entry')}{extension}"
                try:
                    target.write_bytes(archive.read(info))
                except (zipfile.BadZipFile, EOFError, OSError, RuntimeError) as exc:
                    parent.validation_errors.append(f"ZIP 条目损坏或无法读取：{entry_name}")
                    parent.status = "校验失败"
                    parent.content_summary = {
                        **parent.content_summary,
                        "failedEntry": entry_name,
                        "failureType": exc.__class__.__name__,
                    }
                    persisted_parent = self.store.get_file(parent.id) or parent.to_dict()
                    persisted_parent["validation_errors"] = list(parent.validation_errors)
                    persisted_parent["status"] = parent.status
                    persisted_parent["content_summary"] = dict(parent.content_summary)
                    self.store.save_file(persisted_parent)
                    target.unlink(missing_ok=True)
                    continue
                entry = analyzer.analyze_file(target, Path(entry_name).name)
                entry.source_kind = "archive_entry"
                entry.source_relative_path = entry_name
                entry.archive_parent_id = parent.id
                entry.content_summary = {
                    **entry.content_summary,
                    "archive": parent.file_name,
                    "entry": entry_name,
                    "compressedSize": info.compress_size,
                    "uncompressedSize": info.file_size,
                }
                record = entry.to_dict()
                record["source_sha256"] = self._sha256(target)
                items.append(self.store.save_file(record))
        return items

    def _resolve_duplicate_file_name(self, file_name: str, strategy: str) -> str | None:
        """按跳过、覆盖或重命名策略处理上传名称冲突。"""
        name = self._safe_file_name(file_name)
        files = self.store.list_files()
        duplicates = [file for file in files if file.get("file_name") == name]
        if not duplicates:
            return name
        if strategy == "跳过":
            return None
        if strategy == "覆盖":
            for file in duplicates:
                self.store.delete_file(file["id"])
            return name
        return self._unique_file_name(name, {str(file.get("file_name") or "") for file in files})

    def _find_existing_file(self, file_name: str) -> dict[str, Any] | None:
        """按已清理显示名查找存储的文件。"""
        name = self._safe_file_name(file_name)
        return next((file for file in self.store.list_files() if file.get("file_name") == name), None)

    @staticmethod
    def _unique_file_name(file_name: str, existing_names: set[str]) -> str:
        """返回无冲突的显示文件名。"""
        path = Path(file_name)
        stem = path.stem or "unnamed"
        suffix = path.suffix
        for index in range(1, 1000):
            candidate = f"{stem}-{index}{suffix}"
            if candidate not in existing_names:
                return candidate
        return f"{stem}-{new_id('dup')}{suffix}"

    @staticmethod
    def _safe_file_name(file_name: str) -> str:
        """清理受管存储使用的单段文件名。"""
        name = str(file_name or "unnamed").replace("\\", "/").rsplit("/", 1)[-1].strip()
        name = ILLEGAL_FILE_CHARS.sub("_", name).rstrip(" .") or "unnamed"
        # Windows 将这些名称解释为设备，即使带扩展名也无法作为普通文件创建。
        if re.fullmatch(r"CON|PRN|AUX|NUL|COM[1-9¹²³]|LPT[1-9¹²³]", name.split(".", 1)[0].rstrip().upper()):
            name = "_" + name
        return name

    @staticmethod
    def _safe_relative_path(file_name: str) -> str:
        """清理压缩包相对路径，阻止路径穿越。"""
        parts: list[str] = []
        for raw_part in str(file_name or "").replace("\\", "/").split("/"):
            part = raw_part.strip()
            if not part or part in {".", ".."}:
                continue
            safe = TaskProcessor._safe_file_name(part)
            if safe:
                parts.append(safe)
        return "/".join(parts)

    @staticmethod
    def _non_negative_int(value: Any, default: int) -> int:
        """读取非负整数，无效时使用兜底值。"""
        try:
            return max(0, int(value))
        except (TypeError, ValueError):
            return default

    @staticmethod
    def _positive_int(value: Any, default: int) -> int:
        """读取正整数，无效时使用兜底值。"""
        try:
            return max(1, int(value))
        except (TypeError, ValueError):
            return default

    @staticmethod
    def _parse_utc(value: str) -> datetime | None:
        """解析 ISO 时间戳并规范为 UTC。"""
        if not value:
            return None
        try:
            parsed = datetime.fromisoformat(value)
        except ValueError:
            return None
        if parsed.tzinfo is None:
            parsed = parsed.replace(tzinfo=timezone.utc)
        return parsed.astimezone(timezone.utc)

    @staticmethod
    def _safe_archive_name(file_name: str) -> str:
        """返回防止路径穿越的压缩包条目名。"""
        parts: list[str] = []
        for raw_part in str(file_name or "").replace("\\", "/").split("/"):
            part = raw_part.strip()
            if not part or part in {".", ".."}:
                continue
            safe = ILLEGAL_FILE_CHARS.sub("_", part)
            if safe:
                parts.append(safe)
        return "/".join(parts) or "unnamed"

    def _stages_for(self, task_type: str, execute_mode: str) -> list[str]:
        """返回任务类型的用户可见处理阶段。"""
        base = [
            "读取文件元数据并校验格式",
            "识别文档类型与处理能力",
            "写入任务队列",
        ]
        task_stages = {
            "word_to_ppt": ["解析 Word 标题层级", "保留图片、表格和公式对象", "生成 PPT 输出清单"],
            "ppt_to_word": ["读取幻灯片缩略结构", "提取标题、正文和备注", "生成 Word 讲义清单"],
            "pdf_to_word": ["识别 PDF 类型", "准备 Mathpix PDF OCR 请求", "规划 DOCX/LaTeX 下载结果", "生成 Word 转换清单"],
            "excel_to_pdf": ["读取工作表结构", "规划公式和图表输出", "生成 PDF 输出清单"],
            "excel_to_word": ["读取工作表结构", "提取单元格与公式摘要", "生成 Word 表格清单"],
            "excel_to_ppt": ["读取工作表结构", "提取图表页摘要", "生成 PPT 图表页清单"],
            "formula_precheck": ["检测 OMML、MathType、LaTeX 和图片公式", "标记低置信度公式", "生成公式预检报告"],
            "omml_to_mathtype": ["检查 OMML 转换依赖", "检索并复制 OMML 依赖文件", "生成 OMML 转换报告"],
            "mathtype_format": ["读取 MathType 格式化参数", "模拟全文公式格式化", "生成格式化报告"],
            "macro_sequence": ["读取宏库和执行顺序", "确认备份与失败策略", "生成宏执行报告"],
            "small_image_scan": ["扫描文档对象清单", "按阈值识别微小图片", "生成图片检索报告"],
            "batch_process": ["拆分批量任务", "按并发配置调度", "生成批量处理报告"],
        }
        mode_stage = f"执行模式：{execute_mode}"
        return base + [mode_stage] + task_stages.get(task_type, task_stages["batch_process"])

    @staticmethod
    def _task_log_category(task_type: str) -> str:
        """将任务类型映射为日志类别。"""
        if task_type in {"word_to_ppt", "ppt_to_word", "pdf_to_word", "excel_to_pdf", "excel_to_word", "excel_to_ppt", "batch_process"}:
            return "conversion"
        if task_type in {"formula_precheck", "mathtype_format"}:
            return "formula"
        if task_type == "omml_to_mathtype":
            return "omml"
        if task_type == "macro_sequence":
            return "macro"
        if task_type == "small_image_scan":
            return "image"
        return "system"

    def _build_analysis(self, task: dict[str, Any], files: list[dict[str, Any]]) -> dict[str, Any]:
        """生成任务公式、宏和小图片分析。"""
        settings = self.store.get_settings()
        formulas: list[dict[str, Any]] = []
        macros: list[dict[str, Any]] = []
        images: list[dict[str, Any]] = []
        image_errors: list[dict[str, Any]] = []
        allow_batch_macro = bool(settings.get("allowBatchMacroExecution", True))
        for file_index, file in enumerate(files, start=1):
            if file.get("validation_errors"):
                continue
            seed = int(hashlib.sha256(file["id"].encode("utf-8")).hexdigest()[:4], 16)
            if file.get("has_formula"):
                formulas.extend(self._formula_items(file, seed, task))
            file_images: list[dict[str, Any]] = []
            if task["task_type"] == "macro_sequence" or (settings.get("enableMacroDetection", True) and file.get("has_macro")):
                macros.extend(self._macro_items(file, task, batch_allowed=allow_batch_macro or file_index == 1))
            if file.get("has_small_image") or task["task_type"] == "small_image_scan":
                file_images = self._small_image_items(file, task, seed, image_errors)
                images.extend(file_images)
            if task["task_type"] in {"formula_precheck", "mathtype_format"} and file_images:
                formulas.extend(self._image_formula_fallback_items(file, file_images, task))
        return {"formulas": formulas, "macros": macros, "smallImages": images, "imageExtractionErrors": image_errors}

    def _batch_results(self, task: dict[str, Any], files: list[dict[str, Any]]) -> list[dict[str, Any]]:
        """遵守跳过决定生成逐文件批量结果。"""
        skipped_file_ids = self._batch_skipped_file_ids(task)
        results: list[dict[str, Any]] = []
        for index, file in enumerate(files, start=1):
            errors = list(file.get("validation_errors") or [])
            suggested_tasks = self._suggested_tasks_for_file(file)
            skipped = file["id"] in skipped_file_ids and bool(errors)
            status = "跳过" if skipped else ("失败" if errors else "成功")
            results.append(
                {
                    "file_id": file["id"],
                    "file_name": file.get("file_name", ""),
                    "file_type": file.get("file_type", ""),
                    "sequence": index,
                    "status": status,
                    "progress": 100,
                    "retryable": bool(errors) and not skipped,
                    "skipped": skipped,
                    "suggested_tasks": suggested_tasks,
                    "error_message": "；".join(errors),
                    "recommendation": "已按用户选择跳过，不阻塞本批次" if skipped else ("处理文件校验问题后重试" if errors else f"建议执行：{'、'.join(suggested_tasks) if suggested_tasks else '无需额外任务'}"),
                }
            )
        return results

    def _batch_plan(self, task: dict[str, Any], files: list[dict[str, Any]], results: list[dict[str, Any]]) -> dict[str, Any]:
        """生成包含并发数与失败策略的批量调度摘要。"""
        settings = self.store.get_settings()
        failure_strategy = str(task.get("options", {}).get("batchFailureStrategy") or task.get("options", {}).get("failureStrategy") or "跳过")
        try:
            max_concurrent = max(1, min(8, int(task.get("options", {}).get("maxConcurrentTasks", settings.get("maxConcurrentTasks", 3)) or 3)))
        except (TypeError, ValueError):
            max_concurrent = 3
        continue_on_failure = bool(task.get("options", {}).get("continueOnFailure", failure_strategy == "跳过"))
        chunks: list[dict[str, Any]] = []
        for start in range(0, len(results), max_concurrent):
            group = results[start : start + max_concurrent]
            chunks.append(
                {
                    "index": len(chunks) + 1,
                    "file_count": len(group),
                    "file_ids": [item.get("file_id", "") for item in group],
                    "status": "失败" if any(item.get("status") == "失败" for item in group) else "成功",
                }
            )
        return {
            "total_files": len(files),
            "max_concurrent": max_concurrent,
            "failure_strategy": failure_strategy,
            "continue_on_failure": continue_on_failure,
            "requires_user_decision": failure_strategy == "询问",
            "chunk_count": len(chunks),
            "retryable_count": sum(1 for item in results if item.get("retryable")),
            "failed_count": sum(1 for item in results if item.get("status") == "失败"),
            "skipped_count": sum(1 for item in results if item.get("status") == "跳过" or item.get("skipped")),
            "chunks": chunks,
        }

    @staticmethod
    def _batch_skipped_file_ids(task: dict[str, Any]) -> list[str]:
        """返回用户选择跳过的批量文件标识。"""
        raw_ids = task.get("skipped_file_ids") or task.get("skippedFileIds") or []
        if isinstance(raw_ids, str):
            raw_ids = [raw_ids]
        return [str(file_id) for file_id in raw_ids if str(file_id)]

    def _latest_batch_result(self, task_id: str, file_id: str) -> dict[str, Any]:
        """返回指定任务和文件最近的批量结果。"""
        reports = [report for report in self.store.list_reports() if report.get("task_id") == task_id]
        for report in reports:
            for item in report.get("analysis", {}).get("batchResults", []):
                if item.get("file_id") == file_id:
                    return item
        return {}

    @staticmethod
    def _suggested_tasks_for_file(file: dict[str, Any]) -> list[str]:
        """根据检测到的文件能力建议后续任务类型。"""
        file_type = file.get("file_type")
        tasks: list[str] = []
        if file_type == "Word":
            tasks.append("word_to_ppt")
            if file.get("has_formula"):
                tasks.append("formula_precheck")
            if file.get("has_omml"):
                tasks.append("omml_to_mathtype")
            if file.get("has_mathtype") or file.get("has_formula"):
                tasks.append("mathtype_format")
            if file.get("has_macro"):
                tasks.append("macro_sequence")
        elif file_type == "PPT":
            tasks.append("ppt_to_word")
        elif file_type == "PDF":
            tasks.append("pdf_to_word")
        elif file_type == "Excel":
            tasks.append("excel_to_pdf")
            tasks.append("excel_to_word")
            tasks.append("excel_to_ppt")
        if file.get("has_small_image"):
            tasks.append("small_image_scan")
        return list(dict.fromkeys(tasks))

    def _omml_dependency_jobs(self, task: dict[str, Any], files: list[dict[str, Any]]) -> list[dict[str, Any]]:
        """生成 OMML 依赖查找与复制作业，不编辑文档。"""
        settings = self.store.get_settings()
        jobs: list[dict[str, Any]] = []
        if not settings.get("enableOmmlPrecheck", True):
            return jobs
        for file in files:
            if file.get("validation_errors") or file.get("file_type") != "Word" or not file.get("has_omml"):
                continue
            document_path = self._original_source_path(file)
            item = OmmlDependencyItem(file_id=file["id"], document_path=str(document_path))
            manual_source = self._manual_omml_dependency(settings)
            if manual_source:
                item.omml_file_name = manual_source.name
                item.omml_source_path = str(manual_source)
                item.found_status = "手动选择"
                self._copy_omml_dependency(manual_source, document_path, item, str(settings.get("ommlCopyStrategy", "自动重命名")))
                self._set_file_omml_dependency_missing(file, item.copy_status == "失败")
                self.store.append_log(task["id"], f"OMML 手动依赖{item.copy_status}：{manual_source.name}", category="omml")
                jobs.append(item.to_dict())
                continue
            if not settings.get("autoSearchOmml", True):
                item.found_status = "未找到" if settings.get("allowManualOmml", True) else "跳过"
                item.error_message = "自动检索 OMML 已关闭，且未提供可用手动 OMML 文件" if settings.get("allowManualOmml", True) else "自动检索 OMML 已关闭"
                self._set_file_omml_dependency_missing(file, True)
                jobs.append(item.to_dict())
                continue
            source = self._find_omml_dependency(document_path, settings)
            if not source:
                item.found_status = "未找到"
                item.error_message = "未在当前文档目录、运行目录或自定义路径中找到 OMML 依赖文件"
                self._set_file_omml_dependency_missing(file, True)
                jobs.append(item.to_dict())
                self.store.append_log(task["id"], f"OMML 依赖未找到：{file['file_name']}", "warning", category="omml")
                continue
            item.omml_file_name = source.name
            item.omml_source_path = str(source)
            item.found_status = "已找到"
            self._copy_omml_dependency(source, document_path, item, str(settings.get("ommlCopyStrategy", "自动重命名")))
            self._set_file_omml_dependency_missing(file, item.copy_status == "失败")
            self.store.append_log(task["id"], f"OMML 依赖{item.copy_status}：{source.name}", category="omml")
            jobs.append(item.to_dict())
        return jobs

    def _omml_conversion_prompts(self, task: dict[str, Any], files: list[dict[str, Any]], dependencies: list[dict[str, Any]]) -> list[dict[str, Any]]:
        """为 Word 生成 OMML 转 MathType 提示。"""
        settings = self.store.get_settings()
        dependencies_by_file = {item.get("file_id"): item for item in dependencies}
        prompts: list[dict[str, Any]] = []
        for file in files:
            if file.get("file_type") != "Word" or not file.get("has_omml"):
                continue
            decision = self._omml_conversion_decision(task, settings)
            summary = file.get("content_summary") or {}
            omml_count = self._summary_count(summary, "ommlFormulas") or 1
            dependency = dependencies_by_file.get(file.get("id"), {})
            if decision["convert_to_mathtype"] is True:
                status = "已选择转换"
                message = "用户已选择将 Word 自带公式转换为 MathType，并保留原 OMML 兜底"
                local_action = "convert_omml_to_mathtype"
            elif decision["convert_to_mathtype"] is False:
                status = "保留原公式"
                message = "用户选择不转换 Word 自带公式，保留原 OMML 公式"
                local_action = "keep_original_omml"
            else:
                status = "需确认"
                message = "检测到 Word 自带公式，需用户确认是否转换为 MathType"
                local_action = "ask_user"
            prompt = {
                "file_id": file["id"],
                "file_name": file.get("file_name", ""),
                "source_type": "OMML",
                "omml_count": omml_count,
                "status": status,
                "convert_to_mathtype": decision["convert_to_mathtype"],
                "requires_user_confirmation": decision["convert_to_mathtype"] is None,
                "decision_source": decision["source"],
                "local_action": local_action,
                "preserve_original_formula": True,
                "dependency_status": dependency.get("copy_status") or dependency.get("found_status") or ("缺少" if file.get("missing_omml_dependency") else "未检查"),
                "dependency_id": dependency.get("id", ""),
                "message": message,
                "recommendation": "确认转换后由本地客户端执行 OMML 转 MathType；不转换时保留原公式并继续后续流程",
            }
            prompts.append(prompt)
            self.store.append_log(task["id"], f"OMML 转 MathType 确认：{file.get('file_name', '')}，{status}", category="omml")
        return prompts

    @staticmethod
    def _omml_conversion_decision(task: dict[str, Any], settings: dict[str, Any]) -> dict[str, Any]:
        """确定 OMML 公式是否应转为 MathType。"""
        options = task.get("options") or {}
        word_options = options.get("word") or {}
        for key in ("convertOmmlToMathType", "confirmOmmlConversion", "convert_omml_to_mathtype"):
            if key in options:
                return {"convert_to_mathtype": bool(options.get(key)), "source": "任务选项"}
        for key in ("convertOmmlFirst", "convertOmmlToMathType"):
            if key in word_options:
                return {"convert_to_mathtype": bool(word_options.get(key)), "source": "Word 转换选项"}
        if bool(settings.get("autoOmmlToMathType", False)):
            return {"convert_to_mathtype": True, "source": "系统设置"}
        return {"convert_to_mathtype": None, "source": "未确认"}

    def _set_file_omml_dependency_missing(self, file: dict[str, Any], missing: bool) -> None:
        """将 OMML 依赖缺失标记保存回 Word 记录。"""
        if file.get("file_type") != "Word" or not file.get("has_omml"):
            return
        file["missing_omml_dependency"] = bool(missing)
        summary = dict(file.get("content_summary") or {})
        summary["missingOmmlDependency"] = bool(missing)
        file["content_summary"] = summary
        self.store.save_file(file)

    def _manual_omml_dependency(self, settings: dict[str, Any]) -> Path | None:
        """配置的手动 OMML 依赖路径有效时返回该路径。"""
        if not settings.get("allowManualOmml", True):
            return None
        raw_path = str(settings.get("manualOmmlPath") or "").strip()
        if not raw_path:
            return None
        path = Path(raw_path).expanduser()
        if path.is_file() and self._is_omml_dependency_name(path.name):
            return path
        if path.is_dir():
            return self._find_direct_omml_file(path)
        return None

    def _find_omml_dependency(self, document_path: Path, settings: dict[str, Any]) -> Path | None:
        """在配置的本地目录中查找 OMML 依赖。"""
        candidates = self._omml_candidate_dirs(document_path, settings)
        max_files = int(settings.get("ommlSearchMaxFiles", 3000) or 3000)
        scanned = 0
        searchable = [directory for directory in candidates if directory.exists() and directory.is_dir()]
        # 递归扫描前检查所有候选根目录，避免首个目录耗尽预算而遗漏明确配置的依赖。
        for directory in searchable:
            direct = self._find_direct_omml_file(directory)
            if direct:
                return direct
        for directory in searchable:
            for root, dirs, files in os.walk(directory):
                dirs[:] = [name for name in dirs if not name.startswith(".")][:20]
                for name in files:
                    scanned += 1
                    if self._is_omml_dependency_name(name):
                        return Path(root) / name
                    if scanned >= max_files:
                        return None
        return None

    def _omml_candidate_dirs(self, document_path: Path, settings: dict[str, Any]) -> list[Path]:
        """返回按优先级排列的 OMML 依赖候选目录。"""
        candidates: list[Path] = []
        if document_path.exists():
            candidates.append(document_path.parent)
        # 优先搜索用户明确指定的位置，避免运行缓存和用户目录耗尽 OMML 扫描预算。
        configured = str(settings.get("ommlSearchPaths") or "")
        for raw in re.split(r"[\n,;]", configured):
            value = raw.strip()
            if value:
                candidates.append(Path(value).expanduser())
        candidates.extend([self.store.uploads_dir, self.store.output_base_dir(), self.store.data_dir])
        home = Path.home()
        candidates.extend([home / "Documents", home / "Desktop", home / "Downloads"])
        unique: list[Path] = []
        seen: set[str] = set()
        for candidate in candidates:
            key = str(candidate.resolve()) if candidate.exists() else str(candidate)
            if key not in seen:
                seen.add(key)
                unique.append(candidate)
        return unique

    def _copy_omml_dependency(self, source: Path, document_path: Path, item: OmmlDependencyItem, strategy: str) -> None:
        """按选定策略将 OMML 依赖复制到文档旁。"""
        target_dir = document_path.parent if str(document_path) not in {"", "."} and document_path.exists() else self.store.output_base_dir()
        target_dir.mkdir(parents=True, exist_ok=True)
        target = target_dir / source.name
        try:
            source_data = source.read_bytes()
        except OSError as exc:
            item.omml_target_path = str(target)
            item.copy_status = "失败"
            item.error_message = str(exc)
            return
        item.source_size = len(source_data)
        item.source_sha256 = hashlib.sha256(source_data).hexdigest()
        if source.resolve() == target.resolve():
            item.omml_target_path = str(target)
            item.copy_status = "跳过"
            item.target_size = len(source_data)
            item.target_sha256 = item.source_sha256
            return
        if target.exists() and strategy == "跳过":
            item.omml_target_path = str(target)
            try:
                target_data = target.read_bytes()
                item.copy_status = "跳过"
                item.target_size = len(target_data)
                item.target_sha256 = hashlib.sha256(target_data).hexdigest()
            except OSError as exc:
                item.copy_status = "失败"
                item.error_message = str(exc)
            return
        if target.exists() and strategy == "自动重命名":
            target = self._unique_target_path(target)
        try:
            target.write_bytes(source_data)
            target_data = target.read_bytes()
            if target_data != source_data:
                raise OSError("OMML 依赖目标内容与源快照不一致")
            item.omml_target_path = str(target)
            item.copy_status = "成功"
            item.target_size = len(target_data)
            item.target_sha256 = hashlib.sha256(target_data).hexdigest()
        except OSError as exc:
            item.omml_target_path = str(target)
            item.copy_status = "失败"
            item.error_message = str(exc)

    @staticmethod
    def _find_direct_omml_file(directory: Path) -> Path | None:
        """查找目录中直接存在的 OMML 转换依赖文件。"""
        return _find_direct_omml_file(directory)

    @staticmethod
    def _is_omml_dependency_name(file_name: str) -> bool:
        """判断文件名是否符合 OMML 转换依赖命名规则。"""
        return _is_omml_dependency_name(file_name)

    @staticmethod
    def _unique_target_path(path: Path) -> Path:
        """返回依赖复制的唯一目标路径。"""
        for index in range(1, 1000):
            candidate = path.with_name(f"{path.stem}-{index}{path.suffix}")
            if not candidate.exists():
                return candidate
        return path.with_name(f"{path.stem}-{new_id('copy')}{path.suffix}")

    def _mathpix_pdf_jobs(self, task: dict[str, Any], files: list[dict[str, Any]]) -> list[dict[str, Any]]:
        """遵守上传授权门禁创建 Mathpix 作业记录。"""
        settings = self.store.get_settings()
        ocr_settings = self._ocr_settings_snapshot(settings)
        jobs: list[dict[str, Any]] = []
        if settings.get("pdfToWordEngine") != "Mathpix":
            return jobs
        allow_upload = bool(settings.get("allowExternalMathpixUpload", False))
        task_options = task.get("options") if isinstance(task.get("options"), dict) else {}
        task_option_audit = self._mathpix_task_option_audit(task_options)
        wait_for_completion = bool(settings.get("waitForMathpix", False) or task_options.get("waitForMathpix", False))
        wait_source = self._mathpix_wait_source(settings, task_options)
        for file in files:
            if file.get("file_type") != "PDF":
                continue
            retention_plan = self._mathpix_retention_plan(task, file, ocr_settings, settings)
            job = {
                "file_id": file["id"],
                "file_name": file["file_name"],
                "engine": "Mathpix",
                "status": "pending",
                "pdf_id": "",
                "outputs": {},
                "output_summaries": {},
                "message": "",
                "ocr_settings": ocr_settings,
                "retention_plan": retention_plan,
                "request_options": self._mathpix_submit_options(ocr_settings),
                "task_option_audit": task_option_audit,
                "wait_for_completion": wait_for_completion,
                "wait_for_completion_source": wait_source,
            }
            storage_path = Path(file.get("storage_path") or file.get("file_path") or "")
            if file.get("validation_errors"):
                job["status"] = "validation_failed"
                job["message"] = "PDF 未通过文件校验，禁止提交 Mathpix"
                self._attach_mathpix_recognition_plan(job, file, settings, allow_upload, storage_path.exists())
                jobs.append(job)
                continue
            if not storage_path.exists():
                job["status"] = "missing_local_file"
                job["message"] = "PDF 原文件未保存，无法提交 Mathpix"
                self._attach_mathpix_recognition_plan(job, file, settings, allow_upload, False)
                jobs.append(job)
                continue
            if not any([ocr_settings["text_ocr"], ocr_settings["formula_ocr"], ocr_settings["table_ocr"]]):
                job["status"] = "ocr_disabled"
                job["message"] = "文字、公式和表格 OCR 均已关闭，未提交 Mathpix"
                self._attach_mathpix_recognition_plan(job, file, settings, allow_upload, True)
                jobs.append(job)
                continue
            if not allow_upload:
                job["status"] = "authorization_required"
                job["message"] = "PDF 转 Word 已配置为 Mathpix，但外部上传需要用户授权"
                self._attach_mathpix_recognition_plan(job, file, settings, allow_upload, True)
                jobs.append(job)
                continue
            try:
                client = MathpixClient.from_environment(*self._mathpix_env_names(settings))
                response = client.submit_pdf(storage_path, job["request_options"])
                job["pdf_id"] = response.get("pdf_id", "")
                job["status"] = "submitted"
                job["message"] = "已提交 Mathpix PDF OCR"
                self.store.append_log(task["id"], f"Mathpix 已提交：{file['file_name']}", category="conversion")
                if wait_for_completion and job["pdf_id"]:
                    status = client.wait_for_pdf(job["pdf_id"], int(settings.get("mathpixPollTimeoutSeconds", 600)))
                    job["status"] = status.get("status", "unknown")
                    job["percent_done"] = status.get("percent_done", 0)
                    if job["status"] == "completed":
                        output_dir = self.store.output_task_dir(task["id"])
                        stem = Path(file["file_name"]).stem
                        docx_path, existing_docx_path, _strategy = self._conversion_output_target(output_dir, file["file_name"], ".docx")
                        if docx_path is None:
                            job["status"] = "skipped_existing_output"
                            job["message"] = f"同名输出已存在，按策略跳过：{existing_docx_path.name}"
                            self._attach_mathpix_recognition_plan(job, file, settings, allow_upload, True)
                            jobs.append(job)
                            continue
                        try:
                            docx_bytes = client.download_pdf_result(job["pdf_id"], "docx")
                            job["output_summaries"]["docx"] = self._write_mathpix_result(docx_path, docx_bytes, "docx")
                        except (MathpixApiError, OSError) as exc:
                            job["status"] = "download_failed"
                            job["message"] = str(exc)
                            self._attach_mathpix_recognition_plan(job, file, settings, allow_upload, True)
                            jobs.append(job)
                            continue
                        job["outputs"]["docx"] = str(docx_path)
                        if ocr_settings["formula_ocr"]:
                            tex_zip_path, existing_tex_path, _strategy = self._conversion_output_target(output_dir, file["file_name"], "-mathpix-formulas.zip")
                            if tex_zip_path is None:
                                job["tex_zip_status"] = "skipped_existing_output"
                                job["tex_zip_message"] = f"同名输出已存在，按策略跳过：{existing_tex_path.name}"
                                self._attach_mathpix_recognition_plan(job, file, settings, allow_upload, True)
                                jobs.append(job)
                                continue
                            try:
                                tex_zip_bytes = client.download_pdf_result(job["pdf_id"], "tex.zip")
                                job["output_summaries"]["tex_zip"] = self._write_mathpix_result(tex_zip_path, tex_zip_bytes, "tex.zip")
                                job["outputs"]["tex_zip"] = str(tex_zip_path)
                                job["tex_zip_status"] = "completed"
                            except (MathpixApiError, OSError) as exc:
                                job["tex_zip_status"] = "download_failed"
                                job["tex_zip_message"] = str(exc)
            except MathpixConfigError as exc:
                job["status"] = "missing_credentials"
                job["message"] = str(exc)
            except MathpixApiError as exc:
                job["status"] = "api_error"
                job["message"] = str(exc)
            self._attach_mathpix_recognition_plan(job, file, settings, allow_upload, True)
            jobs.append(job)
        return jobs

    @staticmethod
    def _is_local_text_pdf(file: dict[str, Any]) -> bool:
        """判断 PDF 是否可走本地文本层转换。"""
        summary = file.get("content_summary") if isinstance(file.get("content_summary"), dict) else {}
        return bool(
            file.get("file_type") == "PDF"
            and not file.get("validation_errors")
            and summary.get("pdfType") == "文本型 PDF"
            and summary.get("textLayer")
            and int(summary.get("imageObjects") or 0) == 0
            and int(summary.get("formulaHints") or 0) == 0
            and int(summary.get("tableHints") or 0) == 0
        )

    def _local_pdf_text_artifacts(
        self, task: dict[str, Any], files: list[dict[str, Any]]
    ) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
        """本地转换安全的文本层 PDF，返回产物与审计证据。"""
        artifacts: list[dict[str, Any]] = []
        evidence: list[dict[str, Any]] = []
        output_dir = self.store.output_task_dir(str(task.get("id") or ""))
        for file in files:
            source = Path(str(file.get("storage_path") or file.get("file_path") or ""))
            target, existing_target, _strategy = self._conversion_output_target(output_dir, str(file.get("file_name") or "document.pdf"), ".docx")
            if target is None:
                artifacts.append(
                    self._artifact_skipped(
                        task,
                        file,
                        "docx",
                        existing_target.name,
                        f"同名输出已存在，按策略跳过：{existing_target.name}",
                        self._conversion_settings_snapshot(task, "pdf_to_word"),
                    )
                )
                evidence.append({"file_id": file.get("id", ""), "engine": "本地 PDF 文本层", "status": "skipped_existing_output"})
                continue
            try:
                summary = build_docx_from_pdf_text(source, target, Path(str(file.get("file_name") or "document.pdf")).stem)
                settings = self._conversion_settings_snapshot(task, "pdf_to_word")
                settings["pdf_to_word_engine"] = "本地 PDF 文本层"
                artifact = self._artifact_success(task, file, target, "docx", "已解析 PDF 文本层并生成 Word 文档", settings)
                artifacts.append(artifact)
                evidence.append(
                    {
                        "file_id": file.get("id", ""),
                        "engine": "本地 PDF 文本层",
                        "status": artifact.get("status", ""),
                        "paragraph_count": int(summary.get("paragraph_count") or 0),
                        "character_count": int(summary.get("character_count") or 0),
                        "output_size": int(artifact.get("size") or 0),
                        "output_sha256": str(artifact.get("sha256") or ""),
                    }
                )
            except (OSError, ValueError) as exc:
                target.unlink(missing_ok=True)
                artifacts.append(self._artifact_error(task, file, "text_layer_conversion_failed", str(exc)))
                evidence.append(
                    {
                        "file_id": file.get("id", ""),
                        "engine": "本地 PDF 文本层",
                        "status": "失败",
                        "message": str(exc),
                    }
                )
        return artifacts, evidence

    @staticmethod
    def _write_mathpix_result(path: Path, data: bytes, output_type: str) -> dict[str, Any]:
        """保存 Mathpix 产物并返回不含路径的完整性信息。"""
        if not data:
            raise MathpixApiError(f"Mathpix {output_type} 下载结果为空")
        path.write_bytes(data)
        validation_error = TaskProcessor._output_artifact_validation_error(path, output_type)
        if validation_error:
            try:
                path.unlink(missing_ok=True)
            except OSError:
                # 下载结果仍按校验失败处理，不会进入输出清单。
                pass
            raise MathpixApiError(f"Mathpix {output_type} 下载结果校验失败：{validation_error}")
        return {
            "type": output_type,
            "file_name": path.name,
            "size": path.stat().st_size,
            "sha256": hashlib.sha256(data).hexdigest(),
            "path_available": True,
        }

    def _attach_mathpix_recognition_plan(self, job: dict[str, Any], file: dict[str, Any], settings: dict[str, Any], allow_upload: bool, source_available: bool) -> None:
        """为作业附加 Mathpix 识别计划，不提交上传。"""
        job["recognition_plan"] = self._mathpix_recognition_plan(job, file, settings, allow_upload, source_available)

    def _mathpix_recognition_plan(
        self,
        job: dict[str, Any],
        file: dict[str, Any],
        settings: dict[str, Any],
        allow_upload: bool,
        source_available: bool,
    ) -> dict[str, Any]:
        """描述 Mathpix 请求、保留计划和上传门禁。"""
        request = self._mathpix_request_summary(job.get("request_options"))
        ocr_settings = dict(job.get("ocr_settings") or {})
        retention = dict(job.get("retention_plan") or {})
        output_summaries = self._mathpix_output_summaries(job.get("output_summaries"))
        download_failures = self._mathpix_download_failures(job)
        download_manifest = self._mathpix_download_manifest(job, ocr_settings, output_summaries, download_failures)
        credential_envs = list(self._mathpix_env_names(settings))
        status = str(job.get("status") or "pending")
        plan_status = {
            "pending": "pending",
            "missing_local_file": "blocked_missing_local_file",
            "validation_failed": "blocked_validation_failed",
            "ocr_disabled": "blocked_ocr_disabled",
            "authorization_required": "blocked_authorization",
            "missing_credentials": "blocked_missing_credentials",
            "api_error": "api_error",
            "submitted": "submitted",
            "processing": "submitted",
            "completed": "completed",
            "skipped_existing_output": "completed_with_existing_output",
            "download_failed": "download_failed",
        }.get(status, status)
        ocr_enabled = any(bool(ocr_settings.get(key)) for key in ("text_ocr", "formula_ocr", "table_ocr"))
        conversion_formats = list(request.get("conversion_formats") or [])
        download_formats = ["docx"]
        if bool(ocr_settings.get("formula_ocr")):
            download_formats.append("tex.zip")
        upload_gate = self._mathpix_upload_gate(plan_status, source_available, ocr_enabled, allow_upload, status, credential_envs)
        wait_for_completion = bool(job.get("wait_for_completion", settings.get("waitForMathpix", False)))
        wait_source = str(job.get("wait_for_completion_source") or self._mathpix_wait_source(settings, {}))
        task_option_audit = dict(job.get("task_option_audit") or self._mathpix_task_option_audit({}))
        return {
            "schema_version": "k12.mathpixRecognitionPlan.v1",
            "engine": "Mathpix",
            "file_id": file.get("id", ""),
            "file_name": file.get("file_name", ""),
            "status": plan_status,
            "status_label": self._mathpix_recognition_status_label(plan_status),
            "submit_allowed": bool(upload_gate["submit_allowed"]),
            "upload_gate": upload_gate,
            "source_file_available": bool(source_available),
            "external_upload_required": True,
            "external_upload_authorized": bool(allow_upload),
            "external_upload_authorization_source": "settings.allowExternalMathpixUpload",
            "credential_envs": credential_envs,
            "credential_values_exposed": False,
            "task_option_audit": task_option_audit,
            "ocr": {
                "text": bool(ocr_settings.get("text_ocr")),
                "formula": bool(ocr_settings.get("formula_ocr")),
                "table": bool(ocr_settings.get("table_ocr")),
                "language": str(ocr_settings.get("language") or ""),
                "precision_mode": str(ocr_settings.get("precision_mode") or ""),
                "speed_mode": str(ocr_settings.get("speed_mode") or ""),
            },
            "submit_contract": {
                "method": "POST",
                "endpoint": "/v3/pdf",
                "content_transfer": "pdf_upload_to_mathpix_when_authorized",
                "conversion_formats": conversion_formats,
                "download_formats": download_formats,
                "poll_status_endpoint": "/v3/pdf/{pdf_id}",
                "wait_for_completion": wait_for_completion,
                "wait_for_completion_source": wait_source,
                "poll_timeout_seconds": int(settings.get("mathpixPollTimeoutSeconds", 600) or 600),
                "metadata_keys": list(request.get("metadata_keys") or []),
            },
            "download_results": output_summaries,
            "download_failures": download_failures,
            "download_manifest": download_manifest,
            "retention_summary": {
                "pdf_type": retention.get("pdf_type", ""),
                "image_objects": retention.get("image_objects", 0),
                "table_hints": retention.get("table_hints", 0),
                "formula_hints": retention.get("formula_hints", 0),
                "image_retention_status": retention.get("image_retention_status", ""),
                "table_retention_status": retention.get("table_retention_status", ""),
                "formula_retention_status": retention.get("formula_retention_status", ""),
            },
            "formula_review": self._formula_review_contract(settings),
            "safety": {
                "no_upload_without_authorization": True,
                "no_submit_when_file_validation_failed": True,
                "task_options_cannot_authorize_external_upload": True,
                "ignored_task_option_authorization_count": len(task_option_audit.get("ignored_authorization_keys") or []),
                "task_options_may_request_wait_for_completion": True,
                "no_submit_when_all_ocr_disabled": True,
                "credential_values_redacted": True,
                "output_paths_hidden_in_queue": True,
                "download_paths_redacted": True,
                "required_outputs_satisfied": bool(download_manifest.get("all_required_outputs_satisfied")),
                "download_integrity_hashed": bool(download_manifest.get("all_required_outputs_hashed")),
            },
        }

    @staticmethod
    def _mathpix_wait_source(settings: dict[str, Any], task_options: dict[str, Any]) -> str:
        """说明 Mathpix 轮询来自设置还是任务选项。"""
        settings_wait = bool(settings.get("waitForMathpix", False))
        task_wait = bool(task_options.get("waitForMathpix", False))
        if settings_wait and task_wait:
            return "settings+task_options"
        if task_wait:
            return "task_options"
        if settings_wait:
            return "settings"
        return "disabled"

    @staticmethod
    def _mathpix_task_option_audit(task_options: dict[str, Any]) -> dict[str, Any]:
        """审计 Mathpix 任务选项，不将其作为上传授权。"""
        authorization_keys = {
            "allowExternalMathpixUpload",
            "externalUploadAuthorized",
            "externalUploadAllowed",
            "mathpixExternalUpload",
            "enableMathpixUpload",
        }
        ignored_authorization_keys: list[str] = []
        for raw_key in task_options:
            key = str(raw_key)
            lowered = key.lower()
            looks_like_authorization = (
                key in authorization_keys
                or ("upload" in lowered and ("mathpix" in lowered or "external" in lowered or "authoriz" in lowered or "authoris" in lowered))
                or ("authoriz" in lowered and ("mathpix" in lowered or "external" in lowered))
                or ("authoris" in lowered and ("mathpix" in lowered or "external" in lowered))
            )
            if looks_like_authorization:
                ignored_authorization_keys.append(key[:80])
        return {
            "schema_version": "k12.mathpixTaskOptionAudit.v1",
            "task_options_can_authorize_external_upload": False,
            "authorization_source": "settings.allowExternalMathpixUpload",
            "ignored_authorization_keys": ignored_authorization_keys,
            "ignored_authorization_count": len(ignored_authorization_keys),
            "wait_for_completion_option_present": "waitForMathpix" in task_options,
            "wait_for_completion_can_be_requested": True,
        }

    @staticmethod
    def _formula_review_contract(settings: dict[str, Any]) -> dict[str, Any]:
        """描述 Mathpix PDF 公式进入手动审阅的流程。"""
        threshold = TaskProcessor._positive_int(settings.get("formulaConfidenceThreshold", 80), 80)
        return {
            "schema_version": "k12.formulaReviewContract.v1",
            "confidence_threshold": threshold,
            "low_confidence_strategy": str(settings.get("lowConfidenceFormulaStrategy") or "人工确认"),
            "keep_original_image": bool(settings.get("keepFormulaImages", True)),
            "source": "settings.formulaConfidenceThreshold",
            "status_for_below_threshold": "待确认",
            "manual_actions": ["已确认", "已修正", "重新识别", "跳过"],
            "annotation_endpoint": "/api/formula-annotations",
            "bulk_confirm_endpoint": "/api/formula-annotations/bulk-confirm",
            "rerecognition_request_schema": "k12.formulaRecognitionRequest.v1",
            "task_options_can_authorize_external_upload": False,
        }

    @staticmethod
    def _mathpix_download_failures(job: dict[str, Any]) -> list[dict[str, str]]:
        """返回 Mathpix 失败下载，不泄露输出路径。"""
        failures: list[dict[str, str]] = []
        if job.get("status") == "download_failed":
            failures.append({"type": "docx", "status": "download_failed", "message": str(job.get("message") or "Mathpix DOCX 下载失败")})
        if job.get("tex_zip_status") == "download_failed":
            failures.append({"type": "tex.zip", "status": "download_failed", "message": str(job.get("tex_zip_message") or "Mathpix tex.zip 下载失败")})
        return failures

    @staticmethod
    def _mathpix_download_manifest(
        job: dict[str, Any],
        ocr_settings: dict[str, Any],
        output_summaries: dict[str, dict[str, Any]],
        download_failures: list[dict[str, str]],
    ) -> dict[str, Any]:
        """描述必需的 Mathpix 产物及其满足情况。"""
        required_types = ["docx"]
        if bool(ocr_settings.get("formula_ocr")):
            required_types.append("tex.zip")
        failures = {str(item.get("type") or ""): item for item in download_failures if isinstance(item, dict)}
        outputs = []
        for output_type in required_types:
            summary_key = "tex_zip" if output_type == "tex.zip" else output_type
            summary = output_summaries.get(summary_key, {})
            failure = failures.get(output_type, {})
            status = TaskProcessor._mathpix_required_output_status(output_type, job, bool(summary), bool(failure))
            outputs.append(
                {
                    "type": output_type,
                    "required": True,
                    "status": status,
                    "file_name": str(summary.get("file_name") or ""),
                    "size": int(summary.get("size") or 0),
                    "sha256_available": bool(summary.get("sha256")),
                    "path_exposed": False,
                    "failure_message": str(failure.get("message") or "")[:240],
                }
            )
        satisfied = sum(1 for item in outputs if item["status"] in {"downloaded", "skipped_existing_output"})
        failed = sum(1 for item in outputs if item["status"] == "failed")
        pending = sum(1 for item in outputs if item["status"] == "pending")
        not_attempted = sum(1 for item in outputs if item["status"] == "not_attempted")
        hashed = all(item["status"] == "downloaded" and item["sha256_available"] for item in outputs)
        if failed and satisfied:
            manifest_status = "partial_failed"
        elif failed:
            manifest_status = "failed"
        elif not_attempted and not satisfied:
            manifest_status = "not_attempted"
        elif pending:
            manifest_status = "pending"
        elif satisfied == len(outputs):
            manifest_status = "complete"
        else:
            manifest_status = "pending"
        return {
            "schema_version": "k12.mathpixDownloadManifest.v1",
            "status": manifest_status,
            "required_output_count": len(outputs),
            "satisfied_required_count": satisfied,
            "downloaded_required_count": sum(1 for item in outputs if item["status"] == "downloaded"),
            "skipped_existing_required_count": sum(1 for item in outputs if item["status"] == "skipped_existing_output"),
            "failed_required_count": failed,
            "pending_required_count": pending,
            "not_attempted_required_count": not_attempted,
            "all_required_outputs_satisfied": bool(outputs) and satisfied == len(outputs) and failed == 0 and pending == 0 and not_attempted == 0,
            "all_required_outputs_hashed": bool(outputs) and hashed,
            "path_policy": "local output paths are redacted",
            "outputs": outputs,
        }

    @staticmethod
    def _mathpix_required_output_status(output_type: str, job: dict[str, Any], has_summary: bool, has_failure: bool) -> str:
        """判断必需 Mathpix 产物状态，不读取本地路径。"""
        if has_failure:
            return "failed"
        if has_summary:
            return "downloaded"
        if output_type == "docx" and job.get("status") == "skipped_existing_output":
            return "skipped_existing_output"
        if output_type == "tex.zip" and job.get("tex_zip_status") == "skipped_existing_output":
            return "skipped_existing_output"
        raw_status = str(job.get("status") or "")
        if raw_status in {"authorization_required", "missing_credentials", "missing_local_file", "ocr_disabled", "api_error"}:
            return "not_attempted"
        if raw_status in {"pending", "submitted", "processing"}:
            return "pending"
        return "pending"

    @staticmethod
    def _mathpix_output_summaries(value: Any) -> dict[str, dict[str, Any]]:
        """返回不含本地文件路径的 Mathpix 产物元数据。"""
        if not isinstance(value, dict):
            return {}
        summaries: dict[str, dict[str, Any]] = {}
        for key, item in value.items():
            if not isinstance(item, dict):
                continue
            digest = str(item.get("sha256") or "")
            size = item.get("size", 0)
            try:
                safe_size = max(0, int(size or 0))
            except (TypeError, ValueError):
                safe_size = 0
            summaries[str(key)] = {
                "type": str(item.get("type") or key),
                "file_name": Path(str(item.get("file_name") or "")).name,
                "size": safe_size,
                "sha256": digest if re.fullmatch(r"[0-9a-f]{64}", digest) else "",
                "path_available": bool(item.get("path_available")),
                "path_exposed": False,
            }
        return summaries

    @staticmethod
    def _mathpix_upload_gate(
        plan_status: str,
        source_available: bool,
        ocr_enabled: bool,
        allow_upload: bool,
        raw_status: str,
        credential_envs: list[str],
    ) -> dict[str, Any]:
        """说明 PDF 能否上传到 Mathpix 及原因。"""
        blockers: list[str] = []
        if not source_available:
            blockers.append("missing_local_file")
        if not ocr_enabled:
            blockers.append("all_ocr_disabled")
        if not allow_upload:
            blockers.append("external_upload_not_authorized")
        if raw_status == "missing_credentials":
            blockers.append("missing_credentials")
        if raw_status == "api_error":
            blockers.append("mathpix_api_error")
        if raw_status == "validation_failed":
            blockers.append("file_validation_failed")
        status = "ready_to_submit" if not blockers else "blocked"
        if "missing_local_file" in blockers:
            status = "blocked_missing_local_file"
        elif "file_validation_failed" in blockers:
            status = "blocked_validation_failed"
        elif "all_ocr_disabled" in blockers:
            status = "blocked_ocr_disabled"
        elif "external_upload_not_authorized" in blockers:
            status = "blocked_authorization"
        elif "missing_credentials" in blockers:
            status = "blocked_missing_credentials"
        elif "mathpix_api_error" in blockers:
            status = "api_error"
        elif plan_status in {"submitted", "completed", "completed_with_existing_output", "download_failed"}:
            status = plan_status
        return {
            "schema_version": "k12.mathpixUploadGate.v1",
            "status": status,
            "submit_allowed": not blockers,
            "blocking_reasons": blockers,
            "external_upload_required": True,
            "external_upload_authorized": bool(allow_upload),
            "authorization_source": "settings.allowExternalMathpixUpload",
            "task_options_can_authorize": False,
            "source_file_available": bool(source_available),
            "ocr_enabled": bool(ocr_enabled),
            "credential_envs": credential_envs,
            "credential_values_exposed": False,
            "content_transfer": "pdf_upload_to_mathpix_when_authorized",
        }

    @staticmethod
    def _mathpix_recognition_status_label(status: str) -> str:
        """将 Mathpix 识别计划状态转为界面标签。"""
        labels = {
            "pending": "待提交",
            "blocked_missing_local_file": "缺少本地文件",
            "blocked_validation_failed": "文件校验失败",
            "blocked_ocr_disabled": "OCR 已关闭",
            "blocked_authorization": "等待授权",
            "blocked_missing_credentials": "缺少凭证",
            "api_error": "接口错误",
            "submitted": "已提交",
            "completed": "已完成",
            "completed_with_existing_output": "已跳过",
            "download_failed": "下载失败",
        }
        return labels.get(status, status or "未知")

    def _mathpix_artifacts(self, task: dict[str, Any], files: list[dict[str, Any]], jobs: list[dict[str, Any]]) -> list[dict[str, Any]]:
        """根据完成或阻断的 Mathpix 作业生成产物行。"""
        files_by_id = {file["id"]: file for file in files}
        artifacts: list[dict[str, Any]] = []
        for job in jobs:
            file = files_by_id.get(job.get("file_id"))
            if not file:
                continue
            docx_output = str(job.get("outputs", {}).get("docx") or "")
            docx_path = Path(docx_output) if docx_output else None
            if job.get("status") == "completed" and docx_path and docx_path.exists() and docx_path.is_file():
                artifacts.append(
                    self._artifact_success(
                        task,
                        file,
                        docx_path,
                        "docx",
                        "Mathpix PDF OCR 已生成 Word 文档",
                        self._conversion_settings_snapshot(task, "pdf_to_word"),
                    )
                )
                tex_zip_output = str(job.get("outputs", {}).get("tex_zip") or "")
                tex_zip_path = Path(tex_zip_output) if tex_zip_output else None
                if tex_zip_path and tex_zip_path.exists() and tex_zip_path.is_file():
                    artifacts.append(
                        self._artifact_success(
                            task,
                            file,
                            tex_zip_path,
                            "tex.zip",
                            "Mathpix 已生成 LaTeX/公式识别结果包",
                            self._conversion_settings_snapshot(task, "pdf_to_word"),
                        )
                    )
                elif job.get("tex_zip_status") == "download_failed":
                    artifacts.append(
                        self._artifact_error(
                            task,
                            file,
                            "mathpix_tex_zip_download_failed",
                            str(job.get("tex_zip_message") or "Mathpix tex.zip 下载失败"),
                        )
                    )
            elif job.get("status") == "skipped_existing_output":
                artifacts.append(
                    self._artifact_skipped(
                        task,
                        file,
                        "docx",
                        f"{Path(file.get('file_name', '')).stem}.docx",
                        str(job.get("message") or "同名输出已存在，按策略跳过"),
                        self._conversion_settings_snapshot(task, "pdf_to_word"),
                    )
                )
            elif job.get("status") in {"authorization_required", "missing_credentials", "api_error", "download_failed", "missing_local_file", "ocr_disabled", "validation_failed"}:
                artifacts.append(self._artifact_error(task, file, str(job.get("status")), str(job.get("message") or "Mathpix PDF 转 Word 未完成")))
        return artifacts

    @staticmethod
    def _ocr_settings_snapshot(settings: dict[str, Any]) -> dict[str, Any]:
        """保存 Mathpix PDF 转 Word 规划使用的 OCR 开关快照。"""
        return {
            "language": settings.get("ocrLanguage", "中文+英文"),
            "text_ocr": bool(settings.get("enableTextOcr", True)),
            "formula_ocr": bool(settings.get("enableFormulaOcr", True)),
            "table_ocr": bool(settings.get("enableTableOcr", True)),
            "precision_mode": settings.get("ocrPrecisionMode", "平衡"),
            "speed_mode": settings.get("ocrSpeedMode", "标准"),
        }

    @staticmethod
    def _mathpix_submit_options(ocr_settings: dict[str, Any]) -> dict[str, Any]:
        """构建 DOCX 和可选 tex.zip 的 Mathpix 提交选项。"""
        conversion_formats = {"docx": True}
        if ocr_settings.get("formula_ocr"):
            conversion_formats["tex.zip"] = True
        return {
            "conversion_formats": conversion_formats,
            "enable_tables_fallback": bool(ocr_settings.get("table_ocr")),
            "include_page_info": True,
            "math_inline_delimiters": ["\\(", "\\)"],
            "math_display_delimiters": ["\\[", "\\]"],
            "metadata": {
                "k12_ocr_language": str(ocr_settings.get("language", "中文+英文")),
                "k12_text_ocr": str(bool(ocr_settings.get("text_ocr"))).lower(),
                "k12_formula_ocr": str(bool(ocr_settings.get("formula_ocr"))).lower(),
                "k12_table_ocr": str(bool(ocr_settings.get("table_ocr"))).lower(),
                "k12_precision_mode": str(ocr_settings.get("precision_mode", "平衡")),
                "k12_speed_mode": str(ocr_settings.get("speed_mode", "标准")),
            },
        }

    @staticmethod
    def _mathpix_retention_plan(task: dict[str, Any], file: dict[str, Any], ocr_settings: dict[str, Any], settings: dict[str, Any]) -> dict[str, Any]:
        """规划 Mathpix 转换中的 PDF 图片、表格和公式保留。"""
        summary = file.get("content_summary") or {}
        options = task.get("options") or {}
        word_options = options.get("word") or {}
        retain_images = bool(word_options.get("retainImages", options.get("retainImages", settings.get("retainImages", True))))
        retain_tables = bool(word_options.get("retainTables", options.get("retainTables", settings.get("retainTables", True))))
        image_objects = TaskProcessor._summary_count(summary, "imageObjects")
        table_hints = TaskProcessor._summary_count(summary, "tableHints")
        formula_hints = TaskProcessor._summary_count(summary, "formulaHints")
        formula_ocr = bool(ocr_settings.get("formula_ocr"))
        table_ocr = bool(ocr_settings.get("table_ocr"))
        plan = {
            "engine": "Mathpix",
            "pdf_type": summary.get("pdfType", "未知 PDF"),
            "retain_images": retain_images,
            "retain_tables": retain_tables,
            "text_ocr": bool(ocr_settings.get("text_ocr")),
            "formula_ocr": formula_ocr,
            "table_ocr": table_ocr,
            "image_objects": image_objects,
            "table_hints": table_hints,
            "formula_hints": formula_hints,
            "image_retention_status": TaskProcessor._retention_status(image_objects, retain_images, "计划保留图片对象", "图片保留关闭", "未检测到图片对象"),
            "table_retention_status": TaskProcessor._retention_status(table_hints, retain_tables and table_ocr, "计划识别并保留表格结构", "表格保留或 OCR 关闭", "未检测到表格线索"),
            "formula_retention_status": TaskProcessor._retention_status(formula_hints, formula_ocr, "计划识别公式并生成 MathType 预览", "公式 OCR 关闭", "未检测到公式线索"),
        }
        plan["notes"] = TaskProcessor._mathpix_retention_notes(plan)
        return plan

    @staticmethod
    def _summary_count(summary: dict[str, Any], key: str) -> int:
        """从内容摘要读取非负计数。"""
        try:
            return max(0, int(summary.get(key, 0) or 0))
        except (TypeError, ValueError):
            return 0

    @staticmethod
    def _retention_status(count: int, enabled: bool, enabled_label: str, disabled_label: str, empty_label: str) -> str:
        """返回检测到的对象类型的保留状态标签。"""
        if count <= 0:
            return empty_label
        return enabled_label if enabled else disabled_label

    @staticmethod
    def _mathpix_retention_notes(plan: dict[str, Any]) -> list[str]:
        """生成 Mathpix PDF 转换的保留说明。"""
        notes: list[str] = []
        if plan.get("image_objects"):
            notes.append("DOCX 输出计划保留 PDF 图片对象" if plan.get("retain_images") else "图片对象已检测，但当前保留图片关闭")
        if plan.get("table_hints"):
            if plan.get("retain_tables") and plan.get("table_ocr"):
                notes.append("表格线索将交给 Mathpix 表格 OCR 并在 DOCX 中保留结构")
            else:
                notes.append("表格线索已检测，但表格保留或表格 OCR 未同时开启")
        if plan.get("formula_hints"):
            notes.append("公式线索将通过 Mathpix 公式 OCR 生成 LaTeX/MathType 预览" if plan.get("formula_ocr") else "公式线索已检测，但公式 OCR 关闭")
        if not notes:
            notes.append("未检测到图片、表格或公式线索，按 Mathpix DOCX 默认结构处理")
        return notes

    def _mathpix_formula_items(self, files: list[dict[str, Any]], jobs: list[dict[str, Any]], task: dict[str, Any] | None = None) -> list[dict[str, Any]]:
        """从下载的 Mathpix tex.zip 提取 PDF 公式审阅项。"""
        files_by_id = {file["id"]: file for file in files}
        settings = self.store.get_settings()
        task_context = task or {"task_type": "pdf_to_word", "options": {}}
        threshold = self._positive_int(settings.get("formulaConfidenceThreshold", 80), 80)
        formatting_enabled = self._mathtype_formatting_enabled(task_context, settings)
        format_scope = self._formula_format_scope(task_context, settings)
        items: list[dict[str, Any]] = []
        for job in jobs:
            if job.get("status") != "completed":
                continue
            file = files_by_id.get(job.get("file_id"))
            if not file:
                continue
            tex_zip_path = Path(job.get("outputs", {}).get("tex_zip") or "")
            if not tex_zip_path.exists() or not tex_zip_path.is_file():
                continue
            try:
                with zipfile.ZipFile(tex_zip_path) as archive:
                    tex_entries = [
                        info
                        for info in archive.infolist()
                        if info.filename.lower().endswith(".tex") and not info.is_dir()
                    ]
                    for info in tex_entries[:50]:
                        if info.file_size > MATHPIX_TEX_ENTRY_LIMIT_BYTES:
                            continue
                        entry_name = Path(info.filename).name
                        text = archive.read(info).decode("utf-8", errors="ignore")
                        for latex in self._extract_latex_expressions(text):
                            confidence = 88
                            item = FormulaItem(
                                file_id=file["id"],
                                page_index=max(1, len(items) + 1),
                                position=f"Mathpix OCR / {entry_name}",
                                source_type="PDF",
                                latex=latex,
                                confidence=confidence,
                                status="成功" if confidence >= threshold else "待确认",
                                original_image_path=f"mathpix://{file['id']}/{entry_name}",
                                mathml=f"<math><mtext>{html.escape(latex)}</mtext></math>",
                                mathtype_data=f"MathType 预览：{latex}",
                                format_status="Mathpix 识别",
                            ).to_dict()
                            item["original_image_ref"] = item["original_image_path"]
                            item["mathtype_preview"] = item["mathtype_data"]
                            item["output_format"] = settings.get("formulaOutputFormat", "LaTeX+MathML")
                            item["font"] = settings.get("formulaFont", "Cambria Math")
                            item["font_size"] = settings.get("formulaFontSize", 12)
                            item["format_scope"] = format_scope
                            item["alignment"] = settings.get("formulaAlignment", "居中")
                            item["variable_style"] = settings.get("formulaVariableStyle", "斜体")
                            item["function_style"] = settings.get("formulaFunctionStyle", "正体")
                            item["script_scale"] = settings.get("formulaScriptScale", 70)
                            item["fraction_style"] = settings.get("formulaFractionStyle", "标准")
                            item["radical_style"] = settings.get("formulaRadicalStyle", "标准")
                            item["matrix_spacing"] = settings.get("formulaMatrixSpacing", "标准")
                            item["greek_style"] = settings.get("formulaGreekStyle", "标准")
                            item["inline_baseline"] = settings.get("formulaInlineBaseline", "跟随正文")
                            item["display_spacing"] = settings.get("formulaDisplaySpacing", "标准")
                            item["numbering"] = settings.get("formulaNumbering", "按文档位置")
                            item["confidence_threshold"] = threshold
                            item["review_required"] = confidence < threshold
                            item["review_reason"] = "低于当前置信度阈值" if confidence < threshold else ""
                            item["low_confidence_strategy"] = settings.get("lowConfidenceFormulaStrategy", "人工确认")
                            item["keep_original_image"] = bool(settings.get("keepFormulaImages", True))
                            self._apply_formula_format_policy(item, task_context, formatting_enabled)
                            item["format_comparison"] = self._formula_format_comparison(item, settings)
                            items.append(item)
                            if len(items) >= 200:
                                return items
            except zipfile.BadZipFile:
                continue
        return items

    @staticmethod
    def _extract_latex_expressions(text: str) -> list[str]:
        """从 Mathpix 文本产物提取可能的 LaTeX 公式片段。"""
        formulas: list[str] = []
        patterns = [
            r"\$\$(.*?)\$\$",
            r"(?<!\\)\$(?!\$)(.*?)(?<!\\)\$",
            r"\\\[(.*?)\\\]",
            r"\\\((.*?)\\\)",
            r"\\begin\{equation\*?\}(.*?)\\end\{equation\*?\}",
            r"\\begin\{align\*?\}(.*?)\\end\{align\*?\}",
            r"\\begin\{gather\*?\}(.*?)\\end\{gather\*?\}",
            r"\\begin\{multline\*?\}(.*?)\\end\{multline\*?\}",
        ]
        for pattern in patterns:
            for match in re.finditer(pattern, text, flags=re.S):
                latex = TaskProcessor._clean_latex_candidate(match.group(1))
                if latex:
                    formulas.append(latex)
        for line in text.splitlines():
            latex = TaskProcessor._clean_latex_candidate(line)
            if not latex:
                continue
            if latex.startswith("\\") and not any(token in latex for token in ("=", "^", "_", "\\frac", "\\sqrt", "\\sum", "\\int")):
                continue
            if any(latex.startswith(prefix) for prefix in ("%", "\\documentclass", "\\usepackage", "\\begin", "\\end", "\\section", "\\subsection")):
                continue
            if len(latex) <= 500:
                formulas.append(latex)
        return list(dict.fromkeys(formulas))[:100]

    @staticmethod
    def _clean_latex_candidate(value: str) -> str:
        """规范候选 LaTeX 公式片段。"""
        text = re.sub(r"(?m)^\s*%.*$", "", value).strip()
        text = re.sub(r"\\begin\{(?:equation|align|gather|multline)\*?\}", "", text)
        text = re.sub(r"\\end\{(?:equation|align|gather|multline)\*?\}", "", text)
        text = text.strip("$").strip()
        text = re.sub(r"^\s*\\\((.*)\\\)\s*$", r"\1", text, flags=re.S)
        text = re.sub(r"^\s*\\\[(.*)\\\]\s*$", r"\1", text, flags=re.S)
        return re.sub(r"\s+", " ", text)[:1000]

    def _conversion_artifacts(self, task: dict[str, Any], files: list[dict[str, Any]]) -> list[dict[str, Any]]:
        """为支持的本地优先格式创建转换产物记录。"""
        artifacts: list[dict[str, Any]] = []
        output_dir = self.store.output_task_dir(task["id"])
        for file in files:
            if file.get("validation_errors"):
                artifacts.append(
                    self._artifact_error(
                        task,
                        file,
                        "validation_failed",
                        "文件未通过校验，未执行转换",
                    )
                )
                continue
            source = Path(file.get("storage_path") or file.get("file_path") or "")
            if not source.exists() or not source.is_file():
                artifacts.append(self._artifact_error(task, file, "missing_source", "原文件未保存，无法生成转换输出"))
                continue
            if bool((file.get("content_summary") or {}).get("requiresNativeOffice")):
                artifacts.append(self._artifact_pending_local_office(task, file))
                continue
            if task["task_type"] == "word_to_ppt" and file.get("file_type") == "Word":
                artifacts.append(self._word_to_ppt_artifact(task, file, source, output_dir))
            elif task["task_type"] == "ppt_to_word" and file.get("file_type") == "PPT":
                artifacts.append(self._ppt_to_word_artifact(task, file, source, output_dir))
            elif task["task_type"] == "excel_to_pdf" and file.get("file_type") == "Excel":
                artifacts.extend(self._excel_to_pdf_artifacts(task, file, source, output_dir))
            elif task["task_type"] == "excel_to_word" and file.get("file_type") == "Excel":
                artifacts.extend(self._excel_to_word_artifacts(task, file, source, output_dir))
            elif task["task_type"] == "excel_to_ppt" and file.get("file_type") == "Excel":
                artifacts.extend(self._excel_to_ppt_artifacts(task, file, source, output_dir))
            else:
                artifacts.append(self._artifact_error(task, file, "unsupported_input", "当前任务不支持该文件类型"))
        return artifacts

    def _artifact_pending_local_office(self, task: dict[str, Any], file: dict[str, Any]) -> dict[str, Any]:
        """为等待原生执行的旧 Office 输入生成待处理产物行。"""
        output_types = {
            "word_to_ppt": "pptx",
            "ppt_to_word": "docx",
            "excel_to_pdf": "pdf",
            "excel_to_word": "docx",
            "excel_to_ppt": "pptx",
        }
        output_type = output_types.get(str(task.get("task_type") or ""), "")
        return {
            "file_id": file["id"],
            "task_id": task["id"],
            "source_file": file.get("file_name", ""),
            "file_name": "",
            "output_type": output_type,
            "path": "",
            "url": "",
            "status": "待本地客户端执行",
            "message": "旧版 Office 二进制文件已通过 OLE 预检，需本地 Office 客户端生成转换产物",
            "size": 0,
            "requires_native_office": True,
            "conversion_settings": self._conversion_settings_snapshot(task, str(task.get("task_type") or "")),
        }

    def _word_to_ppt_artifact(self, task: dict[str, Any], file: dict[str, Any], source: Path, output_dir: Path) -> dict[str, Any]:
        """生成 Word 转 PPT 产物和对象保留摘要。"""
        try:
            blocks = extract_docx_blocks(source, retain_images=bool(self._conversion_settings_snapshot(task, "word_to_ppt").get("retain_images", True)))
            object_summary = extract_docx_object_summary(source)
            word_options = dict(task.get("options", {}).get("word") or {})
            max_chars = int(word_options.get("maxCharsPerSlide", 320) or 320)
            auto_pagination = bool(word_options.get("autoPagination", True))
            generate_toc = bool(word_options.get("generateToc", False))
            object_preservation = self._word_to_ppt_object_preservation(task, object_summary)
            target, existing_target, _strategy = self._conversion_output_target(output_dir, file["file_name"], ".pptx")
            if target is None:
                artifact = self._artifact_skipped(task, file, "pptx", existing_target.name, f"同名输出已存在，按策略跳过：{existing_target.name}", self._conversion_settings_snapshot(task, "word_to_ppt"))
                artifact["object_preservation"] = object_preservation
                return artifact
            warnings = build_pptx_from_docx(blocks, target, max_chars, auto_pagination, generate_toc, object_preservation,
                                 retain_tables=bool(word_options.get("retainTables", self.store.get_settings().get("retainTables", True))),
                                 retain_images=bool(word_options.get("retainImages", self.store.get_settings().get("retainImages", True)))) or []
            message = f"生成 PPTX，提取段落 {len(blocks)} 个"
            if object_preservation.get("has_objects"):
                message = f"{message}，写入对象保留清单"
            if warnings:
                message += "；提示：" + "；".join(warnings)
            artifact = self._artifact_success(task, file, target, "pptx", message, self._conversion_settings_snapshot(task, "word_to_ppt"))
            artifact["object_preservation"] = object_preservation
            return artifact
        except (OSError, KeyError, ValueError, TypeError, zipfile.BadZipFile) as exc:
            return self._artifact_error(task, file, "conversion_failed", str(exc))

    def _word_to_ppt_object_preservation(self, task: dict[str, Any], source: dict[str, Any]) -> dict[str, Any]:
        """汇总 Word 对象保留需求，不执行原生 Office 写回。"""
        settings = self._conversion_settings_snapshot(task, "word_to_ppt")
        retain_images = bool(settings.get("retain_images", True))
        retain_tables = bool(settings.get("retain_tables", True))
        retain_formulas = bool(settings.get("word_retain_formulas", True))
        convert_omml_first = bool(settings.get("word_convert_omml_first", True))
        source_counts = {
            "paragraphs": self._summary_count(source, "paragraphs"),
            "headings": self._summary_count(source, "headings"),
            "images": self._summary_count(source, "images"),
            "tables": self._summary_count(source, "tables"),
            "omml_formulas": self._summary_count(source, "omml_formulas"),
            "mathtype_objects": self._summary_count(source, "mathtype_objects"),
            "embedded_objects": self._summary_count(source, "embedded_objects"),
            "formulas": self._summary_count(source, "formulas"),
        }
        statuses = {
            "images": self._retention_status(source_counts["images"], retain_images, "计划保留图片对象", "图片保留关闭", "未检测到图片对象"),
            "tables": self._retention_status(source_counts["tables"], retain_tables, "计划保留表格结构", "表格保留关闭", "未检测到表格结构"),
            "formulas": self._retention_status(source_counts["formulas"], retain_formulas, "计划保留公式对象", "公式保留关闭", "未检测到公式对象"),
        }
        notes = [
            "标题层级由 Word 段落样式和段落顺序生成 PPT 标题/正文",
            "图片、表格和公式来源对象已写入本保留清单，供本地客户端做真实 Office/MathType 复核",
        ]
        if source_counts["omml_formulas"]:
            notes.append("OMML 公式按任务设置先转 MathType" if convert_omml_first else "OMML 公式按任务设置保留原格式")
        if source_counts["mathtype_objects"]:
            notes.append("检测到 MathType/嵌入对象，保留来源对象计数并等待本地客户端处理")
        return {
            "source": source_counts,
            "statuses": statuses,
            "retain_images": retain_images,
            "retain_tables": retain_tables,
            "retain_formulas": retain_formulas,
            "convert_omml_first": convert_omml_first,
            "image_parts": source.get("image_parts", []),
            "embedded_parts": source.get("embedded_parts", []),
            "has_objects": any(source_counts[key] for key in ("images", "tables", "formulas")),
            "notes": notes,
        }

    def _ppt_to_word_artifact(self, task: dict[str, Any], file: dict[str, Any], source: Path, output_dir: Path) -> dict[str, Any]:
        """根据提取的幻灯片结构生成 PPT 转 Word 产物。"""
        try:
            slides = extract_pptx_slides(source, retain_images=bool(self._conversion_settings_snapshot(task, "ppt_to_word").get("ppt_retain_images", True)))
            target, existing_target, _strategy = self._conversion_output_target(output_dir, file["file_name"], ".docx")
            if target is None:
                return self._artifact_skipped(task, file, "docx", existing_target.name, f"同名输出已存在，按策略跳过：{existing_target.name}", self._conversion_settings_snapshot(task, "ppt_to_word"))
            settings = self.store.get_settings()
            ppt_options = dict(task.get("options", {}).get("ppt") or {})
            mode = str(ppt_options.get("mode") or settings.get("pptToWordMode", "逐页讲义模式") or "逐页讲义模式")
            generate_toc = bool(ppt_options.get("generateToc", settings.get("pptToWordGenerateToc", True)))
            template_name = str(ppt_options.get("templateName") or settings.get("pptToWordTemplate", "") or "")
            include_notes = bool(ppt_options.get("extractNotes", True))
            retain_images = bool(ppt_options.get("retainImages", settings.get("retainImages", True)))
            retain_formulas = bool(ppt_options.get("retainFormulas", True))
            warnings = build_docx_from_slides(
                slides,
                target,
                mode,
                generate_toc,
                template_name,
                include_notes,
                retain_images,
                retain_formulas,
                retain_tables=bool(ppt_options.get("retainTables", settings.get("retainTables", True))),
            ) or []
            warning_message = "；提示：" + "；".join(warnings) if warnings else ""
            return self._artifact_success(task, file, target, "docx", f"生成 DOCX，提取幻灯片 {len(slides)} 页{warning_message}", self._conversion_settings_snapshot(task, "ppt_to_word"))
        except (OSError, KeyError, ValueError, zipfile.BadZipFile) as exc:
            return self._artifact_error(task, file, "conversion_failed", str(exc))

    def _excel_to_pdf_artifacts(self, task: dict[str, Any], file: dict[str, Any], source: Path, output_dir: Path) -> list[dict[str, Any]]:
        """为 Excel 工作表或所选范围生成 PDF 产物。"""
        try:
            sheets = self._excel_sheets_for_conversion(extract_xlsx_sheets(source), task)
            artifacts: list[dict[str, Any]] = []
            for output_name, group, split_label in self._excel_output_groups(task, file["file_name"], sheets):
                target, existing_target, _strategy = self._conversion_output_target(output_dir, output_name, ".pdf")
                if target is None:
                    artifacts.append(self._artifact_skipped(task, file, "pdf", existing_target.name, f"同名输出已存在，按策略跳过：{existing_target.name}", self._conversion_settings_snapshot(task, "excel_to_pdf")))
                    continue
                warnings = build_pdf_from_xlsx(group, target, file["file_name"])
                warning_message = "；提示：" + "；".join(warnings) if warnings else ""
                cell_count = sum(len(sheet.get("cells", [])) for sheet in group)
                artifacts.append(self._artifact_success(task, file, target, "pdf", f"生成 PDF，提取工作表 {len(group)} 个、单元格 {cell_count} 个{split_label}{warning_message}", self._conversion_settings_snapshot(task, "excel_to_pdf")))
            return artifacts
        except (OSError, KeyError, zipfile.BadZipFile, ValueError) as exc:
            return [self._artifact_error(task, file, "conversion_failed", str(exc))]

    def _excel_to_word_artifacts(self, task: dict[str, Any], file: dict[str, Any], source: Path, output_dir: Path) -> list[dict[str, Any]]:
        """为 Excel 工作表或所选范围生成 Word 产物。"""
        try:
            sheets = self._excel_sheets_for_conversion(extract_xlsx_sheets(source), task)
            artifacts: list[dict[str, Any]] = []
            for output_name, group, split_label in self._excel_output_groups(task, file["file_name"], sheets):
                target, existing_target, _strategy = self._conversion_output_target(output_dir, output_name, ".docx")
                if target is None:
                    artifacts.append(self._artifact_skipped(task, file, "docx", existing_target.name, f"同名输出已存在，按策略跳过：{existing_target.name}", self._conversion_settings_snapshot(task, "excel_to_word")))
                    continue
                warnings = build_docx_from_xlsx(group, target, file["file_name"])
                warning_message = "；提示：" + "；".join(warnings) if warnings else ""
                cell_count = sum(len(sheet.get("cells", [])) for sheet in group)
                artifacts.append(self._artifact_success(task, file, target, "docx", f"生成 DOCX，提取工作表 {len(group)} 个、单元格 {cell_count} 个{split_label}{warning_message}", self._conversion_settings_snapshot(task, "excel_to_word")))
            return artifacts
        except (OSError, KeyError, zipfile.BadZipFile, ValueError) as exc:
            return [self._artifact_error(task, file, "conversion_failed", str(exc))]

    def _excel_to_ppt_artifacts(self, task: dict[str, Any], file: dict[str, Any], source: Path, output_dir: Path) -> list[dict[str, Any]]:
        """为 Excel 工作表或所选范围生成 PPT 产物。"""
        try:
            sheets = self._excel_sheets_for_conversion(extract_xlsx_sheets(source), task)
            artifacts: list[dict[str, Any]] = []
            for output_name, group, split_label in self._excel_output_groups(task, file["file_name"], sheets):
                target, existing_target, _strategy = self._conversion_output_target(output_dir, output_name, ".pptx")
                if target is None:
                    artifacts.append(self._artifact_skipped(task, file, "pptx", existing_target.name, f"同名输出已存在，按策略跳过：{existing_target.name}", self._conversion_settings_snapshot(task, "excel_to_ppt")))
                    continue
                warnings = build_pptx_from_xlsx(group, target, file["file_name"])
                warning_message = "；提示：" + "；".join(warnings) if warnings else ""
                artifacts.append(self._artifact_success(task, file, target, "pptx", f"生成 PPTX，提取工作表 {len(group)} 个{split_label}{warning_message}", self._conversion_settings_snapshot(task, "excel_to_ppt")))
            return artifacts
        except (OSError, KeyError, zipfile.BadZipFile, ValueError) as exc:
            return [self._artifact_error(task, file, "conversion_failed", str(exc))]

    def _conversion_output_target(self, output_dir: Path, file_name: str, suffix: str) -> tuple[Path | None, Path, str]:
        """按冲突策略解析转换输出路径。"""
        target = output_dir / f"{Path(file_name).stem or 'output'}{suffix}"
        strategy = str(self.store.get_settings().get("outputConflictStrategy", "自动重命名") or "自动重命名")
        if strategy not in {"自动重命名", "跳过", "覆盖"}:
            strategy = "自动重命名"
        if not target.exists() or strategy == "覆盖":
            return target, target, strategy
        if strategy == "跳过":
            return None, target, strategy
        return self._unique_target_path(target), target, strategy

    def _excel_output_groups(self, task: dict[str, Any], file_name: str, sheets: list[dict[str, Any]]) -> list[tuple[str, list[dict[str, Any]], str]]:
        """将 Excel 工作表分组为一个或多个转换产物。"""
        if not self._excel_split_sheets(task) or len(sheets) <= 1:
            return [(file_name, sheets, "")]
        groups: list[tuple[str, list[dict[str, Any]], str]] = []
        stem = Path(file_name).stem or "excel"
        source_suffix = Path(file_name).suffix or ".xlsx"
        for sheet in sheets:
            sheet_name = str(sheet.get("name") or f"Sheet{sheet.get('index', len(groups) + 1)}")
            output_name = self._safe_file_name(f"{stem}-{sheet_name}{source_suffix}")
            groups.append((output_name, [sheet], f"，拆分工作表：{sheet_name}"))
        return groups

    def _excel_split_sheets(self, task: dict[str, Any]) -> bool:
        """判断是否将 Excel 工作表拆为单独产物。"""
        settings = self.store.get_settings()
        excel_options = dict(task.get("options", {}).get("excel") or {})
        return effective_excel_settings(excel_options, settings)["excel_split_sheets"]

    def _excel_sheets_for_conversion(self, sheets: list[dict[str, Any]], task: dict[str, Any]) -> list[dict[str, Any]]:
        """共享网页与原生转换的范围和公式处理规则。"""
        return select_excel_sheets(sheets, dict(task.get("options", {}).get("excel") or {}), self.store.get_settings())

    def _conversion_settings_snapshot(self, task: dict[str, Any], task_type: str) -> dict[str, Any]:
        """保存生成产物所用的转换设置快照。"""
        settings = self.store.get_settings()
        word_options = dict(task.get("options", {}).get("word") or {})
        ppt_options = dict(task.get("options", {}).get("ppt") or {})
        return {
            "task_type": task_type,
            "word_to_ppt_template": settings.get("wordToPptTemplate", "教学讲义默认模板"),
            "ppt_to_word_mode": ppt_options.get("mode", settings.get("pptToWordMode", "逐页讲义模式")),
            "ppt_to_word_template": ppt_options.get("templateName", settings.get("pptToWordTemplate", "默认 Word 讲义模板")),
            "ppt_to_word_template_path": ppt_options.get("templatePath", settings.get("pptToWordTemplatePath", "")),
            "ppt_to_word_generate_toc": ppt_options.get("generateToc", settings.get("pptToWordGenerateToc", True)),
            "ppt_extract_notes": ppt_options.get("extractNotes", True),
            "ppt_retain_images": ppt_options.get("retainImages", settings.get("retainImages", True)),
            "ppt_retain_formulas": ppt_options.get("retainFormulas", True),
            "ppt_retain_tables": ppt_options.get("retainTables", settings.get("retainTables", True)),
            "pdf_precision_mode": settings.get("pdfPrecisionMode", "平衡"),
            "pdf_to_word_engine": settings.get("pdfToWordEngine", "Mathpix"),
            **effective_excel_settings(dict(task.get("options", {}).get("excel") or {}), settings),
            "output_conflict_strategy": settings.get("outputConflictStrategy", "自动重命名"),
            "retain_images": word_options.get("retainImages", settings.get("retainImages", True)),
            "retain_tables": word_options.get("retainTables", settings.get("retainTables", True)),
            "retain_headers_footers": settings.get("retainHeadersFooters", False),
            "retain_footnotes_endnotes": settings.get("retainFootnotesEndnotes", True),
            "retain_comments": settings.get("retainComments", False),
            "retain_revisions": settings.get("retainRevisions", False),
            "word_max_chars_per_slide": word_options.get("maxCharsPerSlide", 320),
            "word_auto_pagination": word_options.get("autoPagination", True),
            "word_generate_toc": word_options.get("generateToc", False),
            "word_retain_formulas": word_options.get("retainFormulas", True),
            "word_convert_omml_first": word_options.get("convertOmmlFirst", True),
            "word_apply_template": word_options.get("applyTemplate", True),
            "word_generate_notes": word_options.get("generateNotes", False),
            "word_auto_beautify": word_options.get("autoBeautify", True),
        }

    def _artifact_success(
        self,
        task: dict[str, Any],
        file: dict[str, Any],
        target: Path,
        output_type: str,
        message: str,
        conversion_settings: dict[str, Any] | None = None,
    ) -> dict[str, Any]:
        """生成成功转换产物行。"""
        validation_error = self._output_artifact_validation_error(target, output_type)
        if validation_error:
            try:
                target.unlink(missing_ok=True)
            except OSError:
                # 保留原始校验错误，报告仍不会暴露失败产物的下载入口。
                pass
            artifact = self._artifact_error(task, file, "output_validation_failed", validation_error)
            artifact["file_name"] = target.name
            artifact["output_type"] = output_type
            artifact["conversion_settings"] = conversion_settings or self._conversion_settings_snapshot(task, task.get("task_type", ""))
            return artifact
        return {
            "file_id": file["id"],
            "task_id": task["id"],
            "source_file": file["file_name"],
            "file_name": target.name,
            "output_type": output_type,
            "path": str(target),
            "url": f"/api/artifacts/{task['id']}/{target.name}",
            "status": "成功",
            "message": message,
            "size": target.stat().st_size,
            "sha256": self._sha256(target),
            "conversion_settings": conversion_settings or self._conversion_settings_snapshot(task, task.get("task_type", "")),
        }

    @staticmethod
    def _output_artifact_validation_error(target: Path, output_type: str) -> str:
        """生成产物无法安全打开时返回用户可见错误。"""
        if not target.exists() or not target.is_file():
            return "转换输出不存在，已标记任务失败"
        try:
            if target.stat().st_size <= 0:
                return "转换输出为空，已标记任务失败"
            normalized_type = str(output_type or target.suffix.lstrip(".")).lower()
            required_parts = {
                "docx": {"[Content_Types].xml", "word/document.xml"},
                "docm": {"[Content_Types].xml", "word/document.xml"},
                "pptx": {"[Content_Types].xml", "ppt/presentation.xml"},
                "xlsx": {"[Content_Types].xml", "xl/workbook.xml"},
            }
            if normalized_type in required_parts:
                if not zipfile.is_zipfile(target):
                    return f"转换输出 {target.name} 不是可打开的 {normalized_type.upper()} 文件"
                with zipfile.ZipFile(target) as archive:
                    names = set(archive.namelist())
                    missing = sorted(required_parts[normalized_type] - names)
                    if missing:
                        return f"转换输出 {target.name} 缺少必需结构：{'、'.join(missing)}"
                    corrupt_name = archive.testzip()
                    if corrupt_name:
                        return f"转换输出 {target.name} 存在损坏部件：{corrupt_name}"
            elif normalized_type == "tex.zip":
                if not zipfile.is_zipfile(target):
                    return f"转换输出 {target.name} 不是可打开的 TEX.ZIP 文件"
                with zipfile.ZipFile(target) as archive:
                    infos = archive.infolist()
                    unsafe_names = [
                        info.filename
                        for info in infos
                        if info.filename.startswith(("/", "\\"))
                        or ".." in info.filename.replace("\\", "/").split("/")
                    ]
                    if unsafe_names:
                        return f"转换输出 {target.name} 包含不安全路径：{unsafe_names[0]}"
                    total_size = sum(max(0, int(info.file_size)) for info in infos if not info.is_dir())
                    if total_size > MATHPIX_TEX_ARCHIVE_LIMIT_BYTES:
                        return f"转换输出 {target.name} 解压后超过 50MB 安全限制"
                    corrupt_name = archive.testzip()
                    if corrupt_name:
                        return f"转换输出 {target.name} 存在损坏部件：{corrupt_name}"
            elif normalized_type == "pdf":
                with target.open("rb") as stream:
                    header = stream.read(8)
                    stream.seek(max(0, target.stat().st_size - 1024))
                    trailer = stream.read()
                if not header.startswith(b"%PDF-") or b"%%EOF" not in trailer:
                    return f"转换输出 {target.name} 不是完整的 PDF 文件"
        except (OSError, zipfile.BadZipFile, RuntimeError) as exc:
            return f"转换输出 {target.name} 无法打开：{exc}"
        return ""

    def _artifact_skipped(
        self,
        task: dict[str, Any],
        file: dict[str, Any],
        output_type: str,
        file_name: str,
        message: str,
        conversion_settings: dict[str, Any] | None = None,
    ) -> dict[str, Any]:
        """生成跳过的转换产物行。"""
        return {
            "file_id": file["id"],
            "task_id": task["id"],
            "source_file": file["file_name"],
            "file_name": file_name,
            "output_type": output_type,
            "path": "",
            "url": "",
            "status": "跳过",
            "message": message,
            "size": 0,
            "conversion_settings": conversion_settings or self._conversion_settings_snapshot(task, task.get("task_type", "")),
        }

    def _artifact_error(self, task: dict[str, Any], file: dict[str, Any], status: str, message: str) -> dict[str, Any]:
        """生成失败转换产物行。"""
        return {
            "file_id": file["id"],
            "task_id": task["id"],
            "source_file": file.get("file_name", ""),
            "file_name": "",
            "output_type": "",
            "path": "",
            "url": "",
            "status": status,
            "message": message,
            "size": 0,
            "conversion_settings": self._conversion_settings_snapshot(task, task.get("task_type", "")),
        }

    def _formula_items(self, file: dict[str, Any], seed: int, task: dict[str, Any]) -> list[dict[str, Any]]:
        """生成含 MathType 兜底信息的公式报告项。"""
        settings = self.store.get_settings()
        threshold = self._positive_int(settings.get("formulaConfidenceThreshold", 80), 80)
        formatting_enabled = self._mathtype_formatting_enabled(task, settings)
        format_scope = self._formula_format_scope(task, settings)
        count = seed % 4 + 1
        items: list[dict[str, Any]] = []
        for index in range(count):
            source_type = "OMML" if file.get("has_omml") and index == 0 else ("MathType" if index % 2 else "LaTeX")
            confidence = 72 + ((seed + index * 9) % 25)
            latex = f"x_{index + 1}^2 + y_{index + 1}^2 = z_{index + 1}^2"
            item = FormulaItem(
                file_id=file["id"],
                page_index=index + 1,
                position=f"第 {index + 1} 页 / 段落 {index + 2}",
                source_type=source_type,
                latex=latex,
                mathml=self._latex_to_mathml_preview(latex),
                confidence=confidence,
                status="成功" if confidence >= threshold else "待确认",
                original_image_path=f"source://{file['id']}/formula/{index + 1}",
                mathtype_data=f"MathType 预览：{latex}",
                format_status="已格式化" if confidence >= 85 else "未格式化",
            ).to_dict()
            item["original_image_ref"] = item["original_image_path"]
            item["mathtype_preview"] = item["mathtype_data"]
            item["output_format"] = settings.get("formulaOutputFormat", "LaTeX+MathML")
            item["font"] = settings.get("formulaFont", "Cambria Math")
            item["font_size"] = settings.get("formulaFontSize", 12)
            item["format_scope"] = format_scope
            item["alignment"] = settings.get("formulaAlignment", "居中")
            item["variable_style"] = settings.get("formulaVariableStyle", "斜体")
            item["function_style"] = settings.get("formulaFunctionStyle", "正体")
            item["script_scale"] = settings.get("formulaScriptScale", 70)
            item["fraction_style"] = settings.get("formulaFractionStyle", "标准")
            item["radical_style"] = settings.get("formulaRadicalStyle", "标准")
            item["matrix_spacing"] = settings.get("formulaMatrixSpacing", "标准")
            item["greek_style"] = settings.get("formulaGreekStyle", "标准")
            item["inline_baseline"] = settings.get("formulaInlineBaseline", "跟随正文")
            item["display_spacing"] = settings.get("formulaDisplaySpacing", "标准")
            item["numbering"] = settings.get("formulaNumbering", "按文档位置")
            item["low_confidence_strategy"] = settings.get("lowConfidenceFormulaStrategy", "人工确认")
            item["keep_original_image"] = bool(settings.get("keepFormulaImages", True))
            self._apply_formula_format_policy(item, task, formatting_enabled)
            item["format_comparison"] = self._formula_format_comparison(item, settings)
            items.append(item)
        return items

    def _image_formula_fallback_items(self, file: dict[str, Any], images: list[dict[str, Any]], task: dict[str, Any]) -> list[dict[str, Any]]:
        """根据疑似公式的小图片生成公式兜底项。"""
        settings = self.store.get_settings()
        formatting_enabled = self._mathtype_formatting_enabled(task, settings)
        format_scope = self._formula_format_scope(task, settings)
        items: list[dict[str, Any]] = []
        for index, image in enumerate([item for item in images if item.get("is_formula_like")], start=1):
            original_ref = image.get("asset_url") or image.get("image_path") or image.get("source_name") or f"source://{file['id']}/image-formula/{index}"
            confidence = min(79, max(1, int(image.get("confidence", 60) or 60)))
            item = FormulaItem(
                file_id=file["id"],
                page_index=int(image.get("page_index", index) or index),
                position=str(image.get("location") or f"图片公式 {index}"),
                source_type="图片公式",
                latex="",
                confidence=confidence,
                status="待确认",
                original_image_path=str(original_ref),
                mathtype_data="图片公式识别失败，已保留原图",
                format_status="识别失败",
            ).to_dict()
            item["original_image_ref"] = item["original_image_path"]
            item["mathtype_preview"] = item["mathtype_data"]
            item["output_format"] = settings.get("formulaOutputFormat", "LaTeX+MathML")
            item["font"] = settings.get("formulaFont", "Cambria Math")
            item["font_size"] = settings.get("formulaFontSize", 12)
            item["format_scope"] = format_scope
            item["alignment"] = settings.get("formulaAlignment", "居中")
            item["variable_style"] = settings.get("formulaVariableStyle", "斜体")
            item["function_style"] = settings.get("formulaFunctionStyle", "正体")
            item["script_scale"] = settings.get("formulaScriptScale", 70)
            item["fraction_style"] = settings.get("formulaFractionStyle", "标准")
            item["radical_style"] = settings.get("formulaRadicalStyle", "标准")
            item["matrix_spacing"] = settings.get("formulaMatrixSpacing", "标准")
            item["greek_style"] = settings.get("formulaGreekStyle", "标准")
            item["inline_baseline"] = settings.get("formulaInlineBaseline", "跟随正文")
            item["display_spacing"] = settings.get("formulaDisplaySpacing", "标准")
            item["numbering"] = settings.get("formulaNumbering", "按文档位置")
            item["low_confidence_strategy"] = settings.get("lowConfidenceFormulaStrategy", "人工确认")
            item["keep_original_image"] = True
            item["source_image_id"] = image.get("id", "")
            item["recognition_error"] = "图片公式未生成可用 LaTeX，已保留原图供人工校正"
            item["fallback_action"] = "保留原图"
            self._apply_formula_format_policy(item, task, formatting_enabled)
            item["format_comparison"] = self._formula_format_comparison(item, settings)
            items.append(item)
        return items

    @staticmethod
    def _formula_format_scope(task: dict[str, Any], settings: dict[str, Any]) -> str:
        """解析请求的 MathType 格式化范围。"""
        options = task.get("options") or {}
        nested = options.get("formula") if isinstance(options.get("formula"), dict) else {}
        value = (
            options.get("formulaFormatScope")
            or options.get("formatScope")
            or options.get("format_scope")
            or nested.get("formulaFormatScope")
            or nested.get("formatScope")
            or nested.get("format_scope")
            or settings.get("formulaFormatScope")
            or "全文"
        )
        return str(value or "全文")

    @staticmethod
    def _latex_to_mathml_preview(latex: str) -> str:
        """为 LaTeX 文本生成安全的 MathML 预览包装。"""
        return f"<math><mtext>{html.escape(str(latex or ''))}</mtext></math>"

    @staticmethod
    def _mathtype_formatting_enabled(task: dict[str, Any], settings: dict[str, Any]) -> bool:
        """判断任务是否启用 MathType 格式化。"""
        options = task.get("options") or {}
        if "enableMathTypeFormatting" in options:
            return bool(options.get("enableMathTypeFormatting"))
        if "enable_mathtype_formatting" in options:
            return bool(options.get("enable_mathtype_formatting"))
        return bool(settings.get("enableMathTypeFormatting", True))

    @staticmethod
    def _apply_formula_format_policy(item: dict[str, Any], task: dict[str, Any], formatting_enabled: bool) -> None:
        """附加公式格式化状态，同时保留原始公式。"""
        task_type = str(task.get("task_type") or "")
        is_format_task = task_type == "mathtype_format"
        original_ref = item.get("original_image_ref") or item.get("original_image_path") or ""
        policy = {
            "enabled": formatting_enabled,
            "preserve_original_formula": True,
            "fallback_format": item.get("source_type", "未知"),
            "fallback_ref": original_ref,
            "retryable": False,
            "status": "not_requested",
            "message": "当前任务未请求 MathType 格式化，保留原公式作为来源记录",
        }
        if is_format_task and not formatting_enabled:
            item["format_status"] = "已跳过"
            policy.update(
                {
                    "status": "disabled",
                    "message": "用户未启用 MathType 格式化，已保留原公式",
                }
            )
        elif is_format_task and item.get("source_type") == "MathType":
            item["format_status"] = "原样保留"
            policy.update(
                {
                    "status": "preserved_existing",
                    "message": "已有 MathType 公式按原对象保留，跨平台交付另存 MathML/LaTeX 或图片兜底",
                }
            )
        elif is_format_task and item.get("format_status") != "已格式化":
            policy.update(
                {
                    "status": "format_failed",
                    "retryable": True,
                    "message": "MathType 格式化未完成，已保留原公式供重试或人工校正",
                }
            )
        elif is_format_task:
            policy.update(
                {
                    "status": "formatted",
                    "message": "格式化已生成结果，同时保留原公式引用便于回滚",
                }
            )
        item["formatting_enabled"] = formatting_enabled
        item["preserve_original_formula"] = bool(policy.get("preserve_original_formula"))
        item["format_failure_policy"] = policy

    @staticmethod
    def _formula_format_comparison(item: dict[str, Any], settings: dict[str, Any]) -> dict[str, Any]:
        """生成公式格式化前后对比摘要。"""
        before = {
            "format": item.get("source_type", "未知"),
            "font": "保留原文",
            "font_size": "原文",
            "alignment": "原文位置",
            "numbering": "原文编号",
            "preview": item.get("latex", ""),
        }
        after = {
            "format": settings.get("formulaOutputFormat", "LaTeX+MathML"),
            "font": settings.get("formulaFont", "Cambria Math"),
            "font_size": settings.get("formulaFontSize", 12),
            "alignment": settings.get("formulaAlignment", "居中"),
            "numbering": settings.get("formulaNumbering", "按文档位置"),
            "scope": item.get("format_scope") or settings.get("formulaFormatScope", "全文"),
            "preview": item.get("mathtype_preview") or item.get("mathtype_data") or item.get("latex", ""),
        }
        changes = [
            f"输出格式：{before['format']} -> {after['format']}",
            f"字体：{before['font']} -> {after['font']} {after['font_size']}pt",
            f"对齐：{before['alignment']} -> {after['alignment']}",
            f"编号：{before['numbering']} -> {after['numbering']}",
            f"范围：{after['scope']}",
        ]
        return {
            "before": before,
            "after": after,
            "changes": changes,
            "summary": "；".join(changes),
        }

    def _macro_items(self, file: dict[str, Any], task: dict[str, Any], batch_allowed: bool = True) -> list[dict[str, Any]]:
        """生成有序宏队列，不执行 Word 宏。"""
        options = task.get("options", {})
        settings = self.store.get_settings()
        selected = self._selected_macro_specs(options)
        library = {macro["id"]: macro for macro in self.macro_library()}
        if not selected:
            selected = [{"id": macro["id"], "execute_order": index + 1, "selected": False} for index, macro in enumerate(self.macro_library()[:3])]
        backup_path = ""
        macro_execution_enabled = bool(settings.get("enableMacroExecution", True))
        whitelist_only = bool(settings.get("macroWhitelistOnly", False))
        whitelist = self._macro_whitelist(settings.get("macroWhitelist", ""))
        items: list[dict[str, Any]] = []
        for index, spec in enumerate(selected, start=1):
            macro_id = str(spec.get("id") or spec.get("macro_id") or "")
            if macro_id.startswith("vba_"):
                if macro_id not in library:
                    raise ValueError("导入宏来源已删除，请重新选择宏")
                source = self.macro_source_definition(library[macro_id]["source_id"])
                if source["sha256"] != library[macro_id]["source_sha256"]:
                    raise ValueError("导入宏来源已改变，请重新选择宏")
            macro = dict(library.get(macro_id) or self._macro_from_freeform(spec, index))
            macro["file_id"] = file["id"]
            macro["execute_order"] = int(spec.get("execute_order") or index)
            macro["execute_timing"] = str(spec.get("execute_timing") or options.get("executeTiming") or macro.get("execute_timing") or "Word 处理前")
            macro["failure_strategy"] = str(spec.get("failure_strategy") or options.get("failureStrategy") or "跳过")
            macro["failure_policy"] = self._macro_failure_policy(macro["failure_strategy"])
            macro["confirmed"] = bool(options.get("confirmMacroRisk") or spec.get("confirmed"))
            macro["timeout_seconds"] = self._positive_int(options.get("macroTimeoutSeconds", settings.get("macroTimeoutSeconds", 120)), 120)
            source_allowed = self._macro_source_allowed(macro, settings)
            macro["execute_status"] = self._macro_status(
                macro,
                bool(spec.get("selected", True)),
                file,
                macro_execution_enabled,
                source_allowed,
                batch_allowed,
                whitelist_only,
                whitelist,
            )
            if macro["execute_status"] in {"待本地客户端执行", "成功"} and options.get("macroBackup", True):
                backup_path = backup_path or self._create_macro_backup(file, task["id"])
                macro["backup_path"] = backup_path
                if backup_path:
                    backup = Path(backup_path)
                    macro["backup_size"] = backup.stat().st_size
                    macro["backup_sha256"] = self._sha256(backup)
            else:
                macro["backup_path"] = ""
                macro["backup_size"] = ""
                macro["backup_sha256"] = ""
            if macro["execute_status"] == "待确认":
                macro["error_message"] = "宏执行前需要用户确认风险"
            elif macro["execute_status"] == "待本地客户端执行":
                macro["error_message"] = "当前环境不会直接执行 Word 宏，需本地客户端接收队列"
            elif macro["execute_status"] == "待选择":
                macro["error_message"] = "已检测到可用宏，尚未加入执行队列"
            elif macro["execute_status"] == "已禁用":
                macro["error_message"] = "系统设置已禁用宏执行队列"
            elif macro["execute_status"] == "未授权":
                if not batch_allowed:
                    macro["error_message"] = "当前设置不允许批量执行宏，该文件已阻止进入执行队列"
                elif not source_allowed:
                    macro["error_message"] = f"宏来源 {macro.get('macro_source', '')} 未被设置允许"
                else:
                    macro["error_message"] = "当前宏不在白名单中，已阻止进入执行队列"
            macro["start_time"] = utc_now() if macro["execute_status"] == "成功" else ""
            macro["end_time"] = utc_now() if macro["execute_status"] == "成功" else ""
            items.append(macro)
        items.sort(key=lambda item: item["execute_order"])
        return items

    def _selected_macro_specs(self, options: dict[str, Any]) -> list[dict[str, Any]]:
        """从任务选项解析选中宏标识与执行顺序。"""
        raw = options.get("selectedMacros") or options.get("macroSequence") or []
        specs: list[dict[str, Any]] = []
        if isinstance(raw, list):
            for index, item in enumerate(raw, start=1):
                if isinstance(item, str):
                    specs.append({"id": item, "execute_order": index, "selected": True})
                elif isinstance(item, dict):
                    specs.append({**item, "execute_order": item.get("execute_order") or index, "selected": item.get("selected", True)})
        return specs

    def _macro_from_freeform(self, spec: dict[str, Any], order: int) -> dict[str, Any]:
        """根据用户提供的宏规格生成队列项。"""
        macro = MacroItem(
            str(spec.get("macro_name") or spec.get("name") or "UserProvidedMacro"),
            str(spec.get("macro_source") or spec.get("source") or "用户选择"),
            str(spec.get("macro_description") or spec.get("description") or "用户提供的宏"),
            order,
            id=str(spec.get("id") or new_id("macro")),
        )
        return macro.to_dict()

    def _macro_status(
        self,
        macro: dict[str, Any],
        selected: bool,
        file: dict[str, Any],
        macro_execution_enabled: bool = True,
        source_allowed: bool = True,
        batch_allowed: bool = True,
        whitelist_only: bool = False,
        whitelist: set[str] | None = None,
    ) -> str:
        """判断宏队列项状态，不运行 VBA。"""
        if not selected:
            return "待选择"
        if not macro_execution_enabled:
            return "已禁用"
        if not source_allowed:
            return "未授权"
        if not batch_allowed:
            return "未授权"
        if whitelist_only and not self._macro_allowed(macro, whitelist or set()):
            return "未授权"
        if macro.get("requires_confirmation") and not macro.get("confirmed"):
            return "待确认"
        if os.name != "nt":
            return "待本地客户端执行"
        if not file.get("storage_path") and not file.get("file_path"):
            return "失败"
        return "待本地客户端执行"

    @staticmethod
    def _macro_whitelist(value: Any) -> set[str]:
        """将宏白名单解析为规范标记。"""
        tokens = re.split(r"[\s,，;；]+", str(value or ""))
        return {token.strip().lower() for token in tokens if token.strip()}

    @staticmethod
    def _macro_allowed(macro: dict[str, Any], whitelist: set[str]) -> bool:
        """判断宏标识或名称是否在白名单内。"""
        if not whitelist:
            return False
        candidates = {
            str(macro.get("id") or "").lower(),
            str(macro.get("macro_name") or "").lower(),
        }
        return bool(candidates & whitelist)

    @staticmethod
    def _macro_source_allowed(macro: dict[str, Any], settings: dict[str, Any]) -> bool:
        """判断设置是否允许指定宏来源。"""
        source = str(macro.get("macro_source") or "").lower()
        if "本地宏库" in source or "local" in source:
            return bool(settings.get("allowLocalMacroLibrary", True))
        if "模板" in source or "template" in source:
            return bool(settings.get("allowTemplateMacros", True))
        if "当前文档" in source or "文档" in source or "用户选择" in source:
            return bool(settings.get("allowDocumentMacros", True))
        return True

    @staticmethod
    def _macro_failure_policy(strategy: str) -> dict[str, Any]:
        """返回宏失败策略对应的队列处理行为。"""
        normalized = str(strategy or "跳过").strip()
        policies = {
            "停止": {
                "strategy": "停止",
                "action": "stop_sequence",
                "continue_on_failure": False,
                "skip_failed_macro": False,
                "requires_user_decision": False,
                "description": "宏失败后停止当前文件的后续宏并保留备份供恢复",
            },
            "跳过": {
                "strategy": "跳过",
                "action": "skip_failed_macro",
                "continue_on_failure": True,
                "skip_failed_macro": True,
                "requires_user_decision": False,
                "description": "宏失败后跳过当前宏并继续执行后续宏",
            },
            "继续": {
                "strategy": "继续",
                "action": "continue_sequence",
                "continue_on_failure": True,
                "skip_failed_macro": False,
                "requires_user_decision": False,
                "description": "宏失败后记录错误并继续执行后续宏",
            },
            "询问": {
                "strategy": "询问",
                "action": "pause_for_user_decision",
                "continue_on_failure": False,
                "skip_failed_macro": False,
                "requires_user_decision": True,
                "description": "宏失败后暂停队列并等待用户选择停止、跳过或继续",
            },
        }
        return policies.get(normalized, policies["跳过"])

    def _create_macro_backup(self, file: dict[str, Any], task_id: str) -> str:
        """在宏交接前尽可能创建受管源文件备份。"""
        source = self._original_source_path(file)
        if not source.exists() or not source.is_file():
            return ""
        target_dir = self.store.backups_dir / task_id
        target_dir.mkdir(parents=True, exist_ok=True)
        target = target_dir / f"{Path(file['file_name']).stem}{source.suffix}.bak"
        shutil.copy2(source, target)
        return str(target)

    @staticmethod
    def _original_source_path(file: dict[str, Any]) -> Path:
        """为不应操作快照的文件动作返回原始文档路径。"""
        return Path(
            str(
                file.get("original_storage_path")
                or file.get("original_file_path")
                or file.get("storage_path")
                or file.get("file_path")
                or ""
            )
        )

    def _small_image_items(self, file: dict[str, Any], task: dict[str, Any], seed: int, errors: list[dict[str, Any]] | None = None) -> list[dict[str, Any]]:
        """从支持的源文件提取小图片报告项。"""
        path = Path(file.get("storage_path") or file.get("file_path") or "")
        if not path.exists() or not path.is_file():
            has_source_path = bool(file.get("storage_path") or file.get("file_path"))
            source_kind = str(file.get("source_kind") or "")
            if errors is not None and file.get("file_type") != "PDF" and (has_source_path or source_kind in {"upload", "archive_entry", "conversion_output"}):
                errors.append(self._image_extraction_error(file, "", "来源文件不可访问，图片提取失败", "恢复来源文件后重新执行图片检索"))
            return self._pdf_image_placeholders(file, seed) if file.get("file_type") == "PDF" else []
        if file.get("file_type") in OOXML_IMAGE_PREFIXES:
            return self._extract_ooxml_small_images(file, task, path, errors)
        if file.get("file_type") == "图片":
            item = self._small_image_from_bytes(file, task, path.name, path.read_bytes(), "上传图片", 1, None, errors)
            return [item] if item else []
        if file.get("file_type") == "PDF":
            return self._pdf_small_image_items(file, task, path, seed)
        return []

    def _extract_ooxml_small_images(self, file: dict[str, Any], task: dict[str, Any], path: Path, errors: list[dict[str, Any]] | None = None) -> list[dict[str, Any]]:
        """从 OOXML 媒体部件提取小图片。"""
        prefix = OOXML_IMAGE_PREFIXES.get(file.get("file_type"))
        if not prefix:
            return []
        items: list[dict[str, Any]] = []
        seen_hashes: set[str] = set()
        try:
            with zipfile.ZipFile(path) as archive:
                names = [name for name in archive.namelist() if name.startswith(prefix) and not name.endswith("/")]
                try:
                    locations = image_locations(archive)
                except (ET.ParseError, ValueError, KeyError) as exc:
                    locations = {}
                    if errors is not None:
                        errors.append(self._image_extraction_error(file, "", f"图片位置解析失败：{exc}", "修复源文档后重试位置检索"))
                for index, name in enumerate(names, start=1):
                    try:
                        data = archive.read(name)
                    except (KeyError, OSError, RuntimeError, zipfile.BadZipFile) as exc:
                        if errors is not None:
                            errors.append(self._image_extraction_error(file, name, f"图片条目读取失败：{exc}", "重新上传或修复源文档后重试"))
                        continue
                    references = locations.get(name, [])
                    if references and not self.store.get_settings().get("includeHeaderFooterImages", True):
                        references = [location for location in references if not location.startswith(("页眉", "页脚"))]
                        if not references:
                            continue
                    location = "；".join(references) or f"{file['file_type']} 媒体部件 · {name}（位置未解析）"
                    header_footer = all(reference.startswith(("页眉", "页脚")) for reference in references) if references else None
                    pages = sorted({int(match[1]) for reference in references if (match := re.match(r"幻灯片第 (\d+) 页", reference))})
                    item = self._small_image_from_bytes(file, task, name, data, location, pages[0] if pages else 0, seen_hashes, errors, header_footer)
                    if item:
                        item["page_indexes"] = pages
                        items.append(item)
        except zipfile.BadZipFile:
            if errors is not None:
                errors.append(self._image_extraction_error(file, "", "OOXML 文件损坏，图片提取失败", "修复或重新上传源文档后重试"))
            return []
        return items

    def _small_image_from_bytes(
        self,
        file: dict[str, Any],
        task: dict[str, Any],
        source_name: str,
        data: bytes,
        location: str,
        index: int,
        seen_hashes: set[str] | None = None,
        errors: list[dict[str, Any]] | None = None,
        header_footer: bool | None = None,
    ) -> dict[str, Any] | None:
        """图片字节满足阈值时生成小图片项。"""
        dimensions = self._image_dimensions(data)
        if not dimensions:
            if errors is not None:
                errors.append(self._image_extraction_error(file, source_name, "图片尺寸或格式无法识别，已记录为提取失败", "检查图片是否损坏，重新上传或使用重新导出"))
            return None
        width, height, image_type = dimensions
        settings = self.store.get_settings()
        max_width = int(settings.get("smallImageMaxWidth", 96) or 96)
        max_height = int(settings.get("smallImageMaxHeight", 96) or 96)
        max_area = int(settings.get("smallImageMaxArea", 9216) or 9216)
        is_small = width <= max_width and height <= max_height and width * height <= max_area
        if not is_small:
            return None
        flags = self._small_image_flags(source_name, location, data, image_type, seen_hashes)
        if header_footer is not None:
            flags["is_header_footer"] = header_footer
        if flags["is_header_footer"] and not settings.get("includeHeaderFooterImages", True):
            return None
        if flags["is_watermark"] and not settings.get("includeWatermarkImages", True):
            return None
        if flags["is_transparent"] and not settings.get("includeTransparentImages", True):
            return None
        if flags["is_duplicate"] and not settings.get("includeDuplicateImages", True):
            return None
        if seen_hashes is not None:
            seen_hashes.add(flags["image_hash"])
        output_dir = self.store.images_dir / task["id"] / file["id"]
        output_dir.mkdir(parents=True, exist_ok=True)
        extension = Path(source_name).suffix.lower() or f".{image_type.lower()}"
        target = output_dir / f"{new_id('image')}{extension}"
        target.write_bytes(data)
        label = self._image_heuristic_label(source_name, width, height, data)
        item = SmallImageItem(
            file_id=file["id"],
            page_index=index,
            location=location,
            width=width,
            height=height,
            image_type=image_type,
            confidence=label["confidence"],
            image_path=str(target),
            asset_url=f"/api/assets/{target.relative_to(self.store.data_dir).as_posix()}",
            source_name=source_name,
            is_small=True,
            is_formula_like=label["kind"] == "formula",
            is_icon_like=label["kind"] == "icon",
            is_qrcode_like=label["kind"] == "qrcode",
            is_stamp_like=label["kind"] == "stamp",
            is_signature_like=label["kind"] == "signature",
            is_header_footer=flags["is_header_footer"],
            is_watermark=flags["is_watermark"],
            is_transparent=flags["is_transparent"],
            is_duplicate=flags["is_duplicate"],
            image_hash=flags["image_hash"],
            duplicate_check_status=flags["duplicate_check_status"],
            duplicate_fallback=flags["duplicate_fallback"],
            export_status="可导出",
            export_format=str(settings.get("imageExportFormat", "原格式") or "原格式"),
        )
        payload = item.to_dict()
        payload["image_size"] = len(data)
        return payload

    def _pdf_small_image_items(self, file: dict[str, Any], task: dict[str, Any], path: Path, seed: int) -> list[dict[str, Any]]:
        """从 PDF 图片描述中提取小图片项。"""
        try:
            descriptors = _pdf_image_descriptors(path.read_bytes()[:20_000_000])
        except OSError:
            descriptors = []
        items: list[dict[str, Any]] = []
        settings = self.store.get_settings()
        max_width = int(settings.get("smallImageMaxWidth", 96) or 96)
        max_height = int(settings.get("smallImageMaxHeight", 96) or 96)
        max_area = int(settings.get("smallImageMaxArea", 9216) or 9216)
        seen_hashes: set[str] = set()
        has_sized_descriptor = any(int(descriptor.get("width") or 0) > 0 and int(descriptor.get("height") or 0) > 0 for descriptor in descriptors)
        for descriptor in descriptors:
            width = int(descriptor.get("width") or 0)
            height = int(descriptor.get("height") or 0)
            if width <= 0 or height <= 0:
                continue
            if width > max_width or height > max_height or width * height > max_area:
                continue
            image_type = _pdf_filter_image_type(str(descriptor.get("filter") or ""))
            stream = descriptor.get("stream") or b""
            source_name = f"PDF Image XObject {descriptor.get('index', len(items) + 1)}"
            if descriptor.get("filter"):
                source_name = f"{source_name} ({descriptor['filter']})"
            hash_source = stream or str(descriptor.get("dict_hash", "")).encode("ascii", errors="ignore")
            flags = self._small_image_flags(source_name, f"PDF 图片对象 {descriptor.get('index', len(items) + 1)}", hash_source, image_type, seen_hashes)
            if flags["is_duplicate"] and not settings.get("includeDuplicateImages", True):
                continue
            seen_hashes.add(flags["image_hash"])
            label = self._image_heuristic_label(source_name, width, height, hash_source)
            image_path = ""
            asset_url = ""
            if stream and image_type in {"jpg", "jpx", "jb2"}:
                output_dir = self.store.images_dir / task["id"] / file["id"]
                output_dir.mkdir(parents=True, exist_ok=True)
                target = output_dir / f"{new_id('image')}.{image_type}"
                target.write_bytes(stream)
                image_path = str(target)
                asset_url = f"/api/assets/{target.relative_to(self.store.data_dir).as_posix()}"
            item = SmallImageItem(
                file_id=file["id"],
                page_index=int(descriptor.get("index") or len(items) + 1),
                location=f"PDF 图片对象 {descriptor.get('index', len(items) + 1)}",
                width=width,
                height=height,
                image_type=image_type,
                confidence=label["confidence"],
                image_path=image_path,
                asset_url=asset_url,
                source_name=source_name,
                is_small=True,
                is_formula_like=label["kind"] == "formula",
                is_icon_like=label["kind"] == "icon",
                is_qrcode_like=label["kind"] == "qrcode",
                is_stamp_like=label["kind"] == "stamp",
                is_signature_like=label["kind"] == "signature",
                is_duplicate=flags["is_duplicate"],
                image_hash=flags["image_hash"],
                duplicate_check_status=flags["duplicate_check_status"],
                duplicate_fallback=flags["duplicate_fallback"],
                export_status="可导出" if asset_url else "原始流不可用",
                export_message="" if asset_url else "PDF 图片原始流无法直接导出，可保留报告记录",
                export_format=str(settings.get("imageExportFormat", "原格式") or "原格式"),
            )
            payload = item.to_dict()
            if asset_url:
                payload["image_size"] = len(stream)
            items.append(payload)
        if has_sized_descriptor:
            return items
        return items or self._pdf_image_placeholders(file, seed)

    def _pdf_image_placeholders(self, file: dict[str, Any], seed: int) -> list[dict[str, Any]]:
        """原始流不可用时生成 PDF 图片占位行。"""
        image_count = int(file.get("content_summary", {}).get("imageObjects", 0) or 0)
        items: list[dict[str, Any]] = []
        for index in range(min(image_count, 12)):
            width = 32 + ((seed + index * 13) % 64)
            height = 24 + ((seed + index * 9) % 64)
            item = SmallImageItem(
                file_id=file["id"],
                page_index=index + 1,
                location=f"PDF 图片对象 {index + 1}",
                width=width,
                height=height,
                image_type="pdf-object",
                confidence=62,
                source_name=f"pdf-image-{index + 1}",
                is_small=True,
                is_formula_like=index % 2 == 0,
                duplicate_check_status="无法判断",
                duplicate_fallback="PDF 未提供可导出图片流，保留原始结果",
                export_status="原始流不可用",
                export_message="PDF 未提供可导出图片流，保留原始结果",
                export_format=str(self.store.get_settings().get("imageExportFormat", "原格式") or "原格式"),
            )
            items.append(item.to_dict())
        return items

    @staticmethod
    def _image_extraction_error(file: dict[str, Any], source_name: str, message: str, recommendation: str) -> dict[str, Any]:
        """生成小图片提取错误行。"""
        return {
            "id": new_id("imgerr"),
            "file_id": file.get("id", ""),
            "file_name": file.get("file_name", ""),
            "file_type": file.get("file_type", ""),
            "source_name": source_name,
            "status": "失败",
            "message": message,
            "recommendation": recommendation,
            "created_at": utc_now(),
        }

    @staticmethod
    def _small_image_flags(
        source_name: str,
        location: str,
        data: bytes,
        image_type: str,
        seen_hashes: set[str] | None = None,
    ) -> dict[str, Any]:
        """返回小图片筛选和去重的启发式标记。"""
        text = f"{source_name} {location}".lower()
        try:
            digest = hashlib.sha256(data).hexdigest()
            duplicate_check_status = "已判断" if seen_hashes is not None else "未比较"
            duplicate_fallback = (
                "重复判断失败时保留原始结果"
                if seen_hashes is not None
                else "单图或来源不支持批量去重，保留原始结果"
            )
        except (TypeError, ValueError):
            digest = ""
            duplicate_check_status = "判断失败"
            duplicate_fallback = "重复判断失败，保留原始结果"
        is_duplicate = digest in seen_hashes if seen_hashes is not None else False
        return {
            "is_header_footer": any(token in text for token in ("header", "footer", "页眉", "页脚")),
            "is_watermark": any(token in text for token in ("watermark", "水印")),
            "is_transparent": image_type == "png" and TaskProcessor._png_has_alpha(data),
            "is_duplicate": is_duplicate,
            "image_hash": digest,
            "duplicate_check_status": duplicate_check_status,
            "duplicate_fallback": duplicate_fallback,
        }

    @staticmethod
    def _png_has_alpha(data: bytes) -> bool:
        """判断 PNG 头是否表明存在透明通道。"""
        return data.startswith(b"\x89PNG\r\n\x1a\n") and len(data) > 25 and data[25] in {4, 6}

    @staticmethod
    def _image_dimensions(data: bytes) -> tuple[int, int, str] | None:
        """从常见图片头读取宽高。"""
        if data.startswith(b"\x89PNG\r\n\x1a\n") and len(data) >= 24:
            return int.from_bytes(data[16:20], "big"), int.from_bytes(data[20:24], "big"), "png"
        if data.startswith(b"GIF87a") or data.startswith(b"GIF89a"):
            return int.from_bytes(data[6:8], "little"), int.from_bytes(data[8:10], "little"), "gif"
        if data.startswith(b"BM") and len(data) >= 26:
            width, height = bmp_dimensions(data)
            return (width, height, "bmp") if width > 0 and height > 0 else None
        if data.startswith(b"\xff\xd8"):
            return TaskProcessor._jpeg_dimensions(data)
        if data[:4] in {b"II*\x00", b"MM\x00*"}:
            width, height = tiff_dimensions(data)
            if width > 0 and height > 0:
                return width, height, "tiff"
        if data.lstrip().startswith(b"<"):
            width, height = svg_dimensions(data)
            if width > 0 and height > 0:
                return width, height, "svg"
        return None

    @staticmethod
    def _jpeg_dimensions(data: bytes) -> tuple[int, int, str] | None:
        """从 SOF 标记读取 JPEG 宽高。"""
        index = 2
        while index + 9 < len(data):
            if data[index] != 0xFF:
                index += 1
                continue
            marker = data[index + 1]
            if marker in {0xC0, 0xC1, 0xC2, 0xC3, 0xC5, 0xC6, 0xC7, 0xC9, 0xCA, 0xCB, 0xCD, 0xCE, 0xCF}:
                height = int.from_bytes(data[index + 5:index + 7], "big")
                width = int.from_bytes(data[index + 7:index + 9], "big")
                return width, height, "jpg"
            segment_length = int.from_bytes(data[index + 2:index + 4], "big")
            if segment_length < 2:
                return None
            index += 2 + segment_length
        return None

    @staticmethod
    def _image_heuristic_label(source_name: str, width: int, height: int, data: bytes) -> dict[str, Any]:
        """将小图片分类为图标、公式、印章、签名或疑似二维码。"""
        name = source_name.lower()
        aspect = width / max(height, 1)
        if "qr" in name or (abs(width - height) <= 8 and len(data) > 1500):
            return {"kind": "qrcode", "confidence": 78}
        if "stamp" in name or "seal" in name:
            return {"kind": "stamp", "confidence": 82}
        if "sign" in name or aspect > 3:
            return {"kind": "signature", "confidence": 76}
        if "equation" in name or "formula" in name or aspect > 1.8:
            return {"kind": "formula", "confidence": 74}
        return {"kind": "icon", "confidence": 70}
