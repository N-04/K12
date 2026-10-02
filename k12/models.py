"""定义 K12 的核心数据模型和任务标签。规范上传、任务、公式、宏、图片和报告记录；持久化及接口脱敏分别由存储层和服务层负责。"""

from __future__ import annotations

from dataclasses import asdict, dataclass, field
from datetime import datetime, timezone
from typing import Any
from uuid import uuid4


SUPPORTED_EXTENSIONS = {
    "word": {".doc", ".docx", ".docm", ".dot", ".dotm"},
    "excel": {".xls", ".xlsx", ".xlsm"},
    "ppt": {".ppt", ".pptx", ".pptm"},
    "pdf": {".pdf"},
    "image": {".jpg", ".jpeg", ".png", ".bmp", ".tiff"},
    "archive": {".zip"},
}


TASK_LABELS = {
    "word_to_ppt": "Word 转 PPT",
    "ppt_to_word": "PPT 转 Word",
    "pdf_to_word": "PDF 转 Word",
    "excel_to_pdf": "Excel 转 PDF",
    "excel_to_word": "Excel 转 Word",
    "excel_to_ppt": "Excel 转 PPT",
    "formula_precheck": "Word 公式预检",
    "omml_to_mathtype": "OMML 转 MathType",
    "mathtype_format": "MathType 格式化",
    "macro_sequence": "Word 宏顺序执行",
    "small_image_scan": "微小图片检索",
    "batch_process": "批量处理",
}


LOCAL_REQUIRED_TASKS = {
    "formula_precheck",
    "omml_to_mathtype",
    "mathtype_format",
    "macro_sequence",
}


def utc_now() -> str:
    """返回用于持久化记录的 UTC ISO 时间戳。"""
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def new_id(prefix: str) -> str:
    """生成带稳定记录前缀的简短不透明标识。"""
    return f"{prefix}_{uuid4().hex[:12]}"


@dataclass(slots=True)
class FileItem:
    """用于任务规划的上传或发现文件元数据。"""

    file_name: str
    file_type: str
    file_size: int
    extension: str
    file_path: str = ""
    storage_path: str = ""
    source_kind: str = "metadata"
    source_relative_path: str = ""
    archive_parent_id: str = ""
    archive_entry_count: int = 0
    encrypted: bool = False
    content_summary: dict[str, Any] = field(default_factory=dict)
    id: str = field(default_factory=lambda: new_id("file"))
    page_count: int = 0
    slide_count: int = 0
    sheet_count: int = 0
    status: str = "待处理"
    has_formula: bool = False
    has_mathtype: bool = False
    has_omml: bool = False
    missing_omml_dependency: bool = False
    has_macro: bool = False
    has_image: bool = False
    has_small_image: bool = False
    validation_errors: list[str] = field(default_factory=list)
    created_at: str = field(default_factory=utc_now)

    def to_dict(self) -> dict[str, Any]:
        """将当前记录及计算字段序列化为存储、报告和接口使用的字典。"""
        return asdict(self)


@dataclass(slots=True)
class Task:
    """队列、报告与界面共享的处理任务元数据。"""

    task_type: str
    execute_mode: str
    file_ids: list[str]
    options: dict[str, Any] = field(default_factory=dict)
    id: str = field(default_factory=lambda: new_id("task"))
    status: str = "待处理"
    progress: int = 0
    input_path: str = ""
    output_path: str = ""
    error_message: str = ""
    start_time: str = ""
    end_time: str = ""
    duration_seconds: int = 0
    duration_label: str = "-"
    success_count: int = 0
    fail_count: int = 0
    failure_count: int = 0
    retryable_count: int = 0
    created_at: str = field(default_factory=utc_now)

    def to_dict(self) -> dict[str, Any]:
        """将当前记录及计算字段序列化为存储、报告和接口使用的字典。"""
        data = asdict(self)
        data["task_label"] = TASK_LABELS.get(self.task_type, self.task_type)
        return data


@dataclass(slots=True)
class FormulaItem:
    """来自 Word、PPT、PDF、LaTeX 或 Mathpix 的规范公式记录。"""

    file_id: str
    page_index: int
    position: str
    source_type: str
    latex: str
    confidence: int
    status: str
    id: str = field(default_factory=lambda: new_id("formula"))
    original_image_path: str = ""
    mathml: str = ""
    mathtype_data: str = ""
    format_status: str = "未格式化"
    position_status: str = "已记录"
    position_issue: str = ""
    fallback_position: str = ""

    def to_dict(self) -> dict[str, Any]:
        """将当前记录及计算字段序列化为存储、报告和接口使用的字典。"""
        return asdict(self)


@dataclass(slots=True)
class MacroItem:
    """供有序本地执行计划使用的 Word 宏选择元数据。"""

    macro_name: str
    macro_source: str
    macro_description: str
    execute_order: int
    file_id: str = ""
    execute_timing: str = "Word 处理前"
    execute_status: str = "待执行"
    failure_strategy: str = "跳过"
    requires_confirmation: bool = True
    confirmed: bool = False
    local_only: bool = True
    error_message: str = ""
    start_time: str = ""
    end_time: str = ""
    backup_path: str = ""
    id: str = field(default_factory=lambda: new_id("macro"))

    def to_dict(self) -> dict[str, Any]:
        """将当前记录及计算字段序列化为存储、报告和接口使用的字典。"""
        return asdict(self)


@dataclass(slots=True)
class OmmlDependencyItem:
    """本地 MathType 交接所需的 OMML 依赖查找与复制状态。"""

    file_id: str
    document_path: str
    omml_file_name: str = ""
    omml_source_path: str = ""
    omml_target_path: str = ""
    found_status: str = "未找到"
    copy_status: str = "跳过"
    error_message: str = ""
    source_size: int = 0
    source_sha256: str = ""
    target_size: int = 0
    target_sha256: str = ""
    id: str = field(default_factory=lambda: new_id("omml"))
    created_at: str = field(default_factory=utc_now)

    def to_dict(self) -> dict[str, Any]:
        """将当前记录及计算字段序列化为存储、报告和接口使用的字典。"""
        return asdict(self)


@dataclass(slots=True)
class SmallImageItem:
    """图片报告与手动审阅使用的小图片元数据。"""

    file_id: str
    page_index: int
    location: str
    width: int
    height: int
    image_type: str
    confidence: int
    id: str = field(default_factory=lambda: new_id("smallimg"))
    image_path: str = ""
    asset_url: str = ""
    source_name: str = ""
    is_small: bool = True
    is_formula_like: bool = False
    is_icon_like: bool = False
    is_qrcode_like: bool = False
    is_stamp_like: bool = False
    is_signature_like: bool = False
    is_header_footer: bool = False
    is_watermark: bool = False
    is_transparent: bool = False
    is_duplicate: bool = False
    image_hash: str = ""
    duplicate_check_status: str = "未比较"
    duplicate_fallback: str = "重复判断失败时保留原始结果"
    export_status: str = "可导出"
    export_message: str = ""
    reexported_at: str = ""
    export_format: str = "原格式"

    @property
    def area(self) -> int:
        """返回小图片筛选所用的像素面积。"""
        return self.width * self.height

    def to_dict(self) -> dict[str, Any]:
        """将当前记录及计算字段序列化为存储、报告和接口使用的字典。"""
        data = asdict(self)
        data["area"] = self.area
        return data


@dataclass(slots=True)
class ReportItem:
    """完成任务的已保存报告计数与产物引用。"""

    task_id: str
    file_id: str
    report_type: str
    report_path: str
    success_count: int = 0
    fail_count: int = 0
    formula_count: int = 0
    omml_count: int = 0
    omml_converted_count: int = 0
    macro_count: int = 0
    macro_success_count: int = 0
    macro_fail_count: int = 0
    formatted_formula_count: int = 0
    small_image_count: int = 0
    error_count: int = 0
    id: str = field(default_factory=lambda: new_id("report"))
    created_at: str = field(default_factory=utc_now)

    def to_dict(self) -> dict[str, Any]:
        """将当前记录及计算字段序列化为存储、报告和接口使用的字典。"""
        return asdict(self)
