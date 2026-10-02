"""定义各平台的本地客户端安装配置。Windows 与 macOS 的安装包及 MathType 原生对象格式不同；同平台可交接原生对象，跨平台必须提供 MathML、LaTeX 或图片兜底。"""

from __future__ import annotations

import copy
import os
import platform
from typing import Any


FORMULA_FALLBACK_FORMATS = ["MathML", "LaTeX", "图片"]


TASK_ROUTING = {
    "formula_precheck": "本地优先，网页端只展示识别结果和确认状态",
    "omml_to_mathtype": "需要本地客户端和对应平台 MathType/Office 能力",
    "mathtype_format": "需要本地客户端，跨平台交付前生成 MathML/LaTeX 或图片兜底",
    "macro_sequence": "仅本地授权执行，网页端不直接运行宏",
    "word_to_ppt": "网页负责任务编排，复杂排版由本地客户端补强",
    "pdf_to_word": "Mathpix OCR 需用户授权外部上传，结果回写本地输出目录",
}


# 按平台区分安装包与 MathType 对象，避免把不兼容对象误认为可跨平台使用。
PROFILES: dict[str, dict[str, Any]] = {
    "Windows": {
        "platform": "Windows",
        "installerKind": "windows-msi",
        "recommendedInstaller": "K12 Windows 本地客户端",
        "downloadLabel": "下载 Windows 安装包",
        "officeAutomation": "Windows PowerShell + Office COM 可执行旧 Word/PPT 规范化，无需 pywin32",
        "mathType": "Windows MathType/OLE 对象可作为本机编辑格式",
        "mathtypeObjectFormat": "Windows OLE / Equation Native",
        "formulaPortability": "Windows 与 macOS MathType 公式对象不通用，跨平台任务需同时生成 MathML、LaTeX 或图片兜底。",
        "requiredComponents": [
            "Microsoft Office 桌面版",
            "MathType Windows 版本",
            "K12 Windows 本地客户端",
            "pywin32/Office COM 自动化桥接",
        ],
        "optionalComponents": [
            "Mathpix 凭证，用于扫描型 PDF 转 Word",
            "LibreOffice，用于非核心格式兜底转换",
        ],
        "preflightChecks": [
            "确认 Word、PowerPoint、Excel 可由当前用户打开",
            "确认 MathType 加载项可在 Office 中编辑公式",
            "确认宏执行前已启用本地授权和备份",
            "跨平台交付前生成公式兜底格式",
        ],
        "capabilities": {
            "officeAutomation": True,
            "mathTypeAutomation": True,
            "macroExecution": True,
            "ommlDependencySearch": True,
        },
        "taskRouting": TASK_ROUTING,
    },
    "macOS": {
        "platform": "macOS",
        "installerKind": "macos-pkg",
        "recommendedInstaller": "K12 macOS 本地客户端",
        "downloadLabel": "下载 macOS 安装包",
        "officeAutomation": "macOS Office 自动化受限，优先走本地文件分析与人工确认",
        "mathType": "macOS MathType 对象与 Windows 不通用，不能直接作为 Windows 可编辑公式交付",
        "mathtypeObjectFormat": "macOS MathType 对象 / MathML / LaTeX / 图片兜底",
        "formulaPortability": "macOS 与 Windows 的 MathType 对象不通用，不能互认为同一编辑格式，跨系统流转时优先导出 MathML、LaTeX 或图片备份。",
        "requiredComponents": [
            "Microsoft Office for Mac",
            "MathType macOS 版本",
            "K12 macOS 本地客户端",
            "文件系统权限和本地服务授权",
        ],
        "optionalComponents": [
            "Mathpix 凭证，用于扫描型 PDF 转 Word",
            "手动公式复核流程，用于替代 Windows COM 自动化",
        ],
        "preflightChecks": [
            "确认 Office for Mac 可打开目标文档",
            "确认 MathType macOS 版本已安装且能编辑本机公式",
            "不要把 macOS MathType 对象直接交给 Windows 自动化链路",
            "跨平台交付前生成 MathML/LaTeX 或图片兜底",
        ],
        "capabilities": {
            "officeAutomation": False,
            "mathTypeAutomation": False,
            "macroExecution": False,
            "ommlDependencySearch": True,
        },
        "taskRouting": TASK_ROUTING,
    },
    "Unknown": {
        "platform": "Unknown",
        "installerKind": "manual",
        "recommendedInstaller": "K12 通用网页/本地服务",
        "downloadLabel": "按实际系统选择安装包",
        "officeAutomation": "当前平台未声明 Office 自动化能力",
        "mathType": "MathType 自动化未启用",
        "mathtypeObjectFormat": "MathML / LaTeX / 图片兜底",
        "formulaPortability": "无法确认平台对象兼容性时，不直接交付 MathType 原生对象，优先使用 MathML、LaTeX 或图片兜底。",
        "requiredComponents": [
            "现代浏览器",
            "K12 本地 API 服务",
        ],
        "optionalComponents": [
            "按实际系统安装 Windows 或 macOS 本地客户端",
            "Mathpix 凭证，用于扫描型 PDF 转 Word",
        ],
        "preflightChecks": [
            "先在设置中明确选择 Windows 或 macOS",
            "确认本地客户端能访问待处理文件",
            "公式任务使用 MathML/LaTeX 或图片兜底",
        ],
        "capabilities": {
            "officeAutomation": False,
            "mathTypeAutomation": False,
            "macroExecution": False,
            "ommlDependencySearch": True,
        },
        "taskRouting": TASK_ROUTING,
    },
}


def detect_platform(configured: str | None = None) -> str:
    """解析显式指定或自动检测的客户端平台。"""
    value = (configured or "auto").strip()
    if value and value.lower() != "auto":
        return normalize_platform(value)
    system_name = platform.system()
    if system_name == "Darwin":
        return "macOS"
    if system_name == "Windows" or os.name == "nt":
        return "Windows"
    return normalize_platform(system_name)


def normalize_platform(value: str | None) -> str:
    """将常见平台别名规范为安装配置使用的名称。"""
    lowered = (value or "").strip().lower()
    if lowered in {"windows", "win", "win32", "win64"}:
        return "Windows"
    if lowered in {"macos", "mac", "darwin", "osx", "os x"}:
        return "macOS"
    return value.strip() if value and value.strip() else "Unknown"


def install_profile(platform_name: str | None = None, settings: dict[str, Any] | None = None) -> dict[str, Any]:
    """返回所选平台安装配置的副本，包含公式互操作规则。"""
    configured = platform_name if platform_name is not None else str((settings or {}).get("localClientPlatform") or "auto")
    system_name = detect_platform(configured)
    profile = copy.deepcopy(PROFILES.get(system_name, PROFILES["Unknown"]))
    profile["platform"] = system_name if system_name in PROFILES else "Unknown"
    profile["selectedCompatibilityMode"] = (settings or {}).get("mathtypeCompatibilityMode", "platform-specific")
    profile["sensitiveFilesPreferLocal"] = bool((settings or {}).get("sensitiveFilesPreferLocal", True))
    profile["formulaObjectInterop"] = _formula_object_interop(profile, str(profile["selectedCompatibilityMode"]))
    return profile


def _formula_object_interop(profile: dict[str, Any], compatibility_mode: str) -> dict[str, Any]:
    """说明原生 MathType 对象能否交接给客户端。"""
    platform_name = normalize_platform(str(profile.get("platform") or "Unknown"))
    platform_known = platform_name in {"Windows", "macOS"}
    native_object_allowed = platform_known and compatibility_mode == "platform-specific"
    return {
        "schema_version": "k12.formulaObjectInterop.v1",
        "platform": platform_name,
        "compatibility_mode": compatibility_mode,
        "native_object_format": str(profile.get("mathtypeObjectFormat") or ""),
        "platform_objects_cross_compatible": False,
        "same_platform_required_for_native_objects": native_object_allowed,
        "native_mathtype_object_allowed": native_object_allowed,
        "fallback_formats": list(FORMULA_FALLBACK_FORMATS),
        "fallback_required_for_cross_platform": True,
        "message": str(profile.get("formulaPortability") or ""),
    }
