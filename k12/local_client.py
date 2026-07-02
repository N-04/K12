"""Local companion CLI contracts for K12 desktop-only document actions.

The CLI is intentionally conservative: it may inspect payloads, send heartbeats,
sync dry-run status, and execute explicitly authorized safe file copies. It does
not call Office, MathType, OMML writeback, or Word macros directly.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import platform
import shutil
import socket
import sys
from pathlib import Path
from typing import Any
from urllib.error import HTTPError, URLError
from urllib.parse import quote
from urllib.request import Request, urlopen

from .install_profiles import normalize_platform


LOCAL_CLIENT_VERSION = "0.1.0"
SAFE_CAPABILITY_KEYS = {"officeAutomation", "mathTypeAutomation", "macroExecution", "ommlDependencySearch"}
SAFE_COMPONENT_KEYS = {"office", "word", "powerpoint", "excel", "mathtype", "libreoffice", "omml_dependency"}
OMML_DEPENDENCY_SUFFIXES = {".xsl", ".xslt", ".xml", ".mml"}


class LocalClientError(RuntimeError):
    """Raised when the companion CLI cannot complete a safe handoff step."""

    code = "local_client_error"


def normalize_origin(origin: str) -> str:
    """Validate and normalize the local K12 API origin used by the CLI."""
    value = str(origin or "").strip().rstrip("/")
    if not value:
        raise LocalClientError("请提供本地 API 地址")
    if not value.startswith(("http://", "https://")):
        raise LocalClientError("本地 API 地址必须以 http:// 或 https:// 开头")
    return value


def build_heartbeat(
    client_id: str = "",
    platform_name: str = "",
    status: str = "online",
    active_task_id: str = "",
    message: str = "",
    capabilities: dict[str, Any] | None = None,
    preflight: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """Build a path-redacted heartbeat with platform capability bits.

    The payload tells the web side whether Windows/macOS desktop capabilities
    appear available, but never sends component paths or token material.
    """
    safe_preflight = preflight if preflight is not None else build_component_preflight(platform_name)
    raw_capabilities = capabilities or {
        **dict(safe_preflight.get("capabilities") or {}),
    }
    safe_capabilities = {
        key: value
        for key, value in raw_capabilities.items()
        if key in SAFE_CAPABILITY_KEYS and isinstance(value, (bool, int, float, str))
    }
    return {
        "schema_version": "k12.localClientHeartbeat.v1",
        "client_id": client_id or f"k12-cli-{socket.gethostname()}",
        "status": status,
        "platform": normalize_platform(platform_name or platform.system()),
        "version": LOCAL_CLIENT_VERSION,
        "active_task_id": active_task_id,
        "message": message or "K12 本地伴随 CLI 已连接；默认 dry-run，不执行 Office、MathType 或 Word 宏",
        "capabilities": safe_capabilities,
        "preflight": sanitize_component_preflight(safe_preflight),
    }


def build_component_preflight(platform_name: str = "") -> dict[str, Any]:
    """Return a path-safe desktop component summary for the current platform."""
    detected = normalize_platform(platform_name or platform.system())
    components = {
        "word": _component_status(_component_available(detected, "word"), "Word 桌面组件"),
        "powerpoint": _component_status(_component_available(detected, "powerpoint"), "PowerPoint 桌面组件"),
        "excel": _component_status(_component_available(detected, "excel"), "Excel 桌面组件"),
        "mathtype": _component_status(_component_available(detected, "mathtype"), "MathType 组件"),
        "libreoffice": _component_status(_component_available(detected, "libreoffice"), "LibreOffice 兜底组件"),
        "omml_dependency": _component_status(_component_available(detected, "omml_dependency"), "OMML 依赖文件"),
    }
    office_available = any(components[key]["available"] for key in ("word", "powerpoint", "excel"))
    mathtype_available = components["mathtype"]["available"]
    windows_native = detected == "Windows"
    return {
        "schema_version": "k12.localClientPreflight.v1",
        "platform": detected,
        "components": components,
        "capabilities": {
            "officeAutomation": bool(windows_native and office_available),
            "mathTypeAutomation": bool(windows_native and mathtype_available),
            "macroExecution": bool(windows_native and components["word"]["available"]),
            "ommlDependencySearch": True,
        },
        "executes_native_documents": False,
        "path_policy": "component paths are not reported",
    }


def sanitize_component_preflight(preflight: dict[str, Any]) -> dict[str, Any]:
    """Return only public component status fields from a local preflight."""
    raw_components = preflight.get("components") if isinstance(preflight.get("components"), dict) else {}
    components: dict[str, dict[str, Any]] = {}
    for key, value in raw_components.items():
        if key not in SAFE_COMPONENT_KEYS or not isinstance(value, dict):
            continue
        components[key] = {
            "label": str(value.get("label") or key)[:80],
            "available": bool(value.get("available")),
            "status": str(value.get("status") or "")[:40],
        }
    raw_capabilities = preflight.get("capabilities") if isinstance(preflight.get("capabilities"), dict) else {}
    capabilities = {
        key: value
        for key, value in raw_capabilities.items()
        if key in SAFE_CAPABILITY_KEYS and isinstance(value, (bool, int, float, str))
    }
    return {
        "schema_version": str(preflight.get("schema_version") or "k12.localClientPreflight.v1"),
        "platform": normalize_platform(str(preflight.get("platform") or platform.system())),
        "components": components,
        "capabilities": capabilities,
        "executes_native_documents": False,
        "path_policy": "component paths are not reported",
    }


def _component_status(available: bool, label: str) -> dict[str, Any]:
    """Format one desktop component preflight result without path details."""
    return {"label": label, "available": bool(available), "status": "available" if available else "missing"}


def _component_available(platform_name: str, component: str) -> bool:
    """Check whether a desktop component appears available on this platform."""
    commands = {
        "word": ["winword", "Microsoft Word"],
        "powerpoint": ["powerpnt", "Microsoft PowerPoint"],
        "excel": ["excel", "Microsoft Excel"],
        "mathtype": ["MathType"],
        "libreoffice": ["soffice", "libreoffice"],
    }
    if component in commands and any(shutil.which(name) for name in commands[component]):
        return True
    for candidate in _component_candidates(platform_name, component):
        try:
            if candidate.exists():
                return True
        except OSError:
            continue
    return False


def _component_candidates(platform_name: str, component: str) -> list[Path]:
    """Return conservative platform-specific component locations to probe."""
    env_omml = os.environ.get("K12_OMML_DEPENDENCY", "")
    mac_apps = {
        "word": ["/Applications/Microsoft Word.app"],
        "powerpoint": ["/Applications/Microsoft PowerPoint.app"],
        "excel": ["/Applications/Microsoft Excel.app"],
        "mathtype": ["/Applications/MathType.app"],
        "libreoffice": ["/Applications/LibreOffice.app"],
    }
    windows_roots = [
        os.environ.get("ProgramFiles", r"C:\Program Files"),
        os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)"),
    ]
    windows_paths = {
        "word": [r"Microsoft Office\root\Office16\WINWORD.EXE"],
        "powerpoint": [r"Microsoft Office\root\Office16\POWERPNT.EXE"],
        "excel": [r"Microsoft Office\root\Office16\EXCEL.EXE"],
        "mathtype": [r"MathType\MathType.exe"],
        "libreoffice": [r"LibreOffice\program\soffice.exe"],
    }
    if component == "omml_dependency":
        values = [env_omml, "OMML2MML.XSL", "omml2mml.xsl"]
        return [Path(value) for value in values if value]
    if platform_name == "macOS":
        return [Path(value) for value in mac_apps.get(component, [])]
    if platform_name == "Windows":
        return [Path(root) / suffix for root in windows_roots if root for suffix in windows_paths.get(component, [])]
    return []


def summarize_payload(payload: dict[str, Any]) -> dict[str, Any]:
    """Summarize a local task payload without leaking local paths or tokens."""
    task = payload.get("task") if isinstance(payload.get("task"), dict) else {}
    handoff = payload.get("handoff") if isinstance(payload.get("handoff"), dict) else {}
    workflow_plan = payload.get("workflow_plan") if isinstance(payload.get("workflow_plan"), dict) else {}
    sync = payload.get("sync") if isinstance(payload.get("sync"), dict) else {}
    files = payload.get("files") if isinstance(payload.get("files"), list) else []
    actions = payload.get("local_actions") if isinstance(payload.get("local_actions"), list) else []
    execution_plan = payload.get("desktop_execution_plan") if isinstance(payload.get("desktop_execution_plan"), dict) else {}
    safe_files = [
        {
            "id": str(item.get("id") or ""),
            "file_name": str(item.get("file_name") or ""),
            "file_type": str(item.get("file_type") or ""),
            "input_path_exists": bool(item.get("input_path_exists")),
            "validation_error_count": len(item.get("validation_errors") or []),
            "encrypted": bool(item.get("encrypted")),
            "has_omml": bool(item.get("has_omml")),
            "has_macro": bool(item.get("has_macro")),
        }
        for item in files
        if isinstance(item, dict)
    ]
    safe_actions = [
        {
            "type": str(item.get("type") or ""),
            "status": str(item.get("status") or ""),
            "macro_count": len(item.get("macros") or []) if isinstance(item.get("macros"), list) else 0,
            "dependency_count": len(item.get("dependencies") or []) if isinstance(item.get("dependencies"), list) else 0,
            "retry_request_count": len(item.get("retry_requests") or []) if isinstance(item.get("retry_requests"), list) else 0,
            "recognition_request_count": len(item.get("recognition_requests") or []) if isinstance(item.get("recognition_requests"), list) else 0,
            "artifact_count": len(item.get("artifacts") or []) if isinstance(item.get("artifacts"), list) else 0,
        }
        for item in actions
        if isinstance(item, dict)
    ]
    plan_actions = execution_plan.get("actions") if isinstance(execution_plan.get("actions"), list) else []
    plan_platform = _safe_plan_platform(execution_plan.get("platform"))
    formula_delivery = _safe_formula_delivery_contract(_raw_formula_delivery(payload, execution_plan), plan_platform["expected"])
    safe_plan_actions = []
    for item in plan_actions:
        if not isinstance(item, dict):
            continue
        steps = item.get("steps") if isinstance(item.get("steps"), list) else []
        operations = [
            str(step.get("operation") or "")[:120]
            for step in steps
            if isinstance(step, dict) and step.get("operation")
        ][:20]
        required_capabilities = _safe_required_capabilities(item.get("required_capabilities"))
        action_contract = _safe_formula_delivery_contract(
            item.get("formula_delivery") if isinstance(item.get("formula_delivery"), dict) else formula_delivery,
            plan_platform["expected"],
        )
        safe_plan_actions.append(
            {
                "type": str(item.get("type") or ""),
                "label": str(item.get("label") or item.get("type") or ""),
                "gate_status": str(item.get("gate_status") or ""),
                "step_count": len(steps),
                "recognition_request_count": int(item.get("recognition_request_count") or len(item.get("recognition_requests") or [])),
                "required_capabilities": [capability["key"] for capability in required_capabilities],
                "formula_delivery": _action_formula_delivery_summary(
                    str(item.get("type") or ""),
                    required_capabilities,
                    operations,
                    action_contract,
                ),
            }
        )
    return {
        "schema_version": "k12.localClientDryRunSummary.v1",
        "task": {
            "id": str(task.get("id") or ""),
            "task_type": str(task.get("task_type") or ""),
            "task_label": str(task.get("task_label") or ""),
            "execute_mode": str(task.get("execute_mode") or ""),
            "status": str(task.get("status") or ""),
            "progress": task.get("progress", 0),
        },
        "handoff": {
            "requires_local_client": bool(handoff.get("requires_local_client")),
            "status": str(handoff.get("status") or ""),
            "message": str(handoff.get("message") or ""),
        },
        "workflow_plan": {
            "schema_version": str(workflow_plan.get("schema_version") or ""),
            "current_task": str(workflow_plan.get("current_task") or ""),
            "current_index": int(workflow_plan.get("current_index") or 0),
            "order": [str(item) for item in workflow_plan.get("order") or []],
            "labels": [str(item) for item in workflow_plan.get("labels") or []],
        },
        "files": safe_files,
        "file_count": len(safe_files),
        "missing_input_count": sum(1 for item in files if isinstance(item, dict) and not item.get("input_path_exists")),
        "actions": safe_actions,
        "action_count": len(safe_actions),
        "desktop_execution_plan": {
            "schema_version": str(execution_plan.get("schema_version") or ""),
            "status": str(execution_plan.get("status") or ""),
            "native_execution_allowed": bool(execution_plan.get("native_execution_allowed")),
            "web_executes_native_documents": bool(execution_plan.get("web_executes_native_documents")),
            "current_companion_cli_executes_native_documents": bool(execution_plan.get("current_companion_cli_executes_native_documents")),
            "native_action_count": int(execution_plan.get("native_action_count") or 0),
            "platform": plan_platform,
            "formula_delivery": formula_delivery,
            "actions": safe_plan_actions,
        },
        "sync": {
            "task_status_cloud_sync_allowed": bool(sync.get("task_status_cloud_sync_allowed")),
            "result_upload_allowed": bool(sync.get("result_upload_allowed")),
            "policy": str(sync.get("policy") or ""),
        },
        "guardrails": [
            "dry-run 只验证本地交接和状态同步",
            "不会执行 Office、MathType、OMML 写回或 Word 宏",
            "摘要不输出 input_path、output_path 或安全令牌",
        ],
    }


def build_dry_run_execution_summary(payload: dict[str, Any]) -> dict[str, Any]:
    """Report desktop-plan readiness and same-platform formula delivery contract without executing native document actions."""
    task = payload.get("task") if isinstance(payload.get("task"), dict) else {}
    execution_plan = payload.get("desktop_execution_plan") if isinstance(payload.get("desktop_execution_plan"), dict) else {}
    raw_actions = execution_plan.get("actions") if isinstance(execution_plan.get("actions"), list) else []
    plan_platform = _safe_plan_platform(execution_plan.get("platform"))
    formula_delivery = _safe_formula_delivery_contract(_raw_formula_delivery(payload, execution_plan), plan_platform["expected"])
    actions = [
        _dry_run_action_summary(action, formula_delivery, plan_platform["actual"], plan_platform["expected"])
        for action in raw_actions
        if isinstance(action, dict)
    ]
    return {
        "schema_version": "k12.localDryRunExecution.v1",
        "task": {
            "id": str(task.get("id") or execution_plan.get("task_id") or ""),
            "task_type": str(task.get("task_type") or execution_plan.get("task_type") or ""),
            "task_label": str(task.get("task_label") or execution_plan.get("task_label") or ""),
        },
        "plan_status": str(execution_plan.get("status") or ""),
        "native_execution_allowed": bool(execution_plan.get("native_execution_allowed")),
        "companion_cli_native_execution": False,
        "action_count": len(actions),
        "ready_action_count": sum(1 for action in actions if action["dry_run_status"] == "ready_for_native_executor"),
        "blocked_action_count": sum(1 for action in actions if action["dry_run_status"].startswith("blocked")),
        "waiting_action_count": sum(1 for action in actions if action["dry_run_status"] == "waiting_for_heartbeat"),
        "platform": plan_platform,
        "formula_delivery": formula_delivery,
        "actions": actions,
        "guardrails": [
            "dry-run 只校验桌面执行计划、平台公式合同、能力门槛和输出合同",
            "不会执行 Office、MathType、OMML 写回或 Word 宏",
            "不回传 input_path、output_directory、backup_path 或安全令牌",
        ],
    }


def _dry_run_action_summary(
    action: dict[str, Any],
    plan_formula_delivery: dict[str, Any],
    detected_platform: str,
    expected_platform: str,
) -> dict[str, Any]:
    """Summarize one desktop action for dry-run without exposing local paths."""
    gate_status = str(action.get("gate_status") or "")
    required_capabilities = _safe_required_capabilities(action.get("required_capabilities"))
    steps = action.get("steps") if isinstance(action.get("steps"), list) else []
    operations = [
        str(step.get("operation") or "")[:120]
        for step in steps
        if isinstance(step, dict) and step.get("operation")
    ][:20]
    formula_delivery = _safe_formula_delivery_contract(
        action.get("formula_delivery") if isinstance(action.get("formula_delivery"), dict) else plan_formula_delivery,
        expected_platform,
    )
    blockers = _dry_run_blockers(gate_status, required_capabilities, action, formula_delivery, detected_platform, expected_platform)
    output_contract = action.get("output_contract") if isinstance(action.get("output_contract"), dict) else {}
    return {
        "action_id": str(action.get("action_id") or "")[:120],
        "type": str(action.get("type") or "")[:80],
        "label": str(action.get("label") or action.get("type") or "")[:120],
        "payload_status": str(action.get("status") or "")[:80],
        "gate_status": gate_status[:80],
        "dry_run_status": _dry_run_status(gate_status, blockers),
        "native_execution_performed": False,
        "required_capabilities": required_capabilities,
        "step_count": len(steps),
        "required_step_count": sum(1 for step in steps if isinstance(step, dict) and bool(step.get("required"))),
        "recognition_request_count": int(action.get("recognition_request_count") or len(action.get("recognition_requests") or [])),
        "output_artifact_types": [str(item)[:40] for item in output_contract.get("artifact_types") or [] if isinstance(item, str)],
        "result_upload_optional": bool(output_contract.get("result_upload_optional")),
        "formula_delivery": _action_formula_delivery_summary(
            str(action.get("type") or ""),
            required_capabilities,
            operations,
            formula_delivery,
        ),
        "blockers": blockers,
    }


def _safe_required_capabilities(value: Any) -> list[dict[str, Any]]:
    """Keep only known local-client capability gates from an action payload."""
    if not isinstance(value, list):
        return []
    capabilities: list[dict[str, Any]] = []
    for item in value:
        if not isinstance(item, dict):
            continue
        key = str(item.get("key") or "")
        if key not in SAFE_CAPABILITY_KEYS:
            continue
        capabilities.append(
            {
                "key": key,
                "label": str(item.get("label") or key)[:80],
                "available": bool(item.get("available")),
            }
        )
    return capabilities


def _dry_run_blockers(
    gate_status: str,
    required_capabilities: list[dict[str, Any]],
    action: dict[str, Any],
    formula_delivery: dict[str, Any],
    detected_platform: str,
    expected_platform: str,
) -> list[str]:
    """Merge gate, capability, and formula-platform checks into redacted dry-run blockers."""
    blockers: list[str] = []
    if gate_status == "waiting_for_heartbeat":
        blockers.append("local_client_heartbeat_required")
    elif gate_status == "blocked_by_platform":
        blockers.append("platform_mismatch")
    elif gate_status == "blocked_by_capability":
        missing = [item["key"] for item in required_capabilities if not item.get("available")]
        blockers.extend([f"missing_capability:{key}" for key in missing] or ["missing_capability"])
    elif gate_status:
        blockers.append(gate_status)
    else:
        blockers.append("pending_preflight")
    if _action_uses_mathtype_contract(action, required_capabilities):
        blockers.extend(_formula_delivery_blockers(formula_delivery, detected_platform, expected_platform))
    return list(dict.fromkeys(blockers))


def _dry_run_status(gate_status: str, blockers: list[str]) -> str:
    """Collapse gate and blocker details into one dry-run status value."""
    platform_blockers = {
        "formula_delivery_contract_missing",
        "platform_mismatch",
        "formula_delivery_platform_mismatch",
        "formula_contract_expected_platform_mismatch",
    }
    if platform_blockers.intersection(blockers):
        return "blocked_by_platform"
    if any(item.startswith("missing_capability:") for item in blockers) or "missing_capability" in blockers:
        return "blocked_by_capability"
    if gate_status == "ready":
        return "ready_for_native_executor"
    if gate_status == "waiting_for_heartbeat":
        return "waiting_for_heartbeat"
    if gate_status == "blocked_by_platform":
        return "blocked_by_platform"
    if gate_status == "blocked_by_capability":
        return "blocked_by_capability"
    if blockers:
        return "pending_preflight"
    return "pending_preflight"


def _safe_plan_platform(value: Any) -> dict[str, Any]:
    """Return a path-free platform summary from a desktop execution plan."""
    raw = value if isinstance(value, dict) else {}
    expected = normalize_platform(str(raw.get("expected") or ""))
    actual = normalize_platform(str(raw.get("actual") or ""))
    compatible_default = not (
        expected in {"Windows", "macOS"}
        and actual in {"Windows", "macOS"}
        and expected != actual
    )
    return {
        "expected": expected,
        "actual": actual,
        "compatible": bool(raw.get("compatible", compatible_default)),
        "same_platform_required_for_native_mathtype": bool(raw.get("same_platform_required_for_native_mathtype")),
        "message": str(raw.get("message") or "")[:240],
    }


def _raw_formula_delivery(payload: dict[str, Any], execution_plan: dict[str, Any]) -> dict[str, Any]:
    """Prefer the desktop plan formula contract, then the payload contract."""
    if isinstance(execution_plan.get("formula_delivery"), dict):
        return execution_plan["formula_delivery"]
    if isinstance(payload.get("formula_delivery"), dict):
        return payload["formula_delivery"]
    return {}


def build_dry_run_sync_payload(payload: dict[str, Any], result_upload_requested: bool = False) -> dict[str, Any]:
    """Build a local-sync payload for dry-run handoff validation only."""
    summary = summarize_payload(payload)
    execution = build_dry_run_execution_summary(payload)
    progress = summary["task"].get("progress", 0)
    try:
        progress_value = max(int(progress or 0), 5)
    except (TypeError, ValueError):
        progress_value = 5
    action_count = execution["action_count"]
    ready_count = execution["ready_action_count"]
    blocked_count = execution["blocked_action_count"] + execution["waiting_action_count"]
    return {
        "status": "running",
        "progress": min(progress_value, 95),
        "message": f"本地客户端已接收任务载荷；dry-run 校验 {ready_count}/{action_count} 个动作可交接，{blocked_count} 个动作等待或阻塞；不执行 Office、MathType、OMML 写回或 Word 宏",
        "outputs": [],
        "dryRunExecution": execution,
        "resultUploadRequested": bool(result_upload_requested),
    }


def build_native_execution_request(payload: dict[str, Any], platform_name: str = "", allow_native_execution: bool = False) -> dict[str, Any]:
    """Build a handoff contract for a future same-platform native runner.

    The request lists operations and blockers for a real runner, but this CLI
    still reports ``native_execution_performed=False`` and never touches the
    document through Office, MathType, OMML writeback, or Word macro APIs.
    """
    task = payload.get("task") if isinstance(payload.get("task"), dict) else {}
    execution_plan = payload.get("desktop_execution_plan") if isinstance(payload.get("desktop_execution_plan"), dict) else {}
    plan_platform = execution_plan.get("platform") if isinstance(execution_plan.get("platform"), dict) else {}
    detected_platform = normalize_platform(platform_name or platform.system())
    expected_platform = normalize_platform(str(plan_platform.get("expected") or detected_platform))
    raw_formula_delivery = execution_plan.get("formula_delivery")
    if not isinstance(raw_formula_delivery, dict):
        raw_formula_delivery = (
            payload.get("formula_delivery")
            if isinstance(payload.get("formula_delivery"), dict)
            else {}
        )
    formula_delivery = _safe_formula_delivery_contract(raw_formula_delivery, expected_platform)
    same_platform_required = bool(formula_delivery.get("native_object_requires_same_platform")) or expected_platform in {"Windows", "macOS"}
    runner_profile = _native_runner_profile(detected_platform)
    raw_actions = execution_plan.get("actions") if isinstance(execution_plan.get("actions"), list) else []
    action_requests = [
        _native_action_request(
            action,
            detected_platform,
            expected_platform,
            bool(allow_native_execution),
            bool(execution_plan.get("native_execution_allowed")),
            formula_delivery,
            runner_profile,
        )
        for action in raw_actions
        if isinstance(action, dict)
    ]
    ready_count = sum(1 for action in action_requests if action["native_request_status"] == "ready_for_native_runner")
    blocked_count = sum(1 for action in action_requests if action["native_request_status"].startswith("blocked"))
    return {
        "schema_version": "k12.localNativeExecutionRequest.v1",
        "task": {
            "id": str(task.get("id") or execution_plan.get("task_id") or ""),
            "task_type": str(task.get("task_type") or execution_plan.get("task_type") or ""),
            "task_label": str(task.get("task_label") or execution_plan.get("task_label") or ""),
        },
        "platform": {
            "detected": detected_platform,
            "expected": expected_platform,
            "compatible": bool(plan_platform.get("compatible", detected_platform == expected_platform or expected_platform == "auto")),
            "windows_native_runner": detected_platform == "Windows",
            "macos_native_runner": detected_platform == "macOS",
            "runner_profile": runner_profile,
            "mathtype_objects_cross_platform_compatible": False,
            "same_platform_required_for_native_mathtype": same_platform_required,
            "fallback_formula_formats": ["MathML", "LaTeX", "image"],
            "message": "Windows 原生执行请求可交给 pywin32/Office COM 适配器；macOS 暂按兜底合同处理；MathType 原生对象必须按同平台交接。",
        },
        "formula_delivery": formula_delivery,
        "native_execution_requested": bool(allow_native_execution),
        "native_execution_allowed_by_payload": bool(execution_plan.get("native_execution_allowed")),
        "companion_cli_executes_native_documents": False,
        "action_count": len(action_requests),
        "ready_action_count": ready_count,
        "blocked_action_count": blocked_count,
        "actions": action_requests,
        "guardrails": [
            "本函数只生成原生执行请求合同，不直接调用 Office、MathType、OMML 写回或 Word 宏",
            "真实执行必须由后续同平台本地执行器显式接管，并保留任务状态同步",
            "请求摘要不输出 input_path、output_directory、backup_path 或安全令牌",
        ],
    }


def _native_action_request(
    action: dict[str, Any],
    detected_platform: str,
    expected_platform: str,
    allow_native_execution: bool,
    plan_allows_native_execution: bool,
    formula_delivery: dict[str, Any],
    runner_profile: dict[str, Any],
) -> dict[str, Any]:
    """Build one future native-runner request while keeping execution disabled."""
    action_type = str(action.get("type") or "")[:80]
    required_capabilities = _safe_required_capabilities(action.get("required_capabilities"))
    steps = action.get("steps") if isinstance(action.get("steps"), list) else []
    output_contract = action.get("output_contract") if isinstance(action.get("output_contract"), dict) else {}
    operations = [
        str(step.get("operation") or "")[:120]
        for step in steps
        if isinstance(step, dict) and step.get("operation")
    ][:20]
    blockers = _native_action_blockers(
        action,
        required_capabilities,
        detected_platform,
        expected_platform,
        allow_native_execution,
        plan_allows_native_execution,
        formula_delivery,
        runner_profile,
    )
    status = "ready_for_native_runner" if not blockers else _native_blocked_status(blockers)
    return {
        "action_id": str(action.get("action_id") or "")[:120],
        "type": action_type,
        "label": str(action.get("label") or action_type)[:120],
        "gate_status": str(action.get("gate_status") or "")[:80],
        "native_request_status": status,
        "required_capabilities": required_capabilities,
        "operation_count": len(operations),
        "recognition_request_count": int(action.get("recognition_request_count") or len(action.get("recognition_requests") or [])),
        "operations": operations,
        "output_artifact_types": [
            str(item)[:40]
            for item in output_contract.get("artifact_types", [])
            if isinstance(item, str)
        ][:12],
        "sensitive_path_present": bool(action.get("input_paths") or action.get("backup_paths") or output_contract.get("output_directory")),
        "native_execution_performed": False,
        "formula_delivery": _action_formula_delivery_summary(action_type, required_capabilities, operations, formula_delivery),
        "blockers": blockers,
    }


def _native_action_blockers(
    action: dict[str, Any],
    required_capabilities: list[dict[str, Any]],
    detected_platform: str,
    expected_platform: str,
    allow_native_execution: bool,
    plan_allows_native_execution: bool,
    formula_delivery: dict[str, Any],
    runner_profile: dict[str, Any],
) -> list[str]:
    """Collect why a native action request cannot be handed to a runner yet."""
    blockers: list[str] = []
    gate_status = str(action.get("gate_status") or "")
    if not allow_native_execution:
        blockers.append("native_execution_not_requested")
    if not plan_allows_native_execution:
        blockers.append("payload_not_ready_for_native_execution")
    if gate_status != "ready":
        blockers.append(f"gate_status:{gate_status or 'pending'}")
    if expected_platform in {"Windows", "macOS"} and detected_platform in {"Windows", "macOS"} and detected_platform != expected_platform:
        blockers.append("platform_mismatch")
    if _action_uses_mathtype_contract(action, required_capabilities):
        blockers.extend(_formula_delivery_blockers(formula_delivery, detected_platform, expected_platform))
    if detected_platform == "macOS" and _action_requires_native_document_runner(action):
        blockers.append("macos_native_runner_limited")
    elif detected_platform != "Windows" and _action_requires_native_document_runner(action):
        blockers.append("native_runner_not_available_for_platform")
    elif detected_platform not in {"Windows", "macOS"}:
        blockers.append("native_runner_not_available_for_platform")
    if action.get("type") == "open_output_directory" and "shell.open_output_directory" not in runner_profile.get("supported_operations", []):
        blockers.append("native_runner_not_available_for_platform")
    blockers.extend(f"missing_capability:{item['key']}" for item in required_capabilities if not item.get("available"))
    return list(dict.fromkeys(blockers))


def _native_blocked_status(blockers: list[str]) -> str:
    """Return the highest-level native request block category."""
    platform_blockers = {
        "formula_delivery_contract_missing",
        "platform_mismatch",
        "formula_delivery_platform_mismatch",
        "formula_contract_expected_platform_mismatch",
        "macos_native_runner_limited",
        "native_runner_not_available_for_platform",
    }
    if platform_blockers.intersection(blockers):
        return "blocked_by_platform"
    if any(item.startswith("missing_capability:") for item in blockers):
        return "blocked_by_capability"
    if "native_execution_not_requested" in blockers:
        return "blocked_until_explicit_native_request"
    if "payload_not_ready_for_native_execution" in blockers:
        return "blocked_by_payload"
    return "blocked_preflight"


def _native_runner_profile(detected_platform: str) -> dict[str, Any]:
    """Describe which future native adapter may consume the request contract."""
    platform_name = normalize_platform(detected_platform)
    if platform_name == "Windows":
        return {
            "schema_version": "k12.nativeRunnerProfile.v1",
            "platform": "Windows",
            "support_level": "windows_office_com_adapter",
            "native_document_runner_available": True,
            "office_automation_adapter": "pywin32 / Office COM",
            "mathtype_adapter": "Windows MathType OLE / Equation Native",
            "macro_adapter": "Word VBA / COM",
            "supported_operations": [
                "office.open_source",
                "office.convert",
                "mathtype.convert",
                "macro.run_ordered",
                "shell.open_output_directory",
            ],
            "guardrail": "仅同平台交付 Windows MathType 原生对象，跨平台必须附带 MathML、LaTeX 或图片兜底。",
        }
    if platform_name == "macOS":
        return {
            "schema_version": "k12.nativeRunnerProfile.v1",
            "platform": "macOS",
            "support_level": "macos_limited_handoff",
            "native_document_runner_available": False,
            "office_automation_adapter": "受限：不声明已接入 Office for Mac 自动化",
            "mathtype_adapter": "受限：仅登记 macOS MathType 同平台合同和兜底格式",
            "macro_adapter": "不可直接执行 Word 宏",
            "supported_operations": [
                "formula.export_fallbacks",
                "manual_review",
                "shell.open_output_directory",
            ],
            "guardrail": "macOS MathType 对象不能交给 Windows 链路；当前 CLI 只生成受限交接合同，不执行原生写回。",
        }
    return {
        "schema_version": "k12.nativeRunnerProfile.v1",
        "platform": platform_name or "Unknown",
        "support_level": "unsupported_platform",
        "native_document_runner_available": False,
        "office_automation_adapter": "",
        "mathtype_adapter": "未知平台仅允许 MathML、LaTeX 或图片兜底",
        "macro_adapter": "",
        "supported_operations": [],
        "guardrail": "请先选择 Windows 或 macOS 安装画像，再生成平台专属执行请求。",
    }


def _action_requires_native_document_runner(action: dict[str, Any]) -> bool:
    """Return True for actions that need Office, MathType, OMML, or macro APIs."""
    action_type = str(action.get("type") or "")
    if action_type in {"macro_sequence", "omml_mathtype", "pdf_formula_mathtype", "office_conversion"}:
        return True
    steps = action.get("steps") if isinstance(action.get("steps"), list) else []
    operations = [str(step.get("operation") or "").lower() for step in steps if isinstance(step, dict)]
    return any(
        marker in operation
        for operation in operations
        for marker in ("office.", "mathtype.", "macro.", "omml.", "formula.merge_into_docx")
    )


def _safe_formula_delivery_contract(value: dict[str, Any], fallback_platform: str) -> dict[str, Any]:
    """Return a path-safe MathType delivery contract for native runners."""
    contract_present = bool(value.get("contract_present", bool(value)))
    platform_name = normalize_platform(str(value.get("platform") or fallback_platform or "Unknown"))
    fallback_formats = _safe_formula_format_list(value.get("fallback_formats"), ["MathML", "LaTeX", "图片"])
    output_priority = _safe_formula_format_list(value.get("output_priority"), fallback_formats)
    mode = str(value.get("compatibility_mode") or "platform-specific")[:80]
    raw_native_allowed = bool(value.get("native_mathtype_object_allowed", platform_name in {"Windows", "macOS"}))
    native_requires_same_platform = bool(value.get("native_object_requires_same_platform", platform_name in {"Windows", "macOS"}))
    if raw_native_allowed and platform_name in {"Windows", "macOS"}:
        native_requires_same_platform = True
    native_allowed = bool(contract_present and raw_native_allowed and platform_name in {"Windows", "macOS"})
    return {
        "schema_version": "k12.formulaDeliveryContract.v1",
        "contract_present": contract_present,
        "platform": platform_name,
        "compatibility_mode": mode,
        "platform_object_format": str(value.get("platform_object_format") or "")[:120],
        "platform_objects_cross_compatible": False,
        "native_mathtype_object_allowed": native_allowed and platform_name in {"Windows", "macOS"},
        "native_object_requires_same_platform": native_requires_same_platform,
        "cross_platform_safe": bool(value.get("cross_platform_safe", mode in {"mathml-latex", "image-fallback"})),
        "fallback_formats": fallback_formats,
        "output_priority": output_priority,
        "message": str(value.get("message") or "Windows 与 macOS MathType 原生对象不跨平台通用，原生对象必须同平台交接。")[:240],
    }


def _safe_formula_format_list(value: Any, default: list[str]) -> list[str]:
    """Keep formula format lists compact and serializable."""
    if not isinstance(value, list):
        return list(default)
    formats = [str(item)[:40] for item in value if isinstance(item, str) and str(item).strip()]
    return formats[:8] or list(default)


def _action_uses_mathtype_contract(action: dict[str, Any], required_capabilities: list[dict[str, Any]]) -> bool:
    """Detect actions that must respect same-platform MathType object rules."""
    action_type = str(action.get("type") or "")
    if action_type in {"omml_mathtype", "pdf_formula_mathtype", "office_conversion"}:
        return True
    if any(item.get("key") == "mathTypeAutomation" for item in required_capabilities):
        return True
    steps = action.get("steps") if isinstance(action.get("steps"), list) else []
    operations = [str(step.get("operation") or "") for step in steps if isinstance(step, dict)]
    return any("mathtype" in operation.lower() or "formula" in operation.lower() for operation in operations)


def _formula_delivery_blockers(formula_delivery: dict[str, Any], detected_platform: str, expected_platform: str) -> list[str]:
    """Block native requests that cross the declared formula delivery platform."""
    blockers: list[str] = []
    if not bool(formula_delivery.get("contract_present")):
        blockers.append("formula_delivery_contract_missing")
    if not bool(formula_delivery.get("native_object_requires_same_platform")):
        return blockers
    contract_platform = normalize_platform(str(formula_delivery.get("platform") or ""))
    if contract_platform in {"Windows", "macOS"} and detected_platform in {"Windows", "macOS"} and detected_platform != contract_platform:
        blockers.append("formula_delivery_platform_mismatch")
    if contract_platform in {"Windows", "macOS"} and expected_platform in {"Windows", "macOS"} and expected_platform != contract_platform:
        blockers.append("formula_contract_expected_platform_mismatch")
    return blockers


def _action_formula_delivery_summary(
    action_type: str,
    required_capabilities: list[dict[str, Any]],
    operations: list[str],
    formula_delivery: dict[str, Any],
) -> dict[str, Any]:
    """Attach the relevant formula handoff decision to one native action."""
    uses_contract = action_type in {"omml_mathtype", "pdf_formula_mathtype", "office_conversion"}
    uses_contract = uses_contract or any(item.get("key") == "mathTypeAutomation" for item in required_capabilities)
    uses_contract = uses_contract or any("mathtype" in operation.lower() or "formula" in operation.lower() for operation in operations)
    return {
        "required": uses_contract,
        "platform": str(formula_delivery.get("platform") or ""),
        "native_mathtype_object_allowed": bool(formula_delivery.get("native_mathtype_object_allowed")),
        "native_object_requires_same_platform": bool(formula_delivery.get("native_object_requires_same_platform")),
        "platform_objects_cross_compatible": False,
        "fallback_formats": list(formula_delivery.get("fallback_formats") or []),
    }


def execute_local_file_actions(payload: dict[str, Any], allow_file_actions: bool = False) -> dict[str, Any]:
    """Execute only explicitly allowed local file actions from a task payload.

    At this stage the allowlist is deliberately tiny: OMML dependency files may
    be copied into the source document directory, while every document-native
    action remains a handoff contract for a separate desktop runner.
    """
    actions = payload.get("local_actions") if isinstance(payload.get("local_actions"), list) else []
    results: list[dict[str, Any]] = []
    for action in actions:
        if not isinstance(action, dict) or str(action.get("type") or "") != "omml_mathtype":
            continue
        copy_strategy = str(action.get("copy_strategy") or "自动重命名")
        dependencies = action.get("dependencies") if isinstance(action.get("dependencies"), list) else []
        for dependency in dependencies:
            if isinstance(dependency, dict):
                results.append(_execute_omml_dependency_copy(dependency, copy_strategy, bool(allow_file_actions)))
    performed_count = sum(1 for item in results if item["status"] == "copied")
    blocked_count = sum(1 for item in results if item["status"].startswith("blocked"))
    skipped_count = sum(1 for item in results if item["status"].startswith("skipped"))
    return {
        "schema_version": "k12.localFileActionExecution.v1",
        "file_actions_requested": bool(allow_file_actions),
        "action_count": len(results),
        "performed_count": performed_count,
        "blocked_count": blocked_count,
        "skipped_count": skipped_count,
        "actions": results,
        "guardrails": [
            "只执行 OMML 依赖复制等本地文件动作",
            "仅允许 .xsl、.xslt、.xml、.mml 依赖文件",
            "目标必须是当前文档所在目录",
            "目标文档目录必须具备写入权限",
            "不会执行 Office、MathType、OMML 写回或 Word 宏",
            "结果不输出 document_path、source_path 或 target_path",
        ],
    }


def _execute_omml_dependency_copy(dependency: dict[str, Any], copy_strategy: str, allow_file_actions: bool) -> dict[str, Any]:
    """Copy one OMML dependency without exposing local paths in the result."""
    dependency_id = str(dependency.get("id") or "")[:80]
    document_path = Path(str(dependency.get("document_path") or ""))
    source_path = Path(str(dependency.get("omml_source_path") or ""))
    target_path = Path(str(dependency.get("omml_target_path") or ""))
    if not target_path.name and source_path.name:
        target_path = document_path.parent / source_path.name
    result = {
        "type": "omml_dependency_copy",
        "dependency_id": dependency_id,
        "file_id": str(dependency.get("file_id") or "")[:80],
        "file_name": str(dependency.get("file_name") or dependency.get("omml_file_name") or "")[:160],
        "copy_strategy": copy_strategy,
        "status": "pending",
        "copy_performed": False,
        "source_available": False,
        "dependency_extension_allowed": False,
        "target_directory_allowed": False,
        "target_directory_writeable": False,
        "source_sha256": "",
        "target_sha256": "",
        "blockers": [],
    }
    blockers: list[str] = []
    if not allow_file_actions:
        blockers.append("file_actions_not_requested")
    if source_path.suffix.lower() in OMML_DEPENDENCY_SUFFIXES:
        result["dependency_extension_allowed"] = True
    else:
        blockers.append("unsupported_dependency_extension")
    if not source_path.is_file():
        blockers.append("source_missing")
    else:
        result["source_available"] = True
    # Keep the copy target beside the source document so a malicious payload
    # cannot turn the companion CLI into a general-purpose file writer.
    try:
        document_dir = document_path.parent.resolve()
        target_dir = target_path.parent.resolve()
        result["target_directory_allowed"] = target_dir == document_dir
    except OSError:
        result["target_directory_allowed"] = False
    if not result["target_directory_allowed"]:
        blockers.append("target_not_document_directory")
    elif target_path.parent.exists() and os.access(target_path.parent, os.W_OK):
        result["target_directory_writeable"] = True
    else:
        blockers.append("target_directory_not_writeable")
    if blockers:
        result["status"] = _file_action_blocked_status(blockers)
        result["blockers"] = blockers
        return result

    target_path.parent.mkdir(parents=True, exist_ok=True)
    if target_path.exists() and copy_strategy == "跳过":
        result["status"] = "skipped_existing_target"
        result["target_sha256"] = _sha256_file(target_path)
        return result
    if target_path.exists() and copy_strategy == "自动重命名":
        target_path = _unique_local_target_path(target_path)
    try:
        # This is the only write operation in the companion CLI today.
        shutil.copy2(source_path, target_path)
        result["status"] = "copied"
        result["copy_performed"] = True
        result["source_sha256"] = _sha256_file(source_path)
        result["target_sha256"] = _sha256_file(target_path)
    except OSError as exc:
        result["status"] = "blocked_copy_failed"
        result["blockers"] = [f"copy_failed:{exc.__class__.__name__}"]
    return result


def _file_action_blocked_status(blockers: list[str]) -> str:
    """Choose the user-facing status for a blocked safe file action."""
    if "file_actions_not_requested" in blockers:
        return "blocked_until_explicit_file_action_request"
    if "unsupported_dependency_extension" in blockers:
        return "blocked_unsupported_dependency_extension"
    if "source_missing" in blockers:
        return "blocked_source_missing"
    if "target_not_document_directory" in blockers:
        return "blocked_target_directory"
    if "target_directory_not_writeable" in blockers:
        return "blocked_target_directory_not_writeable"
    return "blocked_preflight"


def _sha256_file(path: Path) -> str:
    """Hash a local file after an explicitly authorized safe file action."""
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _unique_local_target_path(path: Path) -> Path:
    """Create a non-conflicting OMML dependency copy target path."""
    stem = path.stem
    suffix = path.suffix
    for index in range(1, 10_000):
        candidate = path.with_name(f"{stem}-{index}{suffix}")
        if not candidate.exists():
            return candidate
    raise OSError("无法生成唯一 OMML 目标文件名")


def request_json(origin: str, path: str, method: str = "GET", token: str = "", payload: dict[str, Any] | None = None, timeout: float = 10.0) -> dict[str, Any]:
    """Call one local API JSON endpoint with an optional security token."""
    body = None if payload is None else json.dumps(payload, ensure_ascii=False).encode("utf-8")
    headers = {"Accept": "application/json"}
    if body is not None:
        headers["Content-Type"] = "application/json; charset=utf-8"
    if token:
        headers["X-K12-Token"] = token
    request = Request(f"{normalize_origin(origin)}{path}", data=body, headers=headers, method=method)
    try:
        with urlopen(request, timeout=timeout) as response:
            raw = response.read().decode("utf-8")
    except HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        raise LocalClientError(f"本地 API 返回 {exc.code}: {detail}") from exc
    except URLError as exc:
        raise LocalClientError(f"无法连接本地 API: {exc.reason}") from exc
    if not raw.strip():
        return {}
    return json.loads(raw)


def run_once(
    origin: str,
    token: str = "",
    task_id: str = "",
    sync_dry_run: bool = False,
    result_upload_requested: bool = False,
    native_plan: bool = False,
    allow_native_execution: bool = False,
    execute_file_actions: bool = False,
    timeout: float = 10.0,
    client_id: str = "",
) -> dict[str, Any]:
    """Run one safe companion-client cycle against the local K12 API.

    The cycle can send a heartbeat, fetch a token-gated task payload, produce a
    dry-run sync payload, and optionally emit a native execution request
    contract. It does not run Office, MathType, OMML writeback, or Word macros.
    """
    safe_origin = normalize_origin(origin)
    manifest_response = request_json(safe_origin, "/api/local-client/manifest", token=token, timeout=timeout)
    manifest = manifest_response.get("localClientManifest", manifest_response)
    security = manifest.get("security") if isinstance(manifest.get("security"), dict) else {}
    if security.get("heartbeat_requires_configured_token") and not security.get("token_configured"):
        raise LocalClientError("本地 API 尚未配置安全令牌；请先在设置页保存本地安全令牌，再通过 --token 或 K12_LOCAL_TOKEN 运行本地伴随 CLI")
    selected_task_id = task_id or str((manifest.get("queue") or {}).get("next_task_id") or "")
    heartbeat_payload = build_heartbeat(client_id=client_id, active_task_id=selected_task_id)
    heartbeat_response = request_json(safe_origin, "/api/local-client/heartbeat", method="POST", token=token, payload=heartbeat_payload, timeout=timeout)
    result: dict[str, Any] = {
        "schema_version": "k12.localClientRunResult.v1",
        "origin": safe_origin,
        "manifest": {
            "schema_version": manifest.get("schema_version"),
            "platform": (manifest.get("platform") or {}).get("selected", ""),
            "pending_local_task_count": (manifest.get("queue") or {}).get("pending_local_task_count", 0),
            "next_task_id": (manifest.get("queue") or {}).get("next_task_id", ""),
        },
        "heartbeat": heartbeat_response.get("heartbeat", heartbeat_response),
        "task": None,
        "sync": None,
        "native_execution_request": None,
        "file_action_execution": None,
    }
    if not selected_task_id:
        return result
    payload_path = f"/api/tasks/{quote(selected_task_id)}/local-payload"
    payload_response = request_json(safe_origin, payload_path, token=token, timeout=timeout)
    local_payload = payload_response.get("localPayload", payload_response)
    result["task"] = summarize_payload(local_payload)
    if native_plan:
        result["native_execution_request"] = build_native_execution_request(local_payload, allow_native_execution=allow_native_execution)
    if execute_file_actions:
        result["file_action_execution"] = execute_local_file_actions(local_payload, allow_file_actions=True)
    if sync_dry_run:
        sync_payload = build_dry_run_sync_payload(local_payload, result_upload_requested=result_upload_requested)
        sync_path = f"/api/tasks/{quote(selected_task_id)}/local-sync"
        result["sync"] = request_json(safe_origin, sync_path, method="POST", token=token, payload=sync_payload, timeout=timeout)
    return result


def build_parser() -> argparse.ArgumentParser:
    """Create the CLI parser for safe local-client handoff commands."""
    parser = argparse.ArgumentParser(description="Run the K12 local companion client in safe dry-run mode.")
    parser.add_argument("--origin", default="http://127.0.0.1:8765", help="K12 local API origin")
    parser.add_argument("--token", default=os.environ.get("K12_LOCAL_TOKEN", ""), help="local API security token or K12_LOCAL_TOKEN")
    parser.add_argument("--task-id", default="", help="task id to fetch; defaults to manifest next_task_id")
    parser.add_argument("--sync-dry-run", action="store_true", help="sync a running status without executing native document actions")
    parser.add_argument("--native-plan", action="store_true", help="include a native execution request contract without executing native document actions")
    parser.add_argument("--allow-native-execution", action="store_true", help="mark native execution as explicitly requested in the request contract; no native action is executed by this CLI")
    parser.add_argument("--execute-file-actions", action="store_true", help="execute safe local file actions such as OMML dependency copy; does not execute Office, MathType, or Word macros")
    parser.add_argument("--result-upload-requested", action="store_true", help="mark result upload intent during dry-run sync")
    parser.add_argument("--timeout", type=float, default=10.0, help="HTTP timeout in seconds")
    parser.add_argument("--client-id", default="", help="client id reported in heartbeat")
    parser.add_argument("--compact", action="store_true", help="print compact JSON")
    return parser


def main(argv: list[str] | None = None) -> int:
    """CLI entrypoint that prints a redacted JSON handoff result."""
    parser = build_parser()
    args = parser.parse_args(argv)
    try:
        result = run_once(
            origin=args.origin,
            token=args.token,
            task_id=args.task_id,
            sync_dry_run=args.sync_dry_run,
            result_upload_requested=args.result_upload_requested,
            native_plan=args.native_plan,
            allow_native_execution=args.allow_native_execution,
            execute_file_actions=args.execute_file_actions,
            timeout=args.timeout,
            client_id=args.client_id,
        )
    except (LocalClientError, json.JSONDecodeError) as exc:
        print(f"K12 local client failed: {exc}", file=sys.stderr)
        return 2
    json.dump(result, sys.stdout, ensure_ascii=False, indent=None if args.compact else 2)
    sys.stdout.write("\n")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
