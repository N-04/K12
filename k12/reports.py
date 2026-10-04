"""为任务结果和验收证据生成报告。将分析结果输出为 JSON、HTML、文本、PDF、XLSX 和失败 CSV，记录质量检查、Mathpix 状态、本地客户端交接及可处理的失败原因。"""

from __future__ import annotations

import csv
import hashlib
import html
import json
import zipfile
from io import StringIO
from pathlib import Path
from typing import Any

from .converters import build_text_pdf
from .models import ReportItem
from .previews import build_file_preview


PENDING_ARTIFACT_STATUSES = {"待本地客户端执行"}
CANCELLED_ARTIFACT_STATUSES = {"local_execution_cancelled"}


class ReportBuilder:
    """根据规范化任务分析生成和重写报告产物。"""

    def __init__(self, store: Any) -> None:
        """初始化当前对象所需的配置、依赖与运行状态。"""
        self.store = store

    def build(self, task: dict[str, Any], files: list[dict[str, Any]], analysis: dict[str, Any]) -> dict[str, Any]:
        """创建报告载荷并写入所有可下载报告格式。"""
        formulas = analysis.get("formulas", [])
        macros = analysis.get("macros", [])
        small_images = analysis.get("smallImages", [])
        image_errors = analysis.get("imageExtractionErrors", [])
        mathpix_jobs = analysis.get("mathpix", [])
        omml_dependencies = analysis.get("ommlDependencies", [])
        omml_prompts = analysis.get("ommlConversionPrompts", [])
        artifacts = analysis.get("artifacts", [])
        batch_results = analysis.get("batchResults", [])
        batch_plan = analysis.get("batchPlan", {})
        batch_success_count = sum(1 for item in batch_results if item.get("status") == "成功")
        batch_fail_count = sum(1 for item in batch_results if item.get("status") == "失败")
        batch_skipped_count = sum(1 for item in batch_results if item.get("status") == "跳过" or item.get("skipped"))
        validation_failure_file_ids = {
            str(file.get("id") or "")
            for file in files
            if file.get("validation_errors")
        }
        artifact_failure_file_ids = {
            str(item.get("file_id") or "")
            for item in artifacts
            if item.get("status") not in {"成功", "跳过", *PENDING_ARTIFACT_STATUSES, *CANCELLED_ARTIFACT_STATUSES}
        }
        artifact_pending_file_ids = {
            str(item.get("file_id") or "")
            for item in artifacts
            if item.get("status") in PENDING_ARTIFACT_STATUSES
        }
        artifact_cancelled_file_ids = {
            str(item.get("file_id") or "")
            for item in artifacts
            if item.get("status") in CANCELLED_ARTIFACT_STATUSES
        }
        image_failure_file_ids = {
            str(item.get("file_id") or "")
            for item in image_errors
            if item.get("file_id")
        }
        failed_file_ids = {
            file_id
            for file_id in validation_failure_file_ids | artifact_failure_file_ids | image_failure_file_ids
            if file_id
        }
        fail_count = len(failed_file_ids)
        pending_file_ids = {file_id for file_id in artifact_pending_file_ids - failed_file_ids if file_id}
        pending_count = len(pending_file_ids)
        cancelled_file_ids = {
            file_id
            for file_id in artifact_cancelled_file_ids - failed_file_ids - pending_file_ids
            if file_id
        }
        cancelled_count = len(cancelled_file_ids)
        success_count = max(0, len(files) - fail_count - pending_count - cancelled_count)
        if task.get("task_type") == "batch_process" and batch_results:
            success_count = batch_success_count
            fail_count = batch_fail_count
        report_type = self._report_type(task["task_type"])
        report = ReportItem(
            task_id=task["id"],
            file_id=files[0]["id"] if len(files) == 1 else "batch",
            report_type=report_type,
            report_path="",
            success_count=success_count,
            fail_count=fail_count,
            formula_count=len(formulas),
            omml_count=sum(1 for item in formulas if item.get("source_type") == "OMML"),
            omml_converted_count=sum(1 for item in formulas if item.get("source_type") == "OMML" and item.get("confidence", 0) >= 80),
            macro_count=len(macros),
            macro_success_count=sum(1 for item in macros if item.get("execute_status") == "成功"),
            macro_fail_count=sum(1 for item in macros if item.get("execute_status") == "失败"),
            formatted_formula_count=sum(1 for item in formulas if item.get("format_status") == "已格式化"),
            small_image_count=len(small_images),
            error_count=fail_count,
        )
        payload = report.to_dict()
        payload["files"] = files
        payload["analysis"] = analysis
        self._normalize_formula_positions(payload)
        payload["mathpix_job_count"] = len(mathpix_jobs)
        payload["omml_dependency_count"] = len(omml_dependencies)
        payload["omml_dependency_copied_count"] = sum(1 for item in omml_dependencies if item.get("copy_status") == "成功")
        payload["omml_conversion_prompt_count"] = len(omml_prompts)
        payload["omml_conversion_waiting_count"] = sum(1 for item in omml_prompts if item.get("status") == "需确认")
        payload["file_omml_missing_count"] = sum(1 for file in files if file.get("missing_omml_dependency"))
        payload["macro_queued_count"] = sum(1 for item in macros if item.get("execute_status") in {"待本地客户端执行", "待确认"})
        payload["artifact_count"] = sum(1 for item in artifacts if item.get("status") == "成功")
        payload["artifact_failure_count"] = sum(1 for item in artifacts if item.get("status") not in {"成功", "跳过", *PENDING_ARTIFACT_STATUSES, *CANCELLED_ARTIFACT_STATUSES})
        payload["artifact_pending_count"] = sum(1 for item in artifacts if item.get("status") in PENDING_ARTIFACT_STATUSES)
        payload["artifact_cancelled_count"] = sum(1 for item in artifacts if item.get("status") in CANCELLED_ARTIFACT_STATUSES)
        payload["failed_file_ids"] = sorted(failed_file_ids)
        payload["pending_file_ids"] = sorted(pending_file_ids)
        payload["pending_count"] = pending_count
        payload["cancelled_file_ids"] = sorted(cancelled_file_ids)
        payload["cancelled_count"] = cancelled_count
        payload["batch_total_count"] = len(batch_results)
        payload["batch_success_count"] = batch_success_count
        payload["batch_fail_count"] = batch_fail_count
        payload["batch_skipped_count"] = batch_skipped_count
        payload["batch_max_concurrent"] = int(batch_plan.get("max_concurrent") or 0)
        payload["batch_failure_strategy"] = str(batch_plan.get("failure_strategy") or "")
        payload["batch_requires_user_decision"] = bool(batch_plan.get("requires_user_decision", False))
        payload["batch_retryable_count"] = int(batch_plan.get("retryable_count") or 0)
        payload["batch_continue_on_failure"] = bool(batch_plan.get("continue_on_failure", True))
        payload["image_extraction_error_count"] = len(image_errors)
        payload["formula_position_issue_count"] = sum(1 for item in payload["analysis"].get("formulas", []) if item.get("position_status") == "异常位置")
        payload["qualityChecks"] = self._quality_checks(task, files, analysis, payload)
        payload["quality_issue_count"] = sum(1 for item in payload["qualityChecks"] if item.get("status") != "通过")
        payload["quality_blocker_count"] = sum(1 for item in payload["qualityChecks"] if item.get("status") == "失败")
        payload["failureRows"] = self._failure_rows(payload)
        payload["failure_count"] = len(payload["failureRows"])
        payload["qualitySummary"] = self._quality_summary(files, analysis, payload)
        payload["summary"] = self._summary(payload)
        json_path = self.store.reports_dir / f"{payload['id']}.json"
        html_path = self.store.reports_dir / f"{payload['id']}.html"
        pdf_path = self.store.reports_dir / f"{payload['id']}.pdf"
        xlsx_path = self.store.reports_dir / f"{payload['id']}.xlsx"
        txt_path = self.store.reports_dir / f"{payload['id']}.txt"
        failure_csv_path = self.store.reports_dir / f"{payload['id']}-failures.csv"
        payload["json_path"] = str(json_path)
        payload["html_path"] = str(html_path)
        payload["pdf_path"] = str(pdf_path)
        payload["xlsx_path"] = str(xlsx_path)
        payload["txt_path"] = str(txt_path)
        payload["failure_csv_path"] = str(failure_csv_path)
        payload["report_path"] = str(html_path)
        return self.rewrite(payload)

    def rewrite(self, payload: dict[str, Any]) -> dict[str, Any]:
        """将已有报告载荷重写为磁盘文件。"""
        payload.pop("report_file_integrity", None)
        json_path = Path(payload.get("json_path") or self.store.reports_dir / f"{payload['id']}.json")
        html_path = Path(payload.get("html_path") or self.store.reports_dir / f"{payload['id']}.html")
        pdf_path = Path(payload.get("pdf_path") or self.store.reports_dir / f"{payload['id']}.pdf")
        xlsx_path = Path(payload.get("xlsx_path") or self.store.reports_dir / f"{payload['id']}.xlsx")
        txt_path = Path(payload.get("txt_path") or self.store.reports_dir / f"{payload['id']}.txt")
        failure_csv_path = Path(payload.get("failure_csv_path") or self.store.reports_dir / f"{payload['id']}-failures.csv")
        payload["json_path"] = str(json_path)
        payload["html_path"] = str(html_path)
        payload["pdf_path"] = str(pdf_path)
        payload["xlsx_path"] = str(xlsx_path)
        payload["txt_path"] = str(txt_path)
        payload["failure_csv_path"] = str(failure_csv_path)
        payload["report_path"] = str(html_path)
        html_path.write_text(self._html(payload), encoding="utf-8")
        txt_path.write_text(self._text(payload), encoding="utf-8")
        failure_csv_path.write_text(self._failure_csv(payload.get("failureRows", [])), encoding="utf-8")
        build_text_pdf(self._pdf_lines(payload), pdf_path, payload["report_type"])
        self._write_xlsx(payload, xlsx_path)
        integrity = {
            key: self._file_integrity(path)
            for key, path in (
                ("html_path", html_path),
                ("pdf_path", pdf_path),
                ("xlsx_path", xlsx_path),
                ("txt_path", txt_path),
                ("failure_csv_path", failure_csv_path),
            )
        }
        payload["report_file_integrity"] = integrity
        json_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        integrity["json_path"] = self._file_integrity(json_path)
        return payload

    @staticmethod
    def _file_integrity(path: Path) -> dict[str, Any]:
        """返回服务端测量的报告导出完整性信息。"""
        data = path.read_bytes()
        return {"size": len(data), "sha256": hashlib.sha256(data).hexdigest()}

    @staticmethod
    def _report_type(task_type: str) -> str:
        """将任务类型映射为用户可见报告标题。"""
        mapping = {
            "formula_precheck": "公式预检报告",
            "omml_to_mathtype": "OMML 转换报告",
            "mathtype_format": "MathType 格式化报告",
            "macro_sequence": "宏执行报告",
            "small_image_scan": "微小图片报告",
            "batch_process": "批量处理报告",
        }
        return mapping.get(task_type, "处理报告")

    @staticmethod
    def _normalize_formula_positions(payload: dict[str, Any]) -> None:
        """确保每个公式具有可报告位置或明确兜底。"""
        file_names = {file.get("id"): file.get("file_name", "") for file in payload.get("files", [])}
        for index, item in enumerate(payload.get("analysis", {}).get("formulas", []), start=1):
            raw_position = str(item.get("position") or "").strip()
            if raw_position:
                item.setdefault("position_status", "已记录")
                item.setdefault("position_issue", "")
                item.setdefault("fallback_position", raw_position)
                continue
            fallback = f"异常位置 / {file_names.get(item.get('file_id'), item.get('file_id', '未知文件'))} / 公式 {index}"
            item["position"] = fallback
            item["position_status"] = "异常位置"
            item["position_issue"] = "公式位置丢失，已记录异常位置"
            item["fallback_position"] = fallback

    @staticmethod
    def _summary(payload: dict[str, Any]) -> list[str]:
        """根据载荷计数生成顶层报告摘要。"""
        return [
            f"成功文件 {payload['success_count']} 个，待本地处理 {payload.get('pending_count', 0)} 个，已取消 {payload.get('cancelled_count', 0)} 个，失败文件 {payload['fail_count']} 个。",
            f"公式 {payload['formula_count']} 个，OMML {payload['omml_count']} 个，已格式化 {payload['formatted_formula_count']} 个。",
            f"宏 {payload['macro_count']} 个，成功执行 {payload['macro_success_count']} 个，待本地处理 {payload.get('macro_queued_count', 0)} 个。",
            f"微小图片 {payload['small_image_count']} 个，错误 {payload['error_count']} 个。",
            f"Mathpix PDF 作业 {payload.get('mathpix_job_count', 0)} 个。",
            f"OMML 依赖 {payload.get('omml_dependency_count', 0)} 个，复制成功 {payload.get('omml_dependency_copied_count', 0)} 个，转换确认 {payload.get('omml_conversion_prompt_count', 0)} 个，待确认 {payload.get('omml_conversion_waiting_count', 0)} 个，文件缺失 {payload.get('file_omml_missing_count', 0)} 个。",
            f"转换输出 {payload.get('artifact_count', 0)} 个。",
            f"批量结果 {payload.get('batch_total_count', 0)} 个，失败 {payload.get('batch_fail_count', 0)} 个，跳过 {payload.get('batch_skipped_count', 0)} 个，最大并发 {payload.get('batch_max_concurrent', 0)}，失败策略 {payload.get('batch_failure_strategy') or '默认'}，可重试 {payload.get('batch_retryable_count', 0)} 个。",
            f"质量检查 {len(payload.get('qualityChecks', []))} 项，需处理 {payload.get('quality_issue_count', 0)} 项。",
            f"失败清单 {payload.get('failure_count', 0)} 条。",
        ]

    @staticmethod
    def _failure_rows(payload: dict[str, Any]) -> list[dict[str, Any]]:
        """收集文件、转换、Mathpix、OMML、宏和预检失败。"""
        files = payload.get("files", [])
        file_names = {file.get("id"): file.get("file_name", "") for file in files}
        skipped_file_ids = {
            item.get("file_id")
            for item in payload.get("analysis", {}).get("batchResults", [])
            if item.get("status") == "跳过" or item.get("skipped")
        }
        rows: list[dict[str, Any]] = []
        for file in files:
            if file.get("id") in skipped_file_ids:
                continue
            for error in file.get("validation_errors") or []:
                rows.append(
                    ReportBuilder._failure_row(
                        "文件校验",
                        file.get("id", ""),
                        file.get("file_name", ""),
                        "失败",
                        error,
                        "处理文件损坏、加密、格式或大小问题后重试",
                    )
                )
        for item in payload.get("analysis", {}).get("imageExtractionErrors", []):
            rows.append(
                ReportBuilder._failure_row(
                    "图片提取",
                    item.get("file_id", ""),
                    item.get("file_name", ""),
                    item.get("status", "失败"),
                    item.get("message") or "图片提取失败",
                    item.get("recommendation") or "修复源文档后重新执行图片检索",
                    item.get("source_name") or item.get("id", ""),
                )
            )
        for item in payload.get("analysis", {}).get("formulas", []):
            if item.get("position_status") == "异常位置":
                rows.append(
                    ReportBuilder._failure_row(
                        "公式位置",
                        item.get("file_id", ""),
                        file_names.get(item.get("file_id"), item.get("file_id", "")),
                        item.get("position_status", ""),
                        item.get("position_issue") or "公式位置丢失",
                        "在排版校正中补充位置，或重新识别公式后重建报告",
                        item.get("id", ""),
                    )
                )
        for item in payload.get("analysis", {}).get("batchResults", []):
            if item.get("status") == "失败":
                rows.append(
                    ReportBuilder._failure_row(
                        "批量处理",
                        item.get("file_id", ""),
                        item.get("file_name", ""),
                        item.get("status", ""),
                        item.get("error_message") or "批量处理失败",
                        item.get("recommendation", ""),
                        item.get("sequence", ""),
                    )
                )
        omml_failure_file_ids: set[Any] = set()
        for item in payload.get("analysis", {}).get("ommlDependencies", []):
            if item.get("found_status") == "未找到" or item.get("copy_status") == "失败":
                omml_failure_file_ids.add(item.get("file_id"))
                rows.append(
                    ReportBuilder._failure_row(
                        "OMML 依赖",
                        item.get("file_id", ""),
                        file_names.get(item.get("file_id"), item.get("file_id", "")),
                        item.get("copy_status") or item.get("found_status", ""),
                        item.get("error_message") or f"检索：{item.get('found_status', '')}，复制：{item.get('copy_status', '')}",
                        "配置 OMML 检索路径或手动选择依赖文件",
                        item.get("id", ""),
                    )
                )
        for file in files:
            if file.get("missing_omml_dependency") and file.get("id") not in omml_failure_file_ids:
                rows.append(
                    ReportBuilder._failure_row(
                        "OMML 依赖",
                        file.get("id", ""),
                        file.get("file_name", ""),
                        "缺少",
                        "文件含 OMML 公式，但当前目录未检测到 OMML 依赖文件",
                        "配置 OMML 检索路径、手动选择依赖文件，或重新运行 OMML 预检任务",
                    )
                )
        for item in payload.get("analysis", {}).get("macros", []):
            if item.get("execute_status") == "失败":
                rows.append(
                    ReportBuilder._failure_row(
                        "宏执行",
                        item.get("file_id", ""),
                        file_names.get(item.get("file_id"), item.get("file_id", "")),
                        item.get("execute_status", ""),
                        item.get("error_message") or "宏执行失败",
                        "检查宏权限、失败策略和备份后重试",
                        item.get("id", ""),
                    )
                )
        for item in payload.get("analysis", {}).get("artifacts", []):
            if item.get("status") not in {"成功", "跳过", *PENDING_ARTIFACT_STATUSES, *CANCELLED_ARTIFACT_STATUSES}:
                rows.append(
                    ReportBuilder._failure_row(
                        "转换输出",
                        item.get("file_id", ""),
                        item.get("source_file", ""),
                        item.get("status", ""),
                        item.get("message") or "转换输出未生成",
                        "查看任务日志并重试失败文件",
                    )
                )
        for item in payload.get("analysis", {}).get("mathpix", []):
            if item.get("status") in {"authorization_required", "missing_credentials", "api_error", "missing_local_file", "timeout", "error", "ocr_disabled"}:
                rows.append(
                    ReportBuilder._failure_row(
                        "Mathpix PDF",
                        item.get("file_id", ""),
                        item.get("file_name", ""),
                        item.get("status", ""),
                        item.get("message") or "Mathpix PDF 处理未完成",
                        "开启 OCR、授权外部上传并配置 Mathpix 凭证后重试",
                        item.get("pdf_id", ""),
                    )
                )
        for item in payload.get("analysis", {}).get("preflightChecks", []):
            if item.get("status") == "失败":
                rows.append(
                    ReportBuilder._failure_row(
                        "系统预检",
                        "",
                        "",
                        item.get("status", ""),
                        item.get("message") or item.get("label") or "系统预检失败",
                        item.get("recommendation", ""),
                        item.get("id", ""),
                    )
                )
        return rows

    def _quality_summary(self, files: list[dict[str, Any]], analysis: dict[str, Any], payload: dict[str, Any]) -> list[dict[str, str]]:
        """汇总报告概览的需求质量指标。"""
        formulas = analysis.get("formulas", [])
        macros = analysis.get("macros", [])
        small_images = analysis.get("smallImages", [])
        image_errors = analysis.get("imageExtractionErrors", [])
        omml_prompts = analysis.get("ommlConversionPrompts", [])
        omml_dependencies = analysis.get("ommlDependencies", [])
        artifacts = analysis.get("artifacts", [])
        source_objects = ReportBuilder._source_object_counts(files)
        omml_formula_count = max(
            ReportBuilder._safe_int(payload.get("omml_count")),
            sum(ReportBuilder._safe_int(item.get("omml_count")) for item in omml_prompts),
            sum(ReportBuilder._safe_int((file.get("content_summary") or {}).get("ommlFormulas")) for file in files),
        )
        mathtype_formula_count = max(
            sum(1 for item in formulas if item.get("source_type") == "MathType"),
            sum(1 for file in files if file.get("has_mathtype")),
        )
        omml_selected_count = sum(1 for item in omml_prompts if item.get("status") == "已选择转换")
        omml_waiting_count = sum(1 for item in omml_prompts if item.get("status") == "需确认")
        omml_failure_count = sum(1 for item in omml_dependencies if item.get("found_status") == "未找到" or item.get("copy_status") == "失败")
        low_confidence_count = sum(1 for item in formulas if item.get("status") == "待确认" or ReportBuilder._safe_int(item.get("confidence")) < 80)
        formula_position_issues = sum(1 for item in formulas if item.get("position_status") == "异常位置")
        possible_loss = self._possible_loss_notes(payload, analysis)
        return [
            self._quality_summary_row(
                "file_total",
                "处理文件总数",
                str(len(files)),
                "失败" if payload.get("fail_count") else ("需确认" if payload.get("pending_count") or payload.get("cancelled_count") else "通过"),
                f"成功 {payload.get('success_count', 0)} 个，待本地处理 {payload.get('pending_count', 0)} 个，已取消 {payload.get('cancelled_count', 0)} 个，失败 {payload.get('fail_count', 0)} 个",
                "启动本地客户端处理等待文件" if payload.get("pending_count") and not payload.get("fail_count") else ("如需继续，请重试已取消文件" if payload.get("cancelled_count") and not payload.get("fail_count") else "处理失败文件后重试"),
            ),
            self._quality_summary_row("omml_formulas", "OMML 公式数量", str(omml_formula_count), "通过" if omml_formula_count == 0 else "需确认", "来自 Word 原生公式检测和 OMML 转换确认记录", "确认是否转为 MathType 或保留原公式"),
            self._quality_summary_row("omml_conversion", "OMML 转 MathType", f"已确认 {omml_selected_count} 个，待确认 {omml_waiting_count} 个，异常 {omml_failure_count} 个", "失败" if omml_failure_count else ("需确认" if omml_waiting_count else "通过"), "当前记录转换决策和依赖状态，真实转换交给本地客户端", "修复依赖并完成用户确认"),
            self._quality_summary_row("mathtype_formulas", "MathType 公式数量", str(mathtype_formula_count), "通过", "来自公式报告和文件对象检测", "跨平台交付时保留 MathML/LaTeX 或图片兜底"),
            self._quality_summary_row("formatted_formulas", "已格式化公式数量", str(payload.get("formatted_formula_count", 0)), "通过", "按当前公式格式化设置记录", "低置信度或失败项进入人工校正"),
            self._quality_summary_row("macro_execution", "宏执行数量", f"总数 {payload.get('macro_count', 0)}，成功 {payload.get('macro_success_count', 0)}，失败 {payload.get('macro_fail_count', 0)}", "失败" if payload.get("macro_fail_count") else ("需确认" if payload.get("macro_queued_count") else "通过"), "网页端只记录宏队列和本地客户端执行状态", "检查宏授权、白名单和失败策略"),
            self._quality_summary_row("low_confidence", "低置信度公式数量", str(low_confidence_count), "需确认" if low_confidence_count else "通过", "按公式状态和 80% 默认阈值统计", "在公式页确认、修正或重识别"),
            self._quality_summary_row("formula_positions", "公式位置异常", str(formula_position_issues), "失败" if formula_position_issues else "通过", "公式位置丢失会写入异常位置" if formula_position_issues else "公式位置已记录", "在排版校正中补充位置或重新识别公式" if formula_position_issues else "继续处理"),
            self._quality_summary_row(
                "image_extraction",
                "图片提取数量",
                f"成功 {len(small_images)} 个，失败 {len(image_errors)} 个",
                "失败" if image_errors else ("需确认" if small_images else "通过"),
                "存在图片条目读取或格式识别失败" if image_errors else "微小图片、图标、二维码、印章和签名进入图片报告",
                "修复源文档或重新上传后重试图片检索" if image_errors else "查看图片页并确认类型、处理误判、删除或替换",
            ),
            self._quality_summary_row("small_images", "微小图片数量", str(payload.get("small_image_count", 0)), "需确认" if payload.get("small_image_count") else "通过", "按尺寸、面积、透明、水印和重复设置筛选", "按图片清单导出或标记"),
            self._quality_summary_row("table_extraction", "表格提取数量", str(source_objects["tables"]), "通过" if source_objects["tables"] == 0 else "需确认", "来自 Word/PPT/Excel/PDF 表格结构或线索统计", "转换后在预览对比中复核表格结构"),
            self._quality_summary_row("possible_loss", "可能丢失对象", str(len(possible_loss)), "需确认" if possible_loss else "通过", "；".join(possible_loss) if possible_loss else "未发现明显丢失风险", "调整保留设置、修复依赖或使用本地 Office 引擎复核"),
            self._quality_summary_row("error_details", "错误详情", str(payload.get("failure_count", 0)), "失败" if payload.get("failure_count") else "通过", "失败清单汇总文件校验、OMML、宏、转换输出和 Mathpix 问题", "下载失败清单 CSV 并逐项处理"),
        ]

    @staticmethod
    def _quality_summary_row(row_id: str, label: str, metric: str, status: str, detail: str, recommendation: str) -> dict[str, str]:
        """生成规范质量摘要行。"""
        return {
            "id": row_id,
            "label": label,
            "metric": metric,
            "status": status,
            "detail": detail,
            "recommendation": recommendation,
        }

    def _possible_loss_notes(self, payload: dict[str, Any], analysis: dict[str, Any]) -> list[str]:
        """从质量检查和产物提取对象保留风险。"""
        notes: list[str] = []
        for item in payload.get("qualityChecks", []):
            if item.get("id") in {"content_preservation", "preview_consistency", "conversion_output"} and item.get("status") != "通过":
                notes.append(str(item.get("message") or item.get("metric") or item.get("label")))
        for item in analysis.get("artifacts", []):
            plan = item.get("object_preservation") or {}
            statuses = plan.get("statuses") or {}
            for label, status in statuses.items():
                status_text = str(status or "")
                if status_text and "关闭" in status_text:
                    notes.append(f"{item.get('source_file', '')} {label}：{status_text}")
        return list(dict.fromkeys(note for note in notes if note))

    @staticmethod
    def _failure_row(category: str, file_id: object, file_name: object, status: object, message: object, recommendation: object, source_id: object = "") -> dict[str, Any]:
        """生成 CSV 和接口使用的规范失败清单行。"""
        return {
            "category": category,
            "file_id": file_id,
            "file_name": file_name,
            "status": status,
            "message": message,
            "recommendation": recommendation,
            "source_id": source_id,
        }

    @staticmethod
    def _failure_csv(rows: list[dict[str, Any]]) -> str:
        """将失败行渲染为 UTF-8 CSV 字符串。"""
        buffer = StringIO()
        writer = csv.DictWriter(buffer, fieldnames=["category", "file_id", "file_name", "status", "message", "recommendation", "source_id"])
        writer.writeheader()
        for row in rows:
            writer.writerow(row)
        return buffer.getvalue()

    def _quality_checks(self, task: dict[str, Any], files: list[dict[str, Any]], analysis: dict[str, Any], payload: dict[str, Any]) -> list[dict[str, Any]]:
        """生成校验、保留、Mathpix 和交接的详细质量检查。"""
        formulas = analysis.get("formulas", [])
        macros = analysis.get("macros", [])
        small_images = analysis.get("smallImages", [])
        image_errors = analysis.get("imageExtractionErrors", [])
        mathpix_jobs = analysis.get("mathpix", [])
        omml_dependencies = analysis.get("ommlDependencies", [])
        omml_prompts = analysis.get("ommlConversionPrompts", [])
        artifacts = analysis.get("artifacts", [])
        batch_results = analysis.get("batchResults", [])
        preflight_checks = analysis.get("preflightChecks", [])
        workflow_plan = analysis.get("workflowPlan") if isinstance(analysis.get("workflowPlan"), dict) else {}
        settings = self.store.get_settings()
        try:
            formula_threshold = max(1, int(settings.get("formulaConfidenceThreshold", 80) or 80))
        except (TypeError, ValueError):
            formula_threshold = 80
        skipped_batch_file_ids = {
            item.get("file_id")
            for item in batch_results
            if item.get("status") == "跳过" or item.get("skipped")
        }
        validation_failures = [file for file in files if file.get("validation_errors") and file.get("id") not in skipped_batch_file_ids]
        low_confidence = [item for item in formulas if item.get("status") == "待确认" or int(item.get("confidence", 0)) < formula_threshold]
        formula_position_issues = [item for item in formulas if item.get("position_status") == "异常位置"]
        omml_failures = [item for item in omml_dependencies if item.get("found_status") == "未找到" or item.get("copy_status") == "失败"]
        omml_waiting_prompts = [item for item in omml_prompts if item.get("status") == "需确认"]
        omml_missing_files = [file for file in files if file.get("missing_omml_dependency")]
        omml_issue_file_ids = {item.get("file_id") for item in omml_failures}
        omml_issue_file_ids.update(file.get("id") for file in omml_missing_files)
        omml_file_total = sum(1 for file in files if file.get("file_type") == "Word" and file.get("has_omml"))
        macro_failures = [item for item in macros if item.get("execute_status") == "失败"]
        blocked_macros = [item for item in macros if item.get("execute_status") in {"已禁用", "未授权"}]
        queued_macros = [item for item in macros if item.get("execute_status") in {"待本地客户端执行", "待确认"}]
        artifact_errors = [item for item in artifacts if item.get("status") not in {"成功", "跳过", *PENDING_ARTIFACT_STATUSES, *CANCELLED_ARTIFACT_STATUSES}]
        pending_artifacts = [item for item in artifacts if item.get("status") in PENDING_ARTIFACT_STATUSES]
        cancelled_artifacts = [item for item in artifacts if item.get("status") in CANCELLED_ARTIFACT_STATUSES]
        skipped_artifacts = [item for item in artifacts if item.get("status") == "跳过"]
        mathpix_waiting = [item for item in mathpix_jobs if item.get("status") in {"authorization_required", "missing_credentials", "api_error", "missing_local_file", "ocr_disabled"}]
        batch_failures = [item for item in batch_results if item.get("status") == "失败"]
        batch_skipped = [item for item in batch_results if item.get("status") == "跳过" or item.get("skipped")]
        pdf_files = [file for file in files if file.get("file_type") == "PDF"]
        pdf_ocr_files = [file for file in pdf_files if (file.get("content_summary") or {}).get("pdfType") in {"扫描型 PDF", "混合型 PDF"}]
        pdf_formula_files = [file for file in pdf_files if file.get("has_formula") or int((file.get("content_summary") or {}).get("formulaHints", 0) or 0) > 0]
        pdf_table_files = [file for file in pdf_files if int((file.get("content_summary") or {}).get("tableHints", 0) or 0) > 0]
        structural_counts = sum(
            int(file.get("page_count") or 0) + int(file.get("slide_count") or 0) + int(file.get("sheet_count") or 0)
            for file in files
        )
        source_objects = ReportBuilder._source_object_counts(files)
        preservation_risks = ReportBuilder._content_preservation_risks(task, settings, source_objects)
        conversion_tasks = {"word_to_ppt", "ppt_to_word", "pdf_to_word", "excel_to_pdf", "excel_to_word", "excel_to_ppt"}
        expected_artifacts = len(files) if task.get("task_type") in conversion_tasks else 0
        successful_artifacts = sum(1 for item in artifacts if item.get("status") == "成功")
        accounted_artifacts = successful_artifacts + len(skipped_artifacts) + len(pending_artifacts) + len(cancelled_artifacts)
        checks = [
            *preflight_checks,
            ReportBuilder._quality_item(
                "workflow_order",
                "流程顺序",
                "需确认" if workflow_plan.get("invalid_items") or workflow_plan.get("duplicate_items") else "通过",
                ReportBuilder._workflow_plan_metric(workflow_plan),
                ReportBuilder._workflow_plan_message(workflow_plan),
                "检查处理配置中的流程顺序，移除未知或重复任务" if workflow_plan.get("invalid_items") or workflow_plan.get("duplicate_items") else "按当前拖拽顺序继续",
            ),
            ReportBuilder._quality_item(
                "file_validation",
                "文件校验",
                "失败" if validation_failures else "通过",
                f"{len(validation_failures)} / {len(files)} 个文件异常",
                "存在文件损坏、加密、过大或格式异常" if validation_failures else "所有输入文件通过基础校验",
                "处理校验失败文件后重试任务" if validation_failures else "可进入后续处理",
            ),
            ReportBuilder._quality_item(
                "structure_counts",
                "结构数量",
                "需确认" if files and structural_counts == 0 else "通过",
                f"页/幻灯片/工作表合计 {structural_counts}",
                "未读取到页数、幻灯片或工作表数量" if files and structural_counts == 0 else "已记录基础结构数量",
                "打开预览检查文件是否损坏或需要真实 Office 引擎" if files and structural_counts == 0 else "继续转换或导出报告",
            ),
            ReportBuilder._quality_item(
                "content_preservation",
                "内容保留",
                "需确认" if preservation_risks else "通过",
                (
                    f"图片 {source_objects['images']} 个，表格 {source_objects['tables']} 个，公式 {source_objects['formulas']} 个，"
                    f"标题/结构 {source_objects['structures']} 个"
                ),
                "存在可能不保留的对象设置：" + "；".join(preservation_risks) if preservation_risks else "源文档对象数量已记录，当前设置未发现明显保留风险",
                "调整保留设置后重试，或查看转换前后对比" if preservation_risks else "继续转换并在结果预览中复核关键对象",
            ),
            self._preview_consistency_quality(files, artifacts),
            ReportBuilder._quality_item(
                "formula_confidence",
                "公式置信度",
                "需确认" if low_confidence else "通过",
                f"低置信度 {len(low_confidence)} / {len(formulas)} 个，阈值 {formula_threshold}",
                "存在待人工确认的公式" if low_confidence else "公式置信度通过当前阈值",
                "在公式页确认、重识别或修正 LaTeX" if low_confidence else "保留当前公式结果",
            ),
            ReportBuilder._quality_item(
                "formula_positions",
                "公式位置",
                "失败" if formula_position_issues else "通过",
                f"异常 {len(formula_position_issues)} / {len(formulas)} 个",
                "存在公式位置丢失，已记录异常位置" if formula_position_issues else "公式位置已记录",
                "在排版校正中补充位置，或重新识别公式后重建报告" if formula_position_issues else "继续处理",
            ),
            ReportBuilder._quality_item(
                "omml_dependencies",
                "OMML 依赖",
                "失败" if omml_issue_file_ids else "通过",
                f"异常 {len(omml_issue_file_ids)} / {omml_file_total or len(omml_dependencies)} 个文件",
                "OMML 依赖缺失、未找到或复制失败" if omml_issue_file_ids else "OMML 依赖检查通过",
                "配置 OMML 检索路径或手动选择依赖文件" if omml_issue_file_ids else "可继续 OMML 转 MathType 流程",
            ),
            ReportBuilder._quality_item(
                "omml_conversion_prompt",
                "OMML 转换确认",
                "需确认" if omml_waiting_prompts else "通过",
                f"待确认 {len(omml_waiting_prompts)} / {len(omml_prompts)} 个",
                "检测到 Word 自带公式，需要用户确认是否转换为 MathType" if omml_waiting_prompts else "OMML 转 MathType 决策已记录或当前任务不涉及 OMML",
                "在 Word 转 PPT 参数中选择先转 MathType，或在任务选项中确认保留原公式" if omml_waiting_prompts else "按当前决策继续",
            ),
            ReportBuilder._quality_item(
                "macro_results",
                "宏执行结果",
                "失败" if macro_failures else ("需确认" if queued_macros or blocked_macros else "通过"),
                f"失败 {len(macro_failures)} 个，待本地处理 {len(queued_macros)} 个，安全阻止 {len(blocked_macros)} 个",
                "存在失败宏" if macro_failures else ("宏队列等待本地客户端执行、风险确认或安全策略确认" if queued_macros or blocked_macros else "宏执行状态正常"),
                "检查宏权限、失败策略和备份后重试" if macro_failures else ("检查宏风险确认、白名单或禁用设置" if queued_macros or blocked_macros else "保留当前宏结果"),
            ),
            ReportBuilder._quality_item(
                "small_images",
                "微小图片",
                "失败" if image_errors else ("需确认" if small_images else "通过"),
                f"发现 {len(small_images)} 个，提取失败 {len(image_errors)} 个",
                "存在图片条目读取或格式识别失败" if image_errors else ("发现微小图片，需要确认是否误判或导出" if small_images else "未发现需处理的微小图片"),
                "下载失败清单，修复源文档或重新上传后重试图片检索" if image_errors else ("在微小图片页查看、确认类型、导出或标记误判" if small_images else "无需处理"),
            ),
            ReportBuilder._quality_item(
                "conversion_output",
                "转换输出",
                "失败" if artifact_errors or accounted_artifacts < expected_artifacts else ("需确认" if skipped_artifacts or pending_artifacts or cancelled_artifacts else "通过"),
                f"成功 {successful_artifacts} 个，待本地执行 {len(pending_artifacts)} 个，已取消 {len(cancelled_artifacts)} 个，跳过 {len(skipped_artifacts)} 个，预期 {expected_artifacts or accounted_artifacts} 个",
                "部分转换输出未生成" if artifact_errors or accounted_artifacts < expected_artifacts else ("旧版 Office 输入等待本地客户端生成输出" if pending_artifacts else ("本地 Office 转换已取消" if cancelled_artifacts else ("存在按同名策略跳过的输出" if skipped_artifacts else "转换输出已生成"))),
                "查看任务日志并重试失败文件" if artifact_errors or accounted_artifacts < expected_artifacts else ("启动本地 Office 客户端执行并回传结果" if pending_artifacts else ("如需继续转换，请重试任务" if cancelled_artifacts else ("如需生成被跳过文件，请调整输出同名策略后重试" if skipped_artifacts else "可下载转换结果"))),
            ),
            ReportBuilder._quality_item(
                "mathpix_pdf",
                "Mathpix PDF",
                "需确认" if mathpix_waiting else "通过",
                f"需处理 {len(mathpix_waiting)} / {len(mathpix_jobs)} 个",
                "PDF 转 Word 等待 Mathpix 授权、凭证或接口结果" if mathpix_waiting else "Mathpix 作业状态正常或当前任务不涉及 PDF OCR",
                "在设置中授权外部上传并配置 Mathpix 凭证" if mathpix_waiting else "无需额外处理",
            ),
            ReportBuilder._quality_item(
                "pdf_type",
                "PDF 类型识别",
                "需确认" if pdf_ocr_files or pdf_formula_files or pdf_table_files else "通过",
                f"OCR {len(pdf_ocr_files)} 个，公式 {len(pdf_formula_files)} 个，表格 {len(pdf_table_files)} 个",
                "PDF 包含扫描页、图片层、公式或表格线索" if pdf_ocr_files or pdf_formula_files or pdf_table_files else "PDF 类型识别未发现额外 OCR 风险",
                "使用 Mathpix 并按需开启文字、公式和表格 OCR" if pdf_ocr_files or pdf_formula_files or pdf_table_files else "可按当前设置转换",
            ),
            ReportBuilder._quality_item(
                "batch_results",
                "批量单文件结果",
                "失败" if batch_failures else "通过",
                f"失败 {len(batch_failures)} / {len(batch_results)} 个，跳过 {len(batch_skipped)} 个",
                "批量任务存在失败文件" if batch_failures else ("失败文件已按用户选择跳过" if batch_skipped else "批量任务单文件状态正常或当前任务不是批量任务"),
                "查看失败清单，处理文件校验问题后重试" if batch_failures else ("如需重新处理跳过文件，可重新上传或重建批量任务" if batch_skipped else "无需额外处理"),
            ),
        ]
        return checks

    @staticmethod
    def _workflow_plan_metric(plan: dict[str, Any]) -> str:
        """格式化规范工作流顺序供质量指标使用。"""
        labels = [str(label) for label in plan.get("labels") or []]
        index = ReportBuilder._safe_int(plan.get("current_index"))
        total = len(labels)
        order_label = " → ".join(labels[:6]) if labels else "未记录"
        suffix = "…" if len(labels) > 6 else ""
        return f"当前 {index or '-'} / {total} · {order_label}{suffix}"

    @staticmethod
    def _workflow_plan_message(plan: dict[str, Any]) -> str:
        """说明工作流顺序已接受、已过滤或缺失。"""
        invalid = [str(item) for item in plan.get("invalid_items") or []]
        duplicates = [str(item) for item in plan.get("duplicate_items") or []]
        if invalid or duplicates:
            parts = []
            if invalid:
                parts.append("已过滤未知任务：" + "、".join(invalid[:4]))
            if duplicates:
                parts.append("已去除重复任务：" + "、".join(duplicates[:4]))
            return "；".join(parts)
        labels = [str(label) for label in plan.get("labels") or []]
        if labels:
            return "流程顺序已随任务参数写入报告：" + " → ".join(labels[:6])
        return "当前任务未记录流程顺序"

    def _preview_consistency_quality(self, files: list[dict[str, Any]], artifacts: list[dict[str, Any]]) -> dict[str, str]:
        """比较源和输出预览以检测较大的转换偏差。"""
        source_files = {file.get("id"): file for file in files}
        comparable: list[dict[str, Any]] = []
        for artifact in artifacts:
            if artifact.get("status") != "成功":
                continue
            source = source_files.get(artifact.get("file_id"))
            output = self._artifact_preview_file(artifact)
            if not source or not output:
                continue
            try:
                diff = self._preview_diff(build_file_preview(source), build_file_preview(output))
            except (OSError, KeyError, ValueError, zipfile.BadZipFile):
                continue
            comparable.append({"artifact": artifact, "diff": diff})
        page_delta_total = sum(abs(item["diff"]["page_delta"]) for item in comparable)
        object_loss_total = sum(max(0, -change["delta"]) for item in comparable for change in item["diff"]["object_changes"])
        warning_total = sum(len(item["diff"]["warnings"]) for item in comparable)
        status = "需确认" if comparable and (page_delta_total or object_loss_total or warning_total) else "通过"
        message = "当前任务暂无可比对转换输出"
        if comparable:
            parts = [self._preview_consistency_note(item["artifact"], item["diff"]) for item in comparable[:4]]
            message = "；".join(part for part in parts if part) or "转换前后预览未发现明显差异"
        return ReportBuilder._quality_item(
            "preview_consistency",
            "转换前后预览一致性",
            status,
            f"比对 {len(comparable)} 个，页段差异 {page_delta_total}，对象减少 {object_loss_total}，告警 {warning_total}",
            message,
            "打开转换前后对比预览，复核页数、表格、图片和公式对象" if status != "通过" else "保留当前转换结果",
        )

    @staticmethod
    def _artifact_preview_file(artifact: dict[str, Any]) -> dict[str, Any] | None:
        """将成功产物转换为兼容预览的文件载荷。"""
        output_type = str(artifact.get("output_type") or "").lower()
        file_type = {"docx": "Word", "docm": "Word", "pptx": "PPT", "xlsx": "Excel", "pdf": "PDF"}.get(output_type)
        path = Path(str(artifact.get("path") or ""))
        if not file_type or not path.exists() or not path.is_file():
            return None
        return {
            "id": artifact.get("artifact_id") or artifact.get("file_id") or path.stem,
            "file_name": artifact.get("file_name") or path.name,
            "file_type": file_type,
            "file_size": path.stat().st_size,
            "storage_path": str(path),
            "file_path": str(path),
            "extension": path.suffix.lower(),
            "status": "成功",
            "content_summary": {
                "sourceFile": artifact.get("source_file", ""),
                "taskId": artifact.get("task_id", ""),
                "outputType": artifact.get("output_type", ""),
            },
        }

    @staticmethod
    def _preview_diff(source_preview: dict[str, Any], output_preview: dict[str, Any]) -> dict[str, Any]:
        """比较源与转换预览，用于报告质量检查。"""
        source_pages = len(source_preview.get("pages") or [])
        output_pages = len(output_preview.get("pages") or [])
        source_objects = {item.get("label") or item.get("type"): ReportBuilder._safe_int(item.get("count")) for item in source_preview.get("objects", [])}
        output_objects = {item.get("label") or item.get("type"): ReportBuilder._safe_int(item.get("count")) for item in output_preview.get("objects", [])}
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

    @staticmethod
    def _preview_consistency_note(artifact: dict[str, Any], diff: dict[str, Any]) -> str:
        """用简短用户可见文本说明预览对比。"""
        losses = [f"{change['label']} {change['source']}→{change['output']}" for change in diff.get("object_changes", []) if change.get("delta", 0) < 0]
        loss_text = "，对象减少 " + "、".join(losses[:3]) if losses else ""
        return (
            f"{artifact.get('source_file') or artifact.get('file_id', '')} → {artifact.get('file_name', '')}："
            f"页段 {diff.get('source_page_count', 0)}→{diff.get('output_page_count', 0)}，"
            f"对象 {diff.get('source_object_count', 0)}→{diff.get('output_object_count', 0)}{loss_text}"
        )

    @staticmethod
    def _source_object_counts(files: list[dict[str, Any]]) -> dict[str, int]:
        """统计源文件中的图片、表格、公式和结构。"""
        counts = {"images": 0, "tables": 0, "formulas": 0, "structures": 0}
        for file in files:
            summary = file.get("content_summary") or {}
            counts["images"] += ReportBuilder._safe_int(summary.get("images")) + ReportBuilder._safe_int(summary.get("imageObjects"))
            counts["tables"] += ReportBuilder._safe_int(summary.get("tables")) + ReportBuilder._safe_int(summary.get("tableHints"))
            formula_count = (
                ReportBuilder._safe_int(summary.get("ommlFormulas"))
                + ReportBuilder._safe_int(summary.get("formulas"))
                + ReportBuilder._safe_int(summary.get("formulaHints"))
            )
            if file.get("has_mathtype"):
                formula_count += 1
            elif file.get("has_formula") and formula_count == 0:
                formula_count = 1
            counts["formulas"] += formula_count
            counts["structures"] += (
                ReportBuilder._safe_int(summary.get("headings"))
                + ReportBuilder._safe_int(summary.get("slides"))
                + ReportBuilder._safe_int(summary.get("sheets"))
                + ReportBuilder._safe_int(summary.get("pages"))
            )
        return counts

    @staticmethod
    def _content_preservation_risks(task: dict[str, Any], settings: dict[str, Any], counts: dict[str, int]) -> list[str]:
        """返回可能丢失源文档对象的转换设置。"""
        task_type = str(task.get("task_type") or "")
        conversion_tasks = {"word_to_ppt", "ppt_to_word", "pdf_to_word", "excel_to_pdf", "excel_to_word", "excel_to_ppt"}
        if task_type not in conversion_tasks:
            return []
        options = task.get("options") or {}
        word_options = options.get("word") or {}
        risks: list[str] = []
        if counts["images"] and not bool(word_options.get("retainImages", settings.get("retainImages", True))):
            risks.append("图片保留关闭")
        if counts["tables"] and not bool(word_options.get("retainTables", settings.get("retainTables", True))):
            risks.append("表格保留关闭")
        if counts["formulas"]:
            if task_type == "word_to_ppt" and not bool(word_options.get("retainFormulas", True)):
                risks.append("Word 公式保留关闭")
            if task_type.startswith("excel_to") and settings.get("excelFormulaMode") == "仅保留计算结果":
                risks.append("Excel 公式仅保留计算结果")
            if task_type == "pdf_to_word" and not bool(settings.get("enableFormulaOcr", True)):
                risks.append("PDF 公式 OCR 关闭")
        return risks

    @staticmethod
    def _safe_int(value: Any) -> int:
        """将宽松计数值转换为非负整数。"""
        try:
            return max(0, int(value or 0))
        except (TypeError, ValueError):
            return 0

    @staticmethod
    def _quality_item(check_id: str, label: str, status: str, metric: str, message: str, recommendation: str) -> dict[str, str]:
        """生成带推导严重程度的详细质量检查项。"""
        severity = {"通过": "info", "需确认": "warning", "失败": "error"}.get(status, "info")
        return {
            "id": check_id,
            "label": label,
            "status": status,
            "severity": severity,
            "metric": metric,
            "message": message,
            "recommendation": recommendation,
        }

    @staticmethod
    def _html(payload: dict[str, Any]) -> str:
        """将报告载荷渲染为独立 HTML 文档。"""
        summary = "".join(f"<li>{html.escape(line)}</li>" for line in payload["summary"])
        file_names = {file.get("id"): file.get("file_name", "") for file in payload.get("files", [])}
        file_rows = "".join(
            "<tr>"
            f"<td>{html.escape(file.get('file_name', ''))}</td>"
            f"<td>{html.escape(file.get('file_type', ''))}</td>"
            f"<td>{html.escape(file.get('status', ''))}</td>"
            f"<td>{html.escape(', '.join(file.get('validation_errors') or ['无']))}</td>"
            "</tr>"
            for file in payload.get("files", [])
        )
        quality_summary_rows = "".join(
            "<tr>"
            f"<td>{html.escape(item.get('label', ''))}</td>"
            f"<td>{html.escape(item.get('status', ''))}</td>"
            f"<td>{html.escape(item.get('metric', ''))}</td>"
            f"<td>{html.escape(item.get('detail', ''))}</td>"
            f"<td>{html.escape(item.get('recommendation', ''))}</td>"
            "</tr>"
            for item in payload.get("qualitySummary", [])
        )
        quality_summary_table = (
            "<h2>质量摘要</h2>"
            "<table><thead><tr><th>项目</th><th>状态</th><th>指标</th><th>说明</th><th>建议</th></tr></thead>"
            f"<tbody>{quality_summary_rows}</tbody></table>"
            if quality_summary_rows
            else ""
        )
        formula_rows = "".join(
            "<tr>"
            f"<td>{html.escape(file_names.get(item.get('file_id'), item.get('file_id', '')))}</td>"
            f"<td>{html.escape(str(item.get('page_index', '')))}</td>"
            f"<td>{html.escape(item.get('position', ''))}</td>"
            f"<td>{html.escape(item.get('position_status', '已记录'))}</td>"
            f"<td>{html.escape(item.get('source_type', ''))}</td>"
            f"<td>{html.escape(item.get('original_image_ref') or item.get('original_image_path') or '-')}</td>"
            f"<td>{html.escape(item.get('latex') or '-')}</td>"
            f"<td>{html.escape(item.get('mathml') or '-')}</td>"
            f"<td>{html.escape(item.get('mathtype_preview') or item.get('mathtype_data') or '-')}</td>"
            f"<td>{html.escape(item.get('format_status', ''))}</td>"
            f"<td>{html.escape(ReportBuilder._formula_style_label(item))}</td>"
            f"<td>{html.escape(ReportBuilder._formula_format_comparison_label(item))}</td>"
            f"<td>{html.escape(ReportBuilder._formula_format_policy_label(item))}</td>"
            f"<td>{html.escape(str(item.get('confidence', '')))}</td>"
            f"<td>{html.escape(item.get('status', ''))}</td>"
            "</tr>"
            for item in payload.get("analysis", {}).get("formulas", [])
        )
        formula_table = (
            "<h2>公式明细</h2>"
            "<table><thead><tr><th>文件</th><th>页码</th><th>位置</th><th>位置状态</th><th>来源</th><th>原始截图</th><th>LaTeX</th><th>MathML</th><th>MathType 预览</th><th>格式化</th><th>样式</th><th>格式化对比</th><th>失败兜底</th><th>置信度</th><th>状态</th></tr></thead>"
            f"<tbody>{formula_rows}</tbody></table>"
            if formula_rows
            else ""
        )
        dependency_rows = "".join(
            "<tr>"
            f"<td>{html.escape(item.get('omml_file_name') or '-')}</td>"
            f"<td>{html.escape(item.get('found_status', ''))}</td>"
            f"<td>{html.escape(item.get('copy_status', ''))}</td>"
            f"<td>{html.escape(item.get('error_message') or '无')}</td>"
            "</tr>"
            for item in payload.get("analysis", {}).get("ommlDependencies", [])
        )
        dependency_table = (
            "<h2>OMML 依赖</h2>"
            "<table><thead><tr><th>文件</th><th>检索</th><th>复制</th><th>错误</th></tr></thead>"
            f"<tbody>{dependency_rows}</tbody></table>"
            if dependency_rows
            else ""
        )
        omml_prompt_rows = "".join(
            "<tr>"
            f"<td>{html.escape(item.get('file_name') or '-')}</td>"
            f"<td>{html.escape(str(item.get('omml_count', '')))}</td>"
            f"<td>{html.escape(item.get('status', ''))}</td>"
            f"<td>{html.escape(str(item.get('convert_to_mathtype')))}</td>"
            f"<td>{html.escape(item.get('decision_source', ''))}</td>"
            f"<td>{html.escape(item.get('dependency_status', ''))}</td>"
            f"<td>{html.escape(item.get('message', ''))}</td>"
            "</tr>"
            for item in payload.get("analysis", {}).get("ommlConversionPrompts", [])
        )
        omml_prompt_table = (
            "<h2>OMML 转换确认</h2>"
            "<table><thead><tr><th>文件</th><th>OMML 数量</th><th>状态</th><th>转换 MathType</th><th>决策来源</th><th>依赖状态</th><th>说明</th></tr></thead>"
            f"<tbody>{omml_prompt_rows}</tbody></table>"
            if omml_prompt_rows
            else ""
        )
        macro_rows = "".join(
            "<tr>"
            f"<td>{html.escape(item.get('macro_name', ''))}</td>"
            f"<td>{html.escape(str(item.get('execute_order', '')))}</td>"
            f"<td>{html.escape(item.get('execute_status', ''))}</td>"
            f"<td>{html.escape(item.get('failure_strategy', ''))}</td>"
            f"<td>{html.escape(ReportBuilder._macro_failure_policy_label(item))}</td>"
            f"<td>{html.escape(item.get('backup_path') or '无')}</td>"
            "</tr>"
            for item in payload.get("analysis", {}).get("macros", [])
        )
        macro_table = (
            "<h2>宏执行队列</h2>"
            "<table><thead><tr><th>宏</th><th>顺序</th><th>状态</th><th>失败策略</th><th>失败动作</th><th>备份</th></tr></thead>"
            f"<tbody>{macro_rows}</tbody></table>"
            if macro_rows
            else ""
        )
        image_rows = "".join(
            "<tr>"
            f"<td>{html.escape(item.get('source_name') or '-')}</td>"
            f"<td>{html.escape(item.get('location', ''))}</td>"
            f"<td>{html.escape(str(item.get('width', '')))} x {html.escape(str(item.get('height', '')))}</td>"
            f"<td>{html.escape(str(item.get('area', '')))}</td>"
            f"<td>{html.escape(item.get('image_type') or '-')}</td>"
            f"<td>{html.escape(ReportBuilder._small_image_label(item))}</td>"
            f"<td>{html.escape(ReportBuilder._small_image_flags_label(item))}</td>"
            f"<td>{html.escape(str(item.get('duplicate_check_status') or '未比较'))}</td>"
            f"<td>{html.escape(str(item.get('export_status') or '可导出'))}</td>"
            f"<td>{html.escape(str(item.get('confidence', '')))}</td>"
            "</tr>"
            for item in payload.get("analysis", {}).get("smallImages", [])
        )
        image_table = (
            "<h2>微小图片</h2>"
            "<table><thead><tr><th>来源</th><th>位置</th><th>尺寸</th><th>面积</th><th>图片格式</th><th>疑似类型</th><th>判定标记</th><th>重复判断</th><th>导出状态</th><th>置信度</th></tr></thead>"
            f"<tbody>{image_rows}</tbody></table>"
            if image_rows
            else ""
        )
        mathpix_rows = "".join(
            "<tr>"
            f"<td>{html.escape(item.get('file_name') or '-')}</td>"
            f"<td>{html.escape(item.get('status', ''))}</td>"
            f"<td>{html.escape(ReportBuilder._mathpix_ocr_label(item))}</td>"
            f"<td>{html.escape(ReportBuilder._mathpix_hint_label(item))}</td>"
            f"<td>{html.escape(ReportBuilder._mathpix_retention_label(item))}</td>"
            f"<td>{html.escape(ReportBuilder._mathpix_note_label(item))}</td>"
            "</tr>"
            for item in payload.get("analysis", {}).get("mathpix", [])
        )
        mathpix_table = (
            "<h2>Mathpix 保留计划</h2>"
            "<table><thead><tr><th>文件</th><th>状态</th><th>OCR</th><th>来源线索</th><th>保留计划</th><th>说明</th></tr></thead>"
            f"<tbody>{mathpix_rows}</tbody></table>"
            if mathpix_rows
            else ""
        )
        artifact_rows = "".join(
            "<tr>"
            f"<td>{html.escape(item.get('source_file') or '-')}</td>"
            f"<td>{html.escape(item.get('file_name') or '-')}</td>"
            f"<td>{html.escape(item.get('output_type') or '-')}</td>"
            f"<td>{html.escape(item.get('status', ''))}</td>"
            f"<td>{html.escape(item.get('message') or '无')}</td>"
            f"<td>{html.escape(ReportBuilder._conversion_settings_label(item.get('conversion_settings') or {}))}</td>"
            f"<td>{html.escape(ReportBuilder._artifact_preservation_label(item))}</td>"
            "</tr>"
            for item in payload.get("analysis", {}).get("artifacts", [])
        )
        artifact_table = (
            "<h2>转换输出</h2>"
            "<table><thead><tr><th>来源</th><th>文件</th><th>类型</th><th>状态</th><th>说明</th><th>设置</th><th>对象保留</th></tr></thead>"
            f"<tbody>{artifact_rows}</tbody></table>"
            if artifact_rows
            else ""
        )
        batch_rows = "".join(
            "<tr>"
            f"<td>{html.escape(str(item.get('sequence', '')))}</td>"
            f"<td>{html.escape(item.get('file_name', ''))}</td>"
            f"<td>{html.escape(item.get('file_type', ''))}</td>"
            f"<td>{html.escape(item.get('status', ''))}</td>"
            f"<td>{html.escape(item.get('error_message') or '无')}</td>"
            f"<td>{html.escape(item.get('recommendation') or '')}</td>"
            "</tr>"
            for item in payload.get("analysis", {}).get("batchResults", [])
        )
        batch_table = (
            "<h2>批量单文件结果</h2>"
            "<table><thead><tr><th>序号</th><th>文件</th><th>类型</th><th>状态</th><th>失败原因</th><th>建议</th></tr></thead>"
            f"<tbody>{batch_rows}</tbody></table>"
            if batch_rows
            else ""
        )
        quality_rows = "".join(
            "<tr>"
            f"<td>{html.escape(item.get('label', ''))}</td>"
            f"<td>{html.escape(item.get('status', ''))}</td>"
            f"<td>{html.escape(item.get('metric', ''))}</td>"
            f"<td>{html.escape(item.get('message', ''))}</td>"
            f"<td>{html.escape(item.get('recommendation', ''))}</td>"
            "</tr>"
            for item in payload.get("qualityChecks", [])
        )
        quality_table = (
            "<h2>转换质量检查</h2>"
            "<table><thead><tr><th>检查项</th><th>状态</th><th>指标</th><th>说明</th><th>建议</th></tr></thead>"
            f"<tbody>{quality_rows}</tbody></table>"
            if quality_rows
            else ""
        )
        return f"""<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <title>{html.escape(payload['report_type'])}</title>
  <style>
    body {{ font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; margin: 32px; color: #172033; }}
    h1 {{ font-size: 24px; }}
    table {{ border-collapse: collapse; width: 100%; margin-top: 16px; }}
    th, td {{ border: 1px solid #d6dde8; padding: 10px; text-align: left; }}
    th {{ background: #f3f6fa; }}
  </style>
</head>
<body>
  <h1>{html.escape(payload['report_type'])}</h1>
  <p>报告 ID：{html.escape(payload['id'])}</p>
  <ul>{summary}</ul>
  {quality_summary_table}
  <table>
    <thead><tr><th>文件名</th><th>类型</th><th>状态</th><th>校验</th></tr></thead>
    <tbody>{file_rows}</tbody>
  </table>
  {formula_table}
  {dependency_table}
  {omml_prompt_table}
  {macro_table}
  {image_table}
  {mathpix_table}
  {artifact_table}
  {batch_table}
  {quality_table}
</body>
</html>"""

    @staticmethod
    def _text(payload: dict[str, Any]) -> str:
        """将报告载荷渲染为纯文本审计摘要。"""
        lines = [payload["report_type"], f"报告 ID：{payload['id']}", ""]
        lines.extend(payload.get("summary", []))
        quality_summary = payload.get("qualitySummary", [])
        if quality_summary:
            lines.append("")
            lines.append("质量摘要")
            for item in quality_summary:
                lines.append(f"- {item.get('label', '')} / {item.get('status', '')} / {item.get('metric', '')} / {item.get('detail', '')} / {item.get('recommendation', '')}")
        lines.append("")
        lines.append("文件清单")
        file_names = {file.get("id"): file.get("file_name", "") for file in payload.get("files", [])}
        for file in payload.get("files", []):
            errors = "；".join(file.get("validation_errors") or ["无"])
            lines.append(f"- {file.get('file_name', '')} / {file.get('file_type', '')} / {file.get('status', '')} / {errors}")
        formulas = payload.get("analysis", {}).get("formulas", [])
        if formulas:
            lines.append("")
            lines.append("公式明细")
            for item in formulas:
                source_file = file_names.get(item.get("file_id"), item.get("file_id", ""))
                lines.append(
                    f"- {source_file} / {item.get('position', '')} / 位置状态 {item.get('position_status', '已记录')} / {item.get('source_type', '')} / {item.get('status', '')} / {item.get('confidence', '')}% / "
                    f"{ReportBuilder._formula_style_label(item)} / 原始截图 {item.get('original_image_ref') or item.get('original_image_path') or '-'} / "
                    f"MathType {item.get('mathtype_preview') or item.get('mathtype_data') or '-'} / 格式化对比 {ReportBuilder._formula_format_comparison_label(item)} / "
                    f"失败兜底 {ReportBuilder._formula_format_policy_label(item)} / {item.get('latex', '')}"
                )
        omml_prompts = payload.get("analysis", {}).get("ommlConversionPrompts", [])
        if omml_prompts:
            lines.append("")
            lines.append("OMML 转换确认")
            for item in omml_prompts:
                lines.append(
                    f"- {item.get('file_name', '')} / {item.get('status', '')} / OMML {item.get('omml_count', '')} 个 / "
                    f"转换 MathType {item.get('convert_to_mathtype')} / 决策 {item.get('decision_source', '')} / "
                    f"依赖 {item.get('dependency_status', '')} / {item.get('message', '')}"
                )
        small_images = payload.get("analysis", {}).get("smallImages", [])
        if small_images:
            lines.append("")
            lines.append("微小图片")
            for item in small_images:
                source_file = file_names.get(item.get("file_id"), item.get("file_id", ""))
                lines.append(
                    f"- {source_file} / {item.get('source_name', '')} / {item.get('location', '')} / "
                    f"{item.get('width', '')}x{item.get('height', '')} / 面积 {item.get('area', '')} / 图片格式 {item.get('image_type') or '-'} / "
                    f"疑似类型 {ReportBuilder._small_image_label(item)} / 判定标记 {ReportBuilder._small_image_flags_label(item)} / "
                    f"重复判断 {item.get('duplicate_check_status') or '未比较'} / "
                    f"导出状态 {item.get('export_status') or '可导出'} / "
                    f"置信度 {item.get('confidence', '')}%"
                )
        macros = payload.get("analysis", {}).get("macros", [])
        if macros:
            lines.append("")
            lines.append("宏执行队列")
            for item in macros:
                source_file = file_names.get(item.get("file_id"), item.get("file_id", ""))
                lines.append(
                    f"- {source_file} / {item.get('macro_name', '')} / 顺序 {item.get('execute_order', '')} / "
                    f"{item.get('execute_status', '')} / 失败策略 {item.get('failure_strategy', '')} / "
                    f"失败动作 {ReportBuilder._macro_failure_policy_label(item)}"
                )
        mathpix_jobs = payload.get("analysis", {}).get("mathpix", [])
        if mathpix_jobs:
            lines.append("")
            lines.append("Mathpix 保留计划")
            for item in mathpix_jobs:
                lines.append(
                    f"- {item.get('file_name', '')} / {item.get('status', '')} / "
                    f"{ReportBuilder._mathpix_ocr_label(item)} / {ReportBuilder._mathpix_hint_label(item)} / "
                    f"{ReportBuilder._mathpix_retention_label(item)} / {ReportBuilder._mathpix_note_label(item)}"
                )
        artifacts = payload.get("analysis", {}).get("artifacts", [])
        if artifacts:
            lines.append("")
            lines.append("转换输出")
            for item in artifacts:
                settings_label = ReportBuilder._conversion_settings_label(item.get("conversion_settings") or {})
                preservation_label = ReportBuilder._artifact_preservation_label(item)
                lines.append(f"- {item.get('source_file', '')} -> {item.get('file_name') or '-'} / {item.get('status', '')} / {item.get('message', '')} / {settings_label} / 对象保留 {preservation_label}")
        batch_results = payload.get("analysis", {}).get("batchResults", [])
        if batch_results:
            lines.append("")
            lines.append("批量单文件结果")
            for item in batch_results:
                error = item.get("error_message") or "无"
                lines.append(f"- {item.get('sequence', '')}. {item.get('file_name', '')} / {item.get('status', '')} / {error} / {item.get('recommendation', '')}")
        quality_checks = payload.get("qualityChecks", [])
        if quality_checks:
            lines.append("")
            lines.append("转换质量检查")
            for item in quality_checks:
                lines.append(f"- {item.get('label', '')} / {item.get('status', '')} / {item.get('metric', '')} / {item.get('recommendation', '')}")
        return "\n".join(lines) + "\n"

    @staticmethod
    def _pdf_lines(payload: dict[str, Any]) -> list[str]:
        """返回轻量 PDF 渲染器使用的简短文本行。"""
        lines = [payload["report_type"], f"Report ID: {payload['id']}", ""]
        lines.extend(payload.get("summary", []))
        lines.append("")
        lines.append("Files")
        file_names = {file.get("id"): file.get("file_name", "") for file in payload.get("files", [])}
        for file in payload.get("files", []):
            lines.append(f"{file.get('file_name', '')} | {file.get('file_type', '')} | {file.get('status', '')}")
        formulas = payload.get("analysis", {}).get("formulas", [])
        if formulas:
            lines.append("")
            lines.append("Formulas")
            for item in formulas:
                source_file = file_names.get(item.get("file_id"), item.get("file_id", ""))
                lines.append(
                    f"{source_file} | {item.get('position', '')} | {item.get('source_type', '')} | {item.get('confidence', '')}% | "
                    f"{item.get('status', '')} | image={item.get('original_image_ref') or item.get('original_image_path') or '-'} | "
                    f"mathtype={item.get('mathtype_preview') or item.get('mathtype_data') or '-'} | fallback={ReportBuilder._formula_format_policy_label(item)}"
                )
        omml_prompts = payload.get("analysis", {}).get("ommlConversionPrompts", [])
        if omml_prompts:
            lines.append("")
            lines.append("OMML Conversion Prompts")
            for item in omml_prompts:
                lines.append(
                    f"{item.get('file_name', '')} | {item.get('status', '')} | omml={item.get('omml_count', '')} | "
                    f"convert={item.get('convert_to_mathtype')} | source={item.get('decision_source', '')} | "
                    f"dependency={item.get('dependency_status', '')}"
                )
        small_images = payload.get("analysis", {}).get("smallImages", [])
        if small_images:
            lines.append("")
            lines.append("Small Images")
            for item in small_images:
                source_file = file_names.get(item.get("file_id"), item.get("file_id", ""))
                lines.append(
                    f"{source_file} | {item.get('source_name', '')} | {item.get('location', '')} | "
                    f"{item.get('width', '')}x{item.get('height', '')} | format={item.get('image_type') or '-'} | "
                    f"type={ReportBuilder._small_image_label(item)} | confidence={item.get('confidence', '')}%"
                )
        artifacts = payload.get("analysis", {}).get("artifacts", [])
        if artifacts:
            lines.append("")
            lines.append("Artifacts")
            for item in artifacts:
                lines.append(f"{item.get('source_file', '')} -> {item.get('file_name') or '-'} | {item.get('status', '')} | Object preservation: {ReportBuilder._artifact_preservation_label(item)}")
        batch_results = payload.get("analysis", {}).get("batchResults", [])
        if batch_results:
            lines.append("")
            lines.append("Batch Results")
            for item in batch_results:
                lines.append(f"{item.get('sequence', '')}. {item.get('file_name', '')} | {item.get('status', '')} | {item.get('error_message') or 'OK'}")
        quality_checks = payload.get("qualityChecks", [])
        if quality_checks:
            lines.append("")
            lines.append("Quality")
            for item in quality_checks:
                lines.append(f"{item.get('label', '')} | {item.get('status', '')} | {item.get('metric', '')}")
        return lines

    @staticmethod
    def _write_xlsx(payload: dict[str, Any], path: Any) -> None:
        """使用最小 XLSX 包写入报告工作簿。"""
        rows = [
            ["报告类型", payload.get("report_type", "")],
            ["报告 ID", payload.get("id", "")],
            ["任务 ID", payload.get("task_id", "")],
            [],
            ["摘要"],
        ]
        rows.extend([[line] for line in payload.get("summary", [])])
        quality_summary = payload.get("qualitySummary", [])
        if quality_summary:
            rows.extend([[], ["质量摘要", "状态", "指标", "说明", "建议"]])
            for item in quality_summary:
                rows.append([item.get("label", ""), item.get("status", ""), item.get("metric", ""), item.get("detail", ""), item.get("recommendation", "")])
        rows.extend([[], ["文件名", "类型", "状态", "校验"]])
        file_names = {file.get("id"): file.get("file_name", "") for file in payload.get("files", [])}
        for file in payload.get("files", []):
            rows.append(
                [
                    file.get("file_name", ""),
                    file.get("file_type", ""),
                    file.get("status", ""),
                    "；".join(file.get("validation_errors") or ["无"]),
                ]
            )
        formulas = payload.get("analysis", {}).get("formulas", [])
        if formulas:
            rows.extend([[], ["公式明细", "页码", "位置", "位置状态", "位置异常", "来源", "原始截图", "LaTeX", "MathML", "MathType 预览", "格式化", "样式", "格式化对比", "失败兜底", "置信度", "状态"]])
            for item in formulas:
                rows.append(
                    [
                        file_names.get(item.get("file_id"), item.get("file_id", "")),
                        item.get("page_index", ""),
                        item.get("position", ""),
                        item.get("position_status", ""),
                        item.get("position_issue", ""),
                        item.get("source_type", ""),
                        item.get("original_image_ref") or item.get("original_image_path", ""),
                        item.get("latex", ""),
                        item.get("mathml", ""),
                        item.get("mathtype_preview") or item.get("mathtype_data", ""),
                        item.get("format_status", ""),
                        ReportBuilder._formula_style_label(item),
                        ReportBuilder._formula_format_comparison_label(item),
                        ReportBuilder._formula_format_policy_label(item),
                        item.get("confidence", ""),
                        item.get("status", ""),
                    ]
                )
        omml_prompts = payload.get("analysis", {}).get("ommlConversionPrompts", [])
        if omml_prompts:
            rows.extend([[], ["OMML 转换确认", "OMML 数量", "状态", "转换 MathType", "决策来源", "依赖状态", "说明"]])
            for item in omml_prompts:
                rows.append(
                    [
                        item.get("file_name", ""),
                        item.get("omml_count", ""),
                        item.get("status", ""),
                        item.get("convert_to_mathtype", ""),
                        item.get("decision_source", ""),
                        item.get("dependency_status", ""),
                        item.get("message", ""),
                    ]
                )
        small_images = payload.get("analysis", {}).get("smallImages", [])
        if small_images:
            rows.extend([[], ["微小图片", "来源文件", "位置", "尺寸", "面积", "图片格式", "疑似类型", "判定标记", "重复判断", "异常兜底", "导出状态", "导出信息", "重新导出时间", "置信度", "导出格式"]])
            for item in small_images:
                rows.append(
                    [
                        file_names.get(item.get("file_id"), item.get("file_id", "")),
                        item.get("source_name", ""),
                        item.get("location", ""),
                        f"{item.get('width', '')} x {item.get('height', '')}",
                        item.get("area", ""),
                        item.get("image_type", ""),
                        ReportBuilder._small_image_label(item),
                        ReportBuilder._small_image_flags_label(item),
                        item.get("duplicate_check_status", ""),
                        item.get("duplicate_fallback", ""),
                        item.get("export_status", ""),
                        item.get("export_message", ""),
                        item.get("reexported_at", ""),
                        item.get("confidence", ""),
                        item.get("export_format", ""),
                    ]
                )
        macros = payload.get("analysis", {}).get("macros", [])
        if macros:
            rows.extend([[], ["宏执行队列", "来源文件", "顺序", "状态", "失败策略", "失败动作", "备份"]])
            for item in macros:
                rows.append(
                    [
                        item.get("macro_name", ""),
                        file_names.get(item.get("file_id"), item.get("file_id", "")),
                        item.get("execute_order", ""),
                        item.get("execute_status", ""),
                        item.get("failure_strategy", ""),
                        ReportBuilder._macro_failure_policy_label(item),
                        item.get("backup_path", ""),
                    ]
                )
        mathpix_jobs = payload.get("analysis", {}).get("mathpix", [])
        if mathpix_jobs:
            rows.extend([[], ["Mathpix 保留计划", "状态", "OCR", "来源线索", "保留计划", "说明"]])
            for item in mathpix_jobs:
                rows.append(
                    [
                        item.get("file_name", ""),
                        item.get("status", ""),
                        ReportBuilder._mathpix_ocr_label(item),
                        ReportBuilder._mathpix_hint_label(item),
                        ReportBuilder._mathpix_retention_label(item),
                        ReportBuilder._mathpix_note_label(item),
                    ]
                )
        artifacts = payload.get("analysis", {}).get("artifacts", [])
        if artifacts:
            rows.extend([[], ["转换输出", "文件", "类型", "状态", "说明", "设置", "对象保留"]])
            for item in artifacts:
                rows.append([
                    item.get("source_file", ""),
                    item.get("file_name", ""),
                    item.get("output_type", ""),
                    item.get("status", ""),
                    item.get("message", ""),
                    ReportBuilder._conversion_settings_label(item.get("conversion_settings") or {}),
                    ReportBuilder._artifact_preservation_label(item),
                ])
        batch_results = payload.get("analysis", {}).get("batchResults", [])
        if batch_results:
            rows.extend([[], ["批量单文件结果", "文件", "类型", "状态", "失败原因", "建议"]])
            for item in batch_results:
                rows.append([item.get("sequence", ""), item.get("file_name", ""), item.get("file_type", ""), item.get("status", ""), item.get("error_message", ""), item.get("recommendation", "")])
        quality_checks = payload.get("qualityChecks", [])
        if quality_checks:
            rows.extend([[], ["质量检查", "状态", "指标", "说明", "建议"]])
            for item in quality_checks:
                rows.append([item.get("label", ""), item.get("status", ""), item.get("metric", ""), item.get("message", ""), item.get("recommendation", "")])
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
            archive.writestr("[Content_Types].xml", ReportBuilder._xlsx_content_types())
            archive.writestr("_rels/.rels", ReportBuilder._xlsx_package_rels())
            archive.writestr("xl/workbook.xml", ReportBuilder._xlsx_workbook())
            archive.writestr("xl/_rels/workbook.xml.rels", ReportBuilder._xlsx_workbook_rels())
            archive.writestr("xl/worksheets/sheet1.xml", ReportBuilder._xlsx_sheet(rows))

    @staticmethod
    def _conversion_settings_label(settings: dict[str, Any]) -> str:
        """格式化产物表中的转换设置。"""
        if not settings:
            return "默认"
        return "；".join(
            [
                f"模板 {settings.get('word_to_ppt_template', '-')}",
                f"PPT转Word {settings.get('ppt_to_word_mode', '-')}",
                f"Word模板 {settings.get('ppt_to_word_template', '-')}",
                f"Word目录 {'是' if settings.get('ppt_to_word_generate_toc') else '否'}",
                f"PPT备注 {'是' if settings.get('ppt_extract_notes', True) else '否'}",
                f"PPT图片 {'是' if settings.get('ppt_retain_images', settings.get('retain_images', True)) else '否'}",
                f"PPT表格 {'是' if settings.get('ppt_retain_tables', settings.get('retain_tables', True)) else '否'}",
                f"PPT公式 {'是' if settings.get('ppt_retain_formulas', True) else '否'}",
                f"PDF {settings.get('pdf_precision_mode', '-')}/{settings.get('pdf_to_word_engine', '-')}",
                f"Excel {settings.get('excel_conversion_range', '-')}",
                f"Excel公式 {settings.get('excel_formula_mode', '-')}",
                f"Excel批注 {'是' if settings.get('excel_retain_comments', settings.get('retain_comments', False)) else '否'}",
                f"拆分工作表 {'是' if settings.get('excel_split_sheets') else '否'}",
                f"输出同名 {settings.get('output_conflict_strategy', '-')}",
                f"图片 {'是' if settings.get('retain_images') else '否'}",
                f"表格 {'是' if settings.get('retain_tables') else '否'}",
                f"页眉页脚 {'是' if settings.get('retain_headers_footers') else '否'}",
                f"脚注尾注 {'是' if settings.get('retain_footnotes_endnotes') else '否'}",
                f"批注 {'是' if settings.get('retain_comments') else '否'}",
                f"修订 {'是' if settings.get('retain_revisions') else '否'}",
                f"每页字数 {settings.get('word_max_chars_per_slide', '-')}",
                f"自动分页 {'是' if settings.get('word_auto_pagination') else '否'}",
                f"目录页 {'是' if settings.get('word_generate_toc') else '否'}",
                f"保留公式 {'是' if settings.get('word_retain_formulas') else '否'}",
                f"先转MathType {'是' if settings.get('word_convert_omml_first') else '否'}",
                f"套模板 {'是' if settings.get('word_apply_template') else '否'}",
                f"演讲备注 {'是' if settings.get('word_generate_notes') else '否'}",
                f"自动美化 {'是' if settings.get('word_auto_beautify') else '否'}",
            ]
        )

    @staticmethod
    def _artifact_preservation_label(item: dict[str, Any]) -> str:
        """说明图片、表格、公式、OMML 和 MathType 保留情况。"""
        plan = item.get("object_preservation") or {}
        if not plan:
            return "未记录"
        source = plan.get("source") or {}
        statuses = plan.get("statuses") or {}
        return "；".join(
            [
                f"图片 {ReportBuilder._safe_int(source.get('images'))} 个/{statuses.get('images') or '未记录'}",
                f"表格 {ReportBuilder._safe_int(source.get('tables'))} 个/{statuses.get('tables') or '未记录'}",
                f"公式 {ReportBuilder._safe_int(source.get('formulas'))} 个/{statuses.get('formulas') or '未记录'}",
                f"OMML {ReportBuilder._safe_int(source.get('omml_formulas'))} 个",
                f"MathType {ReportBuilder._safe_int(source.get('mathtype_objects'))} 个",
            ]
        )

    @staticmethod
    def _mathpix_ocr_label(item: dict[str, Any]) -> str:
        """格式化 Mathpix 文字、公式和表格 OCR 开关。"""
        plan = item.get("retention_plan") or {}
        ocr = item.get("ocr_settings") or {}
        return "；".join(
            [
                f"文字 {'开' if ReportBuilder._plan_bool(plan, ocr, 'text_ocr') else '关'}",
                f"公式 {'开' if ReportBuilder._plan_bool(plan, ocr, 'formula_ocr') else '关'}",
                f"表格 {'开' if ReportBuilder._plan_bool(plan, ocr, 'table_ocr') else '关'}",
            ]
        )

    @staticmethod
    def _mathpix_hint_label(item: dict[str, Any]) -> str:
        """格式化 PDF 类型和检测到的 Mathpix 源线索。"""
        plan = item.get("retention_plan") or {}
        return "；".join(
            [
                str(plan.get("pdf_type") or "未知 PDF"),
                f"图片 {plan.get('image_objects', 0)} 个",
                f"表格 {plan.get('table_hints', 0)} 个",
                f"公式 {plan.get('formula_hints', 0)} 个",
            ]
        )

    @staticmethod
    def _mathpix_retention_label(item: dict[str, Any]) -> str:
        """格式化 Mathpix 图片、表格和公式保留状态。"""
        plan = item.get("retention_plan") or {}
        if not plan:
            return "未记录"
        return "；".join(
            [
                str(plan.get("image_retention_status") or "图片未记录"),
                str(plan.get("table_retention_status") or "表格未记录"),
                str(plan.get("formula_retention_status") or "公式未记录"),
            ]
        )

    @staticmethod
    def _mathpix_note_label(item: dict[str, Any]) -> str:
        """将 Mathpix 保留说明合并到单个表格单元格。"""
        plan = item.get("retention_plan") or {}
        notes = plan.get("notes") or []
        if isinstance(notes, list):
            return "；".join(str(note) for note in notes)
        return str(notes or "")

    @staticmethod
    def _plan_bool(plan: dict[str, Any], fallback: dict[str, Any], key: str) -> bool:
        """读取计划布尔开关，缺失时使用 OCR 设置。"""
        if key in plan:
            return bool(plan.get(key))
        return bool(fallback.get(key))

    @staticmethod
    def _formula_style_label(item: dict[str, Any]) -> str:
        """格式化报告行中的公式字体和间距设置。"""
        return "；".join(
            [
                f"范围 {item.get('format_scope', '-')}",
                f"字体 {item.get('font', '-')}",
                f"字号 {item.get('font_size', '-')}",
                f"对齐 {item.get('alignment', '-')}",
                f"变量 {item.get('variable_style', '-')}",
                f"函数 {item.get('function_style', '-')}",
                f"上下标 {item.get('script_scale', '-')}%",
                f"分式 {item.get('fraction_style', '-')}",
                f"根式 {item.get('radical_style', '-')}",
                f"矩阵 {item.get('matrix_spacing', '-')}",
                f"希腊 {item.get('greek_style', '-')}",
                f"行内 {item.get('inline_baseline', '-')}",
                f"段距 {item.get('display_spacing', '-')}",
                f"编号 {item.get('numbering', '-')}",
            ]
        )

    @staticmethod
    def _formula_format_comparison_label(item: dict[str, Any]) -> str:
        """说明公式格式化前后对比。"""
        comparison = item.get("format_comparison") or {}
        summary = comparison.get("summary")
        if summary:
            return str(summary)
        before = comparison.get("before") or {}
        after = comparison.get("after") or {}
        if before or after:
            return f"{before.get('format', '-')} -> {after.get('format', '-')}"
        return "-"

    @staticmethod
    def _formula_format_policy_label(item: dict[str, Any]) -> str:
        """说明公式格式化兜底与重试策略。"""
        policy = item.get("format_failure_policy") or {}
        if not policy:
            return "未记录"
        preserve = "保留原公式" if policy.get("preserve_original_formula") else "不保留原公式"
        retry = "可重试" if policy.get("retryable") else "无需重试"
        return "；".join(
            [
                str(policy.get("status") or "-"),
                preserve,
                f"兜底格式 {policy.get('fallback_format', '-')}",
                retry,
                str(policy.get("message") or ""),
            ]
        )

    @staticmethod
    def _macro_failure_policy_label(item: dict[str, Any]) -> str:
        """说明本地客户端交接使用的宏失败处理动作。"""
        policy = item.get("failure_policy") or {}
        description = policy.get("description")
        if description:
            return str(description)
        action = policy.get("action")
        return str(action) if action else "-"

    @staticmethod
    def _xlsx_sheet(rows: list[list[Any]]) -> str:
        """生成使用内联字符串单元格的工作表 XML。"""
        row_xml = []
        for row_index, row in enumerate(rows, start=1):
            cells = []
            for column_index, value in enumerate(row, start=1):
                ref = f"{ReportBuilder._column_name(column_index)}{row_index}"
                cells.append(f'<c r="{ref}" t="inlineStr"><is><t>{html.escape(str(value))}</t></is></c>')
            row_xml.append(f'<row r="{row_index}">{"".join(cells)}</row>')
        return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>{"".join(row_xml)}</sheetData></worksheet>"""

    @staticmethod
    def _column_name(index: int) -> str:
        """将从 1 开始的列序号转为 Excel 列名。"""
        name = ""
        while index:
            index, remainder = divmod(index - 1, 26)
            name = chr(65 + remainder) + name
        return name or "A"

    @staticmethod
    def _xlsx_content_types() -> str:
        """返回生成的 XLSX 包内容类型 XML。"""
        return """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>"""

    @staticmethod
    def _xlsx_package_rels() -> str:
        """返回指向工作簿入口的包关系 XML。"""
        return """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"""

    @staticmethod
    def _xlsx_workbook() -> str:
        """返回含安全工作表名称的工作簿 XML。"""
        return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="K12 Report" sheetId="1" r:id="rId1"/></sheets></workbook>"""

    @staticmethod
    def _xlsx_workbook_rels() -> str:
        """返回工作簿指向单工作表的关系 XML。"""
        return """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"""

    @staticmethod
    def _small_image_label(item: dict[str, Any]) -> str:
        """将小图片分类为公式、二维码、印章、签名或图标。"""
        if item.get("is_formula_like"):
            return "公式"
        if item.get("is_qrcode_like"):
            return "二维码"
        if item.get("is_stamp_like"):
            return "印章"
        if item.get("is_signature_like"):
            return "签名"
        return "图标"

    @staticmethod
    def _small_image_flags_label(item: dict[str, Any]) -> str:
        """格式化水印、重复、透明等小图片标记。"""
        flags = []
        if item.get("is_header_footer"):
            flags.append("页眉页脚")
        if item.get("is_watermark"):
            flags.append("水印")
        if item.get("is_transparent"):
            flags.append("透明")
        if item.get("is_duplicate"):
            flags.append("重复")
        return "、".join(flags) if flags else "无"
