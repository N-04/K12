import ast
import base64
import hashlib
import json
import os
import re
import tempfile
import unittest
import zipfile
import struct
import zlib
from io import BytesIO
from pathlib import Path
from types import SimpleNamespace
from urllib.error import HTTPError
from unittest.mock import patch

import k12.converters as converters_module
import k12.exports as exports_module
import k12.local_client as local_client_module
import k12.models as models_module
import k12.previews as previews_module
import k12.processor as processor_module
import k12.reports as reports_module
import k12.server as server_module
import k12.store as store_module
from k12.converters import extract_pptx_slides, extract_xlsx_sheets
from k12.exports import build_formula_xlsx, build_formula_zip
from k12.local_client import (
    LocalClientError,
    build_component_preflight,
    build_dry_run_execution_summary,
    build_dry_run_sync_payload,
    build_heartbeat,
    build_native_execution_request,
    execute_local_file_actions,
    run_once,
    sanitize_component_preflight,
    summarize_payload,
)
from k12.mathpix import MATHPIX_ERROR_DETAIL_LIMIT, MathpixApiError, MathpixClient, MathpixConfigError
from k12.models import Task
from k12.processor import DocumentAnalyzer, TaskProcessor
from k12.server import JsonError, K12RequestHandler
from k12.store import AppStore


class DocumentAnalyzerTests(unittest.TestCase):
    def test_classifies_supported_files(self) -> None:
        analyzer = DocumentAnalyzer()
        item = analyzer.analyze_metadata("math-lesson.omml.docx", 2048)
        self.assertEqual(item.file_type, "Word")
        self.assertTrue(item.has_omml)
        self.assertTrue(item.missing_omml_dependency)
        self.assertFalse(item.validation_errors)

    def test_rejects_unsupported_files(self) -> None:
        analyzer = DocumentAnalyzer(single_file_limit_mb=1)
        item = analyzer.analyze_metadata("bad.exe", 10)
        self.assertIn("文件格式不支持", item.validation_errors)
        self.assertEqual(item.status, "校验失败")

    def test_detects_word_omml_dependency_from_sibling_file(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "lesson.docx"
            path.write_bytes(make_docx_bytes())
            analyzer = DocumentAnalyzer()

            missing = analyzer.analyze_file(path)
            self.assertTrue(missing.has_omml)
            self.assertTrue(missing.missing_omml_dependency)
            self.assertTrue(missing.content_summary["missingOmmlDependency"])

            (Path(tmp) / "OMML2MML.XSL").write_text("<xsl:stylesheet />", encoding="utf-8")
            ready = analyzer.analyze_file(path)
            self.assertTrue(ready.has_omml)
            self.assertFalse(ready.missing_omml_dependency)
            self.assertFalse(ready.content_summary["missingOmmlDependency"])

    def test_detects_image_presence_separately_from_small_images(self) -> None:
        analyzer = DocumentAnalyzer()
        image = analyzer.analyze_metadata("diagram.png", 2048)
        self.assertEqual(image.file_type, "图片")
        self.assertTrue(image.has_image)
        self.assertFalse(image.has_small_image)

        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "with-image.docx"
            path.write_bytes(make_docx_bytes({"word/media/photo.png": make_png_bytes(40, 30)}))
            word = analyzer.analyze_file(path)
            self.assertTrue(word.has_image)
            self.assertTrue(word.has_small_image)

    def test_detects_word_review_and_reference_structures(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "reviewed.docx"
            path.write_bytes(
                make_docx_bytes(
                    {
                        "word/header1.xml": b"<w:hdr><w:p><w:r><w:t>Header</w:t></w:r></w:p></w:hdr>",
                        "word/footer1.xml": b"<w:ftr><w:p><w:r><w:t>Footer</w:t></w:r></w:p></w:ftr>",
                        "word/footnotes.xml": b'<w:footnotes><w:footnote w:id="2"><w:p><w:r><w:t>Note</w:t></w:r></w:p></w:footnote></w:footnotes>',
                        "word/endnotes.xml": b'<w:endnotes><w:endnote w:id="3"><w:p><w:r><w:t>End</w:t></w:r></w:p></w:endnote></w:endnotes>',
                        "word/comments.xml": b'<w:comments><w:comment w:id="0"><w:p><w:r><w:t>Comment</w:t></w:r></w:p></w:comment></w:comments>',
                    },
                    document_xml='<w:document><w:p><w:ins><w:r><w:t>新增</w:t></w:r></w:ins><w:del><w:r><w:delText>删除</w:delText></w:r></w:del></w:p></w:document>',
                )
            )

            word = DocumentAnalyzer().analyze_file(path)
            summary = word.content_summary

            self.assertEqual(summary["headers"], 1)
            self.assertEqual(summary["footers"], 1)
            self.assertEqual(summary["footnotes"], 1)
            self.assertEqual(summary["endnotes"], 1)
            self.assertEqual(summary["comments"], 1)
            self.assertEqual(summary["revisions"], 2)

    def test_classifies_pdf_type_and_ocr_hints(self) -> None:
        analyzer = DocumentAnalyzer()
        with tempfile.TemporaryDirectory() as tmp:
            text_pdf = Path(tmp) / "text.pdf"
            text_pdf.write_bytes(b"%PDF-1.4\n1 0 obj<< /Type /Page >>(Hello classroom) Tj endobj\n")
            scanned_pdf = Path(tmp) / "scan.pdf"
            scanned_pdf.write_bytes(b"%PDF-1.4\n1 0 obj<< /Type /Page /Subtype /Image >>endobj\n")
            mixed_pdf = Path(tmp) / "mixed.pdf"
            mixed_pdf.write_bytes(b"%PDF-1.4\n1 0 obj<< /Type /Page /Subtype /Image /Table >>(Equation Formula) Tj endobj\n")

            text_item = analyzer.analyze_file(text_pdf)
            scanned_item = analyzer.analyze_file(scanned_pdf)
            mixed_item = analyzer.analyze_file(mixed_pdf)

            self.assertEqual(text_item.content_summary["pdfType"], "文本型 PDF")
            self.assertEqual(text_item.content_summary["ocrRecommendation"], "可解析文本层")
            self.assertEqual(scanned_item.content_summary["pdfType"], "扫描型 PDF")
            self.assertIn("文字 OCR", scanned_item.content_summary["ocrRecommendation"])
            self.assertEqual(mixed_item.content_summary["pdfType"], "混合型 PDF")
            self.assertTrue(mixed_item.has_formula)
            self.assertGreater(mixed_item.content_summary["tableHints"], 0)
            self.assertIn("公式 OCR", mixed_item.content_summary["ocrRecommendation"])


class TaskProcessorTests(unittest.TestCase):
    def test_core_prd_modules_keep_boundary_docstrings(self) -> None:
        root = Path(__file__).resolve().parents[1]
        for module_path in sorted((root / "k12").glob("*.py")):
            tree = ast.parse(module_path.read_text(encoding="utf-8"))
            self.assertTrue(ast.get_docstring(tree), f"{module_path.name} should have a module docstring")
            missing_public_docstrings = [
                f"{module_path.name}:{node.lineno}:{node.name}"
                for node in tree.body
                if isinstance(node, (ast.ClassDef, ast.FunctionDef, ast.AsyncFunctionDef))
                and not node.name.startswith("_")
                and ast.get_docstring(node) is None
            ]
            self.assertFalse(missing_public_docstrings, f"public docstrings missing: {missing_public_docstrings}")
        required_method_docstrings = {
            "models.py": {"FileItem", "Task", "FormulaItem", "MacroItem", "OmmlDependencyItem", "SmallImageItem", "ReportItem"},
            "processor.py": {"DocumentAnalyzer", "TaskProcessor"},
            "server.py": {"K12RequestHandler"},
            "store.py": {"AppStore"},
        }
        for module_name, class_names in required_method_docstrings.items():
            module_path = root / "k12" / module_name
            tree = ast.parse(module_path.read_text(encoding="utf-8"))
            missing_method_docstrings: list[str] = []
            for node in tree.body:
                if not isinstance(node, ast.ClassDef) or node.name not in class_names:
                    continue
                for member in node.body:
                    if isinstance(member, (ast.FunctionDef, ast.AsyncFunctionDef)) and not member.name.startswith("_") and ast.get_docstring(member) is None:
                        missing_method_docstrings.append(f"{module_name}:{node.name}.{member.name}:{member.lineno}")
            self.assertFalse(missing_method_docstrings, f"public method docstrings missing: {missing_method_docstrings}")
        processor_method_docstrings = {
            "preflight_checks",
            "resolve_execute_mode",
            "capabilities",
            "architecture_blueprint",
            "api_catalog",
            "data_dictionary",
            "product_summary",
            "enhancement_plan",
            "install_plan",
            "installer_download_info",
            "local_client_manifest",
            "record_local_client_heartbeat",
            "local_result_upload_queue",
            "local_result_upload_manifest",
            "register_local_result_upload",
            "mathpix_job_queue",
            "create_local_launch_request",
            "local_task_readiness",
            "install_profile",
        }
        processor_tree = ast.parse((root / "k12" / "processor.py").read_text(encoding="utf-8"))
        processor_class = next(node for node in processor_tree.body if isinstance(node, ast.ClassDef) and node.name == "TaskProcessor")
        processor_methods = {node.name: node for node in processor_class.body if isinstance(node, ast.FunctionDef)}
        missing_processor_docstrings = [
            name for name in sorted(processor_method_docstrings)
            if ast.get_docstring(processor_methods[name]) is None
        ]
        self.assertFalse(missing_processor_docstrings, f"TaskProcessor docstrings missing: {missing_processor_docstrings}")
        self.assertIn("local-first baseline", converters_module.__doc__ or "")
        self.assertIn("Office, MathType, OMML writeback", converters_module.__doc__ or "")
        self.assertIn("downloadable reports", exports_module.__doc__ or "")
        self.assertIn("without reaching back into local source paths", exports_module.__doc__ or "")
        self.assertIn("editing user documents", previews_module.__doc__ or "")
        self.assertIn("token-gated sensitive payloads", server_module.__doc__ or "")
        self.assertIn("redacts local paths", server_module.__doc__ or "")
        self.assertIn("SQLite-backed runtime store", store_module.__doc__ or "")
        self.assertIn("public redaction happens at the API layer", store_module.__doc__ or "")
        self.assertIn("Core data models", models_module.__doc__ or "")
        self.assertIn("Business orchestration", processor_module.__doc__ or "")
        self.assertIn("does not claim real Office", processor_module.__doc__ or "")
        self.assertIn("Report generation", reports_module.__doc__ or "")
        self.assertIn("local-client handoff", reports_module.__doc__ or "")
        self.assertIn("safe handoff step", local_client_module.LocalClientError.__doc__ or "")
        self.assertEqual(local_client_module.LocalClientError.code, "local_client_error")
        self.assertEqual(MathpixConfigError.code, "mathpix_config_error")
        self.assertEqual(MathpixApiError.code, "mathpix_api_error")
        self.assertIn("managed upload cache", store_module.AppStore._cleanup_uploaded_payloads.__doc__ or "")
        self.assertIn("managed reports directory", store_module.AppStore._cleanup_report_files.__doc__ or "")
        self.assertIn("managed image directory", store_module.AppStore._cleanup_report_image_files.__doc__ or "")
        self.assertIn("task-scoped output", store_module.AppStore._cleanup_runtime_dirs.__doc__ or "")
        self.assertIn("local K12 API origin", local_client_module.normalize_origin.__doc__ or "")
        self.assertIn("path-redacted heartbeat", local_client_module.build_heartbeat.__doc__ or "")
        self.assertIn("never sends component paths or token material", local_client_module.build_heartbeat.__doc__ or "")
        self.assertIn("without leaking local paths or tokens", local_client_module.summarize_payload.__doc__ or "")
        self.assertIn("without executing native document actions", local_client_module.build_dry_run_execution_summary.__doc__ or "")
        self.assertIn("same-platform formula delivery contract", local_client_module.build_dry_run_execution_summary.__doc__ or "")
        self.assertIn("formula-platform checks", local_client_module._dry_run_blockers.__doc__ or "")
        self.assertIn("dry-run handoff validation only", local_client_module.build_dry_run_sync_payload.__doc__ or "")
        self.assertIn("same-platform native runner", local_client_module.build_native_execution_request.__doc__ or "")
        self.assertIn("does not run Office, MathType, OMML writeback, or Word macros", local_client_module.run_once.__doc__ or "")
        self.assertIn("redacted JSON handoff result", local_client_module.main.__doc__ or "")
        self.assertIn("Extract visible DOCX", converters_module.extract_docx_blocks.__doc__ or "")
        self.assertIn("Package formulas", exports_module.build_formula_zip.__doc__ or "")
        self.assertIn("safe preview payload", previews_module.build_file_preview.__doc__ or "")
        self.assertIn("local security token", server_module.K12RequestHandler._ensure_authorized.__doc__ or "")
        self.assertIn("runtime directories", store_module.AppStore.__doc__ or "")
        self.assertIn("Uploaded or discovered file metadata", models_module.FileItem.__doc__ or "")
        self.assertIn("Classify uploads", processor_module.DocumentAnalyzer.__doc__ or "")
        self.assertIn("PRD acceptance evidence", processor_module.TaskProcessor.acceptance_matrix.__doc__ or "")
        self.assertIn("downloadable report formats", reports_module.ReportBuilder.build.__doc__ or "")

    def test_codebase_does_not_use_divider_comments(self) -> None:
        root = Path(__file__).resolve().parents[1]
        divider_pattern = re.compile(r"#\s*(?:-{3,}|={3,}|\u2014{2,})")
        paths = [root / "README.md"]
        for folder in ("k12", "static", "tests", "docs"):
            paths.extend(path for path in (root / folder).rglob("*") if path.is_file())
        for path in paths:
            if path.suffix in {".py", ".js", ".css", ".html", ".md"}:
                text = path.read_text(encoding="utf-8", errors="ignore")
                self.assertIsNone(divider_pattern.search(text), f"divider comment found in {path}")

    def test_app_store_methods_are_documented_and_timestamp_parsers_are_distinct(self) -> None:
        source = Path(store_module.__file__).read_text(encoding="utf-8")
        tree = ast.parse(source)
        app_store_node = next(node for node in tree.body if isinstance(node, ast.ClassDef) and node.name == "AppStore")
        methods = [node for node in app_store_node.body if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef))]
        missing_docstrings = [node.name for node in methods if ast.get_docstring(node) is None]
        self.assertFalse(missing_docstrings, f"AppStore docstrings missing: {missing_docstrings}")
        method_names = [node.name for node in methods]
        self.assertEqual(method_names.count("_parse_timestamp"), 1)
        self.assertEqual(method_names.count("_parse_cleanup_timestamp"), 1)
        self.assertIn("missing or malformed", store_module.AppStore._parse_timestamp.__doc__ or "")
        self.assertIn("retention comparisons", store_module.AppStore._parse_cleanup_timestamp.__doc__ or "")
        cutoff = store_module.datetime.now(store_module.timezone.utc)
        self.assertIsNone(store_module.AppStore._parse_timestamp(None))
        self.assertFalse(store_module.AppStore._timestamp_at_or_before("not-a-date", cutoff))

    def test_api_handler_methods_are_documented_for_sensitive_routes(self) -> None:
        source = Path(server_module.__file__).read_text(encoding="utf-8")
        tree = ast.parse(source)
        class_names = {"JsonError", "K12RequestHandler", "K12Server"}
        missing_docstrings: list[str] = []
        for node in tree.body:
            if not isinstance(node, ast.ClassDef) or node.name not in class_names:
                continue
            for member in node.body:
                if isinstance(member, (ast.FunctionDef, ast.AsyncFunctionDef)) and ast.get_docstring(member) is None:
                    missing_docstrings.append(f"{node.name}.{member.name}:{member.lineno}")
        self.assertFalse(missing_docstrings, f"server method docstrings missing: {missing_docstrings}")
        self.assertIn("Windows or macOS installer", server_module.K12RequestHandler._send_installer.__doc__ or "")
        self.assertIn("MathType boundary headers", server_module.K12RequestHandler._send_installer.__doc__ or "")
        self.assertIn("local security token", server_module.K12RequestHandler._ensure_authorized.__doc__ or "")
        self.assertIn("local security token", server_module.K12RequestHandler._ensure_local_payload_token_configured.__doc__ or "")
        self.assertIn("without trusting URL paths", server_module.K12RequestHandler._find_artifact_path.__doc__ or "")
        self.assertIn("macOS", server_module.K12RequestHandler._redact_text_paths.__doc__ or "")
        self.assertIn("Windows", server_module.K12RequestHandler._redact_text_paths.__doc__ or "")

    def test_report_builder_methods_are_documented_for_prd_outputs(self) -> None:
        source = Path(reports_module.__file__).read_text(encoding="utf-8")
        tree = ast.parse(source)
        report_builder = next(node for node in tree.body if isinstance(node, ast.ClassDef) and node.name == "ReportBuilder")
        missing_docstrings = [
            f"ReportBuilder.{member.name}:{member.lineno}"
            for member in report_builder.body
            if isinstance(member, (ast.FunctionDef, ast.AsyncFunctionDef)) and ast.get_docstring(member) is None
        ]
        self.assertFalse(missing_docstrings, f"ReportBuilder docstrings missing: {missing_docstrings}")
        self.assertIn("Mathpix", reports_module.ReportBuilder._failure_rows.__doc__ or "")
        self.assertIn("local-client handoff", reports_module.ReportBuilder._macro_failure_policy_label.__doc__ or "")
        self.assertIn("workflow order", reports_module.ReportBuilder._workflow_plan_metric.__doc__ or "")
        self.assertIn("formula", reports_module.ReportBuilder._mathpix_ocr_label.__doc__ or "")
        self.assertIn("XLSX", reports_module.ReportBuilder._write_xlsx.__doc__ or "")
        self.assertIn("PDF renderer", reports_module.ReportBuilder._pdf_lines.__doc__ or "")

    def test_export_helpers_are_documented_for_download_artifacts(self) -> None:
        source = Path(exports_module.__file__).read_text(encoding="utf-8")
        tree = ast.parse(source)
        missing_docstrings = [
            f"{member.name}:{member.lineno}"
            for member in tree.body
            if isinstance(member, (ast.FunctionDef, ast.AsyncFunctionDef)) and ast.get_docstring(member) is None
        ]
        self.assertFalse(missing_docstrings, f"export helper docstrings missing: {missing_docstrings}")
        self.assertIn("macro failure CSV", exports_module._macro_failure_row_needed.__doc__ or "")
        self.assertIn("OMML dependency", exports_module._omml_failure_row.__doc__ or "")
        self.assertIn("Hide local paths", exports_module._export_path.__doc__ or "")
        self.assertIn("path-free JSON", exports_module._formula_json.__doc__ or "")
        self.assertIn("safe MathML text fallback", exports_module._formula_mathml.__doc__ or "")
        self.assertIn("minimal XLSX workbook", exports_module._xlsx_bytes.__doc__ or "")
        self.assertIn("deterministic export output", exports_module._csv_cell.__doc__ or "")

    def test_converter_helpers_are_documented_for_local_first_outputs(self) -> None:
        source = Path(converters_module.__file__).read_text(encoding="utf-8")
        tree = ast.parse(source)
        missing_docstrings = [
            f"{member.name}:{member.lineno}"
            for member in tree.body
            if isinstance(member, (ast.FunctionDef, ast.AsyncFunctionDef)) and ast.get_docstring(member) is None
        ]
        self.assertFalse(missing_docstrings, f"converter helper docstrings missing: {missing_docstrings}")
        self.assertIn("lightweight slides", converters_module._slides_from_blocks.__doc__ or "")
        self.assertIn("without performing native writeback", converters_module._docx_object_preservation_lines.__doc__ or "")
        self.assertIn("without claiming real rendering", converters_module._pptx_slide_summary.__doc__ or "")
        self.assertIn("formula-like PPTX markers", converters_module._pptx_formula_count.__doc__ or "")
        self.assertIn("lightweight PDF writer", converters_module._pdf_pages.__doc__ or "")
        self.assertIn("minimal PPTX slide", converters_module._slide_xml.__doc__ or "")

    def test_preview_helpers_are_documented_for_safe_ui_payloads(self) -> None:
        source = Path(previews_module.__file__).read_text(encoding="utf-8")
        tree = ast.parse(source)
        missing_docstrings = [
            f"{member.name}:{member.lineno}"
            for member in tree.body
            if isinstance(member, (ast.FunctionDef, ast.AsyncFunctionDef)) and ast.get_docstring(member) is None
        ]
        self.assertFalse(missing_docstrings, f"preview helper docstrings missing: {missing_docstrings}")
        self.assertIn("metadata-first PDF preview", previews_module._pdf_preview.__doc__ or "")
        self.assertIn("OCR and Mathpix hints", previews_module._pdf_preview.__doc__ or "")
        self.assertIn("PRD-sensitive preview findings", previews_module._preview_markers.__doc__ or "")
        self.assertIn("low-confidence warning marker", previews_module._has_low_confidence_hint.__doc__ or "")
        self.assertIn("header bytes", previews_module._image_dimensions.__doc__ or "")

    def test_document_analyzer_helpers_are_documented_for_prd_boundaries(self) -> None:
        source = Path(processor_module.__file__).read_text(encoding="utf-8")
        tree = ast.parse(source)
        top_level_missing = [
            f"{member.name}:{member.lineno}"
            for member in tree.body
            if isinstance(member, (ast.FunctionDef, ast.AsyncFunctionDef)) and ast.get_docstring(member) is None
        ]
        self.assertFalse(top_level_missing, f"processor module helper docstrings missing: {top_level_missing}")
        analyzer = next(node for node in tree.body if isinstance(node, ast.ClassDef) and node.name == "DocumentAnalyzer")
        analyzer_missing = [
            f"DocumentAnalyzer.{member.name}:{member.lineno}"
            for member in analyzer.body
            if isinstance(member, (ast.FunctionDef, ast.AsyncFunctionDef)) and ast.get_docstring(member) is None
        ]
        self.assertFalse(analyzer_missing, f"DocumentAnalyzer helper docstrings missing: {analyzer_missing}")
        self.assertIn("without rendering page pixels", processor_module._pdf_image_descriptors.__doc__ or "")
        self.assertIn("OMML conversion dependency", processor_module._find_direct_omml_file.__doc__ or "")
        self.assertIn("Mathpix-oriented OCR planning", processor_module.DocumentAnalyzer._analyze_pdf.__doc__ or "")
        self.assertIn("without uploading it", processor_module.DocumentAnalyzer._analyze_pdf.__doc__ or "")
        self.assertIn("Mathpix OCR toggles", processor_module.DocumentAnalyzer._pdf_ocr_recommendation.__doc__ or "")

    def test_task_processor_core_helpers_are_documented_for_local_handoff(self) -> None:
        source = Path(processor_module.__file__).read_text(encoding="utf-8")
        tree = ast.parse(source)
        processor = next(node for node in tree.body if isinstance(node, ast.ClassDef) and node.name == "TaskProcessor")
        missing_docstrings = [
            f"TaskProcessor.{member.name}:{member.lineno}"
            for member in processor.body
            if isinstance(member, (ast.FunctionDef, ast.AsyncFunctionDef))
            and member.lineno < 1900
            and ast.get_docstring(member) is None
        ]
        self.assertFalse(missing_docstrings, f"TaskProcessor core helper docstrings missing: {missing_docstrings}")
        self.assertIn("draggable workflow order", processor_module.TaskProcessor._workflow_plan_from_options.__doc__ or "")
        self.assertIn("dry-run execution summary", processor_module.TaskProcessor._local_sync_dry_run_execution.__doc__ or "")
        self.assertIn("Windows/macOS client heartbeats", processor_module.TaskProcessor._preflight_local_client_platform.__doc__ or "")
        self.assertIn("Mathpix upload authorization", processor_module.TaskProcessor._preflight_mathpix.__doc__ or "")
        self.assertIn("same-platform delivery", processor_module.TaskProcessor._preflight_mathtype_compatibility.__doc__ or "")

    def test_task_processor_prd_evidence_helpers_are_documented(self) -> None:
        source = Path(processor_module.__file__).read_text(encoding="utf-8")
        tree = ast.parse(source)
        processor = next(node for node in tree.body if isinstance(node, ast.ClassDef) and node.name == "TaskProcessor")
        methods = {node.name: node for node in processor.body if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef))}
        required = {
            "_preflight_item",
            "_dictionary_object",
            "_api_endpoint",
            "_acceptance_group",
            "_acceptance_item",
            "_acceptance_verification_context",
            "_acceptance_current",
            "_task_presence_current",
            "_conversion_capability_probes",
            "_probe_word_to_ppt_conversion",
            "_probe_ppt_to_word_conversion",
            "_conversion_probe_failure",
            "_conversion_probe_current",
            "_word_object_preservation_probe",
            "_macro_batch_sequence_probe",
            "_macro_ordered_execution_probe",
            "_latex_mathtype_probe",
            "_mathtype_preservation_probe",
            "_omml_mathtype_handoff_probe",
            "_local_api_security_probe",
            "_local_result_upload_probe",
            "_local_file_action_execution_probe",
            "_formula_format_scope_probe",
            "_macro_detection_probe",
            "_ppt_formula_probe",
            "_omml_dependency_probe",
            "_omml_dependency_probe_current",
            "_mathpix_contract_probe",
            "_pdf_ocr_contract_probe",
            "_pdf_retention_probe",
            "_pdf_formula_handoff_probe",
            "_mathpix_contract_probe_current",
        }
        missing_docstrings = [name for name in sorted(required) if ast.get_docstring(methods[name]) is None]
        self.assertFalse(missing_docstrings, f"TaskProcessor PRD evidence helper docstrings missing: {missing_docstrings}")
        self.assertIn("PRD acceptance group", processor_module.TaskProcessor._acceptance_group.__doc__ or "")
        self.assertIn("without overstating local or external coverage", processor_module.TaskProcessor._acceptance_verification_context.__doc__ or "")
        self.assertIn("without external upload", processor_module.TaskProcessor._mathpix_contract_probe.__doc__ or "")
        self.assertIn("blocked upload risk", processor_module.TaskProcessor._pdf_ocr_contract_probe.__doc__ or "")
        self.assertIn("OMML dependency", processor_module.TaskProcessor._local_file_action_execution_probe.__doc__ or "")

    def test_task_processor_install_mathpix_helpers_are_documented(self) -> None:
        source = Path(processor_module.__file__).read_text(encoding="utf-8")
        tree = ast.parse(source)
        processor = next(node for node in tree.body if isinstance(node, ast.ClassDef) and node.name == "TaskProcessor")
        methods = {node.name: node for node in processor.body if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef))}
        required = {
            "_install_formula_compatibility_contract",
            "_heartbeat_platform",
            "_installer_heartbeat_readiness",
            "_safe_local_client_preflight",
            "_mathpix_queue_item",
            "_mathpix_request_summary",
            "_mathpix_status_label",
            "_local_upload_item",
            "_local_upload_outputs",
            "_receive_local_upload_files",
            "_unique_cloud_upload_path",
            "_local_upload_manifest_files",
            "_local_upload_manifest_outputs",
            "_safe_sha256",
            "_safe_output_name",
            "_local_launch_origin",
            "_product_capability",
            "_product_phase",
            "_enhancement_item",
            "_installer_specs",
            "_installer_spec",
            "_installer_steps",
            "_sha256",
        }
        missing_docstrings = [name for name in sorted(required) if ast.get_docstring(methods[name]) is None]
        self.assertFalse(missing_docstrings, f"TaskProcessor install/Mathpix helper docstrings missing: {missing_docstrings}")
        self.assertIn("Windows/macOS boundaries", processor_module.TaskProcessor._install_formula_compatibility_contract.__doc__ or "")
        self.assertIn("heartbeat platform", processor_module.TaskProcessor._installer_heartbeat_readiness.__doc__ or "")
        self.assertIn("without credentials or paths", processor_module.TaskProcessor._mathpix_request_summary.__doc__ or "")
        self.assertIn("managed runtime storage", processor_module.TaskProcessor._receive_local_upload_files.__doc__ or "")
        self.assertIn("Windows and macOS installer specifications", processor_module.TaskProcessor._installer_specs.__doc__ or "")

    def test_task_processor_all_helpers_are_documented(self) -> None:
        source = Path(processor_module.__file__).read_text(encoding="utf-8")
        tree = ast.parse(source)
        processor = next(node for node in tree.body if isinstance(node, ast.ClassDef) and node.name == "TaskProcessor")
        missing_docstrings = [
            f"TaskProcessor.{member.name}:{member.lineno}"
            for member in processor.body
            if isinstance(member, (ast.FunctionDef, ast.AsyncFunctionDef))
            and ast.get_docstring(member) is None
        ]
        self.assertFalse(missing_docstrings, f"TaskProcessor helper docstrings missing: {missing_docstrings}")
        self.assertIn("without performing native document actions", processor_module.TaskProcessor._local_desktop_execution_plan.__doc__ or "")
        self.assertIn("without executing Word macros", processor_module.TaskProcessor._macro_items.__doc__ or "")
        self.assertIn("Mathpix PDF-to-Word planning", processor_module.TaskProcessor._ocr_settings_snapshot.__doc__ or "")
        self.assertIn("PDF image placeholder rows", processor_module.TaskProcessor._pdf_image_placeholders.__doc__ or "")

    def test_python_codebase_definitions_are_documented(self) -> None:
        root = Path(__file__).resolve().parents[1] / "k12"
        missing_docstrings: list[str] = []
        for module_path in sorted(root.glob("*.py")):
            tree = ast.parse(module_path.read_text(encoding="utf-8"))
            for node in ast.walk(tree):
                if isinstance(node, (ast.ClassDef, ast.FunctionDef, ast.AsyncFunctionDef)) and ast.get_docstring(node) is None:
                    missing_docstrings.append(f"{module_path.name}:{node.lineno}:{node.name}")
        self.assertFalse(missing_docstrings, f"Python definition docstrings missing: {missing_docstrings}")

    def test_task_generates_report_and_logs(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            file_item = processor.create_file({"file_name": "formula-macro.docm", "file_size": 4096})
            task = processor.create_task(
                {
                    "task_type": "macro_sequence",
                    "file_ids": [file_item["id"]],
                    "options": {"failureStrategy": "跳过"},
                }
            )
            reports = store.list_reports()
            logs = store.list_logs(task["id"])
            self.assertEqual(task["status"], "成功")
            self.assertEqual(len(reports), 1)
            self.assertGreaterEqual(reports[0]["macro_count"], 1)
            self.assertTrue(any("生成报告" in log["message"] for log in logs))
            self.assertIn("duration_seconds", task)
            self.assertIn("duration_label", task)
            self.assertGreaterEqual(task["duration_seconds"], 0)
            self.assertTrue(task["duration_label"].endswith("秒"))

    def test_task_duration_is_enriched_for_task_center(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            task = Task(task_type="word_to_ppt", execute_mode="local", file_ids=["file_demo"]).to_dict()
            task["start_time"] = "2026-06-25T00:00:00+00:00"
            task["end_time"] = "2026-06-25T01:02:03+00:00"

            saved = store.save_task(task)
            listed = store.list_tasks()[0]
            loaded = store.get_task(task["id"])

            self.assertEqual(saved["duration_seconds"], 3723)
            self.assertEqual(saved["duration_label"], "1小时2分3秒")
            self.assertEqual(listed["duration_seconds"], 3723)
            self.assertEqual(loaded["duration_label"], "1小时2分3秒")

            legacy_task = Task(task_type="word_to_ppt", execute_mode="local", file_ids=["file_legacy"]).to_dict()
            legacy_task["start_time"] = "2026-06-25T00:00:00"
            legacy_task["end_time"] = "2026-06-25T00:00:02Z"
            legacy_saved = store.save_task(legacy_task)
            self.assertEqual(legacy_saved["duration_seconds"], 2)
            self.assertEqual(legacy_saved["duration_label"], "2秒")

    def test_log_category_settings_filter_business_logs(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            store.update_settings({"logUploadEvents": False, "logConversionEvents": False})
            processor = TaskProcessor(store)
            word = processor.create_uploaded_file("lesson.docx", make_docx_bytes())[0]
            processor.create_task({"task_type": "word_to_ppt", "file_ids": [word["id"]]})
            logs = store.list_logs(limit=1000)
            categories = {log.get("category") for log in logs}

            self.assertNotIn("upload", categories)
            self.assertNotIn("conversion", categories)
            self.assertIn("system", categories)

    def test_log_categories_are_recorded_when_enabled(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            word = processor.create_uploaded_file("lesson.docx", make_docx_bytes())[0]
            processor.create_task({"task_type": "word_to_ppt", "file_ids": [word["id"]]})
            logs = store.list_logs(limit=1000)
            categories = {log.get("category") for log in logs}

            self.assertIn("upload", categories)
            self.assertIn("conversion", categories)

    def test_image_log_category_can_be_disabled(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            store.update_settings({"logImageEvents": False})
            processor = TaskProcessor(store)
            word = processor.create_uploaded_file("with-image.docx", make_docx_bytes({"word/media/icon.png": make_png_bytes(32, 24)}))[0]
            processor.create_task({"task_type": "small_image_scan", "file_ids": [word["id"]]})
            categories = {log.get("category") for log in store.list_logs(limit=1000)}
            self.assertNotIn("image", categories)

        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            word = processor.create_uploaded_file("with-image.docx", make_docx_bytes({"word/media/icon.png": make_png_bytes(32, 24)}))[0]
            processor.create_task({"task_type": "small_image_scan", "file_ids": [word["id"]]})
            categories = {log.get("category") for log in store.list_logs(limit=1000)}
            self.assertIn("image", categories)

    def test_error_logs_have_category_and_can_be_disabled(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            store.append_log("task_demo", "输出文件打不开", "error")
            logs = store.list_logs("task_demo")
            self.assertEqual(logs[0]["category"], "error")
            self.assertEqual(logs[0]["level"], "error")

        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            store.update_settings({"logErrorEvents": False})
            store.append_log("task_demo", "输出文件打不开", "error")
            self.assertEqual(store.list_logs("task_demo"), [])

    def test_cleanup_runtime_history_preserves_uploaded_sources(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            word = processor.create_uploaded_file("cleanup.docx", make_docx_bytes())[0]
            task = processor.create_task({"task_type": "word_to_ppt", "file_ids": [word["id"]]})
            report = store.list_reports()[0]
            report_paths = [Path(report[key]) for key in ("html_path", "json_path", "pdf_path", "xlsx_path", "txt_path", "failure_csv_path")]
            output_dir = Path(task["output_path"])
            legacy_output_dir = store.outputs_dir / task["id"]
            backup_dir = store.backups_dir / task["id"]
            image_dir = store.images_dir / task["id"]
            for directory in (legacy_output_dir, backup_dir, image_dir):
                directory.mkdir(parents=True, exist_ok=True)
                (directory / "marker.txt").write_text("runtime cache", encoding="utf-8")

            result = processor.cleanup_runtime_history({"retentionDays": 0, "logRetentionDays": 0})

            self.assertEqual(result["tasks_deleted"], 1)
            self.assertEqual(result["reports_deleted"], 1)
            self.assertGreaterEqual(result["logs_deleted"], 1)
            self.assertEqual(len(store.list_tasks()), 0)
            self.assertEqual(len(store.list_reports()), 0)
            self.assertEqual(len(store.list_files()), 1)
            self.assertTrue(Path(word["storage_path"]).exists())
            self.assertTrue(all(not path.exists() for path in report_paths))
            self.assertFalse(output_dir.exists())
            self.assertFalse(legacy_output_dir.exists())
            self.assertFalse(backup_dir.exists())
            self.assertFalse(image_dir.exists())

    def test_delete_report_removes_managed_small_image_cache_only(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            managed_image = store.images_dir / "task_image_cleanup" / "small.png"
            external_image = Path(tmp) / "external-small.png"
            managed_image.parent.mkdir(parents=True, exist_ok=True)
            managed_image.write_bytes(make_png_bytes(16, 16))
            external_image.write_bytes(make_png_bytes(16, 16))
            store.save_report(
                {
                    "id": "report_image_cleanup",
                    "task_id": "task_image_cleanup",
                    "created_at": "2026-01-01T00:00:00+00:00",
                    "analysis": {
                        "smallImages": [
                            {"image_path": str(managed_image)},
                            {"image_path": str(external_image)},
                        ]
                    },
                }
            )

            result = store.delete_report("report_image_cleanup")

            self.assertTrue(result["deleted"])
            self.assertEqual(result["image_files_deleted"], 1)
            self.assertFalse(managed_image.exists())
            self.assertTrue(external_image.exists())

    def test_disabling_history_removes_completed_task_report_logs_and_outputs(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            store.update_settings({"saveHistory": False})
            processor = TaskProcessor(store)
            word = processor.create_uploaded_file("private.docx", make_docx_bytes())[0]
            task = processor.create_task({"task_type": "word_to_ppt", "file_ids": [word["id"]]})

            self.assertEqual(task["status"], "成功")
            self.assertFalse(task["history_saved"])
            self.assertEqual(store.list_tasks(), [])
            self.assertEqual(store.list_reports(), [])
            self.assertEqual(store.list_logs(limit=1000), [])
            self.assertFalse(Path(task["output_path"]).exists())
            self.assertIsNotNone(store.get_file(word["id"]))
            self.assertTrue(Path(word["storage_path"]).exists())

    def test_auto_cleanup_respects_interval(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            store.update_settings(
                {
                    "autoCleanTemp": True,
                    "cleanupRetentionDays": 0,
                    "logRetentionDays": 0,
                    "cleanupIntervalDays": 7,
                }
            )
            first_file = processor.create_uploaded_file("first.docx", make_docx_bytes())[0]
            first_task = processor.create_task({"task_type": "word_to_ppt", "file_ids": [first_file["id"]]})

            first_result = processor.auto_cleanup_runtime_history()

            self.assertTrue(first_result["auto_cleanup"])
            self.assertEqual(store.get_task(first_task["id"]), None)
            self.assertTrue(store.get_settings()["lastCleanupAt"])
            second_file = processor.create_uploaded_file("second.docx", make_docx_bytes())[0]
            second_task = processor.create_task({"task_type": "word_to_ppt", "file_ids": [second_file["id"]]})

            second_result = processor.auto_cleanup_runtime_history()

            self.assertTrue(second_result["skipped"])
            self.assertEqual(second_result["reason"], "未到自动清理周期")
            self.assertIsNotNone(store.get_task(second_task["id"]))

    def test_macro_sequence_uses_selection_order_and_backup(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            word = processor.create_uploaded_file("macro.docm", make_docx_bytes())[0]
            task = processor.create_task(
                {
                    "task_type": "macro_sequence",
                    "file_ids": [word["id"]],
                    "options": {
                        "selectedMacros": [
                            {"id": "macro_repair_equation_anchors", "execute_order": 1},
                            {"id": "macro_clean_empty_paragraphs", "execute_order": 2},
                        ],
                        "confirmMacroRisk": True,
                        "macroBackup": True,
                        "macroTimeoutSeconds": 45,
                        "failureStrategy": "停止",
                    },
                }
            )
            report = store.list_reports()[0]
            macros = report["analysis"]["macros"]
            self.assertEqual(task["status"], "成功")
            self.assertEqual([macro["id"] for macro in macros], ["macro_repair_equation_anchors", "macro_clean_empty_paragraphs"])
            self.assertEqual(macros[0]["execute_status"], "待本地客户端执行")
            self.assertEqual(macros[0]["failure_strategy"], "停止")
            self.assertEqual(macros[0]["failure_policy"]["action"], "stop_sequence")
            self.assertFalse(macros[0]["failure_policy"]["continue_on_failure"])
            self.assertEqual(macros[0]["timeout_seconds"], 45)
            self.assertTrue(macros[0]["backup_path"])
            self.assertGreaterEqual(report["macro_queued_count"], 2)

            report["analysis"]["macros"][0]["execute_status"] = "失败"
            report["analysis"]["macros"][0]["error_message"] = "本地客户端回传宏执行失败"
            store.save_report(report)
            handler = make_handler(store, processor)
            K12RequestHandler._send_report_macro_failures(handler, report["id"])
            csv_text = handler.wfile.getvalue().decode("utf-8")
            self.assertEqual(handler.status, 200)
            self.assertIn(("Content-Type", "text/csv; charset=utf-8"), handler.output_headers)
            self.assertIn(("Content-Disposition", f'attachment; filename="{report["id"]}-macro-failures.csv"'), handler.output_headers)
            self.assertIn("macro_repair_equation_anchors", csv_text)
            self.assertIn("本地客户端回传宏执行失败", csv_text)
            self.assertIn("本地路径已隐藏", csv_text)
            self.assertNotIn(str(Path(tmp)), csv_text)

            route_handler = make_handler(store, processor)
            K12RequestHandler._handle_api_get(route_handler, f"/api/reports/{report['id']}/macro-failures.csv", {})
            self.assertEqual(route_handler.status, 200)
            self.assertIn("macro_repair_equation_anchors", route_handler.wfile.getvalue().decode("utf-8"))

    def test_macro_ordered_execution_probe_builds_local_handoff_dry_run(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            processor = TaskProcessor(AppStore(tmp))

            probe = processor._macro_ordered_execution_probe()

            self.assertTrue(probe["available"])
            self.assertEqual(probe["schema_version"], "k12.macroOrderedExecutionProbe.v1")
            self.assertEqual(probe["macro_count"], 2)
            self.assertEqual(probe["plan_status"], "ready_for_native_client")
            self.assertEqual(probe["dry_run_status"], "ready_for_native_executor")
            self.assertFalse(probe["native_execution_performed"])
            self.assertIn("macro.run_ordered", probe["evidence"])
            self.assertIn("真实 VBA 执行仍需本地客户端", probe["current"])

    def test_macro_failure_strategies_map_to_local_client_policy(self) -> None:
        cases = {
            "停止": ("stop_sequence", False, False, False),
            "跳过": ("skip_failed_macro", True, True, False),
            "继续": ("continue_sequence", True, False, False),
            "询问": ("pause_for_user_decision", False, False, True),
        }
        for strategy, expected in cases.items():
            with self.subTest(strategy=strategy):
                with tempfile.TemporaryDirectory() as tmp:
                    store = AppStore(tmp)
                    processor = TaskProcessor(store)
                    word = processor.create_uploaded_file("macro.docm", make_docx_bytes())[0]
                    task = processor.create_task(
                        {
                            "task_type": "macro_sequence",
                            "file_ids": [word["id"]],
                            "options": {
                                "selectedMacros": [{"id": "macro_clean_empty_paragraphs", "execute_order": 1}],
                                "confirmMacroRisk": True,
                                "macroBackup": False,
                                "failureStrategy": strategy,
                            },
                        }
                    )
                    report = store.list_reports()[0]
                    macro = report["analysis"]["macros"][0]
                    payload = processor.local_task_payload(task["id"])
                    action = next(item for item in payload["local_actions"] if item["type"] == "macro_sequence")

                    self.assertEqual(macro["failure_strategy"], strategy)
                    self.assertEqual(macro["failure_policy"]["strategy"], strategy)
                    self.assertEqual(macro["failure_policy"]["action"], expected[0])
                    self.assertEqual(macro["failure_policy"]["continue_on_failure"], expected[1])
                    self.assertEqual(macro["failure_policy"]["skip_failed_macro"], expected[2])
                    self.assertEqual(macro["failure_policy"]["requires_user_decision"], expected[3])
                    self.assertEqual(action["macros"][0]["failure_policy"]["action"], expected[0])
                    self.assertIn("失败动作", Path(report["txt_path"]).read_text(encoding="utf-8"))
                    self.assertIn("失败动作", Path(report["html_path"]).read_text(encoding="utf-8"))
                    with zipfile.ZipFile(report["xlsx_path"]) as archive:
                        sheet_xml = archive.read("xl/worksheets/sheet1.xml").decode("utf-8")
                        self.assertIn("失败动作", sheet_xml)

    def test_macro_execution_can_be_disabled(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            store.update_settings({"enableMacroExecution": False})
            processor = TaskProcessor(store)
            word = processor.create_uploaded_file("macro.docm", make_docx_bytes())[0]
            processor.create_task(
                {
                    "task_type": "macro_sequence",
                    "file_ids": [word["id"]],
                    "options": {
                        "workflowOrder": ["pdf_to_word", "macro_sequence", "small_image_scan"],
                        "selectedMacros": [{"id": "macro_clean_empty_paragraphs", "execute_order": 1}],
                        "confirmMacroRisk": True,
                        "macroBackup": True,
                    },
                }
            )
            report = store.list_reports()[0]
            macro = report["analysis"]["macros"][0]
            self.assertEqual(macro["execute_status"], "已禁用")
            self.assertIn("禁用", macro["error_message"])
            self.assertEqual(macro["backup_path"], "")
            self.assertEqual(report["macro_queued_count"], 0)

    def test_macro_whitelist_blocks_unlisted_macros(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            store.update_settings(
                {
                    "macroWhitelistOnly": True,
                    "macroWhitelist": "macro_clean_empty_paragraphs",
                }
            )
            processor = TaskProcessor(store)
            word = processor.create_uploaded_file("macro.docm", make_docx_bytes())[0]
            processor.create_task(
                {
                    "task_type": "macro_sequence",
                    "file_ids": [word["id"]],
                    "options": {
                        "selectedMacros": [
                            {"id": "macro_repair_equation_anchors", "execute_order": 1},
                            {"id": "macro_clean_empty_paragraphs", "execute_order": 2},
                        ],
                        "confirmMacroRisk": True,
                        "macroBackup": True,
                    },
                }
            )
            report = store.list_reports()[0]
            macros = {macro["id"]: macro for macro in report["analysis"]["macros"]}
            self.assertEqual(macros["macro_repair_equation_anchors"]["execute_status"], "未授权")
            self.assertEqual(macros["macro_repair_equation_anchors"]["backup_path"], "")
            self.assertEqual(macros["macro_clean_empty_paragraphs"]["execute_status"], "待本地客户端执行")
            self.assertTrue(macros["macro_clean_empty_paragraphs"]["backup_path"])
            self.assertEqual(report["macro_queued_count"], 1)

    def test_macro_source_permissions_block_local_macro_library(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            store.update_settings({"allowLocalMacroLibrary": False})
            processor = TaskProcessor(store)
            word = processor.create_uploaded_file("macro.docm", make_docx_bytes())[0]
            processor.create_task(
                {
                    "task_type": "macro_sequence",
                    "file_ids": [word["id"]],
                    "options": {
                        "selectedMacros": [
                            {"id": "macro_repair_equation_anchors", "execute_order": 1},
                            {"id": "macro_clean_empty_paragraphs", "execute_order": 2},
                        ],
                        "confirmMacroRisk": True,
                        "macroBackup": True,
                    },
                }
            )
            report = store.list_reports()[0]
            macros = {macro["id"]: macro for macro in report["analysis"]["macros"]}
            self.assertEqual(macros["macro_repair_equation_anchors"]["execute_status"], "未授权")
            self.assertIn("本地宏库", macros["macro_repair_equation_anchors"]["error_message"])
            self.assertEqual(macros["macro_clean_empty_paragraphs"]["execute_status"], "待本地客户端执行")

    def test_macro_batch_execution_can_be_disabled(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            store.update_settings({"allowBatchMacroExecution": False})
            processor = TaskProcessor(store)
            first = processor.create_uploaded_file("first.docm", make_docx_bytes())[0]
            second = processor.create_uploaded_file("second.docm", make_docx_bytes())[0]
            processor.create_task(
                {
                    "task_type": "macro_sequence",
                    "file_ids": [first["id"], second["id"]],
                    "options": {
                        "selectedMacros": [{"id": "macro_clean_empty_paragraphs", "execute_order": 1}],
                        "confirmMacroRisk": True,
                        "macroBackup": True,
                    },
                }
            )
            report = store.list_reports()[0]
            macros = report["analysis"]["macros"]
            first_macro = next(macro for macro in macros if macro["file_id"] == first["id"])
            second_macro = next(macro for macro in macros if macro["file_id"] == second["id"])

            self.assertEqual(first_macro["execute_status"], "待本地客户端执行")
            self.assertEqual(second_macro["execute_status"], "未授权")
            self.assertIn("不允许批量执行宏", second_macro["error_message"])
            self.assertEqual(second_macro["backup_path"], "")

    def test_macro_detection_can_be_disabled_for_non_macro_tasks(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            store.update_settings({"enableMacroDetection": False})
            processor = TaskProcessor(store)
            word = processor.create_uploaded_file("macro.docm", make_docx_bytes())[0]
            processor.create_task({"task_type": "word_to_ppt", "file_ids": [word["id"]]})
            report = store.list_reports()[0]
            self.assertEqual(report["analysis"]["macros"], [])

    def test_macro_library_includes_purpose_and_recent_usage(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            initial = {macro["id"]: macro for macro in processor.macro_library()}
            self.assertEqual(initial["macro_repair_equation_anchors"]["macro_purpose"], "公式修复")
            self.assertFalse(initial["macro_repair_equation_anchors"]["recently_used"])

            word = processor.create_uploaded_file("macro.docm", make_docx_bytes())[0]
            processor.create_task(
                {
                    "task_type": "macro_sequence",
                    "file_ids": [word["id"]],
                    "options": {
                        "selectedMacros": [{"id": "macro_repair_equation_anchors", "execute_order": 1}],
                        "confirmMacroRisk": True,
                    },
                }
            )

            updated = {macro["id"]: macro for macro in processor.macro_library()}
            macro = updated["macro_repair_equation_anchors"]
            self.assertTrue(macro["recently_used"])
            self.assertEqual(macro["usage_count"], 1)
            self.assertTrue(macro["last_used_at"])

    def test_formula_settings_control_confidence_and_output_preferences(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            store.update_settings(
                {
                    "formulaConfidenceThreshold": 95,
                    "formulaOutputFormat": "MathML",
                    "formulaFont": "STIX Two Math",
                    "formulaFontSize": 14,
                    "formulaFormatScope": "当前章节",
                    "formulaAlignment": "右对齐",
                    "formulaVariableStyle": "正体",
                    "formulaFunctionStyle": "斜体",
                    "formulaScriptScale": 65,
                    "formulaFractionStyle": "展示型",
                    "formulaRadicalStyle": "加长根号",
                    "formulaMatrixSpacing": "宽松",
                    "formulaGreekStyle": "斜体",
                    "formulaInlineBaseline": "数学轴对齐",
                    "formulaDisplaySpacing": "宽松",
                    "formulaNumbering": "连续编号",
                    "lowConfidenceFormulaStrategy": "标记重识别",
                    "keepFormulaImages": False,
                }
            )
            processor = TaskProcessor(store)
            word = processor.create_file({"file_name": "formula.docx", "file_size": 4096})
            processor.create_task({"task_type": "formula_precheck", "file_ids": [word["id"]]})
            report = store.list_reports()[0]
            formulas = report["analysis"]["formulas"]
            quality = {item["id"]: item for item in report["qualityChecks"]}
            html_report = Path(report["html_path"]).read_text(encoding="utf-8")
            text_report = Path(report["txt_path"]).read_text(encoding="utf-8")

            self.assertTrue(formulas)
            for formula in formulas:
                expected_status = "成功" if formula["confidence"] >= 95 else "待确认"
                self.assertEqual(formula["status"], expected_status)
                self.assertEqual(formula["output_format"], "MathML")
                self.assertEqual(formula["font"], "STIX Two Math")
                self.assertEqual(formula["font_size"], 14)
                self.assertEqual(formula["format_scope"], "当前章节")
                self.assertEqual(formula["alignment"], "右对齐")
                self.assertEqual(formula["variable_style"], "正体")
                self.assertEqual(formula["function_style"], "斜体")
                self.assertEqual(formula["script_scale"], 65)
                self.assertEqual(formula["fraction_style"], "展示型")
                self.assertEqual(formula["radical_style"], "加长根号")
                self.assertEqual(formula["matrix_spacing"], "宽松")
                self.assertEqual(formula["greek_style"], "斜体")
                self.assertEqual(formula["inline_baseline"], "数学轴对齐")
                self.assertEqual(formula["display_spacing"], "宽松")
                self.assertEqual(formula["numbering"], "连续编号")
                self.assertEqual(formula["low_confidence_strategy"], "标记重识别")
                self.assertFalse(formula["keep_original_image"])
                self.assertTrue(formula["original_image_ref"].startswith(f"source://{word['id']}/formula/"))
                self.assertTrue(formula["mathml"].startswith("<math>"))
                self.assertIn(formula["latex"], formula["mathml"])
                self.assertIn("MathType 预览", formula["mathtype_preview"])
                comparison = formula["format_comparison"]
                self.assertEqual(comparison["before"]["format"], formula["source_type"])
                self.assertEqual(comparison["after"]["format"], "MathML")
                self.assertEqual(comparison["after"]["font"], "STIX Two Math")
                self.assertEqual(comparison["after"]["scope"], "当前章节")
                self.assertIn("输出格式", comparison["summary"])
            self.assertIn("阈值 95", quality["formula_confidence"]["metric"])
            self.assertIn("范围 当前章节", text_report)
            self.assertIn("上下标 65%", text_report)
            self.assertIn("矩阵 宽松", text_report)
            self.assertIn("原始截图 source://", text_report)
            self.assertIn("MathType MathType 预览", text_report)
            self.assertIn("格式化对比 输出格式", text_report)
            self.assertIn("<th>格式化对比</th>", html_report)
            with zipfile.ZipFile(report["xlsx_path"]) as archive:
                sheet_xml = archive.read("xl/worksheets/sheet1.xml").decode("utf-8")
                self.assertIn("样式", sheet_xml)
                self.assertIn("格式化对比", sheet_xml)
                self.assertIn("原始截图", sheet_xml)
                self.assertIn("MathType 预览", sheet_xml)
                self.assertIn("希腊 斜体", sheet_xml)

    def test_image_formula_failure_preserves_original_image(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            word = processor.create_uploaded_file(
                "image-formula.docx",
                make_docx_bytes(
                    extra_files={"word/media/formula-equation.png": make_png_bytes(80, 30)},
                    document_xml="<w:document><w:p><w:t>仅图片公式</w:t></w:p></w:document>",
                ),
            )[0]
            self.assertFalse(word["has_formula"])
            self.assertTrue(word["has_small_image"])

            processor.create_task({"task_type": "formula_precheck", "file_ids": [word["id"]]})
            report = store.list_reports()[0]
            formula = report["analysis"]["formulas"][0]
            small_image = report["analysis"]["smallImages"][0]
            quality = {item["id"]: item for item in report["qualityChecks"]}

            self.assertTrue(small_image["is_formula_like"])
            self.assertEqual(formula["source_type"], "图片公式")
            self.assertEqual(formula["status"], "待确认")
            self.assertEqual(formula["latex"], "")
            self.assertEqual(formula["format_status"], "识别失败")
            self.assertTrue(formula["keep_original_image"])
            self.assertEqual(formula["fallback_action"], "保留原图")
            self.assertEqual(formula["source_image_id"], small_image["id"])
            self.assertTrue(formula["original_image_ref"].startswith("/api/assets/"))
            self.assertIn("图片公式识别失败", formula["mathtype_preview"])
            self.assertIn("已保留原图", formula["recognition_error"])
            self.assertEqual(report["formula_count"], 1)
            self.assertEqual(quality["formula_confidence"]["status"], "需确认")
            self.assertIn("图片公式", Path(report["txt_path"]).read_text(encoding="utf-8"))
            self.assertIn("图片公式识别失败", Path(report["html_path"]).read_text(encoding="utf-8"))
            with zipfile.ZipFile(report["xlsx_path"]) as archive:
                sheet_xml = archive.read("xl/worksheets/sheet1.xml").decode("utf-8")
                self.assertIn("图片公式", sheet_xml)
                self.assertIn("图片公式识别失败", sheet_xml)

    def test_mathtype_format_disabled_preserves_original_formulas(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            store.update_settings({"enableMathTypeFormatting": False})
            processor = TaskProcessor(store)
            word = processor.create_file({"file_name": "formula.docx", "file_size": 4096})
            processor.create_task({"task_type": "mathtype_format", "file_ids": [word["id"]]})
            report = store.list_reports()[0]
            formulas = report["analysis"]["formulas"]
            self.assertTrue(formulas)
            self.assertEqual(report["formatted_formula_count"], 0)

            for formula in formulas:
                policy = formula["format_failure_policy"]
                self.assertFalse(formula["formatting_enabled"])
                self.assertEqual(formula["format_status"], "已跳过")
                self.assertTrue(formula["preserve_original_formula"])
                self.assertEqual(policy["status"], "disabled")
                self.assertTrue(policy["preserve_original_formula"])
                self.assertFalse(policy["retryable"])
                self.assertIn("保留原公式", policy["message"])
                self.assertEqual(policy["fallback_format"], formula["source_type"])

            text_report = Path(report["txt_path"]).read_text(encoding="utf-8")
            html_report = Path(report["html_path"]).read_text(encoding="utf-8")
            self.assertIn("失败兜底", text_report)
            self.assertIn("保留原公式", text_report)
            self.assertIn("<th>失败兜底</th>", html_report)
            with zipfile.ZipFile(report["xlsx_path"]) as archive:
                sheet_xml = archive.read("xl/worksheets/sheet1.xml").decode("utf-8")
                self.assertIn("失败兜底", sheet_xml)
                self.assertIn("保留原公式", sheet_xml)

    def test_mathtype_format_task_option_overrides_format_scope(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            store.update_settings({"formulaFormatScope": "全文"})
            processor = TaskProcessor(store)
            word = processor.create_file({"file_name": "formula.docx", "file_size": 4096})
            processor.create_task(
                {
                    "task_type": "mathtype_format",
                    "file_ids": [word["id"]],
                    "options": {"formatScope": "选中区域"},
                }
            )
            report = store.list_reports()[0]
            formulas = report["analysis"]["formulas"]

            self.assertTrue(formulas)
            for formula in formulas:
                self.assertEqual(formula["format_scope"], "选中区域")
                self.assertEqual(formula["format_comparison"]["after"]["scope"], "选中区域")
                self.assertTrue(formula["format_failure_policy"]["preserve_original_formula"])

    def test_macro_backup_restore_copies_backup_to_original_file(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            original_bytes = make_docx_bytes()
            word = processor.create_uploaded_file("macro.docm", original_bytes)[0]
            task = processor.create_task(
                {
                    "task_type": "macro_sequence",
                    "file_ids": [word["id"]],
                    "options": {
                        "workflowOrder": ["pdf_to_word", "macro_sequence", "small_image_scan"],
                        "selectedMacros": [{"id": "macro_clean_empty_paragraphs", "execute_order": 1}],
                        "confirmMacroRisk": True,
                        "macroBackup": True,
                    },
                }
            )
            storage_path = Path(word["storage_path"])
            storage_path.write_bytes(b"changed by macro")
            result = processor.restore_task_backups(task["id"])
            self.assertEqual(result["restored_count"], 1)
            self.assertEqual(result["failed_count"], 0)
            self.assertEqual(storage_path.read_bytes(), original_bytes)
            reloaded_task = store.get_task(task["id"])
            self.assertEqual(reloaded_task["backup_restore_results"][0]["status"], "成功")
            self.assertTrue(any("恢复宏备份" in log["message"] for log in store.list_logs(task["id"])))

    def test_macro_backup_restore_rejects_backup_outside_backup_dir(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            word = processor.create_uploaded_file("macro.docm", make_docx_bytes())[0]
            task = processor.create_task(
                {
                    "task_type": "macro_sequence",
                    "file_ids": [word["id"]],
                    "options": {
                        "selectedMacros": [{"id": "macro_clean_empty_paragraphs", "execute_order": 1}],
                        "confirmMacroRisk": True,
                        "macroBackup": True,
                    },
                }
            )
            report = store.list_reports()[0]
            report["analysis"]["macros"][0]["backup_path"] = str(Path(tmp) / "outside.bak")
            Path(report["analysis"]["macros"][0]["backup_path"]).write_bytes(b"outside")
            store.save_report(report)
            result = processor.restore_task_backups(task["id"])
            self.assertEqual(result["restored_count"], 0)
            self.assertEqual(result["failed_count"], 1)
            self.assertEqual(result["results"][0]["status"], "失败")

    def test_macro_template_save_list_and_delete(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            template = processor.save_macro_template(
                {
                    "name": "教案转换前清理",
                    "macro_sequence": [
                        {"id": "macro_repair_equation_anchors", "execute_order": 2},
                        {"id": "macro_clean_empty_paragraphs", "execute_order": 1},
                        {"id": "missing_macro", "execute_order": 3},
                    ],
                    "failureStrategy": "停止",
                    "executeTiming": "Word 转 PPT 前",
                    "confirmMacroRisk": True,
                }
            )
            templates = processor.list_macro_templates()
            self.assertEqual(template["name"], "教案转换前清理")
            self.assertEqual(len(template["macro_sequence"]), 2)
            self.assertEqual([item["id"] for item in template["macro_sequence"]], ["macro_clean_empty_paragraphs", "macro_repair_equation_anchors"])
            self.assertEqual(template["defaults"]["failureStrategy"], "停止")
            self.assertEqual(templates[0]["id"], template["id"])
            self.assertTrue(processor.delete_macro_template(template["id"]))
            self.assertEqual(processor.list_macro_templates(), [])

    def test_macro_template_requires_valid_macro(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            with self.assertRaises(ValueError):
                processor.save_macro_template({"name": "空模板", "macro_sequence": [{"id": "missing"}]})

    def test_settings_update_ignores_unknown_keys(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            settings = store.update_settings({"maxConcurrentTasks": 4, "pdfToWordEngine": "LocalOCR", "unknown": True})
            self.assertEqual(settings["maxConcurrentTasks"], 4)
            self.assertEqual(settings["pdfToWordEngine"], "Mathpix")
            self.assertNotIn("unknown", settings)

    def test_user_roles_control_task_creation_permission(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            current = processor.current_user()
            self.assertEqual(current["role"], "管理员")
            self.assertIn("tasks.create", current["effective_permissions"])
            word = processor.create_uploaded_file("lesson.docx", make_docx_bytes())[0]
            visitor = processor.save_user({"name": "只读审阅", "role": "访客", "active": True})
            self.assertTrue(visitor["active"])
            with self.assertRaises(ValueError) as raised:
                processor.create_task({"task_type": "word_to_ppt", "file_ids": [word["id"]]})
            self.assertIn("tasks.create", str(raised.exception))

            active = processor.activate_user("user_admin")
            self.assertEqual(active["id"], "user_admin")
            task = processor.create_task({"task_type": "word_to_ppt", "file_ids": [word["id"]]})
            self.assertEqual(task["status"], "成功")

    def test_task_control_requires_tasks_control_permission(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            word = processor.create_uploaded_file("lesson.docx", make_docx_bytes())[0]
            student = processor.save_user({"name": "学生用户", "role": "学生", "active": True})
            created = processor.create_task({"task_type": "word_to_ppt", "file_ids": [word["id"]]})
            self.assertEqual(created["status"], "成功")

            task = Task(task_type="word_to_ppt", execute_mode="hybrid", file_ids=[word["id"]]).to_dict()
            task["status"] = "处理中"
            task["progress"] = 40
            store.save_task(task)
            blocked_actions = [
                lambda: processor.retry_task(task["id"]),
                lambda: processor.pause_task(task["id"]),
                lambda: processor.resume_task(task["id"]),
                lambda: processor.cancel_task(task["id"]),
                lambda: processor.skip_batch_file(task["id"], word["id"]),
                lambda: processor.restore_task_backups(task["id"]),
            ]
            for action in blocked_actions:
                with self.assertRaises(ValueError) as raised:
                    action()
                self.assertIn("tasks.control", str(raised.exception))

            self.assertEqual(processor.current_user()["id"], student["id"])
            handler = make_handler(store, processor)
            with self.assertRaises(JsonError) as raised:
                K12RequestHandler._handle_api_post(handler, f"/api/tasks/{task['id']}/cancel", {})
            self.assertEqual(raised.exception.status, 400)
            self.assertIn("tasks.control", raised.exception.message)

            processor.activate_user("user_admin")
            paused = processor.pause_task(task["id"])
            self.assertEqual(paused["status"], "已暂停")

    def test_file_and_report_management_require_permissions(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            word = processor.create_uploaded_file("lesson.docx", make_docx_bytes())[0]
            task = processor.create_task({"task_type": "word_to_ppt", "file_ids": [word["id"]]})
            report = store.list_reports()[0]
            visitor = processor.save_user({"name": "只读审阅", "role": "访客", "active": True})
            self.assertEqual(processor.current_user()["id"], visitor["id"])

            blocked_actions = [
                ("files.manage", lambda: processor.create_file({"file_name": "new.docx", "file_size": 1024})),
                ("files.manage", lambda: processor.create_uploaded_file("upload.docx", make_docx_bytes())),
                ("files.manage", lambda: processor.replace_uploaded_file(word["id"], "replace.docx", make_docx_bytes())),
                ("files.manage", lambda: processor.set_file_password(word["id"], "123456")),
                ("files.manage", lambda: processor.delete_file(word["id"])),
                ("reports.manage", lambda: processor.delete_report(report["id"])),
            ]
            for permission, action in blocked_actions:
                with self.assertRaises(ValueError) as raised:
                    action()
                self.assertIn(permission, str(raised.exception))

            handler = make_handler(store, processor)
            with self.assertRaises(JsonError) as raised:
                K12RequestHandler._handle_api_post(handler, "/api/files", {"file_name": "http.docx", "file_size": 1024})
            self.assertEqual(raised.exception.status, 400)
            self.assertIn("files.manage", raised.exception.message)

            processor.activate_user("user_admin")
            self.assertTrue(processor.delete_report(report["id"])["deleted"])
            self.assertTrue(processor.delete_file(word["id"]))
            self.assertEqual(task["status"], "成功")

    def test_file_reports_endpoint_exposes_file_report_entry(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            word = processor.create_uploaded_file("lesson.docx", make_docx_bytes())[0]
            processor.create_task({"task_type": "formula_precheck", "file_ids": [word["id"]]})
            reports = processor.reports_for_file(word["id"])
            self.assertEqual(len(reports), 1)
            self.assertEqual(reports[0]["files"][0]["id"], word["id"])

            handler = make_handler(store, processor)
            K12RequestHandler._handle_api_get(handler, f"/api/files/{word['id']}/reports", {})
            response = json.loads(handler.wfile.getvalue().decode("utf-8"))
            self.assertEqual(handler.status, 200)
            self.assertEqual(response["reports"][0]["id"], reports[0]["id"])
            self.assertIn("本地路径已隐藏", json.dumps(response, ensure_ascii=False))

            with self.assertRaises(KeyError):
                processor.reports_for_file("missing")

    def test_management_permissions_block_non_admin_writes(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            visitor = processor.save_user({"name": "只读审阅", "role": "访客", "active": True})
            self.assertTrue(visitor["active"])

            blocked_actions = [
                lambda: processor.save_user({"name": "学生账号", "role": "学生"}),
                lambda: processor.delete_user("user_admin"),
                lambda: processor.save_template({"name": "访客模板", "applies_to": "word_to_ppt"}),
                lambda: processor.delete_template("tpl_missing"),
                lambda: processor.save_macro_template({"name": "访客宏模板", "macro_sequence": [{"id": "macro_clean_empty_paragraphs"}]}),
                lambda: processor.delete_macro_template("macro_tpl_missing"),
                lambda: processor.save_authorization({"key": "mathpix_external_upload", "enabled": True}),
                lambda: processor.update_settings({"maxConcurrentTasks": 4}),
            ]
            for action in blocked_actions:
                with self.assertRaises(ValueError) as raised:
                    action()
                self.assertIn("权限", str(raised.exception))

            processor.activate_user("user_admin")
            template = processor.save_template({"name": "管理员模板", "applies_to": "word_to_ppt"})
            self.assertEqual(template["name"], "管理员模板")
            settings = processor.update_settings({"maxConcurrentTasks": 4})
            self.assertEqual(settings["maxConcurrentTasks"], 4)

            handler = make_handler(store, processor)
            processor.activate_user(visitor["id"])
            with self.assertRaises(JsonError) as raised:
                K12RequestHandler._handle_api_post(handler, "/api/authorizations", {"key": "mathpix_external_upload", "enabled": True})
            self.assertEqual(raised.exception.status, 400)
            self.assertIn("权限", raised.exception.message)

    def test_require_login_blocks_api_until_session_is_activated(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            store.update_settings({"requireLogin": True})
            processor = TaskProcessor(store)
            handler = make_handler(store, processor)

            with self.assertRaises(JsonError) as raised:
                K12RequestHandler._ensure_logged_in(handler, "/api/files", "GET")
            self.assertEqual(raised.exception.status, 401)

            K12RequestHandler._ensure_logged_in(handler, "/api/health", "GET")
            K12RequestHandler._ensure_logged_in(handler, "/api/users", "GET")
            K12RequestHandler._ensure_logged_in(handler, "/api/session", "POST")
            K12RequestHandler._handle_api_post(handler, "/api/session", {"user_id": "user_admin"})

            self.assertEqual(handler.status, 200)
            self.assertTrue(processor.current_user()["last_login_at"])
            K12RequestHandler._ensure_logged_in(handler, "/api/files", "GET")

    def test_disabled_login_user_cannot_be_activated(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            user = processor.save_user({"id": "user_disabled", "name": "禁用登录", "role": "教师", "login_enabled": False})

            self.assertFalse(user["login_enabled"])
            with self.assertRaises(ValueError) as raised:
                processor.activate_user("user_disabled")
            self.assertIn("不允许登录", str(raised.exception))

    def test_user_management_keeps_one_enabled_admin(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            with self.assertRaises(ValueError):
                processor.delete_user("user_admin")
            processor.save_user({"name": "第二管理员", "role": "管理员"})
            self.assertTrue(processor.delete_user("user_admin"))
            self.assertEqual(processor.current_user()["role"], "管理员")

    def test_authorization_updates_sync_back_to_settings(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            before = {item["key"]: item for item in processor.list_authorizations()}
            self.assertFalse(before["mathpix_external_upload"]["enabled"])
            enabled = processor.save_authorization({"key": "mathpix_external_upload", "enabled": True, "note": "单测授权"})
            self.assertTrue(enabled["enabled"])
            self.assertTrue(store.get_settings()["allowExternalMathpixUpload"])
            disabled = processor.save_authorization({"key": "mathpix_external_upload", "enabled": False})
            self.assertFalse(disabled["enabled"])
            self.assertFalse(store.get_settings()["allowExternalMathpixUpload"])

    def test_generic_template_save_list_delete_and_setting_sync(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            template = processor.save_template(
                {
                    "name": "K12 讲义模板",
                    "template_type": "PPT 模板",
                    "applies_to": "word_to_ppt",
                    "template_path": str(Path(tmp) / "template.pptx"),
                    "description": "数学课件版式",
                    "settings": {"theme": "teaching"},
                }
            )
            templates = processor.list_templates()
            self.assertEqual(template["name"], "K12 讲义模板")
            self.assertEqual(templates[0]["id"], template["id"])
            self.assertEqual(store.get_settings()["wordToPptTemplate"], "K12 讲义模板")
            self.assertTrue(processor.delete_template(template["id"]))
            self.assertEqual(processor.list_templates(), [])

    def test_ppt_to_word_template_syncs_mode_name_and_path(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            template_path = str(Path(tmp) / "handout.dotx")
            template = processor.save_template(
                {
                    "name": "K12 Word 讲义模板",
                    "template_type": "Word 模板",
                    "applies_to": "ppt_to_word",
                    "template_path": template_path,
                    "description": "PPT 转讲义",
                    "settings": {"mode": "大纲模式"},
                }
            )
            settings = store.get_settings()

            self.assertEqual(template["template_path"], template_path)
            self.assertEqual(settings["pptToWordMode"], "大纲模式")
            self.assertEqual(settings["pptToWordTemplate"], "K12 Word 讲义模板")
            self.assertEqual(settings["pptToWordTemplatePath"], template_path)

    def test_preflight_disk_failure_is_reported_as_quality_blocker(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            store.update_settings({"minFreeDiskMb": 10**12})
            processor = TaskProcessor(store)
            word = processor.create_uploaded_file("lesson.docx", make_docx_bytes())[0]
            task = processor.create_task({"task_type": "word_to_ppt", "file_ids": [word["id"]]})
            report = store.list_reports()[0]
            quality = {item["id"]: item for item in report["qualityChecks"]}
            failure_categories = {item["category"] for item in report["failureRows"]}

            self.assertEqual(task["status"], "失败")
            self.assertEqual(quality["disk_space"]["status"], "失败")
            self.assertGreaterEqual(report["quality_blocker_count"], 1)
            self.assertIn("系统预检", failure_categories)

    def test_mathpix_preflight_reports_missing_credentials_without_exposing_values(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            store.update_settings(
                {
                    "allowExternalMathpixUpload": True,
                    "mathpixAppIdEnv": "K12_TEST_MATHPIX_ID",
                    "mathpixAppKeyEnv": "K12_TEST_MATHPIX_KEY",
                }
            )
            processor = TaskProcessor(store)
            with patch.dict(os.environ, {}, clear=False):
                os.environ.pop("K12_TEST_MATHPIX_ID", None)
                os.environ.pop("K12_TEST_MATHPIX_KEY", None)
                checks = processor.preflight_checks({"task_type": "pdf_to_word"})

            mathpix = next(item for item in checks if item["id"] == "mathpix_authorization")
            self.assertEqual(mathpix["status"], "需确认")
            self.assertEqual(mathpix["metric"], "缺少凭证")
            self.assertIn("K12_TEST_MATHPIX_ID", mathpix["message"])
            self.assertIn("K12_TEST_MATHPIX_KEY", mathpix["message"])

    def test_mathpix_preflight_passes_when_credentials_are_configured_and_redacted(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            store.update_settings(
                {
                    "allowExternalMathpixUpload": True,
                    "mathpixAppIdEnv": "K12_TEST_MATHPIX_ID",
                    "mathpixAppKeyEnv": "K12_TEST_MATHPIX_KEY",
                }
            )
            processor = TaskProcessor(store)
            with patch.dict(os.environ, {"K12_TEST_MATHPIX_ID": "app-id", "K12_TEST_MATHPIX_KEY": "super-secret"}, clear=False):
                checks = processor.preflight_checks({"task_type": "pdf_to_word"})

            mathpix = next(item for item in checks if item["id"] == "mathpix_authorization")
            serialized = json.dumps(mathpix, ensure_ascii=False)
            self.assertEqual(mathpix["status"], "通过")
            self.assertIn("K12_TEST_MATHPIX_ID", serialized)
            self.assertIn("K12_TEST_MATHPIX_KEY", serialized)
            self.assertNotIn("app-id", serialized)
            self.assertNotIn("super-secret", serialized)

    def test_logs_can_export_txt_log_csv_and_json(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            store.append_log("task_demo", "转换完成", category="conversion")
            store.append_log("task_other", "其他任务日志", category="conversion")

            log_handler = make_handler(store, processor)
            K12RequestHandler._send_logs_file(log_handler, None, "log")
            log_text = log_handler.wfile.getvalue().decode("utf-8")
            self.assertEqual(log_handler.status, 200)
            self.assertIn(("Content-Type", "text/plain; charset=utf-8"), log_handler.output_headers)
            self.assertIn(("Content-Disposition", 'attachment; filename="k12-logs.log"'), log_handler.output_headers)
            self.assertIn("转换完成", log_text)

            csv_handler = make_handler(store, processor)
            K12RequestHandler._send_logs_file(csv_handler, None, "csv")
            csv_text = csv_handler.wfile.getvalue().decode("utf-8")
            self.assertEqual(csv_handler.status, 200)
            self.assertIn(("Content-Type", "text/csv; charset=utf-8"), csv_handler.output_headers)
            self.assertIn("created_at,category,level,task_id,message", csv_text)
            self.assertIn('"conversion"', csv_text)

            json_handler = make_handler(store, processor)
            K12RequestHandler._send_logs_file(json_handler, "task_demo", "json")
            payload = json.loads(json_handler.wfile.getvalue().decode("utf-8"))
            self.assertEqual(len(payload["logs"]), 1)
            self.assertEqual(payload["logs"][0]["message"], "转换完成")
            self.assertIn(("Content-Disposition", 'attachment; filename="task_demo-logs.json"'), json_handler.output_headers)

            store.update_settings({"logExportFormat": "csv"})
            default_handler = make_handler(store, processor)
            K12RequestHandler._send_logs_file(default_handler)
            self.assertIn(("Content-Disposition", 'attachment; filename="k12-logs.csv"'), default_handler.output_headers)

            unsafe_handler = make_handler(store, processor)
            K12RequestHandler._send_logs_file(unsafe_handler, 'task"\r\nX-Injected: yes', "json")
            disposition = dict(unsafe_handler.output_headers)["Content-Disposition"]
            self.assertNotIn("\r", disposition)
            self.assertNotIn("\n", disposition)
            self.assertIn("task___X-Injected: yes-logs.json", disposition)

    def test_upload_zip_registers_supported_entries(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            files = processor.create_uploaded_file("batch.zip", make_zip_bytes({"lesson.docx": make_docx_bytes(), "ignore.exe": b"x"}))
            names = {file["file_name"] for file in files}
            self.assertIn("batch.zip", names)
            self.assertIn("lesson.docx", names)
            self.assertEqual(len(files), 2)
            self.assertTrue(any(file["source_kind"] == "archive_entry" for file in files))
            child = next(file for file in files if file["source_kind"] == "archive_entry")
            self.assertEqual(child["source_relative_path"], "lesson.docx")

    def test_folder_upload_preserves_browser_relative_path(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            uploaded = processor.create_uploaded_file("七年级/数学/期末试卷.docx", make_docx_bytes())[0]

            self.assertEqual(uploaded["file_name"], "期末试卷.docx")
            self.assertEqual(uploaded["source_relative_path"], "七年级/数学/期末试卷.docx")
            task = processor.create_task({"task_type": "formula_precheck", "file_ids": [uploaded["id"]]})
            payload = processor.local_task_payload(task["id"])
            self.assertEqual(payload["files"][0]["source_relative_path"], "七年级/数学/期末试卷.docx")

    def test_duplicate_upload_strategy_renames_skips_and_overwrites(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            first = processor.create_uploaded_file("lesson.docx", make_docx_bytes())[0]
            renamed = processor.create_uploaded_file("lesson.docx", make_docx_bytes())[0]
            self.assertEqual(first["file_name"], "lesson.docx")
            self.assertEqual(renamed["file_name"], "lesson-1.docx")

        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            store.update_settings({"duplicateFileStrategy": "跳过"})
            processor = TaskProcessor(store)
            first = processor.create_uploaded_file("lesson.docx", make_docx_bytes())[0]
            skipped = processor.create_uploaded_file("lesson.docx", make_docx_bytes())[0]
            self.assertEqual(skipped["id"], first["id"])
            self.assertEqual(len(store.list_files()), 1)

        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            store.update_settings({"duplicateFileStrategy": "覆盖"})
            processor = TaskProcessor(store)
            first = processor.create_uploaded_file("lesson.docx", make_docx_bytes())[0]
            overwritten = processor.create_uploaded_file("lesson.docx", make_docx_bytes())[0]
            self.assertNotEqual(overwritten["id"], first["id"])
            self.assertEqual(overwritten["file_name"], "lesson.docx")
            self.assertEqual(len(store.list_files()), 1)

    def test_delete_file_cleans_upload_cache_but_keeps_external_source(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(Path(tmp) / "data")
            processor = TaskProcessor(store)
            uploaded = processor.create_uploaded_file("lesson.docx", make_docx_bytes())[0]
            upload_path = Path(uploaded["storage_path"])
            self.assertTrue(upload_path.exists())
            self.assertTrue(store.delete_file(uploaded["id"]))
            self.assertIsNone(store.get_file(uploaded["id"]))
            self.assertFalse(upload_path.exists())

            external_path = Path(tmp) / "external.docx"
            external_path.write_bytes(make_docx_bytes())
            external = processor.create_file({"file_name": "external.docx", "file_size": external_path.stat().st_size, "file_path": str(external_path)})
            self.assertTrue(store.delete_file(external["id"]))
            self.assertTrue(external_path.exists())

    def test_file_download_info_supports_upload_and_external_source(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(Path(tmp) / "data")
            processor = TaskProcessor(store)
            uploaded = processor.create_uploaded_file("lesson.docx", make_docx_bytes())[0]
            uploaded_info = processor.file_download_info(uploaded["id"])
            self.assertEqual(uploaded_info["file_name"], "lesson.docx")
            self.assertTrue(Path(uploaded_info["path"]).exists())

            external_path = Path(tmp) / "external.docx"
            external_path.write_bytes(make_docx_bytes())
            external = processor.create_file({"file_name": "external.docx", "file_size": external_path.stat().st_size, "file_path": str(external_path)})
            external_info = processor.file_download_info(external["id"])
            self.assertEqual(external_info["path"], external_path)

            store.delete_file(uploaded["id"])
            with self.assertRaises(KeyError):
                processor.file_download_info(uploaded["id"])

            missing = processor.create_file({"file_name": "missing.docx", "file_size": 2048, "file_path": str(Path(tmp) / "missing.docx")})
            with self.assertRaises(ValueError):
                processor.file_download_info(missing["id"])

    def test_download_content_disposition_sanitizes_attachment_names(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "source.docx"
            source.write_bytes(make_docx_bytes())
            store = AppStore(Path(tmp) / "data")
            processor = TaskProcessor(store)
            file_item = processor.create_file(
                {
                    "file_name": 'lesson"\r\nX-Injected: yes.docx',
                    "file_size": source.stat().st_size,
                    "file_path": str(source),
                }
            )
            handler = make_handler(store, processor)

            K12RequestHandler._send_source_file(handler, file_item["id"])

            headers = dict(handler.output_headers)
            disposition = headers["Content-Disposition"]
            self.assertEqual(handler.status, 200)
            self.assertNotIn("\r", disposition)
            self.assertNotIn("\n", disposition)
            self.assertNotIn('"lesson"', disposition)
            self.assertIn("lesson___X-Injected_ yes.docx", disposition)

    def test_replace_uploaded_file_preserves_id_and_reanalyzes_content(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            original = processor.create_uploaded_file("lesson.docx", make_docx_bytes())[0]
            original_path = Path(original["storage_path"])
            replaced = processor.replace_uploaded_file(original["id"], "scan.pdf", b"%PDF-1.4\n1 0 obj<< /Type /Page /Subtype /Image >>endobj\n")
            self.assertEqual(replaced["id"], original["id"])
            self.assertEqual(replaced["file_name"], "scan.pdf")
            self.assertEqual(replaced["file_type"], "PDF")
            self.assertEqual(replaced["source_kind"], "upload")
            self.assertTrue(replaced["replaced_at"])
            self.assertFalse(original_path.exists())
            self.assertTrue(Path(replaced["storage_path"]).exists())

    def test_replace_zip_parent_removes_old_archive_entries(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            files = processor.create_uploaded_file("batch.zip", make_zip_bytes({"lesson.docx": make_docx_bytes()}))
            parent = next(file for file in files if file["file_name"] == "batch.zip")
            child = next(file for file in files if file.get("archive_parent_id") == parent["id"])
            replaced = processor.replace_uploaded_file(parent["id"], "replacement.docx", make_docx_bytes())
            self.assertEqual(replaced["id"], parent["id"])
            self.assertEqual(replaced["file_type"], "Word")
            self.assertIsNone(store.get_file(child["id"]))
            self.assertFalse((store.uploads_dir / parent["id"]).exists())

    def test_files_bundle_download_includes_selected_files_and_manifest(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            word = processor.create_uploaded_file("lesson.docx", make_docx_bytes())[0]
            pdf = processor.create_uploaded_file("scan.pdf", b"%PDF-1.4\n1 0 obj<< /Type /Page >>endobj\n")[0]
            handler = make_handler(store, processor)
            K12RequestHandler._send_files_bundle(handler, {"ids": [f"{word['id']},{pdf['id']}"]})
            self.assertEqual(handler.status, 200)
            with zipfile.ZipFile(BytesIO(handler.wfile.getvalue())) as archive:
                names = archive.namelist()
                manifest = archive.read("manifest.csv").decode("utf-8")
                self.assertIn("manifest.csv", names)
                self.assertTrue(any(name.endswith("lesson.docx") for name in names))
                self.assertTrue(any(name.endswith("scan.pdf") for name in names))
                self.assertIn(word["id"], manifest)
                self.assertIn(pdf["id"], manifest)

            handler = make_handler(store, processor)
            K12RequestHandler._send_files_bundle(handler, {"type": ["PDF"]})
            with zipfile.ZipFile(BytesIO(handler.wfile.getvalue())) as archive:
                names = archive.namelist()
                self.assertTrue(any(name.endswith("scan.pdf") for name in names))
                self.assertFalse(any(name.endswith("lesson.docx") for name in names))

    def test_delete_archive_parent_removes_extracted_entries(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            files = processor.create_uploaded_file("batch.zip", make_zip_bytes({"lesson.docx": make_docx_bytes()}))
            parent = next(file for file in files if file["file_name"] == "batch.zip")
            child = next(file for file in files if file.get("archive_parent_id") == parent["id"])
            parent_path = Path(parent["storage_path"])
            child_path = Path(child["storage_path"])
            extract_dir = store.uploads_dir / parent["id"]
            self.assertTrue(parent_path.exists())
            self.assertTrue(child_path.exists())
            self.assertTrue(store.delete_file(parent["id"]))
            self.assertIsNone(store.get_file(parent["id"]))
            self.assertIsNone(store.get_file(child["id"]))
            self.assertFalse(parent_path.exists())
            self.assertFalse(extract_dir.exists())

    def test_batch_process_records_per_file_results_and_failure_list(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            store.update_settings({"maxConcurrentTasks": 1})
            processor = TaskProcessor(store)
            good = processor.create_file({"file_name": "lesson.docx", "file_size": 4096})
            bad = processor.create_file({"file_name": "bad.exe", "file_size": 10})
            task = processor.create_task({"task_type": "batch_process", "file_ids": [good["id"], bad["id"]]})
            report = store.list_reports()[0]
            results = report["analysis"]["batchResults"]
            plan = report["analysis"]["batchPlan"]
            quality = {item["id"]: item for item in report["qualityChecks"]}
            self.assertEqual(task["status"], "失败")
            self.assertEqual(report["batch_total_count"], 2)
            self.assertEqual(report["batch_success_count"], 1)
            self.assertEqual(report["batch_fail_count"], 1)
            self.assertEqual(plan["max_concurrent"], 1)
            self.assertEqual(plan["chunk_count"], 2)
            self.assertTrue(plan["continue_on_failure"])
            self.assertEqual(plan["failure_strategy"], "跳过")
            self.assertFalse(plan["requires_user_decision"])
            self.assertEqual(report["batch_failure_strategy"], "跳过")
            self.assertEqual(report["batch_retryable_count"], 1)
            self.assertGreaterEqual(report["failure_count"], 1)
            self.assertEqual(task["success_count"], 1)
            self.assertEqual(task["fail_count"], 1)
            self.assertEqual(task["failure_count"], report["failure_count"])
            self.assertEqual(task["retryable_count"], 1)
            self.assertIn("bad.exe", task["error_message"])
            self.assertIn("文件格式不支持", task["error_message"])
            self.assertEqual([item["status"] for item in results], ["成功", "失败"])
            self.assertEqual([item["progress"] for item in results], [100, 100])
            self.assertTrue(results[1]["retryable"])
            self.assertIn("文件格式不支持", results[1]["error_message"])
            self.assertEqual(quality["batch_results"]["status"], "失败")
            text_report = Path(report["txt_path"]).read_text(encoding="utf-8")
            self.assertIn("批量单文件结果", text_report)
            self.assertIn("最大并发 1", text_report)
            failure_csv = Path(report["failure_csv_path"]).read_text(encoding="utf-8")
            self.assertIn("bad.exe", failure_csv)
            self.assertIn("文件格式不支持", failure_csv)
            handler = make_handler(store, processor)
            K12RequestHandler._send_report_failures(handler, report["id"])
            self.assertEqual(handler.status, 200)
            self.assertIn("bad.exe", handler.wfile.getvalue().decode("utf-8"))
            with zipfile.ZipFile(report["xlsx_path"]) as archive:
                self.assertIn("批量单文件结果", archive.read("xl/worksheets/sheet1.xml").decode("utf-8"))

    def test_batch_process_options_override_concurrency_and_failure_strategy(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            store.update_settings({"maxConcurrentTasks": 1})
            processor = TaskProcessor(store)
            files = [
                processor.create_file({"file_name": "lesson.docx", "file_size": 4096}),
                processor.create_file({"file_name": "slides.pptx", "file_size": 4096}),
                processor.create_file({"file_name": "bad.exe", "file_size": 10}),
            ]
            task = processor.create_task(
                {
                    "task_type": "batch_process",
                    "file_ids": [file["id"] for file in files],
                    "options": {"maxConcurrentTasks": 2, "batchFailureStrategy": "询问", "continueOnFailure": False},
                }
            )
            report = store.list_reports()[0]
            plan = report["analysis"]["batchPlan"]

            self.assertEqual(task["options"]["maxConcurrentTasks"], 2)
            self.assertEqual(plan["max_concurrent"], 2)
            self.assertEqual(plan["chunk_count"], 2)
            self.assertEqual(plan["failure_strategy"], "询问")
            self.assertFalse(plan["continue_on_failure"])
            self.assertTrue(plan["requires_user_decision"])
            self.assertEqual(report["batch_failure_strategy"], "询问")

    def test_task_workflow_order_is_normalized_and_reported(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            word = processor.create_uploaded_file("lesson.docx", make_docx_bytes())[0]

            task = processor.create_task(
                {
                    "task_type": "word_to_ppt",
                    "file_ids": [word["id"]],
                    "options": {
                        "workflowOrder": ["pdf_to_word", "not_a_task", "pdf_to_word", "small_image_scan"],
                        "workflowLabels": ["伪造标签"],
                    },
                }
            )
            report = store.list_reports()[0]
            plan = task["options"]["workflowPlan"]
            quality = {item["id"]: item for item in report["qualityChecks"]}

            self.assertEqual(task["options"]["workflowOrder"], ["word_to_ppt", "pdf_to_word", "small_image_scan"])
            self.assertEqual(task["options"]["workflowLabels"], ["Word 转 PPT", "PDF 转 Word", "微小图片检索"])
            self.assertEqual(plan["version"], "k12.workflowPlan.v1")
            self.assertEqual(plan["current_index"], 1)
            self.assertEqual(plan["invalid_items"], ["not_a_task"])
            self.assertEqual(plan["duplicate_items"], ["pdf_to_word"])
            self.assertEqual(report["analysis"]["workflowPlan"], plan)
            self.assertEqual(quality["workflow_order"]["status"], "需确认")
            self.assertIn("Word 转 PPT → PDF 转 Word → 微小图片检索", quality["workflow_order"]["metric"])
            self.assertIn("已过滤未知任务：not_a_task", quality["workflow_order"]["message"])
            self.assertIn("已去除重复任务：pdf_to_word", quality["workflow_order"]["message"])
            self.assertIn("流程顺序", Path(report["txt_path"]).read_text(encoding="utf-8"))

    def test_batch_failed_file_can_be_skipped_and_report_rebuilt(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            good = processor.create_file({"file_name": "lesson.docx", "file_size": 4096})
            bad = processor.create_file({"file_name": "bad.exe", "file_size": 10})
            task = processor.create_task({"task_type": "batch_process", "file_ids": [good["id"], bad["id"]]})
            self.assertEqual(task["status"], "失败")

            handler = make_handler(store, processor)
            K12RequestHandler._handle_api_post(handler, f"/api/tasks/{task['id']}/skip-file", {"file_id": bad["id"]})
            self.assertEqual(handler.status, 200)
            payload = json.loads(handler.wfile.getvalue().decode("utf-8"))
            updated_task = payload["task"]
            report = store.list_reports()[0]
            results = report["analysis"]["batchResults"]
            plan = report["analysis"]["batchPlan"]
            quality = {item["id"]: item for item in report["qualityChecks"]}

            self.assertEqual(updated_task["status"], "成功")
            self.assertEqual(updated_task["skipped_file_ids"], [bad["id"]])
            self.assertEqual([item["status"] for item in results], ["成功", "跳过"])
            self.assertFalse(results[1]["retryable"])
            self.assertTrue(results[1]["skipped"])
            self.assertIn("已按用户选择跳过", results[1]["recommendation"])
            self.assertEqual(report["batch_fail_count"], 0)
            self.assertEqual(report["batch_skipped_count"], 1)
            self.assertEqual(report["batch_retryable_count"], 0)
            self.assertEqual(report["fail_count"], 0)
            self.assertEqual(plan["failed_count"], 0)
            self.assertEqual(plan["skipped_count"], 1)
            self.assertEqual(report["failure_count"], 0)
            self.assertEqual(quality["file_validation"]["status"], "通过")
            self.assertEqual(quality["batch_results"]["status"], "通过")
            self.assertIn("跳过 1 个", Path(report["txt_path"]).read_text(encoding="utf-8"))
            self.assertNotIn("bad.exe", Path(report["failure_csv_path"]).read_text(encoding="utf-8"))
            self.assertTrue(any("跳过批量失败文件" in log["message"] for log in store.list_logs(task["id"], limit=1000)))

    def test_pdf_to_word_records_mathpix_authorization_requirement(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            files = processor.create_uploaded_file("scan.pdf", b"%PDF-1.4\n1 0 obj<< /Type /Page /Subtype /Image /Table >>(Math Formula) Tj endobj\n")
            pdf_file = next(file for file in files if file["file_type"] == "PDF")
            task = processor.create_task({"task_type": "pdf_to_word", "file_ids": [pdf_file["id"]]})
            report = store.list_reports()[0]
            mathpix = report["analysis"]["mathpix"][0]
            retention_plan = mathpix["retention_plan"]
            recognition_plan = mathpix["recognition_plan"]
            quality = {item["id"]: item for item in report["qualityChecks"]}
            self.assertEqual(task["status"], "成功")
            self.assertEqual(task["execute_mode"], "hybrid")
            self.assertEqual(pdf_file["content_summary"]["pdfType"], "混合型 PDF")
            self.assertEqual(mathpix["engine"], "Mathpix")
            self.assertEqual(mathpix["status"], "authorization_required")
            self.assertEqual(recognition_plan["schema_version"], "k12.mathpixRecognitionPlan.v1")
            self.assertEqual(recognition_plan["engine"], "Mathpix")
            self.assertEqual(recognition_plan["status"], "blocked_authorization")
            self.assertEqual(recognition_plan["status_label"], "等待授权")
            self.assertFalse(recognition_plan["submit_allowed"])
            upload_gate = recognition_plan["upload_gate"]
            self.assertEqual(upload_gate["schema_version"], "k12.mathpixUploadGate.v1")
            self.assertEqual(upload_gate["status"], "blocked_authorization")
            self.assertFalse(upload_gate["submit_allowed"])
            self.assertIn("external_upload_not_authorized", upload_gate["blocking_reasons"])
            self.assertFalse(upload_gate["external_upload_authorized"])
            self.assertTrue(upload_gate["source_file_available"])
            self.assertTrue(upload_gate["ocr_enabled"])
            self.assertFalse(upload_gate["credential_values_exposed"])
            self.assertTrue(recognition_plan["source_file_available"])
            self.assertTrue(recognition_plan["external_upload_required"])
            self.assertFalse(recognition_plan["external_upload_authorized"])
            self.assertFalse(recognition_plan["credential_values_exposed"])
            self.assertEqual(recognition_plan["submit_contract"]["conversion_formats"], ["docx", "tex.zip"])
            self.assertEqual(recognition_plan["submit_contract"]["download_formats"], ["docx", "tex.zip"])
            self.assertTrue(recognition_plan["safety"]["no_upload_without_authorization"])
            self.assertTrue(recognition_plan["safety"]["credential_values_redacted"])
            self.assertEqual(retention_plan["engine"], "Mathpix")
            self.assertEqual(retention_plan["pdf_type"], "混合型 PDF")
            self.assertTrue(retention_plan["retain_images"])
            self.assertTrue(retention_plan["retain_tables"])
            self.assertTrue(retention_plan["formula_ocr"])
            self.assertEqual(retention_plan["image_objects"], 1)
            self.assertGreaterEqual(retention_plan["table_hints"], 1)
            self.assertGreaterEqual(retention_plan["formula_hints"], 1)
            self.assertEqual(retention_plan["image_retention_status"], "计划保留图片对象")
            self.assertEqual(retention_plan["table_retention_status"], "计划识别并保留表格结构")
            self.assertEqual(retention_plan["formula_retention_status"], "计划识别公式并生成 MathType 预览")
            self.assertIn("Mathpix 保留计划", Path(report["txt_path"]).read_text(encoding="utf-8"))
            self.assertIn("计划保留图片对象", Path(report["html_path"]).read_text(encoding="utf-8"))
            with zipfile.ZipFile(report["xlsx_path"]) as archive:
                sheet_xml = archive.read("xl/worksheets/sheet1.xml").decode("utf-8")
                self.assertIn("Mathpix 保留计划", sheet_xml)
                self.assertIn("计划识别并保留表格结构", sheet_xml)
            self.assertEqual(quality["mathpix_pdf"]["status"], "需确认")
            self.assertEqual(quality["pdf_type"]["status"], "需确认")
            self.assertIn("公式 1 个", quality["pdf_type"]["metric"])
            self.assertIn("表格 1 个", quality["pdf_type"]["metric"])
            self.assertEqual(report["analysis"]["artifacts"][0]["status"], "authorization_required")
            queue = processor.mathpix_job_queue()
            self.assertEqual(queue["schema_version"], "k12.mathpixJobQueue.v1")
            self.assertFalse(queue["external_upload_allowed"])
            self.assertEqual(queue["summary"]["blocked"], 1)
            self.assertEqual(queue["jobs"][0]["status"], "authorization_required")
            self.assertEqual(queue["jobs"][0]["status_label"], "等待授权")
            self.assertEqual(queue["jobs"][0]["recognition_plan"]["schema_version"], "k12.mathpixRecognitionPlan.v1")
            self.assertEqual(queue["jobs"][0]["recognition_plan"]["status"], "blocked_authorization")
            self.assertFalse(queue["jobs"][0]["recognition_plan"]["submit_allowed"])
            self.assertEqual(queue["jobs"][0]["recognition_plan"]["upload_gate"]["status"], "blocked_authorization")
            self.assertIn("external_upload_not_authorized", queue["jobs"][0]["recognition_plan"]["upload_gate"]["blocking_reasons"])
            self.assertFalse(queue["jobs"][0]["recognition_plan"]["credential_values_exposed"])
            self.assertEqual(queue["jobs"][0]["recognition_plan"]["conversion_formats"], ["docx", "tex.zip"])
            payload = processor.local_task_payload(task["id"])
            pdf_action = next(action for action in payload["local_actions"] if action["type"] == "pdf_formula_mathtype")
            self.assertEqual(pdf_action["status"], "blocked_mathpix")
            self.assertGreaterEqual(pdf_action["formula_count"], 1)
            self.assertEqual(pdf_action["mathpix_jobs"][0]["recognition_status"], "blocked_authorization")
            plan_action = next(action for action in payload["desktop_execution_plan"]["actions"] if action["type"] == "pdf_formula_mathtype")
            self.assertEqual(plan_action["gate_status"], "waiting_for_heartbeat")
            self.assertEqual(plan_action["required_capabilities"][0]["key"], "mathTypeAutomation")
            self.assertEqual(len(plan_action["steps"]), 4)

    def test_mathpix_jobs_endpoint_is_read_only_and_redacts_outputs(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            store.update_settings({"allowExternalMathpixUpload": False})
            processor = TaskProcessor(store)
            pdf = processor.create_uploaded_file("scan.pdf", b"%PDF-1.4\n1 0 obj<< /Type /Page /Subtype /Image >>(Math Formula) Tj endobj\n")[0]
            processor.create_task({"task_type": "pdf_to_word", "file_ids": [pdf["id"]]})
            handler = make_handler(store, processor)

            with patch("k12.processor.MathpixClient.from_environment") as client_factory:
                K12RequestHandler._handle_api_get(handler, "/api/mathpix-jobs", {})

            client_factory.assert_not_called()
            response = json.loads(handler.wfile.getvalue().decode("utf-8"))
            queue = response["mathpixJobs"]
            job = queue["jobs"][0]

            self.assertEqual(handler.status, 200)
            self.assertEqual(queue["schema_version"], "k12.mathpixJobQueue.v1")
            self.assertEqual(queue["engine"], "Mathpix")
            self.assertFalse(queue["external_upload_allowed"])
            self.assertEqual(job["status"], "authorization_required")
            self.assertFalse(job["recognition_plan"]["submit_allowed"])
            self.assertFalse(job["recognition_plan"]["credential_values_exposed"])
            self.assertEqual(job["outputs"], [])
            self.assertNotIn(str(store.data_dir), json.dumps(queue, ensure_ascii=False))

    def test_pdf_to_word_task_options_cannot_bypass_mathpix_upload_authorization(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            pdf = processor.create_uploaded_file("scan.pdf", b"%PDF-1.4\n1 0 obj<< /Type /Page /Subtype /Image >>(Math Formula) Tj endobj\n")[0]

            with patch("k12.processor.MathpixClient.from_environment") as client_factory:
                task = processor.create_task(
                    {
                        "task_type": "pdf_to_word",
                        "file_ids": [pdf["id"]],
                        "options": {"allowExternalMathpixUpload": True, "externalUploadAuthorized": True, "waitForMathpix": True},
                    }
                )

            client_factory.assert_not_called()
            report = store.list_reports()[0]
            mathpix = report["analysis"]["mathpix"][0]
            recognition_plan = mathpix["recognition_plan"]
            task_option_audit = recognition_plan["task_option_audit"]

            self.assertEqual(task["status"], "成功")
            self.assertEqual(mathpix["status"], "authorization_required")
            self.assertEqual(recognition_plan["status"], "blocked_authorization")
            self.assertFalse(recognition_plan["submit_allowed"])
            self.assertFalse(recognition_plan["external_upload_authorized"])
            self.assertEqual(recognition_plan["external_upload_authorization_source"], "settings.allowExternalMathpixUpload")
            self.assertEqual(recognition_plan["upload_gate"]["status"], "blocked_authorization")
            self.assertIn("external_upload_not_authorized", recognition_plan["upload_gate"]["blocking_reasons"])
            self.assertEqual(recognition_plan["upload_gate"]["authorization_source"], "settings.allowExternalMathpixUpload")
            self.assertFalse(recognition_plan["upload_gate"]["task_options_can_authorize"])
            self.assertTrue(recognition_plan["submit_contract"]["wait_for_completion"])
            self.assertEqual(recognition_plan["submit_contract"]["wait_for_completion_source"], "task_options")
            self.assertEqual(task_option_audit["schema_version"], "k12.mathpixTaskOptionAudit.v1")
            self.assertFalse(task_option_audit["task_options_can_authorize_external_upload"])
            self.assertEqual(task_option_audit["authorization_source"], "settings.allowExternalMathpixUpload")
            self.assertEqual(task_option_audit["ignored_authorization_keys"], ["allowExternalMathpixUpload", "externalUploadAuthorized"])
            self.assertEqual(task_option_audit["ignored_authorization_count"], 2)
            self.assertTrue(task_option_audit["wait_for_completion_option_present"])
            self.assertTrue(task_option_audit["wait_for_completion_can_be_requested"])
            self.assertEqual(recognition_plan["safety"]["ignored_task_option_authorization_count"], 2)
            self.assertTrue(recognition_plan["safety"]["task_options_cannot_authorize_external_upload"])
            self.assertTrue(recognition_plan["safety"]["task_options_may_request_wait_for_completion"])
            self.assertEqual(processor.mathpix_job_queue()["jobs"][0]["recognition_plan"]["task_option_audit"], task_option_audit)

    def test_pdf_ocr_contract_probe_keeps_mathpix_upload_blocked_without_authorization(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            processor = TaskProcessor(AppStore(tmp))

            probe = processor._pdf_ocr_contract_probe()

            self.assertEqual(probe["schema_version"], "k12.pdfOcrContractProbe.v1")
            self.assertTrue(probe["scanned_available"])
            self.assertTrue(probe["formula_available"])
            self.assertEqual(probe["scanned_pdf_type"], "扫描型 PDF")
            self.assertEqual(probe["recognition_status"]["scanned"], "blocked_authorization")
            self.assertEqual(probe["recognition_status"]["formula"], "blocked_authorization")
            self.assertIn("docx", probe["conversion_formats"])
            self.assertIn("tex.zip", probe["conversion_formats"])
            self.assertIn("扫描型 PDF", probe["scanned_evidence"])
            self.assertIn("未发起外部上传", probe["formula_current"])
            self.assertIn("blocked_mathpix", probe["formula_handoff_current"])

    def test_encrypted_file_password_is_session_only(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            pdf = processor.create_uploaded_file("locked.pdf", b"%PDF-1.4\n1 0 obj<< /Type /Page /Encrypt >>endobj\n")[0]
            self.assertTrue(pdf["encrypted"])
            self.assertEqual(pdf["status"], "校验失败")
            self.assertTrue(any("密码" in error for error in pdf["validation_errors"]))

            updated = processor.set_file_password(pdf["id"], "123456")
            self.assertTrue(store.has_file_password(pdf["id"]))
            self.assertEqual(updated["status"], "待处理")
            self.assertFalse(updated["validation_errors"])
            self.assertNotIn("123456", repr(updated))

            reopened = AppStore(tmp)
            reloaded = reopened.get_file(pdf["id"])
            self.assertFalse(reloaded["password_session_active"])
            self.assertFalse(reloaded["password_provided"])
            self.assertTrue(any("密码" in error for error in reloaded["validation_errors"]))
            self.assertFalse(reopened.has_file_password(pdf["id"]))

            with self.assertRaises(ValueError):
                processor.set_file_password(pdf["id"], "")

    def test_pdf_to_word_mathpix_completion_creates_docx_artifact(self) -> None:
        fake_client = None

        class FakeMathpixClient:
            def submit_pdf(self, path: Path, options: dict) -> dict:
                self.submitted_path = path
                self.options = options
                return {"pdf_id": "pdf_fake_123"}

            def wait_for_pdf(self, pdf_id: str, timeout_seconds: int) -> dict:
                return {"status": "completed", "percent_done": 100}

            def download_pdf_result(self, pdf_id: str, extension: str) -> bytes:
                if extension == "tex.zip":
                    return make_zip_bytes({"formula.tex": b"x^2+y^2=z^2\n"})
                return make_docx_bytes()

        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            store.update_settings(
                {
                    "allowExternalMathpixUpload": True,
                    "waitForMathpix": True,
                    "localClientPlatform": "Windows",
                    "ocrLanguage": "中文",
                    "enableTextOcr": True,
                    "enableFormulaOcr": True,
                    "enableTableOcr": False,
                    "ocrPrecisionMode": "高精度",
                    "ocrSpeedMode": "优先速度",
                    "formulaConfidenceThreshold": 95,
                    "lowConfidenceFormulaStrategy": "标记重识别",
                    "keepFormulaImages": False,
                }
            )
            processor = TaskProcessor(store)
            processor.record_local_client_heartbeat(
                {
                    "client_id": "desktop-pdf-formula",
                    "status": "online",
                    "platform": "Windows",
                    "preflight": {
                        "platform": "Windows",
                        "components": {
                            "mathtype": {"label": "MathType 组件", "available": True, "status": "available"},
                        },
                        "capabilities": {"mathTypeAutomation": True},
                    },
                }
            )
            pdf = processor.create_uploaded_file("scan.pdf", b"%PDF-1.4\n1 0 obj<< /Type /Page >>endobj\n")[0]
            fake_client = FakeMathpixClient()
            with patch("k12.processor.MathpixClient.from_environment", return_value=fake_client):
                task = processor.create_task({"task_type": "pdf_to_word", "file_ids": [pdf["id"]]})
            report = store.list_reports()[0]
            mathpix = report["analysis"]["mathpix"][0]
            recognition_plan = mathpix["recognition_plan"]
            formulas = report["analysis"]["formulas"]
            artifact = report["analysis"]["artifacts"][0]
            tex_artifact = report["analysis"]["artifacts"][1]
            quality = {item["id"]: item for item in report["qualityChecks"]}
            self.assertEqual(task["status"], "成功")
            self.assertEqual(task["execute_mode"], "hybrid")
            self.assertEqual(mathpix["status"], "completed")
            self.assertEqual(recognition_plan["schema_version"], "k12.mathpixRecognitionPlan.v1")
            self.assertEqual(recognition_plan["status"], "completed")
            self.assertTrue(recognition_plan["submit_allowed"])
            self.assertEqual(recognition_plan["upload_gate"]["schema_version"], "k12.mathpixUploadGate.v1")
            self.assertEqual(recognition_plan["upload_gate"]["status"], "completed")
            self.assertTrue(recognition_plan["upload_gate"]["submit_allowed"])
            self.assertEqual(recognition_plan["upload_gate"]["blocking_reasons"], [])
            self.assertTrue(recognition_plan["upload_gate"]["external_upload_authorized"])
            self.assertTrue(recognition_plan["external_upload_authorized"])
            self.assertEqual(recognition_plan["submit_contract"]["conversion_formats"], ["docx", "tex.zip"])
            self.assertEqual(recognition_plan["submit_contract"]["download_formats"], ["docx", "tex.zip"])
            self.assertTrue(recognition_plan["submit_contract"]["wait_for_completion"])
            self.assertEqual(recognition_plan["submit_contract"]["wait_for_completion_source"], "settings")
            self.assertEqual(recognition_plan["formula_review"]["schema_version"], "k12.formulaReviewContract.v1")
            self.assertEqual(recognition_plan["formula_review"]["confidence_threshold"], 95)
            self.assertEqual(recognition_plan["formula_review"]["low_confidence_strategy"], "标记重识别")
            self.assertFalse(recognition_plan["formula_review"]["keep_original_image"])
            self.assertIn("重新识别", recognition_plan["formula_review"]["manual_actions"])
            self.assertEqual(recognition_plan["formula_review"]["annotation_endpoint"], "/api/formula-annotations")
            self.assertEqual(recognition_plan["external_upload_authorization_source"], "settings.allowExternalMathpixUpload")
            self.assertEqual(recognition_plan["upload_gate"]["authorization_source"], "settings.allowExternalMathpixUpload")
            self.assertFalse(recognition_plan["upload_gate"]["task_options_can_authorize"])
            self.assertEqual(mathpix["tex_zip_status"], "completed")
            self.assertEqual(mathpix["ocr_settings"]["language"], "中文")
            self.assertFalse(mathpix["ocr_settings"]["table_ocr"])
            self.assertEqual(mathpix["retention_plan"]["engine"], "Mathpix")
            self.assertFalse(mathpix["retention_plan"]["table_ocr"])
            self.assertEqual(mathpix["retention_plan"]["table_retention_status"], "未检测到表格线索")
            self.assertEqual(fake_client.options["conversion_formats"], {"docx": True, "tex.zip": True})
            self.assertFalse(fake_client.options["enable_tables_fallback"])
            self.assertEqual(fake_client.options["metadata"]["k12_ocr_language"], "中文")
            self.assertEqual(fake_client.options["metadata"]["k12_precision_mode"], "高精度")
            self.assertEqual(fake_client.options["metadata"]["k12_speed_mode"], "优先速度")
            self.assertEqual(mathpix["request_options"]["conversion_formats"], {"docx": True, "tex.zip": True})
            self.assertIn("tex_zip", mathpix["outputs"])
            self.assertIn("docx", mathpix["output_summaries"])
            self.assertIn("tex_zip", mathpix["output_summaries"])
            self.assertEqual(artifact["output_type"], "docx")
            self.assertEqual(artifact["status"], "成功")
            self.assertTrue(Path(artifact["path"]).exists())
            self.assertTrue(artifact["url"].startswith(f"/api/artifacts/{task['id']}/"))
            self.assertEqual(tex_artifact["output_type"], "tex.zip")
            self.assertEqual(tex_artifact["status"], "成功")
            self.assertTrue(Path(tex_artifact["path"]).exists())
            docx_summary = mathpix["output_summaries"]["docx"]
            tex_summary = mathpix["output_summaries"]["tex_zip"]
            self.assertEqual(docx_summary["size"], Path(artifact["path"]).stat().st_size)
            self.assertEqual(docx_summary["sha256"], hashlib.sha256(Path(artifact["path"]).read_bytes()).hexdigest())
            self.assertEqual(tex_summary["size"], Path(tex_artifact["path"]).stat().st_size)
            self.assertEqual(tex_summary["sha256"], hashlib.sha256(Path(tex_artifact["path"]).read_bytes()).hexdigest())
            self.assertEqual(recognition_plan["download_results"]["docx"]["sha256"], docx_summary["sha256"])
            self.assertEqual(recognition_plan["download_results"]["tex_zip"]["sha256"], tex_summary["sha256"])
            self.assertFalse(recognition_plan["download_results"]["docx"]["path_exposed"])
            download_manifest = recognition_plan["download_manifest"]
            manifest_outputs = {item["type"]: item for item in download_manifest["outputs"]}
            self.assertEqual(download_manifest["schema_version"], "k12.mathpixDownloadManifest.v1")
            self.assertEqual(download_manifest["status"], "complete")
            self.assertEqual(download_manifest["required_output_count"], 2)
            self.assertEqual(download_manifest["satisfied_required_count"], 2)
            self.assertEqual(download_manifest["failed_required_count"], 0)
            self.assertTrue(download_manifest["all_required_outputs_satisfied"])
            self.assertTrue(download_manifest["all_required_outputs_hashed"])
            self.assertEqual(manifest_outputs["docx"]["status"], "downloaded")
            self.assertTrue(manifest_outputs["docx"]["sha256_available"])
            self.assertFalse(manifest_outputs["docx"]["path_exposed"])
            self.assertEqual(manifest_outputs["tex.zip"]["status"], "downloaded")
            self.assertTrue(recognition_plan["safety"]["download_integrity_hashed"])
            self.assertTrue(recognition_plan["safety"]["required_outputs_satisfied"])
            with zipfile.ZipFile(tex_artifact["path"]) as archive:
                self.assertIn("formula.tex", archive.namelist())
            self.assertTrue(any(item["source_type"] == "PDF" and "x^2+y^2=z^2" in item["latex"] for item in formulas))
            pdf_formula = next(item for item in formulas if item["source_type"] == "PDF")
            self.assertTrue(pdf_formula["original_image_ref"].startswith(f"mathpix://{pdf['id']}/"))
            self.assertIn("MathType 预览", pdf_formula["mathtype_preview"])
            self.assertEqual(pdf_formula["confidence"], 88)
            self.assertEqual(pdf_formula["confidence_threshold"], 95)
            self.assertEqual(pdf_formula["status"], "待确认")
            self.assertTrue(pdf_formula["review_required"])
            self.assertEqual(pdf_formula["review_reason"], "低于当前置信度阈值")
            self.assertEqual(pdf_formula["low_confidence_strategy"], "标记重识别")
            self.assertFalse(pdf_formula["keep_original_image"])
            self.assertEqual(report["formula_count"], len(formulas))
            with zipfile.ZipFile(BytesIO(build_formula_zip(report, []))) as archive:
                self.assertIn("x^2+y^2=z^2", archive.read("formulas.tex").decode("utf-8"))
            self.assertEqual(quality["mathpix_pdf"]["status"], "通过")
            self.assertEqual(quality["formula_confidence"]["status"], "需确认")
            self.assertEqual(quality["conversion_output"]["status"], "通过")
            queue = processor.mathpix_job_queue()
            self.assertEqual(queue["summary"]["completed"], 1)
            self.assertEqual(queue["jobs"][0]["request"]["conversion_formats"], ["docx", "tex.zip"])
            self.assertEqual(queue["jobs"][0]["recognition_plan"]["status"], "completed")
            self.assertTrue(queue["jobs"][0]["recognition_plan"]["submit_allowed"])
            self.assertEqual(queue["jobs"][0]["recognition_plan"]["upload_gate"]["status"], "completed")
            self.assertEqual(queue["jobs"][0]["recognition_plan"]["download_formats"], ["docx", "tex.zip"])
            self.assertEqual(queue["jobs"][0]["recognition_plan"]["download_results"]["docx"]["sha256"], docx_summary["sha256"])
            self.assertEqual(queue["jobs"][0]["recognition_plan"]["download_manifest"]["status"], "complete")
            self.assertTrue(queue["jobs"][0]["recognition_plan"]["download_manifest"]["all_required_outputs_hashed"])
            self.assertEqual(queue["jobs"][0]["recognition_plan"]["formula_review"]["confidence_threshold"], 95)
            self.assertEqual(queue["jobs"][0]["recognition_plan"]["formula_review"]["low_confidence_strategy"], "标记重识别")
            self.assertTrue(queue["jobs"][0]["request"]["formula_zip_requested"])
            self.assertIn("k12_ocr_language", queue["jobs"][0]["request"]["metadata_keys"])
            queue_outputs = {item["type"]: item for item in queue["jobs"][0]["outputs"]}
            self.assertEqual(queue_outputs["docx"]["path_display"], "本地路径已隐藏")
            self.assertEqual(queue_outputs["docx"]["sha256"], docx_summary["sha256"])
            self.assertEqual(queue_outputs["docx"]["file_size"], docx_summary["size"])
            self.assertEqual(queue_outputs["tex_zip"]["sha256"], tex_summary["sha256"])
            self.assertNotIn(str(Path(tex_artifact["path"]).parent), json.dumps(queue, ensure_ascii=False))
            payload = processor.local_task_payload(task["id"])
            pdf_action = next(action for action in payload["local_actions"] if action["type"] == "pdf_formula_mathtype")
            payload_contract = payload["formula_delivery"]
            action_contract = pdf_action["formula_delivery"]
            self.assertEqual(pdf_action["status"], "queued")
            self.assertEqual(pdf_action["formula_count"], 1)
            self.assertTrue(pdf_action["mathpix_jobs"][0]["tex_zip_available"])
            self.assertEqual(payload_contract["platform"], "Windows")
            self.assertEqual(payload_contract["compatibility_mode"], "platform-specific")
            self.assertFalse(payload_contract["platform_objects_cross_compatible"])
            self.assertTrue(payload_contract["native_object_requires_same_platform"])
            self.assertEqual(action_contract["platform"], "Windows")
            self.assertFalse(action_contract["platform_objects_cross_compatible"])
            readiness = payload["client_readiness"]
            self.assertEqual(readiness["status"], "ready_for_handoff")
            plan_action = next(action for action in payload["desktop_execution_plan"]["actions"] if action["type"] == "pdf_formula_mathtype")
            plan_contract = plan_action["formula_delivery"]
            self.assertEqual(plan_action["gate_status"], "ready")
            self.assertEqual(plan_action["required_capabilities"][0]["key"], "mathTypeAutomation")
            self.assertEqual(plan_action["formula_count"], 1)
            self.assertEqual(plan_action["mathpix_job_count"], 1)
            self.assertEqual(plan_action["output_contract"]["artifact_types"], ["docx", "tex.zip", "formula-report"])
            self.assertEqual(plan_contract["platform"], "Windows")
            self.assertFalse(plan_contract["platform_objects_cross_compatible"])
            self.assertTrue(plan_contract["native_object_requires_same_platform"])
            matrix = processor.acceptance_matrix()
            pdf_items = {
                item["key"]: item
                for group in matrix["groups"]
                if group["section"] == "17.6"
                for item in group["items"]
            }
            self.assertEqual(pdf_items["17.6.1"]["status"], "已覆盖")
            self.assertIn("已有 Mathpix DOCX 完成记录", pdf_items["17.6.1"]["evidence"])
            self.assertIn("已有完成 DOCX", pdf_items["17.6.1"]["current"])

    def test_pdf_to_word_task_option_wait_is_reported_without_authorizing_upload(self) -> None:
        class FakeMathpixClient:
            def __init__(self) -> None:
                self.wait_called = False

            def submit_pdf(self, path: Path, options: dict) -> dict:
                return {"pdf_id": "pdf_wait_task_option"}

            def wait_for_pdf(self, pdf_id: str, timeout_seconds: int) -> dict:
                self.wait_called = True
                return {"status": "completed", "percent_done": 100}

            def download_pdf_result(self, pdf_id: str, extension: str) -> bytes:
                return make_docx_bytes()

        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            store.update_settings(
                {
                    "allowExternalMathpixUpload": True,
                    "waitForMathpix": False,
                    "enableFormulaOcr": False,
                }
            )
            processor = TaskProcessor(store)
            pdf = processor.create_uploaded_file("scan.pdf", b"%PDF-1.4\n1 0 obj<< /Type /Page >>endobj\n")[0]
            fake_client = FakeMathpixClient()

            with patch("k12.processor.MathpixClient.from_environment", return_value=fake_client):
                task = processor.create_task(
                    {
                        "task_type": "pdf_to_word",
                        "file_ids": [pdf["id"]],
                        "options": {"waitForMathpix": True},
                    }
                )

            report = store.list_reports()[0]
            mathpix = report["analysis"]["mathpix"][0]
            recognition_plan = mathpix["recognition_plan"]
            queue_job = processor.mathpix_job_queue()["jobs"][0]["recognition_plan"]

            self.assertEqual(task["status"], "成功")
            self.assertTrue(fake_client.wait_called)
            self.assertEqual(mathpix["status"], "completed")
            self.assertTrue(recognition_plan["submit_contract"]["wait_for_completion"])
            self.assertEqual(recognition_plan["submit_contract"]["wait_for_completion_source"], "task_options")
            self.assertTrue(recognition_plan["external_upload_authorized"])
            self.assertEqual(recognition_plan["external_upload_authorization_source"], "settings.allowExternalMathpixUpload")
            self.assertEqual(recognition_plan["upload_gate"]["authorization_source"], "settings.allowExternalMathpixUpload")
            self.assertFalse(recognition_plan["upload_gate"]["task_options_can_authorize"])
            self.assertEqual(queue_job["wait_for_completion_source"], "task_options")
            self.assertEqual(queue_job["external_upload_authorization_source"], "settings.allowExternalMathpixUpload")
            self.assertFalse(queue_job["upload_gate"]["task_options_can_authorize"])

    def test_pdf_to_word_mathpix_tex_zip_download_failure_is_visible(self) -> None:
        class TexZipFailureMathpixClient:
            def submit_pdf(self, path: Path, options: dict) -> dict:
                return {"pdf_id": "pdf_tex_zip_failure"}

            def wait_for_pdf(self, pdf_id: str, timeout_seconds: int) -> dict:
                return {"status": "completed", "percent_done": 100}

            def download_pdf_result(self, pdf_id: str, extension: str) -> bytes:
                if extension == "tex.zip":
                    raise MathpixApiError("tex.zip download failed")
                return make_docx_bytes()

        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            store.update_settings({"allowExternalMathpixUpload": True, "waitForMathpix": True, "enableFormulaOcr": True})
            processor = TaskProcessor(store)
            pdf = processor.create_uploaded_file("formula-result.pdf", b"%PDF-1.4\n1 0 obj<< /Type /Page >>(Formula) Tj endobj\n")[0]

            with patch("k12.processor.MathpixClient.from_environment", return_value=TexZipFailureMathpixClient()):
                processor.create_task({"task_type": "pdf_to_word", "file_ids": [pdf["id"]]})

            report = store.list_reports()[0]
            mathpix = report["analysis"]["mathpix"][0]
            artifacts = report["analysis"]["artifacts"]
            artifact_statuses = [artifact["status"] for artifact in artifacts]
            queue = processor.mathpix_job_queue()

            self.assertEqual(mathpix["status"], "completed")
            self.assertEqual(mathpix["tex_zip_status"], "download_failed")
            self.assertIn("tex.zip download failed", mathpix["tex_zip_message"])
            self.assertIn("docx", mathpix["output_summaries"])
            self.assertNotIn("tex_zip", mathpix["output_summaries"])
            self.assertEqual(mathpix["recognition_plan"]["download_failures"][0]["type"], "tex.zip")
            download_manifest = mathpix["recognition_plan"]["download_manifest"]
            manifest_outputs = {item["type"]: item for item in download_manifest["outputs"]}
            self.assertEqual(download_manifest["status"], "partial_failed")
            self.assertEqual(download_manifest["required_output_count"], 2)
            self.assertEqual(download_manifest["satisfied_required_count"], 1)
            self.assertEqual(download_manifest["failed_required_count"], 1)
            self.assertFalse(download_manifest["all_required_outputs_satisfied"])
            self.assertFalse(download_manifest["all_required_outputs_hashed"])
            self.assertEqual(manifest_outputs["docx"]["status"], "downloaded")
            self.assertEqual(manifest_outputs["tex.zip"]["status"], "failed")
            self.assertIn("tex.zip download failed", manifest_outputs["tex.zip"]["failure_message"])
            self.assertFalse(mathpix["recognition_plan"]["safety"]["download_integrity_hashed"])
            self.assertFalse(mathpix["recognition_plan"]["safety"]["required_outputs_satisfied"])
            self.assertIn("成功", artifact_statuses)
            self.assertIn("mathpix_tex_zip_download_failed", artifact_statuses)
            self.assertEqual(queue["summary"]["completed"], 1)
            self.assertEqual(queue["summary"]["partial_failed"], 1)
            self.assertEqual(queue["summary"]["failed"], 0)
            self.assertEqual(queue["jobs"][0]["tex_zip_status"], "download_failed")
            self.assertEqual(queue["jobs"][0]["recognition_plan"]["download_failures"][0]["type"], "tex.zip")
            self.assertEqual(queue["jobs"][0]["recognition_plan"]["download_manifest"]["status"], "partial_failed")
            self.assertEqual([item["type"] for item in queue["jobs"][0]["outputs"]], ["docx"])

    def test_pdf_to_word_empty_mathpix_docx_download_is_reported_as_failure(self) -> None:
        class EmptyDocxMathpixClient:
            def submit_pdf(self, path: Path, options: dict) -> dict:
                return {"pdf_id": "pdf_empty_docx"}

            def wait_for_pdf(self, pdf_id: str, timeout_seconds: int) -> dict:
                return {"status": "completed", "percent_done": 100}

            def download_pdf_result(self, pdf_id: str, extension: str) -> bytes:
                return b""

        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            store.update_settings(
                {
                    "allowExternalMathpixUpload": True,
                    "waitForMathpix": True,
                    "enableFormulaOcr": False,
                }
            )
            processor = TaskProcessor(store)
            pdf = processor.create_uploaded_file("empty-result.pdf", b"%PDF-1.4\n1 0 obj<< /Type /Page >>endobj\n")[0]

            with patch("k12.processor.MathpixClient.from_environment", return_value=EmptyDocxMathpixClient()):
                processor.create_task({"task_type": "pdf_to_word", "file_ids": [pdf["id"]]})

            report = store.list_reports()[0]
            mathpix = report["analysis"]["mathpix"][0]
            artifact = report["analysis"]["artifacts"][0]
            queue = processor.mathpix_job_queue()

            self.assertEqual(mathpix["status"], "download_failed")
            self.assertEqual(mathpix["recognition_plan"]["status"], "download_failed")
            self.assertEqual(mathpix["recognition_plan"]["upload_gate"]["status"], "download_failed")
            self.assertEqual(mathpix["output_summaries"], {})
            download_manifest = mathpix["recognition_plan"]["download_manifest"]
            self.assertEqual(download_manifest["status"], "failed")
            self.assertEqual(download_manifest["required_output_count"], 1)
            self.assertEqual(download_manifest["failed_required_count"], 1)
            self.assertFalse(download_manifest["all_required_outputs_satisfied"])
            self.assertEqual(download_manifest["outputs"][0]["type"], "docx")
            self.assertEqual(download_manifest["outputs"][0]["status"], "failed")
            self.assertIn("下载结果为空", mathpix["message"])
            self.assertEqual(artifact["status"], "download_failed")
            self.assertIn("下载结果为空", artifact["message"])
            self.assertEqual(queue["summary"]["failed"], 1)
            self.assertEqual(queue["jobs"][0]["outputs"], [])

    def test_mathpix_pdf_options_keep_conversion_formats_separate_from_local_ocr_metadata(self) -> None:
        default_options = MathpixClient._pdf_options({})
        legacy_options = MathpixClient._pdf_options({"docx": True, "tex.zip": True, "lines.json": False})
        options = MathpixClient._pdf_options(
            {
                "docx": True,
                "tex.zip": False,
                "language": "中文",
                "precision_mode": "高精度",
                "table_ocr": True,
                "conversion_formats": {"docx": True, "tex.zip": True, "lines.json": False, "../bad": True, "": True},
                "enable_tables_fallback": True,
                "include_page_info": True,
                "metadata": {"k12_ocr_language": "中文"},
            }
        )

        self.assertEqual(default_options["conversion_formats"], {"docx": True})
        self.assertEqual(legacy_options["conversion_formats"], {"docx": True, "tex.zip": True})
        self.assertEqual(options["conversion_formats"], {"docx": True, "tex.zip": True})
        self.assertNotIn("../bad", options["conversion_formats"])
        self.assertNotIn("", options["conversion_formats"])
        self.assertTrue(options["enable_tables_fallback"])
        self.assertTrue(options["include_page_info"])
        self.assertEqual(options["metadata"]["k12_ocr_language"], "中文")
        self.assertNotIn("language", options)
        self.assertNotIn("precision_mode", options)
        self.assertNotIn("table_ocr", options)

    def test_mathpix_multipart_body_sanitizes_disposition_names(self) -> None:
        body, content_type = MathpixClient._multipart_body(
            {"options_json": "{}"},
            {"file": ('scan"\r\nX-Injected: yes.pdf', b"%PDF-1.4", "application/pdf")},
        )
        text = body.decode("utf-8", errors="ignore")

        self.assertIn("multipart/form-data; boundary=k12-mathpix-", content_type)
        self.assertIn('name="file"; filename="scan___X-Injected: yes.pdf"', text)
        self.assertNotIn('filename="scan"\r\nX-Injected: yes.pdf"', text)
        self.assertNotIn("\r\nX-Injected: yes.pdf", text)

    def test_mathpix_request_json_wraps_invalid_json_responses(self) -> None:
        client = MathpixClient("id", "key")
        with patch.object(client, "_request_bytes", return_value=b"<html>bad gateway</html>"):
            with self.assertRaises(MathpixApiError) as raised:
                client.get_pdf_status("pdf_123")
        self.assertIn("非 JSON", str(raised.exception))

        with patch.object(client, "_request_bytes", return_value=b'["unexpected"]'):
            with self.assertRaises(MathpixApiError) as raised:
                client.get_pdf_status("pdf_123")
        self.assertIn("格式异常", str(raised.exception))

    def test_mathpix_http_errors_are_compact_and_redacted_for_reports(self) -> None:
        client = MathpixClient("app-id", "super-secret-key")
        detail = client._error_detail((f"app-id\nsuper-secret-key\n{'x' * 900}").encode("utf-8"))
        self.assertLessEqual(len(detail), MATHPIX_ERROR_DETAIL_LIMIT + 3)
        self.assertTrue(detail.endswith("..."))
        self.assertNotIn("\n", detail)
        self.assertNotIn("app-id", detail)
        self.assertNotIn("super-secret-key", detail)
        self.assertEqual(client._error_detail(b""), "无响应正文")

        body = BytesIO((f"<html>\nsuper-secret-key\n{'x' * 900}</html>").encode("utf-8"))
        error = HTTPError("https://api.mathpix.com/v3/pdf/pdf_123", 502, "Bad Gateway", {}, body)
        with patch("urllib.request.urlopen", side_effect=error):
            with self.assertRaises(MathpixApiError) as raised:
                client.get_pdf_status("pdf_123")
        message = str(raised.exception)
        self.assertIn("Mathpix API 返回 502:", message)
        self.assertLess(len(message), 700)
        self.assertNotIn("\n", message)
        self.assertNotIn("super-secret-key", message)

    def test_mathpix_status_restricts_pdf_id_path(self) -> None:
        client = MathpixClient("id", "key")
        with patch.object(client, "_request_json", return_value={"status": "completed"}) as request:
            self.assertEqual(client.get_pdf_status(" pdf_123-ABC "), {"status": "completed"})
        request.assert_called_once_with("GET", "/v3/pdf/pdf_123-ABC")

        with self.assertRaises(MathpixApiError):
            client.get_pdf_status("../secret")
        with self.assertRaises(MathpixApiError):
            client.get_pdf_status("pdf/123")

    def test_mathpix_download_result_restricts_result_path_parts(self) -> None:
        client = MathpixClient("id", "key")
        with patch.object(client, "_request_bytes", return_value=b"docx") as request:
            self.assertEqual(client.download_pdf_result("pdf_123-ABC", "DOCX"), b"docx")
        request.assert_called_once_with("GET", "/v3/pdf/pdf_123-ABC.docx")

        with patch.object(client, "_request_bytes", return_value=b"zip") as request:
            self.assertEqual(client.download_pdf_result("pdf_123", "tex.zip"), b"zip")
        request.assert_called_once_with("GET", "/v3/pdf/pdf_123.tex.zip")

        with self.assertRaises(MathpixApiError):
            client.download_pdf_result("../secret", "docx")
        with self.assertRaises(MathpixApiError):
            client.download_pdf_result("pdf_123", "html")

    def test_mathpix_tex_formula_parser_handles_common_mathpix_shapes(self) -> None:
        text = r"""
        \documentclass{article}
        Inline formula \(a^2+b^2=c^2\)
        $$E=mc^2$$
        \begin{align}
        x &= y + z \\
        \alpha &= \frac{1}{2}
        \end{align}
        plain_line=\sqrt{n}
        \section{Not a formula}
        """

        formulas = TaskProcessor._extract_latex_expressions(text)

        self.assertIn("a^2+b^2=c^2", formulas)
        self.assertIn("E=mc^2", formulas)
        self.assertTrue(any("x &= y + z" in item and "\\frac{1}{2}" in item for item in formulas))
        self.assertIn("plain_line=\\sqrt{n}", formulas)
        self.assertFalse(any("\\documentclass" in item for item in formulas))
        self.assertFalse(any("\\section" in item for item in formulas))

    def test_mathpix_tex_formula_parser_skips_oversized_tex_entries(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            pdf = processor.create_uploaded_file("formula.pdf", b"%PDF-1.4\n1 0 obj<< /Type /Page >>(Formula) Tj endobj\n")[0]
            tex_zip = Path(tmp) / "mathpix-formulas.zip"
            with zipfile.ZipFile(tex_zip, "w", zipfile.ZIP_DEFLATED) as archive:
                archive.writestr("too-large.tex", ("x" * (1024 * 1024 + 1)) + "\n$$too_big=1$$")
                archive.writestr("nested/small.tex", "$$small=1$$")
            job = {
                "file_id": pdf["id"],
                "file_name": pdf["file_name"],
                "engine": "Mathpix",
                "status": "completed",
                "outputs": {"tex_zip": str(tex_zip)},
            }

            formulas = processor._mathpix_formula_items([pdf], [job])

            self.assertEqual([item["latex"] for item in formulas], ["small=1"])
            self.assertEqual(formulas[0]["position"], "Mathpix OCR / small.tex")
            self.assertEqual(formulas[0]["original_image_ref"], f"mathpix://{pdf['id']}/small.tex")

    def test_pdf_to_word_skips_mathpix_when_ocr_is_disabled(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            store.update_settings(
                {
                    "allowExternalMathpixUpload": True,
                    "enableTextOcr": False,
                    "enableFormulaOcr": False,
                    "enableTableOcr": False,
                }
            )
            processor = TaskProcessor(store)
            pdf = processor.create_uploaded_file("scan.pdf", b"%PDF-1.4\n1 0 obj<< /Type /Page >>endobj\n")[0]
            with patch("k12.processor.MathpixClient.from_environment") as client_factory:
                processor.create_task({"task_type": "pdf_to_word", "file_ids": [pdf["id"]]})
            report = store.list_reports()[0]
            mathpix = report["analysis"]["mathpix"][0]
            artifact = report["analysis"]["artifacts"][0]
            quality = {item["id"]: item for item in report["qualityChecks"]}

            client_factory.assert_not_called()
            self.assertEqual(mathpix["status"], "ocr_disabled")
            self.assertEqual(mathpix["recognition_plan"]["status"], "blocked_ocr_disabled")
            self.assertFalse(mathpix["recognition_plan"]["submit_allowed"])
            self.assertEqual(mathpix["recognition_plan"]["upload_gate"]["status"], "blocked_ocr_disabled")
            self.assertIn("all_ocr_disabled", mathpix["recognition_plan"]["upload_gate"]["blocking_reasons"])
            self.assertFalse(mathpix["recognition_plan"]["upload_gate"]["submit_allowed"])
            self.assertEqual(mathpix["recognition_plan"]["submit_contract"]["conversion_formats"], ["docx"])
            self.assertEqual(artifact["status"], "ocr_disabled")
            self.assertEqual(quality["mathpix_pdf"]["status"], "需确认")

    def test_pdf_to_word_engine_is_forced_to_mathpix_contract(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            store.update_settings({"pdfToWordEngine": "LocalOCR", "allowExternalMathpixUpload": False})
            processor = TaskProcessor(store)
            pdf = processor.create_uploaded_file("worksheet.pdf", b"%PDF-1.4\n1 0 obj<< /Type /Page >>endobj\n")[0]

            checks = processor.preflight_checks({"task_type": "pdf_to_word", "file_ids": [pdf["id"]]})
            mathpix_preflight = next(item for item in checks if item["id"] == "mathpix_authorization")
            task = processor.create_task({"task_type": "pdf_to_word", "file_ids": [pdf["id"]]})
            report = store.list_reports()[0]
            mathpix = report["analysis"]["mathpix"][0]
            artifact = report["analysis"]["artifacts"][0]

            self.assertEqual(store.get_settings()["pdfToWordEngine"], "Mathpix")
            self.assertEqual(mathpix_preflight["status"], "需确认")
            self.assertIn("Mathpix", mathpix_preflight["message"])
            self.assertEqual(mathpix["engine"], "Mathpix")
            self.assertEqual(mathpix["status"], "authorization_required")
            self.assertEqual(artifact["status"], "authorization_required")
            self.assertEqual(artifact["conversion_settings"]["pdf_to_word_engine"], "Mathpix")

    def test_settings_api_cannot_switch_pdf_to_word_away_from_mathpix(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            handler = make_handler(store, processor)

            settings = processor.update_settings({"pdfToWordEngine": "LocalOCR", "allowExternalMathpixUpload": False})
            self.assertEqual(settings["pdfToWordEngine"], "Mathpix")

            K12RequestHandler._handle_api_get(handler, "/api/settings", {})
            settings_response = json.loads(handler.wfile.getvalue().decode("utf-8"))
            self.assertEqual(settings_response["settings"]["pdfToWordEngine"], "Mathpix")

            pdf = processor.create_uploaded_file("api-worksheet.pdf", b"%PDF-1.4\n1 0 obj<< /Type /Page >>endobj\n")[0]
            task = processor.create_task({"task_type": "pdf_to_word", "file_ids": [pdf["id"]]})
            report = store.list_reports()[0]
            mathpix = report["analysis"]["mathpix"][0]
            artifact = report["analysis"]["artifacts"][0]

            self.assertIn(task["status"], {"成功", "失败"})
            self.assertEqual(mathpix["engine"], "Mathpix")
            self.assertEqual(mathpix["recognition_plan"]["engine"], "Mathpix")
            self.assertFalse(mathpix["recognition_plan"]["submit_allowed"])
            self.assertEqual(mathpix["recognition_plan"]["upload_gate"]["status"], "blocked_authorization")
            self.assertEqual(artifact["status"], "authorization_required")

    def test_settings_api_redacts_local_security_token(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            store.update_settings({"localSecurityToken": "secret-token", "exposeLocalPaths": True})
            processor = TaskProcessor(store)
            handler = make_handler(store, processor)

            K12RequestHandler._handle_api_get(handler, "/api/settings", {})
            settings_response = json.loads(handler.wfile.getvalue().decode("utf-8"))

            self.assertEqual(settings_response["settings"]["localSecurityToken"], "")
            self.assertTrue(settings_response["settings"]["localSecurityTokenConfigured"])
            self.assertNotIn("secret-token", json.dumps(settings_response, ensure_ascii=False))

            body = json.dumps({"localSecurityToken": "next-token"}).encode("utf-8")
            put_handler = make_handler(store, processor)
            put_handler.path = "/api/settings"
            put_handler.headers = {"Content-Length": str(len(body)), "X-K12-Token": "secret-token"}
            put_handler.rfile = BytesIO(body)

            K12RequestHandler.do_PUT(put_handler)
            put_response = json.loads(put_handler.wfile.getvalue().decode("utf-8"))

            self.assertEqual(put_handler.status, 200)
            self.assertEqual(store.get_settings()["localSecurityToken"], "next-token")
            self.assertEqual(put_response["settings"]["localSecurityToken"], "")
            self.assertTrue(put_response["settings"]["localSecurityTokenConfigured"])
            self.assertNotIn("next-token", json.dumps(put_response, ensure_ascii=False))

    def test_omml_dependency_is_found_and_copied(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            dependency_dir = store.data_dir / "dependency-source"
            dependency_dir.mkdir()
            source = dependency_dir / "OMML2MML.XSL"
            source.write_text("<xsl:stylesheet />", encoding="utf-8")
            store.update_settings({"ommlSearchPaths": str(dependency_dir)})
            processor = TaskProcessor(store)
            word = processor.create_uploaded_file("lesson.docx", make_docx_bytes())[0]
            self.assertTrue(word["missing_omml_dependency"])
            task = processor.create_task({"task_type": "omml_to_mathtype", "file_ids": [word["id"]]})
            report = store.list_reports()[0]
            dependency = report["analysis"]["ommlDependencies"][0]
            updated_word = store.get_file(word["id"])
            self.assertEqual(task["status"], "成功")
            self.assertEqual(dependency["found_status"], "已找到")
            self.assertEqual(dependency["copy_status"], "成功")
            self.assertFalse(updated_word["missing_omml_dependency"])
            self.assertTrue((store.uploads_dir / "OMML2MML.XSL").exists())

    def test_missing_omml_dependency_is_reported_as_quality_issue(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            word = processor.create_uploaded_file("lesson.docx", make_docx_bytes())[0]
            processor.create_task({"task_type": "formula_precheck", "file_ids": [word["id"]]})
            report = store.list_reports()[0]
            quality = {item["id"]: item for item in report["qualityChecks"]}
            failure_categories = {item["category"] for item in report["failureRows"]}

            self.assertTrue(store.get_file(word["id"])["missing_omml_dependency"])
            self.assertEqual(report["file_omml_missing_count"], 1)
            self.assertEqual(quality["omml_dependencies"]["status"], "失败")
            self.assertIn("OMML 依赖", failure_categories)

    def test_omml_conversion_prompt_requires_user_decision_by_default(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            word = processor.create_uploaded_file("lesson.docx", make_docx_bytes())[0]
            task = processor.create_task({"task_type": "formula_precheck", "file_ids": [word["id"]]})
            report = store.list_reports()[0]
            prompt = report["analysis"]["ommlConversionPrompts"][0]
            quality = {item["id"]: item for item in report["qualityChecks"]}
            payload = processor.local_task_payload(task["id"])
            omml_action = next(action for action in payload["local_actions"] if action["type"] == "omml_mathtype")

            self.assertEqual(prompt["status"], "需确认")
            self.assertIsNone(prompt["convert_to_mathtype"])
            self.assertTrue(prompt["requires_user_confirmation"])
            self.assertTrue(prompt["preserve_original_formula"])
            self.assertEqual(prompt["local_action"], "ask_user")
            self.assertEqual(prompt["decision_source"], "未确认")
            self.assertEqual(quality["omml_conversion_prompt"]["status"], "需确认")
            self.assertEqual(omml_action["status"], "awaiting_confirmation")
            self.assertEqual(omml_action["conversion_prompts"][0]["status"], "需确认")
            self.assertIn("OMML 转换确认", Path(report["txt_path"]).read_text(encoding="utf-8"))
            self.assertIn("检测到 Word 自带公式", Path(report["html_path"]).read_text(encoding="utf-8"))
            with zipfile.ZipFile(report["xlsx_path"]) as archive:
                sheet_xml = archive.read("xl/worksheets/sheet1.xml").decode("utf-8")
                self.assertIn("OMML 转换确认", sheet_xml)
                self.assertIn("需确认", sheet_xml)

    def test_omml_conversion_prompt_records_confirmed_conversion(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            dependency_dir = store.data_dir / "dependency-source"
            dependency_dir.mkdir()
            (dependency_dir / "OMML2MML.XSL").write_text("<xsl:stylesheet />", encoding="utf-8")
            store.update_settings({"ommlSearchPaths": str(dependency_dir)})
            processor = TaskProcessor(store)
            word = processor.create_uploaded_file("lesson.docx", make_docx_bytes())[0]
            task = processor.create_task(
                {
                    "task_type": "omml_to_mathtype",
                    "file_ids": [word["id"]],
                    "options": {"convertOmmlToMathType": True},
                }
            )
            report = store.list_reports()[0]
            prompt = report["analysis"]["ommlConversionPrompts"][0]
            quality = {item["id"]: item for item in report["qualityChecks"]}
            payload = processor.local_task_payload(task["id"])
            omml_action = next(action for action in payload["local_actions"] if action["type"] == "omml_mathtype")

            self.assertEqual(prompt["status"], "已选择转换")
            self.assertTrue(prompt["convert_to_mathtype"])
            self.assertFalse(prompt["requires_user_confirmation"])
            self.assertEqual(prompt["local_action"], "convert_omml_to_mathtype")
            self.assertEqual(prompt["decision_source"], "任务选项")
            self.assertEqual(prompt["dependency_status"], "成功")
            self.assertEqual(quality["omml_conversion_prompt"]["status"], "通过")
            self.assertEqual(omml_action["status"], "queued")
            self.assertTrue(omml_action["conversion_prompts"][0]["preserve_original_formula"])
            self.assertTrue(any("OMML 转 MathType 确认" in log["message"] for log in store.list_logs(task["id"], limit=1000)))

    def test_omml_retry_conversion_is_handed_to_local_payload(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            dependency_dir = store.data_dir / "dependency-source"
            dependency_dir.mkdir()
            (dependency_dir / "OMML2MML.XSL").write_text("<xsl:stylesheet />", encoding="utf-8")
            store.update_settings(
                {
                    "ommlSearchPaths": str(dependency_dir),
                    "localSecurityToken": "secret-token",
                    "localClientPlatform": "Windows",
                }
            )
            processor = TaskProcessor(store)
            processor.record_local_client_heartbeat(
                {
                    "client_id": "desktop-omml",
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
            word = processor.create_uploaded_file("lesson.docx", make_docx_bytes())[0]
            task = processor.create_task(
                {
                    "task_type": "omml_to_mathtype",
                    "file_ids": [word["id"]],
                    "options": {"convertOmmlToMathType": True},
                }
            )
            report = store.list_reports()[0]
            dependency = report["analysis"]["ommlDependencies"][0]
            annotation = processor.save_omml_annotation(
                {"report_id": report["id"], "dependency_id": dependency["id"], "status": "重新转换", "retry_conversion": True}
            )

            payload = processor.local_task_payload(task["id"])
            omml_action = next(action for action in payload["local_actions"] if action["type"] == "omml_mathtype")
            plan_action = next(action for action in payload["desktop_execution_plan"]["actions"] if action["type"] == "omml_mathtype")
            operations = [step["operation"] for step in plan_action["steps"]]
            summary = summarize_payload(payload)

            self.assertEqual(omml_action["status"], "retry_queued")
            self.assertEqual(omml_action["retry_requests"][0]["annotation_id"], annotation["id"])
            self.assertEqual(omml_action["retry_requests"][0]["dependency_id"], dependency["id"])
            self.assertTrue(omml_action["retry_requests"][0]["retry_conversion"])
            self.assertEqual(plan_action["gate_status"], "ready")
            self.assertEqual(plan_action["retry_request_count"], 1)
            self.assertIn("omml.retry_conversion", operations)
            self.assertEqual(summary["actions"][0]["retry_request_count"], 1)

    def test_omml_annotation_is_persisted_for_report_dependency(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            word = processor.create_uploaded_file("lesson.docx", make_docx_bytes())[0]
            processor.create_task({"task_type": "omml_to_mathtype", "file_ids": [word["id"]]})
            report = store.list_reports()[0]
            dependency = report["analysis"]["ommlDependencies"][0]
            annotation = processor.save_omml_annotation(
                {
                    "report_id": report["id"],
                    "dependency_id": dependency["id"],
                    "status": "手动指定依赖",
                    "keep_omml": True,
                    "retry_conversion": True,
                    "manual_omml_path": str(Path(tmp) / "OMML2MML.XSL"),
                    "note": "单测 OMML 校正",
                }
            )
            annotations = processor.list_omml_annotations(report["id"])

            self.assertEqual(annotation["status"], "手动指定依赖")
            self.assertTrue(annotation["keep_omml"])
            self.assertTrue(annotation["retry_conversion"])
            self.assertEqual(annotations[0]["dependency_id"], dependency["id"])
            with self.assertRaises(ValueError):
                processor.save_omml_annotation({"report_id": report["id"], "dependency_id": "missing", "status": "已修复"})
            with self.assertRaises(ValueError):
                processor.save_omml_annotation({"report_id": report["id"], "dependency_id": dependency["id"], "status": "未知"})

            handler = make_handler(store, processor)
            K12RequestHandler._handle_api_get(handler, "/api/omml-annotations", {"report_id": [report["id"]]})
            response = json.loads(handler.wfile.getvalue().decode("utf-8"))
            self.assertEqual(handler.status, 200)
            self.assertEqual(response["annotations"][0]["id"], annotation["id"])
            self.assertEqual(response["annotations"][0]["manual_omml_path"], "本地路径已隐藏/OMML2MML.XSL")
            self.assertTrue(response["annotations"][0]["manual_omml_path_available"])
            self.assertNotIn(str(Path(tmp)), json.dumps(response, ensure_ascii=False))

            manual_post_handler = make_handler(store, processor)
            K12RequestHandler._handle_api_post(
                manual_post_handler,
                "/api/omml-annotations",
                {
                    "report_id": report["id"],
                    "dependency_id": dependency["id"],
                    "status": "手动指定依赖",
                    "manual_omml_path": str(Path(tmp) / "OMML2MML.XSL"),
                    "retry_conversion": True,
                },
            )
            manual_post_response = json.loads(manual_post_handler.wfile.getvalue().decode("utf-8"))
            self.assertEqual(manual_post_handler.status, 201)
            self.assertEqual(manual_post_response["annotation"]["status"], "手动指定依赖")
            self.assertEqual(manual_post_response["annotation"]["manual_omml_path"], "本地路径已隐藏/OMML2MML.XSL")
            self.assertTrue(manual_post_response["annotation"]["manual_omml_path_available"])
            csv_handler = make_handler(store, processor)
            K12RequestHandler._send_report_omml_failures(csv_handler, report["id"])
            csv_text = csv_handler.wfile.getvalue().decode("utf-8")
            self.assertEqual(csv_handler.status, 200)
            self.assertIn(("Content-Type", "text/csv; charset=utf-8"), csv_handler.output_headers)
            self.assertIn(("Content-Disposition", f'attachment; filename="{report["id"]}-omml-failures.csv"'), csv_handler.output_headers)
            self.assertIn("dependency_id", csv_text)
            self.assertIn("lesson.docx", csv_text)
            self.assertIn("手动指定依赖", csv_text)
            self.assertIn("本地路径已隐藏/OMML2MML.XSL", csv_text)
            self.assertNotIn(str(Path(tmp)), csv_text)

            store.update_settings({"exposeLocalPaths": True})
            exposed_handler = make_handler(store, processor)
            K12RequestHandler._send_report_omml_failures(exposed_handler, report["id"])
            exposed_csv = exposed_handler.wfile.getvalue().decode("utf-8")
            self.assertIn(str(Path(tmp) / "OMML2MML.XSL"), exposed_csv)
            store.update_settings({"exposeLocalPaths": False})

            post_handler = make_handler(store, processor)
            K12RequestHandler._handle_api_post(
                post_handler,
                "/api/omml-annotations",
                {"report_id": report["id"], "dependency_id": dependency["id"], "status": "重新转换", "retry_conversion": True},
            )
            post_response = json.loads(post_handler.wfile.getvalue().decode("utf-8"))
            self.assertEqual(post_handler.status, 201)
            self.assertEqual(post_response["annotation"]["status"], "重新转换")
            self.assertTrue(post_response["annotation"]["retry_conversion"])
            route_handler = make_handler(store, processor)
            K12RequestHandler._handle_api_get(route_handler, f"/api/reports/{report['id']}/omml-failures.csv", {})
            self.assertEqual(route_handler.status, 200)
            self.assertIn("重新转换", route_handler.wfile.getvalue().decode("utf-8"))
            self.assertTrue(processor.delete_omml_annotation(annotation["id"]))

    def test_delete_report_removes_omml_annotations(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            word = processor.create_uploaded_file("lesson.docx", make_docx_bytes())[0]
            processor.create_task({"task_type": "omml_to_mathtype", "file_ids": [word["id"]]})
            report = store.list_reports()[0]
            dependency = report["analysis"]["ommlDependencies"][0]
            processor.save_omml_annotation({"report_id": report["id"], "dependency_id": dependency["id"], "status": "转换失败"})

            result = store.delete_report(report["id"])

            self.assertTrue(result["deleted"])
            self.assertEqual(result["annotations_deleted"], 1)
            self.assertEqual(processor.list_omml_annotations(report["id"]), [])

    def test_manual_omml_dependency_is_used_when_auto_search_is_disabled(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            dependency_dir = store.data_dir / "manual-source"
            dependency_dir.mkdir()
            source = dependency_dir / "OMML2MML.XSL"
            source.write_text("<xsl:stylesheet />", encoding="utf-8")
            store.update_settings({"autoSearchOmml": False, "allowManualOmml": True, "manualOmmlPath": str(source)})
            processor = TaskProcessor(store)
            word = processor.create_uploaded_file("lesson.docx", make_docx_bytes())[0]
            processor.create_task({"task_type": "omml_to_mathtype", "file_ids": [word["id"]]})
            report = store.list_reports()[0]
            dependency = report["analysis"]["ommlDependencies"][0]
            quality = {item["id"]: item for item in report["qualityChecks"]}
            self.assertEqual(dependency["found_status"], "手动选择")
            self.assertEqual(dependency["copy_status"], "成功")
            self.assertEqual(quality["omml_dependencies"]["status"], "通过")
            self.assertTrue((store.uploads_dir / "OMML2MML.XSL").exists())

    def test_manual_omml_dependency_without_source_path_copies_to_outputs(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            dependency_dir = store.data_dir / "manual-source"
            dependency_dir.mkdir()
            source = dependency_dir / "OMML2MML.XSL"
            source.write_text("<xsl:stylesheet />", encoding="utf-8")
            store.update_settings({"autoSearchOmml": False, "allowManualOmml": True, "manualOmmlPath": str(source)})
            processor = TaskProcessor(store)
            word = processor.create_file({"file_name": "metadata.omml.docx", "file_size": 4096})
            processor.create_task({"task_type": "omml_to_mathtype", "file_ids": [word["id"]]})
            report = store.list_reports()[0]
            dependency = report["analysis"]["ommlDependencies"][0]
            self.assertEqual(dependency["found_status"], "手动选择")
            self.assertEqual(dependency["omml_target_path"], str(store.outputs_dir / "OMML2MML.XSL"))
            self.assertTrue((store.outputs_dir / "OMML2MML.XSL").exists())

    def test_small_image_scan_extracts_ooxml_media(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            word = processor.create_uploaded_file("with-image.docx", make_docx_bytes({"word/media/icon.png": make_png_bytes(32, 24)}))[0]
            task = processor.create_task({"task_type": "small_image_scan", "file_ids": [word["id"]]})
            report = store.list_reports()[0]
            image = report["analysis"]["smallImages"][0]
            self.assertEqual(task["status"], "成功")
            self.assertEqual(image["width"], 32)
            self.assertEqual(image["height"], 24)
            self.assertTrue(image["asset_url"].startswith("/api/assets/images/"))
            self.assertTrue(Path(image["image_path"]).exists())
            image_path = Path(image["image_path"])
            image_path.unlink()

            reexported = processor.reexport_image_asset({"report_id": report["id"], "image_id": image["id"]})
            updated_image = reexported["image"]
            updated_report = store.get_report(report["id"])

            self.assertTrue(reexported["reexported"])
            self.assertEqual(updated_image["export_status"], "已重新导出")
            self.assertTrue(updated_image["asset_url"].startswith("/api/assets/images/"))
            self.assertTrue(Path(updated_image["image_path"]).exists())
            self.assertNotEqual(updated_image["image_path"], str(image_path))
            self.assertEqual(updated_report["analysis"]["smallImages"][0]["export_status"], "已重新导出")
            self.assertIn("已重新导出", Path(report["json_path"]).read_text(encoding="utf-8"))

            handler = make_handler(store, processor)
            K12RequestHandler._handle_api_post(handler, "/api/image-assets/reexport", {"report_id": report["id"], "image_id": image["id"]})
            route_response = json.loads(handler.wfile.getvalue().decode("utf-8"))
            self.assertEqual(handler.status, 200)
            self.assertTrue(route_response["reexported"])
            self.assertEqual(route_response["image"]["export_status"], "已重新导出")

    def test_small_image_scan_reads_pdf_image_object_dimensions(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            pdf = processor.create_uploaded_file("scan.pdf", make_pdf_with_image_bytes(40, 30))[0]
            summary = pdf["content_summary"]
            self.assertEqual(summary["imageObjects"], 1)
            self.assertEqual(summary["pdfImages"][0]["width"], 40)
            self.assertEqual(summary["pdfImages"][0]["height"], 30)

            task = processor.create_task({"task_type": "small_image_scan", "file_ids": [pdf["id"]]})
            report = store.list_reports()[0]
            image = report["analysis"]["smallImages"][0]

            self.assertEqual(task["status"], "成功")
            self.assertEqual(image["location"], "PDF 图片对象 1")
            self.assertEqual(image["width"], 40)
            self.assertEqual(image["height"], 30)
            self.assertEqual(image["image_type"], "jpg")
            self.assertTrue(image["asset_url"].startswith("/api/assets/images/"))
            self.assertTrue(Path(image["image_path"]).exists())
            self.assertIn("PDF Image XObject 1", image["source_name"])

            handler = make_handler(store, processor)
            K12RequestHandler._send_report_images(handler, report["id"])
            self.assertEqual(handler.status, 200)
            with zipfile.ZipFile(BytesIO(handler.wfile.getvalue())) as archive:
                manifest = archive.read("manifest.csv").decode("utf-8")
                self.assertIn("PDF 图片对象 1", manifest)
                self.assertIn('"jpg","图标"', manifest)

    def test_pdf_small_image_scan_respects_size_thresholds(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            store.update_settings({"smallImageMaxWidth": 32, "smallImageMaxHeight": 32, "smallImageMaxArea": 1024})
            processor = TaskProcessor(store)
            pdf = processor.create_uploaded_file("scan.pdf", make_pdf_with_image_bytes(40, 30))[0]

            processor.create_task({"task_type": "small_image_scan", "file_ids": [pdf["id"]]})
            report = store.list_reports()[0]

            self.assertEqual(report["analysis"]["smallImages"], [])

    def test_small_image_settings_filter_sources_and_record_export_format(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            store.update_settings(
                {
                    "includeHeaderFooterImages": False,
                    "includeTransparentImages": False,
                    "includeDuplicateImages": False,
                    "imageExportFormat": "PNG",
                }
            )
            processor = TaskProcessor(store)
            icon = make_png_bytes(32, 24)
            alpha = make_png_bytes(32, 24, alpha=True)
            word = processor.create_uploaded_file(
                "with-images.docx",
                make_docx_bytes(
                    {
                        "word/media/header-logo.png": icon,
                        "word/media/icon.png": icon,
                        "word/media/icon-copy.png": icon,
                        "word/media/transparent.png": alpha,
                    }
                ),
            )[0]
            processor.create_task({"task_type": "small_image_scan", "file_ids": [word["id"]]})
            report = store.list_reports()[0]
            images = report["analysis"]["smallImages"]

            self.assertEqual(len(images), 1)
            self.assertEqual(images[0]["source_name"], "word/media/icon.png")
            self.assertEqual(images[0]["image_type"], "png")
            self.assertEqual(images[0]["area"], 768)
            self.assertEqual(images[0]["export_format"], "PNG")
            self.assertFalse(images[0]["is_header_footer"])
            self.assertFalse(images[0]["is_transparent"])
            self.assertFalse(images[0]["is_duplicate"])
            self.assertEqual(images[0]["duplicate_check_status"], "已判断")
            self.assertEqual(images[0]["duplicate_fallback"], "重复判断失败时保留原始结果")
            text_report = Path(report["txt_path"]).read_text(encoding="utf-8")
            html_report = Path(report["html_path"]).read_text(encoding="utf-8")
            self.assertIn("面积 768", text_report)
            self.assertIn("图片格式 png", text_report)
            self.assertIn("重复判断 已判断", text_report)
            self.assertIn("<th>面积</th>", html_report)
            self.assertIn("<th>图片格式</th>", html_report)
            self.assertIn("<th>重复判断</th>", html_report)
            with zipfile.ZipFile(report["xlsx_path"]) as archive:
                sheet_xml = archive.read("xl/worksheets/sheet1.xml").decode("utf-8")
                self.assertIn("面积", sheet_xml)
                self.assertIn("图片格式", sheet_xml)
                self.assertIn("重复判断", sheet_xml)
                self.assertIn("异常兜底", sheet_xml)
                self.assertIn("768", sheet_xml)
                self.assertIn("png", sheet_xml)

            handler = make_handler(store, processor)
            K12RequestHandler._send_report_images(handler, report["id"])
            self.assertEqual(handler.status, 200)
            with zipfile.ZipFile(BytesIO(handler.wfile.getvalue())) as archive:
                manifest = archive.read("manifest.csv").decode("utf-8")
                self.assertIn("width,height,area,image_format,suspected_type", manifest)
                self.assertIn("duplicate_check_status,duplicate_fallback", manifest)
                self.assertIn("export_status,export_message,reexported_at", manifest)
                self.assertIn("word/media/icon.png", manifest)
                self.assertIn('"768","png","图标"', manifest)
                self.assertIn('"已判断","重复判断失败时保留原始结果"', manifest)

    def test_small_image_report_includes_prd_classification_flags(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            qr_image = make_png_bytes(32, 32)
            transparent = make_png_bytes(24, 24, alpha=True)
            word = processor.create_uploaded_file(
                "with-image-flags.docx",
                make_docx_bytes(
                    {
                        "word/media/qr-code.png": qr_image,
                        "word/media/qr-copy.png": qr_image,
                        "word/media/watermark.png": make_png_bytes(30, 24),
                        "word/media/transparent.png": transparent,
                    }
                ),
            )[0]
            processor.create_task({"task_type": "small_image_scan", "file_ids": [word["id"]]})
            report = store.list_reports()[0]
            images = report["analysis"]["smallImages"]
            by_name = {image["source_name"]: image for image in images}

            self.assertEqual(len(images), 4)
            self.assertEqual(by_name["word/media/qr-code.png"]["area"], 1024)
            self.assertTrue(by_name["word/media/qr-code.png"]["is_qrcode_like"])
            self.assertTrue(by_name["word/media/qr-copy.png"]["is_duplicate"])
            self.assertEqual(by_name["word/media/qr-copy.png"]["duplicate_check_status"], "已判断")
            self.assertEqual(by_name["word/media/qr-copy.png"]["duplicate_fallback"], "重复判断失败时保留原始结果")
            self.assertTrue(by_name["word/media/watermark.png"]["is_watermark"])
            self.assertTrue(by_name["word/media/transparent.png"]["is_transparent"])
            text_report = Path(report["txt_path"]).read_text(encoding="utf-8")
            html_report = Path(report["html_path"]).read_text(encoding="utf-8")
            self.assertIn("判定标记", text_report)
            self.assertIn("疑似类型 二维码", text_report)
            self.assertIn("重复", text_report)
            self.assertIn("<th>判定标记</th>", html_report)
            with zipfile.ZipFile(report["xlsx_path"]) as archive:
                sheet_xml = archive.read("xl/worksheets/sheet1.xml").decode("utf-8")
            self.assertIn("判定标记", sheet_xml)
            self.assertIn("水印", sheet_xml)
            self.assertIn("透明", sheet_xml)

    def test_small_image_extraction_failures_are_reported(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            word = processor.create_uploaded_file(
                "with-broken-image.docx",
                make_docx_bytes({"word/media/broken.png": b"not-a-valid-image"}),
            )[0]

            task = processor.create_task({"task_type": "small_image_scan", "file_ids": [word["id"]]})
            report = store.list_reports()[0]
            errors = report["analysis"]["imageExtractionErrors"]
            quality = {item["id"]: item for item in report["qualityChecks"]}
            failure_categories = {item["category"] for item in report["failureRows"]}

            self.assertEqual(task["status"], "成功")
            self.assertEqual(report["small_image_count"], 0)
            self.assertEqual(report["image_extraction_error_count"], 1)
            self.assertEqual(report["error_count"], 1)
            self.assertEqual(errors[0]["source_name"], "word/media/broken.png")
            self.assertIn("图片尺寸或格式无法识别", errors[0]["message"])
            self.assertIn("图片提取", failure_categories)
            self.assertEqual(quality["small_images"]["status"], "失败")
            self.assertIn("提取失败 1", quality["small_images"]["metric"])
            failure_csv = Path(report["failure_csv_path"]).read_text(encoding="utf-8")
            self.assertIn("图片提取", failure_csv)
            self.assertIn("word/media/broken.png", failure_csv)

    def test_image_annotation_is_persisted_for_report_image(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            word = processor.create_uploaded_file("with-image.docx", make_docx_bytes({"word/media/icon.png": make_png_bytes(32, 24)}))[0]
            processor.create_task({"task_type": "small_image_scan", "file_ids": [word["id"]]})
            report = store.list_reports()[0]
            image = report["analysis"]["smallImages"][0]
            annotation = processor.save_image_annotation(
                {
                    "report_id": report["id"],
                    "image_id": image["id"],
                    "status": "误判",
                    "note": "单测标记",
                }
            )
            annotations = processor.list_image_annotations(report["id"])
            self.assertEqual(annotation["status"], "误判")
            self.assertEqual(annotations[0]["image_id"], image["id"])
            confirmed = processor.save_image_annotation(
                {
                    "report_id": report["id"],
                    "image_id": image["id"],
                    "status": "类型已确认",
                    "label": "图标",
                    "note": "单测确认类型",
                }
            )
            self.assertEqual(confirmed["status"], "类型已确认")
            self.assertEqual(confirmed["label"], "图标")
            unknown_location = processor.save_image_annotation(
                {
                    "report_id": report["id"],
                    "image_id": image["id"],
                    "status": "位置未知",
                    "note": "单测定位失败",
                }
            )
            self.assertEqual(unknown_location["status"], "位置未知")
            replacement = processor.save_image_annotation(
                {
                    "report_id": report["id"],
                    "image_id": image["id"],
                    "status": "替换待处理",
                    "replacement_file_name": "new-icon.png",
                    "note": "单测替换",
                }
            )
            replacement_bytes = make_png_bytes(20, 18)
            replacement_upload = processor.save_image_replacement(
                {
                    "report_id": report["id"],
                    "image_id": image["id"],
                    "file_name": "new-icon.png",
                    "mime_type": "image/png",
                    "content_base64": base64.b64encode(replacement_bytes).decode("ascii"),
                }
            )
            delete_mark = processor.save_image_annotation(
                {
                    "report_id": report["id"],
                    "image_id": image["id"],
                    "status": "删除待处理",
                    "note": "单测删除",
                }
            )
            self.assertEqual(replacement["replacement_file_name"], "new-icon.png")
            self.assertEqual(replacement_upload["status"], "替换待处理")
            self.assertEqual(replacement_upload["replacement_file_size"], len(replacement_bytes))
            self.assertEqual(replacement_upload["replacement_sha256"], hashlib.sha256(replacement_bytes).hexdigest())
            self.assertTrue(replacement_upload["replacement_asset_url"].startswith("/api/assets/replacements/"))
            self.assertEqual(delete_mark["id"], annotation["id"])
            self.assertEqual(processor.list_image_annotations(report["id"])[0]["status"], "删除待处理")
            with self.assertRaises(ValueError):
                processor.save_image_annotation({"report_id": report["id"], "image_id": image["id"], "status": "未知"})
            with self.assertRaises(ValueError):
                processor.save_image_replacement(
                    {
                        "report_id": report["id"],
                        "image_id": image["id"],
                        "file_name": "bad.txt",
                        "content_base64": base64.b64encode(b"bad").decode("ascii"),
                    }
                )
            post_handler = make_handler(store, processor)
            K12RequestHandler._handle_api_post(
                post_handler,
                "/api/image-annotations",
                {"report_id": report["id"], "image_id": image["id"], "status": "类型已确认", "label": "公式"},
            )
            post_response = json.loads(post_handler.wfile.getvalue().decode("utf-8"))
            self.assertEqual(post_handler.status, 201)
            self.assertEqual(post_response["annotation"]["label"], "公式")
            replacement_handler = make_handler(store, processor)
            K12RequestHandler._handle_api_post(
                replacement_handler,
                "/api/image-annotations/replacement",
                {
                    "report_id": report["id"],
                    "image_id": image["id"],
                    "file_name": "browser-choice.png",
                    "mime_type": "image/png",
                    "content_base64": "data:image/png;base64," + base64.b64encode(replacement_bytes).decode("ascii"),
                },
            )
            replacement_response = json.loads(replacement_handler.wfile.getvalue().decode("utf-8"))
            self.assertEqual(replacement_handler.status, 201)
            self.assertEqual(replacement_response["annotation"]["replacement_file_name"], "browser-choice.png")
            asset_handler = make_handler(store, processor)
            asset_name = Path(replacement_response["annotation"]["replacement_asset_url"]).name
            K12RequestHandler._send_replacement_asset(asset_handler, asset_name)
            self.assertEqual(asset_handler.status, 200)
            self.assertEqual(asset_handler.wfile.getvalue(), replacement_bytes)
            image_zip_handler = make_handler(store, processor)
            K12RequestHandler._send_report_images(image_zip_handler, report["id"])
            with zipfile.ZipFile(BytesIO(image_zip_handler.wfile.getvalue())) as archive:
                manifest = archive.read("manifest.csv").decode("utf-8")
                self.assertIn("confirmed_type", manifest)
                self.assertIn("duplicate_check_status", manifest)
                self.assertIn("duplicate_fallback", manifest)
                self.assertIn("export_status", manifest)
                self.assertIn("export_message", manifest)
                self.assertIn('"替换待处理"', manifest)
                self.assertIn('"browser-choice.png"', manifest)
            image_xlsx_handler = make_handler(store, processor)
            K12RequestHandler._send_report_image_manifest_xlsx(image_xlsx_handler, report["id"])
            self.assertEqual(image_xlsx_handler.status, 200)
            self.assertIn(("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"), image_xlsx_handler.output_headers)
            self.assertIn(("Content-Disposition", f'attachment; filename="{report["id"]}-images.xlsx"'), image_xlsx_handler.output_headers)
            with zipfile.ZipFile(BytesIO(image_xlsx_handler.wfile.getvalue())) as archive:
                sheet_xml = archive.read("xl/worksheets/sheet1.xml").decode("utf-8")
                workbook_xml = archive.read("xl/workbook.xml").decode("utf-8")
                self.assertIn("Image Manifest", workbook_xml)
                self.assertIn("confirmed_type", sheet_xml)
                self.assertIn("duplicate_check_status", sheet_xml)
                self.assertIn("duplicate_fallback", sheet_xml)
                self.assertIn("export_status", sheet_xml)
                self.assertIn("export_message", sheet_xml)
                self.assertIn("with-image.docx", sheet_xml)
                self.assertIn("word/media/icon.png", sheet_xml)
                self.assertIn("browser-choice.png", sheet_xml)
                self.assertIn("替换待处理", sheet_xml)
            route_handler = make_handler(store, processor)
            K12RequestHandler._handle_api_get(route_handler, f"/api/reports/{report['id']}/images.xlsx", {})
            self.assertEqual(route_handler.status, 200)
            self.assertTrue(processor.delete_image_annotation(annotation["id"]))

    def test_delete_report_removes_report_cache_and_annotations_only(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            word = processor.create_uploaded_file("with-image.docx", make_docx_bytes({"word/media/icon.png": make_png_bytes(32, 24)}))[0]
            task = processor.create_task({"task_type": "small_image_scan", "file_ids": [word["id"]]})
            report = store.list_reports()[0]
            image = report["analysis"]["smallImages"][0]
            formula = report["analysis"]["formulas"][0]
            processor.save_image_annotation({"report_id": report["id"], "image_id": image["id"], "status": "误判"})
            processor.save_formula_annotation({"report_id": report["id"], "formula_id": formula["id"], "status": "已确认"})
            report_paths = [Path(report[key]) for key in ("html_path", "json_path", "pdf_path", "xlsx_path", "txt_path", "failure_csv_path")]
            image_path = Path(image["image_path"])

            result = store.delete_report(report["id"])

            self.assertTrue(result["deleted"])
            self.assertGreaterEqual(result["report_files_deleted"], 6)
            self.assertEqual(result["image_files_deleted"], 1)
            self.assertEqual(result["annotations_deleted"], 2)
            self.assertEqual(store.list_reports(), [])
            self.assertIsNotNone(store.get_task(task["id"]))
            self.assertIsNotNone(store.get_file(word["id"]))
            self.assertTrue(Path(word["storage_path"]).exists())
            self.assertTrue(all(not path.exists() for path in report_paths))
            self.assertFalse(image_path.exists())
            self.assertEqual(processor.list_image_annotations(report["id"]), [])
            self.assertEqual(processor.list_formula_annotations(report["id"]), [])

    def test_formula_annotation_is_persisted_for_report_formula(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            word = processor.create_uploaded_file("formula.docx", make_docx_bytes())[0]
            processor.create_task({"task_type": "formula_precheck", "file_ids": [word["id"]]})
            report = store.list_reports()[0]
            formula = report["analysis"]["formulas"][0]
            annotation = processor.save_formula_annotation(
                {
                    "report_id": report["id"],
                    "formula_id": formula["id"],
                    "status": "已修正",
                    "latex": "a^2+b^2=c^2",
                    "note": "单测修正",
                }
            )
            annotations = processor.list_formula_annotations(report["id"])
            self.assertEqual(annotation["status"], "已修正")
            self.assertEqual(annotation["latex"], "a^2+b^2=c^2")
            self.assertEqual(annotations[0]["formula_id"], formula["id"])
            self.assertTrue(processor.delete_formula_annotation(annotation["id"]))

    def test_formula_rerecognition_records_mathpix_gate_without_uploading(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            store.update_settings({"allowExternalMathpixUpload": False})
            processor = TaskProcessor(store)
            report = store.save_report(
                {
                    "id": "report_formula_retry",
                    "task_id": "task_formula_retry",
                    "report_type": "PDF 公式报告",
                    "created_at": "2026-06-24T00:00:00+00:00",
                    "files": [{"id": "file_pdf", "file_name": "scan.pdf"}],
                    "analysis": {
                        "formulas": [
                            {
                                "id": "formula_pdf_retry",
                                "file_id": "file_pdf",
                                "source_type": "PDF",
                                "position": "第 1 页",
                                "latex": "x+y",
                                "confidence": 62,
                                "status": "待确认",
                                "original_image_ref": "/api/assets/images/task/formula.png",
                            }
                        ]
                    },
                }
            )

            annotation = processor.save_formula_annotation(
                {"report_id": report["id"], "formula_id": "formula_pdf_retry", "status": "重新识别", "note": "重新跑公式 OCR"}
            )
            request = annotation["recognition_request"]

            self.assertTrue(annotation["retry_recognition"])
            self.assertEqual(annotation["recognition_status"], "blocked")
            self.assertEqual(request["schema_version"], "k12.formulaRecognitionRequest.v1")
            self.assertEqual(request["engine"], "Mathpix")
            self.assertTrue(request["external_upload_required"])
            self.assertFalse(request["external_upload_authorized"])
            self.assertFalse(request["external_upload_performed"])
            self.assertFalse(request["task_options_can_authorize_external_upload"])
            self.assertIn("external_mathpix_upload_not_authorized", request["blocking_reasons"])
            self.assertIn("授权外部上传", request["next_step"])

            formula_zip = build_formula_zip(report, [annotation])
            with zipfile.ZipFile(BytesIO(formula_zip)) as archive:
                formula_json = json.loads(archive.read("formulas.json").decode("utf-8"))
                manifest = archive.read("manifest.csv").decode("utf-8")
            exported = formula_json["formulas"][0]
            self.assertTrue(exported["retry_recognition"])
            self.assertEqual(exported["recognition_status"], "blocked")
            self.assertEqual(exported["recognition_request"]["engine"], "Mathpix")
            self.assertIn("recognition_status", manifest)

            formula_xlsx = build_formula_xlsx(report, [annotation])
            with zipfile.ZipFile(BytesIO(formula_xlsx)) as archive:
                sheet_xml = archive.read("xl/worksheets/sheet1.xml").decode("utf-8")
            self.assertIn("retry_recognition", sheet_xml)
            self.assertIn("recognition_status", sheet_xml)

    def test_formula_rerecognition_routes_latex_to_local_client_contract(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            report = store.save_report(
                {
                    "id": "report_latex_retry",
                    "task_id": "task_latex_retry",
                    "report_type": "公式报告",
                    "created_at": "2026-06-24T00:00:00+00:00",
                    "files": [{"id": "file_word", "file_name": "lesson.docx"}],
                    "analysis": {"formulas": [{"id": "formula_latex", "file_id": "file_word", "source_type": "LaTeX", "latex": "a=b", "confidence": 90}]},
                }
            )

            annotation = processor.save_formula_annotation({"report_id": report["id"], "formula_id": "formula_latex", "status": "重识别"})
            request = annotation["recognition_request"]

            self.assertEqual(annotation["status"], "重新识别")
            self.assertEqual(request["engine"], "local-client")
            self.assertEqual(request["status"], "queued_for_local_client")
            self.assertFalse(request["external_upload_required"])
            self.assertIn("同平台 Office/MathType 本地客户端", request["required_environment"])

    def test_formula_rerecognition_is_handed_to_local_payload(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            store.update_settings({"localSecurityToken": "secret-token", "localClientPlatform": "Windows"})
            processor = TaskProcessor(store)
            processor.record_local_client_heartbeat(
                {
                    "client_id": "desktop-formula",
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
            word = processor.create_file({"file_name": "lesson.docx", "file_size": 2048})
            task = store.save_task(Task(task_type="formula_precheck", execute_mode="local", file_ids=[word["id"]]).to_dict())
            report = store.save_report(
                {
                    "id": "report_formula_local_retry",
                    "task_id": task["id"],
                    "report_type": "公式报告",
                    "created_at": "2026-06-24T00:00:00+00:00",
                    "files": [word],
                    "analysis": {"formulas": [{"id": "formula_latex_retry", "file_id": word["id"], "source_type": "LaTeX", "position": "第 1 段", "latex": "a=b", "confidence": 90}]},
                }
            )
            annotation = processor.save_formula_annotation({"report_id": report["id"], "formula_id": "formula_latex_retry", "status": "重新识别"})

            payload = processor.local_task_payload(task["id"])
            omml_action = next(action for action in payload["local_actions"] if action["type"] == "omml_mathtype")
            plan_action = next(action for action in payload["desktop_execution_plan"]["actions"] if action["type"] == "omml_mathtype")
            operations = [step["operation"] for step in plan_action["steps"]]
            summary = summarize_payload(payload)
            execution = build_dry_run_execution_summary(payload)
            native_request = build_native_execution_request(payload, platform_name="Windows", allow_native_execution=True)

            self.assertEqual(omml_action["status"], "recognition_queued")
            self.assertEqual(omml_action["recognition_requests"][0]["annotation_id"], annotation["id"])
            self.assertEqual(omml_action["recognition_requests"][0]["engine"], "local-client")
            self.assertEqual(plan_action["recognition_request_count"], 1)
            self.assertIn("formula.rerecognize_local", operations)
            self.assertEqual(summary["actions"][0]["recognition_request_count"], 1)
            self.assertEqual(summary["desktop_execution_plan"]["actions"][0]["recognition_request_count"], 1)
            self.assertEqual(execution["actions"][0]["recognition_request_count"], 1)
            self.assertEqual(native_request["actions"][0]["recognition_request_count"], 1)

    def test_pdf_formula_rerecognition_request_is_exposed_in_pdf_formula_action(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            store.update_settings({"allowExternalMathpixUpload": False, "localSecurityToken": "secret-token"})
            processor = TaskProcessor(store)
            pdf = processor.create_file({"file_name": "scan.pdf", "file_size": 4096})
            task = store.save_task(Task(task_type="pdf_to_word", execute_mode="hybrid", file_ids=[pdf["id"]]).to_dict())
            report = store.save_report(
                {
                    "id": "report_pdf_formula_retry",
                    "task_id": task["id"],
                    "report_type": "PDF 公式报告",
                    "created_at": "2026-06-24T00:00:00+00:00",
                    "files": [pdf],
                    "analysis": {"formulas": [{"id": "formula_pdf_retry", "file_id": pdf["id"], "source_type": "PDF", "position": "第 1 页", "latex": "x+y", "confidence": 62}]},
                }
            )
            annotation = processor.save_formula_annotation({"report_id": report["id"], "formula_id": "formula_pdf_retry", "status": "重新识别"})

            payload = processor.local_task_payload(task["id"])
            pdf_action = next(action for action in payload["local_actions"] if action["type"] == "pdf_formula_mathtype")
            plan_action = next(action for action in payload["desktop_execution_plan"]["actions"] if action["type"] == "pdf_formula_mathtype")
            operations = [step["operation"] for step in plan_action["steps"]]
            summary = summarize_payload(payload)

            self.assertEqual(pdf_action["recognition_requests"][0]["annotation_id"], annotation["id"])
            self.assertEqual(pdf_action["recognition_requests"][0]["engine"], "Mathpix")
            self.assertFalse(pdf_action["recognition_requests"][0]["external_upload_performed"])
            self.assertIn("external_mathpix_upload_not_authorized", pdf_action["recognition_requests"][0]["blocking_reasons"])
            self.assertEqual(plan_action["recognition_request_count"], 1)
            self.assertIn("mathpix.rerecognize_formula", operations)
            self.assertEqual(summary["actions"][0]["recognition_request_count"], 1)
            self.assertEqual(summary["desktop_execution_plan"]["actions"][0]["recognition_request_count"], 1)

    def test_formula_annotation_rejects_unsupported_status(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            word = processor.create_uploaded_file("formula.docx", make_docx_bytes())[0]
            processor.create_task({"task_type": "formula_precheck", "file_ids": [word["id"]]})
            report = store.list_reports()[0]
            formula = report["analysis"]["formulas"][0]
            with self.assertRaises(ValueError):
                processor.save_formula_annotation({"report_id": report["id"], "formula_id": formula["id"], "status": "自动执行外部识别"})

    def test_formula_zip_exports_json_tex_mml_and_manifest(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            word = processor.create_uploaded_file("formula.docx", make_docx_bytes())[0]
            processor.create_task({"task_type": "formula_precheck", "file_ids": [word["id"]]})
            report = store.list_reports()[0]
            formula = report["analysis"]["formulas"][0]
            annotation = processor.save_formula_annotation(
                {
                    "report_id": report["id"],
                    "formula_id": formula["id"],
                    "status": "已修正",
                    "latex": "a^2+b^2=c^2",
                    "mathml": "<math><mi>a</mi></math>",
                }
            )
            data = build_formula_zip(report, [annotation])
            with zipfile.ZipFile(BytesIO(data)) as archive:
                names = archive.namelist()
                self.assertIn("manifest.csv", names)
                self.assertIn("formulas.json", names)
                self.assertIn("formulas.tex", names)
                self.assertIn("formulas.mml", names)
                self.assertTrue(any(name.startswith("tex/") and name.endswith(".tex") for name in names))
                self.assertTrue(any(name.startswith("mml/") and name.endswith(".mml") for name in names))
                manifest = archive.read("manifest.csv").decode("utf-8")
                formula_json = json.loads(archive.read("formulas.json").decode("utf-8"))
                tex = archive.read("formulas.tex").decode("utf-8")
                mml = archive.read("formulas.mml").decode("utf-8")
                self.assertIn("formula.docx", manifest)
                self.assertEqual(formula_json["formulas"][0]["latex"], "a^2+b^2=c^2")
                self.assertEqual(formula_json["formulas"][0]["mathml"], "<math><mi>a</mi></math>")
                self.assertIn("original_image_ref", formula_json["formulas"][0])
                self.assertIn("mathtype_preview", formula_json["formulas"][0])
                self.assertIn("a^2+b^2=c^2", tex)
                self.assertIn("<math><mi>a</mi></math>", mml)
            formula_xlsx = build_formula_xlsx(report, [annotation])
            with zipfile.ZipFile(BytesIO(formula_xlsx)) as archive:
                sheet_xml = archive.read("xl/worksheets/sheet1.xml").decode("utf-8")
                workbook_xml = archive.read("xl/workbook.xml").decode("utf-8")
                self.assertIn("Formula Results", workbook_xml)
                self.assertIn("formula_id", sheet_xml)
                self.assertIn("formula.docx", sheet_xml)
                self.assertIn("a^2+b^2=c^2", sheet_xml)
                self.assertIn("&lt;math&gt;&lt;mi&gt;a&lt;/mi&gt;&lt;/math&gt;", sheet_xml)
                self.assertIn("已修正", sheet_xml)
            xlsx_handler = make_handler(store, processor)
            K12RequestHandler._send_report_formula_xlsx(xlsx_handler, report["id"])
            self.assertEqual(xlsx_handler.status, 200)
            self.assertIn(("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"), xlsx_handler.output_headers)
            self.assertIn(("Content-Disposition", f'attachment; filename="{report["id"]}-formulas.xlsx"'), xlsx_handler.output_headers)
            route_handler = make_handler(store, processor)
            K12RequestHandler._handle_api_get(route_handler, f"/api/reports/{report['id']}/formulas.xlsx", {})
            self.assertEqual(route_handler.status, 200)

    def test_formula_missing_position_is_recorded_as_exception(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            word = processor.create_uploaded_file("formula.docx", make_docx_bytes())[0]
            task = Task(task_type="formula_precheck", execute_mode="本地模式", file_ids=[word["id"]]).to_dict()
            report = processor.report_builder.build(
                task,
                [word],
                {
                    "formulas": [
                        {
                            "id": "formula-missing-position",
                            "file_id": word["id"],
                            "page_index": 1,
                            "position": "",
                            "source_type": "LaTeX",
                            "latex": "x+y",
                            "confidence": 91,
                            "status": "成功",
                            "format_status": "已格式化",
                        }
                    ],
                    "macros": [],
                    "smallImages": [],
                },
            )
            formula = report["analysis"]["formulas"][0]
            quality = {item["id"]: item for item in report["qualityChecks"]}
            failure_categories = {item["category"] for item in report["failureRows"]}

            self.assertEqual(formula["position_status"], "异常位置")
            self.assertIn("异常位置 / formula.docx", formula["position"])
            self.assertEqual(report["formula_position_issue_count"], 1)
            self.assertIn("公式位置", failure_categories)
            self.assertEqual(quality["formula_positions"]["status"], "失败")
            self.assertIn("公式位置丢失", Path(report["failure_csv_path"]).read_text(encoding="utf-8"))
            self.assertIn("<th>位置状态</th>", Path(report["html_path"]).read_text(encoding="utf-8"))
            self.assertIn("位置状态 异常位置", Path(report["txt_path"]).read_text(encoding="utf-8"))
            data = build_formula_zip(report, [])
            with zipfile.ZipFile(BytesIO(data)) as archive:
                manifest = archive.read("manifest.csv").decode("utf-8")
                formula_json = json.loads(archive.read("formulas.json").decode("utf-8"))
                self.assertIn("position_status,position_issue", manifest)
                self.assertEqual(formula_json["formulas"][0]["position_status"], "异常位置")
            formula_xlsx = build_formula_xlsx(report, [])
            with zipfile.ZipFile(BytesIO(formula_xlsx)) as archive:
                sheet_xml = archive.read("xl/worksheets/sheet1.xml").decode("utf-8")
                self.assertIn("position_status", sheet_xml)
                self.assertIn("异常位置", sheet_xml)

    def test_formula_annotation_requires_formula_in_report(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            word = processor.create_uploaded_file("formula.docx", make_docx_bytes())[0]
            processor.create_task({"task_type": "formula_precheck", "file_ids": [word["id"]]})
            report = store.list_reports()[0]
            with self.assertRaises(ValueError):
                processor.save_formula_annotation({"report_id": report["id"], "formula_id": "missing", "status": "已确认"})

    def test_formula_bulk_confirm_uses_confidence_threshold_and_skips_existing(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            report = store.save_report(
                {
                    "id": "report_formula_bulk",
                    "task_id": "task_formula_bulk",
                    "file_id": "file_formula_bulk",
                    "report_type": "公式预检报告",
                    "report_path": "",
                    "created_at": "2026-06-23T00:00:00+00:00",
                    "analysis": {
                        "formulas": [
                            {"id": "formula_high_1", "latex": "x=1", "mathml": "<math />", "confidence": 96},
                            {"id": "formula_high_2", "latex": "y=2", "mathml": "", "confidence": 88},
                            {"id": "formula_low", "latex": "z=3", "mathml": "", "confidence": 70},
                        ]
                    },
                }
            )
            existing = processor.save_formula_annotation(
                {
                    "report_id": report["id"],
                    "formula_id": "formula_high_2",
                    "status": "已修正",
                    "latex": "y=2",
                }
            )
            result = processor.bulk_confirm_formulas({"report_id": report["id"], "min_confidence": 80, "only_unannotated": True})
            annotations = processor.list_formula_annotations(report["id"])
            self.assertEqual(result["confirmed_count"], 1)
            self.assertEqual(result["annotations"][0]["formula_id"], "formula_high_1")
            self.assertTrue(any(item["id"] == existing["id"] and item["status"] == "已修正" for item in annotations))
            self.assertFalse(any(item["formula_id"] == "formula_low" for item in annotations))

    def test_word_to_ppt_generates_artifact(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            word = processor.create_uploaded_file("lesson.docx", make_docx_bytes())[0]
            task = processor.create_task({"task_type": "word_to_ppt", "file_ids": [word["id"]]})
            report = store.list_reports()[0]
            artifact = report["analysis"]["artifacts"][0]
            self.assertEqual(task["status"], "成功")
            self.assertEqual(artifact["output_type"], "pptx")
            self.assertTrue(Path(artifact["path"]).exists())
            with zipfile.ZipFile(artifact["path"]) as archive:
                self.assertIn("ppt/presentation.xml", archive.namelist())

    def test_word_to_ppt_records_source_object_preservation(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            document_xml = (
                '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
                'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">'
                '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>第一章</w:t></w:r></w:p>'
                '<w:p><w:r><w:t>函数图像</w:t></w:r></w:p>'
                '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>表格内容</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
                '<m:oMath><m:r><m:t>x=1</m:t></m:r></m:oMath>'
                '<w:p><w:r><w:t>MathType 对象</w:t></w:r></w:p>'
                '</w:document>'
            )
            word = processor.create_uploaded_file(
                "objects.docx",
                make_docx_bytes(
                    extra_files={
                        "word/media/photo.png": make_png_bytes(40, 30),
                        "word/embeddings/oleObject1.bin": b"mathtype",
                    },
                    document_xml=document_xml,
                ),
            )[0]
            processor.create_task({"task_type": "word_to_ppt", "file_ids": [word["id"]]})
            report = store.list_reports()[0]
            artifact = report["analysis"]["artifacts"][0]
            preservation = artifact["object_preservation"]

            self.assertTrue(preservation["has_objects"])
            self.assertEqual(preservation["source"]["images"], 1)
            self.assertEqual(preservation["source"]["tables"], 1)
            self.assertEqual(preservation["source"]["omml_formulas"], 1)
            self.assertEqual(preservation["source"]["mathtype_objects"], 1)
            self.assertEqual(preservation["statuses"]["images"], "计划保留图片对象")
            self.assertIn("写入对象保留清单", artifact["message"])
            with zipfile.ZipFile(artifact["path"]) as archive:
                slide_xml = "\n".join(
                    archive.read(name).decode("utf-8")
                    for name in archive.namelist()
                    if name.startswith("ppt/slides/slide") and name.endswith(".xml")
                )
            self.assertIn("对象保留清单", slide_xml)
            self.assertIn("图片 1 个", slide_xml)
            self.assertIn("表格 1 个", slide_xml)
            self.assertIn("OMML 1 个", slide_xml)
            text_report = Path(report["txt_path"]).read_text(encoding="utf-8")
            html_report = Path(report["html_path"]).read_text(encoding="utf-8")
            self.assertIn("对象保留", text_report)
            self.assertIn("计划保留表格结构", html_report)
            with zipfile.ZipFile(report["xlsx_path"]) as archive:
                sheet_xml = archive.read("xl/worksheets/sheet1.xml").decode("utf-8")
            self.assertIn("对象保留", sheet_xml)
            self.assertIn("MathType 1 个", sheet_xml)

    def test_report_comparison_previews_source_and_conversion_output(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            word = processor.create_uploaded_file("lesson.docx", make_docx_bytes())[0]
            processor.create_task({"task_type": "word_to_ppt", "file_ids": [word["id"]]})
            report = store.list_reports()[0]

            comparison = processor.report_comparison(report["id"])

            self.assertEqual(comparison["status"], "ready")
            self.assertEqual(comparison["item_count"], 1)
            item = comparison["items"][0]
            self.assertEqual(item["source_file"], "lesson.docx")
            self.assertEqual(item["output_type"], "pptx")
            self.assertEqual(item["source_preview"]["file_type"], "Word")
            self.assertEqual(item["output_preview"]["file_type"], "PPT")
            self.assertIn("page_delta", item["diff"])
            self.assertTrue(item["artifact_url"].startswith("/api/artifacts/"))
            quality = {entry["id"]: entry for entry in report["qualityChecks"]}
            self.assertIn("preview_consistency", quality)
            self.assertIn("比对 1 个", quality["preview_consistency"]["metric"])
            self.assertIn("lesson.docx", quality["preview_consistency"]["message"])
            self.assertIn("转换前后预览一致性", Path(report["txt_path"]).read_text(encoding="utf-8"))
            self.assertIn("转换前后预览一致性", Path(report["html_path"]).read_text(encoding="utf-8"))

    def test_report_comparison_is_empty_without_successful_artifacts(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            report = store.save_report(
                {
                    "id": "report_empty_comparison",
                    "task_id": "task_empty_comparison",
                    "file_id": "file_empty",
                    "report_type": "处理报告",
                    "report_path": "",
                    "created_at": "2026-06-24T00:00:00+00:00",
                    "files": [],
                    "analysis": {"artifacts": []},
                }
            )

            comparison = processor.report_comparison(report["id"])

            self.assertEqual(comparison["status"], "empty")
            self.assertEqual(comparison["items"], [])

    def test_layout_annotation_is_persisted_for_report(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            word = processor.create_uploaded_file("lesson.docx", make_docx_bytes())[0]
            processor.create_task({"task_type": "word_to_ppt", "file_ids": [word["id"]]})
            report = store.list_reports()[0]
            annotation = processor.save_layout_annotation(
                {
                    "report_id": report["id"],
                    "location": "第 1 页 / 标题 1",
                    "issue_type": "标题层级",
                    "status": "待校正",
                    "page_index": 1,
                    "before": "正文",
                    "after": "一级标题",
                    "recommendation": "恢复标题层级",
                    "note": "单测排版校正",
                }
            )
            updated = processor.save_layout_annotation(
                {
                    "report_id": report["id"],
                    "location": "第 1 页 / 标题 1",
                    "issue_type": "标题层级",
                    "status": "已校正",
                }
            )
            annotations = processor.list_layout_annotations(report["id"])

            self.assertEqual(updated["id"], annotation["id"])
            self.assertEqual(annotations[0]["status"], "已校正")
            self.assertEqual(annotations[0]["issue_type"], "标题层级")
            with self.assertRaises(ValueError):
                processor.save_layout_annotation({"report_id": report["id"], "location": "第 1 页", "issue_type": "未知"})

            handler = make_handler(store, processor)
            K12RequestHandler._handle_api_get(handler, "/api/layout-annotations", {"report_id": [report["id"]]})
            response = json.loads(handler.wfile.getvalue().decode("utf-8"))
            self.assertEqual(handler.status, 200)
            self.assertEqual(response["annotations"][0]["id"], annotation["id"])
            post_handler = make_handler(store, processor)
            K12RequestHandler._handle_api_post(
                post_handler,
                "/api/layout-annotations",
                {
                    "report_id": report["id"],
                    "location": "第 3 页 / 公式 2",
                    "issue_type": "公式位置",
                    "status": "需本地客户端处理",
                    "before": "公式偏移到页脚",
                    "after": "恢复到正文段落",
                    "recommendation": "使用本地 Office 复核锚点",
                },
            )
            post_response = json.loads(post_handler.wfile.getvalue().decode("utf-8"))
            self.assertEqual(post_handler.status, 201)
            self.assertEqual(post_response["annotation"]["issue_type"], "公式位置")
            self.assertEqual(post_response["annotation"]["status"], "需本地客户端处理")
            self.assertTrue(processor.delete_layout_annotation(annotation["id"]))

    def test_delete_report_removes_layout_annotations(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            word = processor.create_uploaded_file("lesson.docx", make_docx_bytes())[0]
            processor.create_task({"task_type": "word_to_ppt", "file_ids": [word["id"]]})
            report = store.list_reports()[0]
            processor.save_layout_annotation({"report_id": report["id"], "location": "第 2 页 / 图片", "issue_type": "图片位置"})

            result = store.delete_report(report["id"])

            self.assertTrue(result["deleted"])
            self.assertEqual(result["annotations_deleted"], 1)
            self.assertEqual(processor.list_layout_annotations(report["id"]), [])

    def test_conversion_quality_reports_content_preservation_risks(self) -> None:
        document_xml = (
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>第一章</w:t></w:r></w:p>'
            "<w:tbl><w:tr><w:tc><w:p>表格</w:p></w:tc></w:tr></w:tbl>"
            "<w:p>公式 $$x+y=z$$</w:p>"
            "</w:document>"
        )
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            word = processor.create_uploaded_file(
                "lesson.docx",
                make_docx_bytes(extra_files={"word/media/image1.png": b"image"}, document_xml=document_xml),
            )[0]
            processor.create_task(
                {
                    "task_type": "word_to_ppt",
                    "file_ids": [word["id"]],
                    "options": {"word": {"retainImages": False, "retainTables": False, "retainFormulas": False}},
                }
            )
            report = store.list_reports()[0]
            quality = {item["id"]: item for item in report["qualityChecks"]}

            self.assertEqual(word["content_summary"]["headings"], 1)
            self.assertEqual(quality["content_preservation"]["status"], "需确认")
            self.assertIn("图片 1 个", quality["content_preservation"]["metric"])
            self.assertIn("表格 1 个", quality["content_preservation"]["metric"])
            self.assertIn("公式 1 个", quality["content_preservation"]["metric"])
            self.assertIn("标题/结构 1 个", quality["content_preservation"]["metric"])
            self.assertIn("图片保留关闭", quality["content_preservation"]["message"])
            self.assertIn("Word 公式保留关闭", quality["content_preservation"]["message"])

    def test_conversion_uses_configured_output_directory(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            custom_output = Path(tmp) / "custom-output"
            store = AppStore(Path(tmp) / "data")
            store.update_settings({"outputDirectory": str(custom_output)})
            processor = TaskProcessor(store)
            word = processor.create_uploaded_file("lesson.docx", make_docx_bytes())[0]
            task = processor.create_task({"task_type": "word_to_ppt", "file_ids": [word["id"]]})
            report = store.list_reports()[0]
            artifact = report["analysis"]["artifacts"][0]
            artifact_path = Path(artifact["path"])
            self.assertTrue(artifact_path.exists())
            self.assertEqual(artifact_path.parent, custom_output / task["id"])
            self.assertEqual(task["output_path"], str(custom_output / task["id"]))

    def test_conversion_output_conflict_strategy_renames_skips_and_overwrites(self) -> None:
        def run_with_strategy(strategy: str | None) -> tuple[dict, list[dict], dict]:
            tmp = tempfile.TemporaryDirectory()
            self.addCleanup(tmp.cleanup)
            root = Path(tmp.name)
            store = AppStore(root / "data")
            if strategy:
                store.update_settings({"outputConflictStrategy": strategy})
            analyzer = DocumentAnalyzer()
            files: list[dict] = []
            for index in range(2):
                source = root / f"source-{index}.docx"
                source.write_bytes(make_docx_bytes(document_xml=f"<w:document><w:p>题目 {index}</w:p></w:document>"))
                files.append(store.save_file(analyzer.analyze_file(source, "lesson.docx").to_dict()))
            processor = TaskProcessor(store)
            task = processor.create_task({"task_type": "word_to_ppt", "file_ids": [file["id"] for file in files]})
            report = store.list_reports()[0]
            return task, report["analysis"]["artifacts"], report

        _task, artifacts, _report = run_with_strategy(None)
        self.assertEqual([item["status"] for item in artifacts], ["成功", "成功"])
        self.assertEqual([item["file_name"] for item in artifacts], ["lesson.pptx", "lesson-1.pptx"])
        self.assertEqual(artifacts[0]["conversion_settings"]["output_conflict_strategy"], "自动重命名")
        self.assertTrue(all(Path(item["path"]).exists() for item in artifacts))

        task, artifacts, report = run_with_strategy("跳过")
        quality = {item["id"]: item for item in report["qualityChecks"]}
        self.assertEqual(task["status"], "成功")
        self.assertEqual([item["status"] for item in artifacts], ["成功", "跳过"])
        self.assertEqual(quality["conversion_output"]["status"], "需确认")
        self.assertIn("同名输出已存在", artifacts[1]["message"])

        _task, artifacts, _report = run_with_strategy("覆盖")
        self.assertEqual([item["status"] for item in artifacts], ["成功", "成功"])
        self.assertEqual([item["file_name"] for item in artifacts], ["lesson.pptx", "lesson.pptx"])
        self.assertEqual(artifacts[0]["path"], artifacts[1]["path"])

    def test_task_records_output_directory_open_request(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            store.update_settings({"autoOpenOutputDirectory": True, "keepOriginalFile": False})
            processor = TaskProcessor(store)
            word = processor.create_uploaded_file("lesson.docx", make_docx_bytes())[0]
            task = processor.create_task({"task_type": "word_to_ppt", "file_ids": [word["id"]]})
            logs = store.list_logs(task["id"])

            self.assertFalse(task["keep_original_file"])
            self.assertEqual(task["output_directory_action"]["status"], "待本地客户端执行")
            self.assertEqual(task["output_directory_action"]["path"], task["output_path"])
            self.assertTrue(any("打开输出目录" in log["message"] for log in logs))

    def test_task_completion_notice_can_be_enabled_or_disabled(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            word = processor.create_uploaded_file("lesson.docx", make_docx_bytes())[0]
            task = processor.create_task({"task_type": "word_to_ppt", "file_ids": [word["id"]]})
            self.assertEqual(task["completion_notice"]["status"], "待前端提示")
            self.assertIn("Word 转 PPT 已完成", task["completion_notice"]["message"])

            handler = make_handler(store, processor)
            K12RequestHandler._handle_api_post(handler, f"/api/tasks/{task['id']}/completion-notice/ack", {})
            payload = json.loads(handler.wfile.getvalue().decode("utf-8"))
            self.assertEqual(payload["task"]["completion_notice"]["status"], "已提示")
            self.assertIn("acknowledged_at", payload["task"]["completion_notice"])
            self.assertTrue(any("确认任务完成提醒" in log["message"] for log in store.list_logs(task["id"], limit=1000)))

        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            store.update_settings({"enableTaskCompletionNotice": False})
            processor = TaskProcessor(store)
            word = processor.create_uploaded_file("lesson.docx", make_docx_bytes())[0]
            task = processor.create_task({"task_type": "word_to_ppt", "file_ids": [word["id"]]})
            self.assertNotIn("completion_notice", task)

    def test_conversion_artifact_records_conversion_settings(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            store.update_settings(
                {
                    "wordToPptTemplate": "K12 讲义模板",
                    "pptToWordMode": "图文混排模式",
                    "pdfPrecisionMode": "高精度",
                    "excelConversionRange": "选区",
                    "excelFormulaMode": "仅保留计算结果",
                    "excelSplitSheets": True,
                    "retainImages": False,
                    "retainTables": True,
                    "retainHeadersFooters": True,
                    "retainFootnotesEndnotes": False,
                    "retainComments": True,
                    "retainRevisions": True,
                }
            )
            processor = TaskProcessor(store)
            word = processor.create_uploaded_file("lesson.docx", make_docx_bytes())[0]
            processor.create_task(
                {
                    "task_type": "word_to_ppt",
                    "file_ids": [word["id"]],
                    "options": {
                        "word": {
                            "maxCharsPerSlide": 180,
                            "autoPagination": False,
                            "generateToc": True,
                            "retainImages": True,
                            "retainTables": False,
                            "retainFormulas": False,
                            "convertOmmlFirst": False,
                            "applyTemplate": False,
                            "generateNotes": True,
                            "autoBeautify": False,
                        }
                    },
                }
            )
            report = store.list_reports()[0]
            artifact = report["analysis"]["artifacts"][0]
            settings = artifact["conversion_settings"]
            text_report = Path(report["txt_path"]).read_text(encoding="utf-8")

            self.assertEqual(settings["word_to_ppt_template"], "K12 讲义模板")
            self.assertEqual(settings["ppt_to_word_mode"], "图文混排模式")
            self.assertEqual(settings["pdf_precision_mode"], "高精度")
            self.assertEqual(settings["excel_conversion_range"], "选区")
            self.assertEqual(settings["excel_formula_mode"], "仅保留计算结果")
            self.assertTrue(settings["excel_split_sheets"])
            self.assertTrue(settings["retain_images"])
            self.assertFalse(settings["retain_tables"])
            self.assertTrue(settings["retain_headers_footers"])
            self.assertFalse(settings["retain_footnotes_endnotes"])
            self.assertTrue(settings["retain_comments"])
            self.assertTrue(settings["retain_revisions"])
            self.assertEqual(settings["word_max_chars_per_slide"], 180)
            self.assertFalse(settings["word_auto_pagination"])
            self.assertTrue(settings["word_generate_toc"])
            self.assertFalse(settings["word_retain_formulas"])
            self.assertFalse(settings["word_convert_omml_first"])
            self.assertFalse(settings["word_apply_template"])
            self.assertTrue(settings["word_generate_notes"])
            self.assertFalse(settings["word_auto_beautify"])
            self.assertIn("K12 讲义模板", text_report)
            self.assertIn("页眉页脚 是", text_report)
            self.assertIn("脚注尾注 否", text_report)
            self.assertIn("修订 是", text_report)
            self.assertIn("目录页 是", text_report)
            with zipfile.ZipFile(artifact["path"]) as archive:
                self.assertIn("目录", archive.read("ppt/slides/slide1.xml").decode("utf-8"))

    def test_settings_page_lists_prd_ppt_to_word_modes(self) -> None:
        html = (Path(__file__).resolve().parents[1] / "static" / "index.html").read_text(encoding="utf-8")
        for mode in ["逐页讲义模式", "内容提取模式", "图文混排模式", "备注优先模式", "大纲模式"]:
            self.assertIn(f"<option>{mode}</option>", html)
        self.assertNotIn("<option>连续文档模式</option>", html)
        self.assertNotIn("<option>仅备注模式</option>", html)

    def test_settings_page_exposes_excel_formula_mode(self) -> None:
        html = (Path(__file__).resolve().parents[1] / "static" / "index.html").read_text(encoding="utf-8")
        app_js = (Path(__file__).resolve().parents[1] / "static" / "app.js").read_text(encoding="utf-8")
        self.assertIn('id="excelFormulaMode"', html)
        self.assertIn('id="excelSplitSheets"', html)
        self.assertIn("<option>保留公式</option>", html)
        self.assertIn("<option>仅保留计算结果</option>", html)
        self.assertIn('"excelFormulaMode"', app_js)
        self.assertIn('"excelSplitSheets"', app_js)

    def test_desktop_main_nav_keeps_admin_tools_under_settings(self) -> None:
        root = Path(__file__).resolve().parents[1]
        html = (root / "static" / "index.html").read_text(encoding="utf-8")
        app_js = (root / "static" / "app.js").read_text(encoding="utf-8")
        styles = (root / "static" / "styles.css").read_text(encoding="utf-8")

        for label in ["工作台", "转换", "公式", "宏", "图片", "报告", "设置"]:
            self.assertIn(f"<b>{label}</b>", html)
        self.assertNotIn('data-view="admin"', html)
        self.assertIn('data-panel="admin"', html)
        self.assertIn('id="userList"', html)
        self.assertIn('id="templateList"', html)
        self.assertIn('id="authorizationList"', html)
        self.assertIn('id="adminReportSummary"', html)
        self.assertIn('id="adminReportList"', html)
        self.assertIn("报告管理", html)
        self.assertIn('id="preflightList"', html)
        self.assertIn('view === "settings" && panelName === "admin"', app_js)
        self.assertIn("function renderAdminReports()", app_js)
        self.assertIn('id="taskDetailPanel"', html)
        self.assertIn('id="taskDetailScope"', html)
        self.assertIn("<th>页/页签</th>", html)
        self.assertIn("<th>能力</th>", html)
        self.assertIn('colspan="11"', html)
        self.assertIn("<td>${escapeHtml(fileMetric(file))}</td>", app_js)
        self.assertIn("<td>${fileCapabilityBadges(file)}</td>", app_js)
        self.assertIn("function fileCapabilityBadges", app_js)
        self.assertIn("function primaryValidationError", app_js)
        self.assertIn("function uploadValidationSummary", app_js)
        self.assertIn("table-note error-note", app_js)
        self.assertIn("需处理：", app_js)
        self.assertIn(".table-note.error-note", styles)
        self.assertIn("OMML缺失", app_js)
        self.assertIn("微小图", app_js)
        self.assertIn("file.k12RelativePath || file.webkitRelativePath || file.name", app_js)
        self.assertIn("function importDroppedItems", app_js)
        self.assertIn("function collectDroppedEntryFiles", app_js)
        self.assertIn("webkitGetAsEntry", app_js)
        self.assertIn("readEntries", app_js)
        self.assertIn("source_relative_path", app_js)
        self.assertIn("来源层级", app_js)
        self.assertIn('data-view="menu"', html)
        self.assertIn('data-panel="menu"', html)
        self.assertIn('id="prdMenuStructure"', html)
        self.assertIn("PRD 菜单结构", html)
        self.assertIn("const menuStructure", app_js)
        self.assertIn("function renderMenuStructure()", app_js)
        self.assertIn("function openMenuTarget", app_js)
        self.assertIn(".menu-structure-grid", styles)
        self.assertIn(".menu-group-head", styles)
        for menu_label in [
            "文件处理",
            "Word 处理",
            "宏处理",
            "公式处理",
            "图片检索",
            "批量任务",
            "本地客户端",
            "网页管理",
            "文档预览",
            "处理报告",
            "系统设置",
        ]:
            self.assertIn(menu_label, app_js)
        self.assertIn('id="completionNoticePanel"', html)
        self.assertIn('class="queue-selection-copy"', html)
        self.assertIn('id="formatSupportStrip"', html)
        self.assertIn('class="quick-task-row home-quick-grid"', html)
        for label in [
            "Word 转 PPT",
            "PPT 转 Word",
            "PDF 转 Word",
            "MathType 公式处理",
            "Word 自带公式转 MathType",
            "Word 公式格式化",
            "Word 宏顺序执行",
            "PDF 公式识别",
            "微小图片检索",
            "批量处理",
        ]:
            self.assertIn(label, html)
        self.assertIn('id="workspaceRecentTasks"', html)
        self.assertIn('class="panel preview-helper"', html)
        self.assertIn('id="filePreviewPanel"', html)
        self.assertIn('id="fileDetail"', html)
        self.assertIn('id="capabilityList"', html)
        self.assertNotIn('class="panel collapsed-helper"', html)
        self.assertIn('id="latestReportDate"', html)
        self.assertIn('queue-extra-action', html)
        self.assertIn('id="exportOmmlFailuresButton"', html)
        self.assertIn('id="exportMacroFailuresButton"', html)
        self.assertIn('id="loadMacroReportSequenceButton"', html)
        self.assertIn('id="macroDetailPanel"', html)
        self.assertIn('id="macroDetailScope"', html)
        self.assertIn("宏详情", html)
        self.assertIn("function renderMacroDetail()", app_js)
        self.assertIn("macro-detail", app_js)
        self.assertIn(".macro-detail-panel", styles)
        self.assertIn('id="macroLogList"', html)
        self.assertIn('id="macroLogScope"', html)
        self.assertIn("宏执行日志", html)
        self.assertIn("function renderMacroLogs()", app_js)
        self.assertIn(".macro-log-list", styles)
        self.assertIn('id="layoutAnnotationList"', html)
        self.assertIn('id="saveLayoutAnnotationButton"', html)
        self.assertIn('id="layoutIssueType"', html)
        self.assertIn('id="exportImageManifestButton"', html)
        self.assertIn('id="imageDetailPanel"', html)
        self.assertIn('class="image-workbench"', html)
        self.assertIn('id="exportFormulaSheetButton"', html)
        self.assertIn('id="formulaDetailPanel"', html)
        self.assertIn('class="formula-workbench"', html)
        self.assertIn('id="pdfExportWordButton"', html)
        self.assertIn('id="pdfLowConfidenceList"', html)
        self.assertIn('id="pdfOcrSettings"', html)
        self.assertIn('id="pptExportWordButton"', html)
        self.assertIn('id="pptFileList"', html)
        self.assertIn('id="pptContentPreview"', html)
        self.assertIn('id="pptWorkspaceMode"', html)
        self.assertIn('id="pptWorkspaceExtractNotes"', html)
        self.assertIn('id="pptWorkspaceExtractImages"', html)
        self.assertIn('id="pptWorkspaceExtractFormulas"', html)
        self.assertIn('id="pptArtifactList"', html)
        self.assertIn('class="panel wide-panel word-workspace"', html)
        self.assertIn('id="wordFilePreviewList"', html)
        self.assertIn('id="wordFormulaDetectionList"', html)
        self.assertIn('id="wordMacroDetectionList"', html)
        self.assertIn('id="wordReportLinkList"', html)
        self.assertIn('id="selectWordFilesButton"', html)
        self.assertIn('id="wordPreviewSelectedButton"', html)
        self.assertIn('id="wordExportPptButton"', html)
        self.assertIn('class="settings-section-nav"', html)
        for section_id in [
            "settings-basic",
            "settings-conversion",
            "settings-ocr",
            "settings-formula",
            "settings-omml",
            "settings-mathtype",
            "settings-macro",
            "settings-image",
            "settings-output",
            "settings-log",
            "settings-cache",
            "settings-local",
            "settings-cloud",
        ]:
            self.assertIn(f'id="{section_id}"', html)
            self.assertIn(f'href="#{section_id}"', html)
        self.assertIn('id="localHandoffPanel"', html)
        self.assertIn('id="localHandoffScope"', html)
        self.assertIn("本地任务交接", html)
        self.assertIn("function renderLocalHandoff()", app_js)
        self.assertIn("function localTaskCandidates()", app_js)
        self.assertIn("function launchLocalClient", app_js)
        self.assertIn("function copyLocalPayload", app_js)
        self.assertIn("function refreshLocalReadiness", app_js)
        self.assertIn("function renderLocalReadinessCard", app_js)
        self.assertIn("function renderLocalExecutionPlan", app_js)
        self.assertIn("桌面执行计划", app_js)
        self.assertIn("desktop_execution_plan", app_js)
        self.assertIn("executionPlanBadge", app_js)
        self.assertIn("clientReadinessBadge", app_js)
        self.assertIn("platform_mismatch", app_js)
        self.assertIn("/local-readiness", app_js)
        self.assertIn("k12-local://open", app_js)
        self.assertIn("copy-local-payload", app_js)
        self.assertIn(".local-handoff-grid", styles)
        self.assertIn(".local-handoff-status", styles)
        self.assertIn(".local-readiness-card", styles)
        self.assertIn(".local-execution-plan", styles)
        self.assertIn(".local-execution-actions", styles)
        self.assertIn(".local-readiness-capabilities", styles)
        self.assertIn('id="riskRegisterPanel"', html)
        self.assertIn('id="riskRegisterScope"', html)
        self.assertIn("风险控制", html)
        self.assertIn("const riskControls", app_js)
        self.assertIn("function renderRiskRegister()", app_js)
        self.assertIn("function riskControlState", app_js)
        for risk_label in [
            "MathType 兼容性复杂",
            "OMML 转 MathType 不稳定",
            "宏执行安全风险",
            "PDF 公式识别准确率不足",
            "本地客户端与网页通信风险",
        ]:
            self.assertIn(risk_label, app_js)
        self.assertIn(".risk-register-grid", styles)
        self.assertIn(".risk-card", styles)
        self.assertIn('id="performanceTargetPanel"', html)
        self.assertIn('id="performanceTargetScope"', html)
        self.assertIn("性能指标", html)
        self.assertIn("const performanceTargets", app_js)
        self.assertIn("function renderPerformanceTargets()", app_js)
        self.assertIn("function performanceTargetState", app_js)
        for performance_label in [
            "文件拖拽响应",
            "单文件大小",
            "并发任务数",
            "100 页 PDF 转 Word",
            "任务恢复",
        ]:
            self.assertIn(performance_label, app_js)
        self.assertIn(".performance-target-grid", styles)
        self.assertIn(".performance-card", styles)
        self.assertIn('id="exceptionPolicyPanel"', html)
        self.assertIn('id="exceptionPolicyScope"', html)
        self.assertIn("异常处理", html)
        self.assertIn("const exceptionPolicies", app_js)
        self.assertIn("function renderExceptionPolicies()", app_js)
        self.assertIn("function exceptionPolicyState", app_js)
        for exception_label in [
            "文件异常",
            "转换异常",
            "公式异常",
            "OMML 文件异常",
            "宏异常",
            "系统异常",
            "磁盘空间不足",
            "客户端平台心跳",
            "Office / MathType 组件",
            "OCR 引擎异常",
        ]:
            self.assertIn(exception_label, app_js)
        self.assertIn("local_client_components", app_js)
        self.assertIn("function renderSystemExceptionDetails()", app_js)
        self.assertIn("function systemExceptionDetails()", app_js)
        self.assertIn("function exceptionPreflightLevel", app_js)
        self.assertIn(".exception-policy-grid", styles)
        self.assertIn(".exception-card", styles)
        self.assertIn(".system-exception-list", styles)
        self.assertIn(".system-exception-row", styles)
        self.assertIn('id="acceptanceOverviewPanel"', html)
        self.assertIn('id="acceptanceOverviewScope"', html)
        self.assertIn("验收总览", html)
        self.assertIn("const acceptanceGroups", app_js)
        self.assertIn("function renderAcceptanceOverview()", app_js)
        self.assertIn("function acceptanceGroupState", app_js)
        for acceptance_label in [
            "17.1",
            "文件上传验收",
            "Word 处理验收",
            "PDF 处理验收",
            "本地 + 网页双模式验收",
        ]:
            self.assertIn(acceptance_label, app_js)
        self.assertIn(".acceptance-overview-grid", styles)
        self.assertIn(".acceptance-card", styles)
        self.assertIn('id="acceptanceMatrixPanel"', html)
        self.assertIn('id="acceptanceMatrixScope"', html)
        self.assertIn('id="acceptanceRiskPanel"', html)
        self.assertIn("验收证据矩阵", html)
        self.assertIn("/api/acceptance-matrix", app_js)
        self.assertIn("acceptanceMatrix", app_js)
        self.assertIn('status: "交接覆盖"', app_js)
        self.assertIn('status: "需本地客户端实测"', app_js)
        self.assertNotIn('status: "本地实测"', app_js)
        self.assertIn("真实 OMML 转 MathType 需本地客户端实测", app_js)
        self.assertIn("真实宏执行需本地客户端实测", app_js)
        self.assertIn("uncovered_risks", app_js)
        self.assertIn("blocking_reasons", app_js)
        self.assertIn("required_environment", app_js)
        self.assertIn("条未覆盖风险", app_js)
        self.assertIn("function renderAcceptanceMatrix()", app_js)
        self.assertIn(".acceptance-risk-grid", styles)
        self.assertIn(".acceptance-risk-card", styles)
        self.assertIn(".acceptance-matrix-grid", styles)
        self.assertIn(".acceptance-item-row", styles)
        self.assertIn('id="versionRoadmapPanel"', html)
        self.assertIn('id="versionRoadmapScope"', html)
        self.assertIn("版本路线图", html)
        self.assertIn("const versionRoadmap", app_js)
        self.assertIn("function renderVersionRoadmap()", app_js)
        self.assertIn("function versionRoadmapState", app_js)
        for version_label in [
            "V1.0",
            "基础本地可用版",
            "V2.0",
            "混合模式版",
            "V3.0",
            "智能增强版",
        ]:
            self.assertIn(version_label, app_js)
        self.assertIn(".version-roadmap-grid", styles)
        self.assertIn(".version-card", styles)
        self.assertIn('id="dataDictionaryPanel"', html)
        self.assertIn('id="dataDictionaryScope"', html)
        self.assertIn("数据字典", html)
        self.assertIn("const dataDictionary", app_js)
        self.assertIn("/api/data-dictionary", app_js)
        self.assertIn("state.dataDictionary", app_js)
        self.assertIn("function renderDataDictionary()", app_js)
        self.assertIn("function dataDictionaryObjects()", app_js)
        self.assertIn("k12.dataDictionary.v1", app_js)
        self.assertIn("function dataDictionaryCount", app_js)
        for model_label in [
            "FileItem",
            "Task",
            "FormulaItem",
            "FormulaAnnotation",
            "OmmlDependencyItem",
            "MacroItem",
            "SmallImageItem",
            "ReportItem",
        ]:
            self.assertIn(model_label, app_js)
        self.assertIn(".data-dictionary-grid", styles)
        self.assertIn(".data-field-list", styles)
        self.assertIn('id="securityPrivacyPanel"', html)
        self.assertIn('id="securityPrivacyScope"', html)
        self.assertIn("安全与隐私", html)
        self.assertIn("const securityPrivacyControls", app_js)
        self.assertIn("function renderSecurityPrivacy()", app_js)
        self.assertIn("function securityPrivacyState", app_js)
        for security_label in [
            "文件安全",
            "宏安全",
            "本地与网页通信安全",
            "隐私保护",
            "离线与私有化",
        ]:
            self.assertIn(security_label, app_js)
        self.assertIn(".security-privacy-grid", styles)
        self.assertIn(".security-card", styles)
        self.assertIn('id="compatibilityMatrixPanel"', html)
        self.assertIn('id="compatibilityMatrixScope"', html)
        self.assertIn("兼容性矩阵", html)
        self.assertIn("const compatibilityMatrix", app_js)
        self.assertIn("function renderCompatibilityMatrix()", app_js)
        self.assertIn("function compatibilityMatrixState", app_js)
        for compatibility_label in [
            "操作系统",
            "Office 兼容",
            "公式兼容",
            "宏兼容",
            "模式兼容",
            "Windows 10",
            "macOS",
            "MathType",
            "Word 原生 OMML",
        ]:
            self.assertIn(compatibility_label, app_js)
        self.assertIn(".compatibility-matrix-grid", styles)
        self.assertIn(".compatibility-card", styles)
        self.assertIn('id="architectureBlueprintPanel"', html)
        self.assertIn('id="architectureBlueprintScope"', html)
        self.assertIn("技术架构", html)
        self.assertIn("/api/architecture", app_js)
        self.assertIn("function renderArchitectureBlueprint()", app_js)
        self.assertIn("architecture-blueprint-grid", styles)
        for architecture_label in [
            "当前运行时",
            "接口合同",
            "architecture-runtime",
            "architecture-contract-list",
        ]:
            self.assertIn(architecture_label, app_js + styles)
        self.assertIn('id="apiCatalogPanel"', html)
        self.assertIn('id="apiCatalogScope"', html)
        self.assertIn("开放 API", html)
        self.assertIn("/api/api-catalog", app_js)
        self.assertIn("function renderApiCatalog()", app_js)
        self.assertIn("api-catalog-grid", styles)
        for api_label in [
            "鉴权边界",
            "本地载荷",
            "api-endpoint-head",
            "api-method",
        ]:
            self.assertIn(api_label, app_js + styles)
        self.assertIn('id="productSummaryPanel"', html)
        self.assertIn('id="productSummaryScope"', html)
        self.assertIn("产品总结", html)
        self.assertIn("/api/product-summary", app_js)
        self.assertIn("function renderProductSummary()", app_js)
        self.assertIn("product-summary-grid", styles)
        for product_label in [
            "产品定位",
            "产品路线",
            "k12.productSummary.v1",
            "product-phase-list",
        ]:
            self.assertIn(product_label, app_js + styles)
        self.assertIn('id="enhancementPlanPanel"', html)
        self.assertIn('id="enhancementPlanScope"', html)
        self.assertIn("智能增强规划", html)
        self.assertIn("/api/enhancement-plan", app_js)
        self.assertIn("function renderEnhancementPlan()", app_js)
        self.assertIn("enhancement-plan-grid", styles)
        for enhancement_label in [
            "V3.0 增强能力规划",
            "k12.enhancementPlan.v1",
            "接入前置条件",
            "enhancement-readiness-list",
            "enhancement-guardrails",
        ]:
            self.assertIn(enhancement_label, app_js + styles)
        self.assertIn("/api/install-plan", app_js)
        self.assertIn("installPlan", app_js)
        for install_label in [
            "安装计划",
            "公式交付合同",
            "install-plan-card",
            "install-contract-card",
            "installer-manifest-card",
            "install-package-grid",
            "install-warning-list",
            "install-readiness-list",
            "install-step-list",
            "formula_compatibility",
            "platform?.installer",
            "k12.localInstallerManifest.v1",
            "download_requires_platform_query",
            "checksum_required",
            "path_policy",
            "install-boundary-grid",
            "同平台要求",
            "原生交接",
            "阻断原因",
            "安装包边界",
            "recommended_action",
            "same_platform_required_for_native_objects",
            "native_handoff_allowed",
            "native_handoff_blocking_reasons",
            "package_boundary",
        ]:
            self.assertIn(install_label, app_js + styles)
        self.assertIn("function installContractBadgeClass", app_js)
        self.assertIn("// Platform mismatch must stay visible because MathType native objects are not portable.", app_js)
        self.assertIn('blockers.includes("platform_mismatch")', app_js)
        self.assertIn("installContractBadgeClass(formulaContract)", app_js)
        self.assertNotIn('formulaContract.native_mathtype_object_allowed ? "warn" : "good"', app_js)
        self.assertIn("/api/local-client/manifest", app_js)
        self.assertIn("/api/local-client/heartbeat", app_js)
        for local_client_label in [
            "本地客户端运行清单",
            "客户端心跳",
            "启动请求",
            "本地结果上传队列",
            "/api/local-client/uploads",
            "local-client-manifest",
            "local-client-actions",
            "local-upload-queue",
            "local-upload-manifest",
            "接收清单",
            "云端接收清单",
            "function uploadBadge",
            "function loadLocalUploadManifest",
            "function renderLocalUploadManifest",
            "function launchUrlWithToken",
            "/local-launch",
            "已登记",
        ]:
            self.assertIn(local_client_label, app_js + styles)
        self.assertIn(".settings-section-nav", styles)
        self.assertIn(".settings-subtitle", styles)
        self.assertIn('draggable="true"', app_js)
        self.assertIn("任务详情", html)
        self.assertIn("task-detail", app_js)
        self.assertIn("task-notice-ack", app_js)
        self.assertIn("/completion-notice/ack", app_js)
        self.assertIn("function renderCompletionNotices()", app_js)
        self.assertIn("function renderQueueTabs()", app_js)
        self.assertIn("function renderWorkspaceRecentTasks()", app_js)
        self.assertIn("formatDate(report.created_at)", app_js)
        self.assertIn("function exportImageManifest()", app_js)
        self.assertIn("images.xlsx", app_js)
        self.assertIn("function renderImageDetail", app_js)
        self.assertIn("function imageDetailBadges", app_js)
        self.assertIn("image-detail", app_js)
        self.assertIn("function markImageReplacement", app_js)
        self.assertIn("function readFileAsDataUrl", app_js)
        self.assertIn('input.accept = "image/png,image/jpeg,image/gif,image/webp,image/bmp"', app_js)
        self.assertIn("/api/image-annotations/replacement", app_js)
        self.assertIn("替换图片已登记", app_js)
        self.assertIn("查看替换素材", app_js)
        self.assertIn("function markImageLocationUnknown", app_js)
        self.assertIn('"位置未知"', app_js)
        self.assertIn("来源预览中未命中图片位置", app_js)
        self.assertIn("button.dataset.report, button.dataset.id", app_js)
        self.assertIn("duplicate_check_status", app_js)
        self.assertIn("duplicate_fallback", app_js)
        self.assertIn("<b>重复判断</b>", app_js)
        self.assertIn("<b>异常兜底</b>", app_js)
        self.assertIn("function reexportImage", app_js)
        self.assertIn("/api/image-assets/reexport", app_js)
        self.assertIn("image-reexport", app_js)
        self.assertIn("<b>导出状态</b>", app_js)
        self.assertIn("<b>导出信息</b>", app_js)
        self.assertIn("imageExtractionErrors", app_js)
        self.assertIn("提取失败", app_js)
        self.assertIn(".image-detail-panel", styles)
        self.assertIn(".image-detail-meta", styles)
        self.assertIn(".format-support-strip", styles)
        self.assertIn(".recent-task-list", styles)
        self.assertIn(".progress-panel .recent-task-list", styles)
        self.assertIn(".home-quick-grid", styles)
        self.assertIn(".preview-helper", styles)
        self.assertIn("function previewMarkerChips", app_js)
        self.assertIn("function previewPageSearchText", app_js)
        self.assertIn("page.markers || []", app_js)
        self.assertIn("preview-marker", app_js)
        self.assertIn(".preview-markers", styles)
        self.assertIn(".preview-marker.warn", styles)
        self.assertIn("function exportFormulaSheet()", app_js)
        self.assertIn("formulas.xlsx", app_js)
        self.assertIn("function renderFormulaDetail", app_js)
        self.assertIn("position_status", app_js)
        self.assertIn("position_issue", app_js)
        self.assertIn("位置异常", app_js)
        self.assertIn("<h4>位置状态</h4>", app_js)
        self.assertIn("formulaOriginalPreview", app_js)
        self.assertIn("formula-detail-latex", app_js)
        self.assertIn("formula-select", app_js)
        self.assertIn(".formula-detail-panel", styles)
        self.assertIn(".formula-param-grid", styles)
        self.assertIn("function renderPdfWorkspace()", app_js)
        self.assertIn("/api/mathpix-jobs", app_js)
        self.assertIn("Mathpix 作业队列", html)
        self.assertIn("function renderPdfMathpixJobs()", app_js)
        self.assertRegex(app_js, r'pdf_to_word:\s*\{[^}]*mode:\s*"hybrid"')
        self.assertIn("Mathpix 识别 PDF，公式 OCR 后进入本地 MathType 后处理", app_js)
        self.assertIn("function mathpixJobBadge", app_js)
        self.assertIn("recognition_plan", app_js)
        self.assertIn("upload_gate", app_js)
        self.assertIn("上传门禁", app_js)
        self.assertIn("识别计划", app_js)
        self.assertIn("conversion_formats", app_js)
        self.assertIn("function previewJumpChip", app_js)
        self.assertIn("function jumpPreviewTo", app_js)
        self.assertIn("preview-jump-chip", app_js)
        self.assertIn(".preview-jump-chip:hover", styles)
        self.assertIn("已定位到预览命中页", app_js)
        self.assertIn("function startPdfToWord()", app_js)
        self.assertIn("function openPdfSettings()", app_js)
        self.assertIn("function renderWordWorkspace()", app_js)
        self.assertIn("function selectWordFiles()", app_js)
        self.assertIn("function previewSelectedWord()", app_js)
        self.assertIn("function startWordToPpt()", app_js)
        self.assertIn('id="pdfToWordEngine" disabled title="固定使用 Mathpix"', html)
        self.assertIn("word-select", app_js)
        self.assertIn(".word-workspace-grid", styles)
        self.assertIn("function renderPptWorkspace()", app_js)
        self.assertIn("function startPptToWord()", app_js)
        self.assertIn("function previewSelectedPpt()", app_js)
        self.assertIn("pptWorkspaceExtractNotes", app_js)
        self.assertIn("function orderedPlannerTasks", app_js)
        self.assertIn("function movePlannerTask", app_js)
        self.assertIn("application/x-k12-planner-task", app_js)
        self.assertIn("planner-drag-handle", app_js)
        self.assertIn('button type="button" class="planner-drag-handle" draggable="true"', app_js)
        self.assertIn("拖拽或用方向键调整流程顺序", app_js)
        self.assertIn('lineIcon("drag-lines", "drag-line-icon")', app_js)
        self.assertIn('"drag-lines": \'<path d="M7 7h10"></path><path d="M7 12h10"></path><path d="M7 17h10"></path>\'', app_js)
        self.assertIn('event.target.closest(".planner-drag-handle")', app_js)
        self.assertIn("item.getBoundingClientRect()", app_js)
        self.assertIn('placement === "after"', app_js)
        self.assertIn("planner-enabled-check", app_js)
        self.assertIn("已加入处理流程", app_js)
        self.assertIn('lineIcon("check", "planner-check-icon")', app_js)
        self.assertIn("function startPlannerPointerDrag", app_js)
        self.assertIn("function updatePlannerPointerDrag", app_js)
        self.assertIn("function finishPlannerPointerDrag", app_js)
        self.assertIn("function handlePlannerHandleKeydown", app_js)
        self.assertIn("event.target instanceof Element", app_js)
        self.assertIn('event.key === "ArrowUp"', app_js)
        self.assertIn('event.key === "ArrowDown"', app_js)
        self.assertIn('event.key === "Home"', app_js)
        self.assertIn('event.key === "End"', app_js)
        self.assertIn("function movePlannerTaskToIndex", app_js)
        self.assertIn("function focusPlannerHandle", app_js)
        self.assertIn("Re-render replaces buttons", app_js)
        self.assertIn("plannerDropPlacement", app_js)
        self.assertIn("function lineIcon", app_js)
        self.assertIn(".line-icon", styles)
        self.assertIn(".planner-drag-handle .drag-line-icon", styles)
        self.assertIn(".planner-enabled-check", styles)
        self.assertIn(".planner-drag-handle:hover", styles)
        self.assertIn(".planner-drag-handle:focus-visible", styles)
        self.assertIn("macro-order-item", app_js)
        self.assertIn("macro-drag-handle", app_js)
        self.assertIn("application/x-k12-macro-order", app_js)
        self.assertIn("function moveMacroBefore", app_js)
        self.assertIn(".order-list .macro-order-item", styles)
        self.assertIn("workflowOrder", app_js)
        self.assertIn("function renderTaskDetail()", app_js)
        self.assertIn("function batchProgressRows(report)", app_js)
        self.assertIn("批量单文件进度", app_js)
        self.assertIn("function hasPermission(permission)", app_js)
        self.assertIn("function permissionButtonAttrs(enabled, title, permission)", app_js)
        self.assertIn("缺少 ${permission} 权限", app_js)
        self.assertIn('permissionButtonAttrs(true, "输入密码", "files.manage")', app_js)
        self.assertIn('permissionButtonAttrs(true, "创建转换任务", "tasks.create")', app_js)
        self.assertIn('permissionButtonAttrs(true, "创建检索任务", "tasks.create")', app_js)
        self.assertIn("file-convert", app_js)
        self.assertIn("file-scan", app_js)
        self.assertIn("file-reports", app_js)
        self.assertIn("function openFileReports", app_js)
        self.assertIn("/reports", app_js)
        self.assertIn("当前文件暂无报告", app_js)
        self.assertIn("已打开 ${reports.length} 份关联报告", app_js)
        self.assertIn("state.activeReportId", app_js)
        self.assertIn("function createFileTask", app_js)
        self.assertIn("function defaultConversionTask", app_js)
        self.assertIn("function defaultSearchTask", app_js)
        self.assertIn("file-download", app_js)
        self.assertIn('permissionButtonAttrs(true, "重新上传替换", "files.manage")', app_js)
        self.assertIn('permissionButtonAttrs(true, "删除文件", "files.manage")', app_js)
        self.assertIn("确定删除这个文件记录吗", app_js)
        self.assertIn('permissionButtonAttrs(true, "删除报告", "reports.manage")', app_js)
        self.assertIn('permissionButtonAttrs(true, "重试", "tasks.control")', app_js)
        self.assertIn('permissionButtonAttrs(canPause, "暂停", "tasks.control")', app_js)
        self.assertIn("batch-skip-file", app_js)
        self.assertIn("<th>耗时</th>", html)
        self.assertIn("<th>失败原因</th>", html)
        self.assertIn("<th>成功</th>", html)
        self.assertIn("<th>失败</th>", html)
        self.assertIn('colspan="13"', app_js)
        self.assertIn("formatDuration(task)", app_js)
        self.assertIn("taskResultSummary(task)", app_js)
        self.assertIn("taskFailureReason(task)", app_js)
        self.assertIn('id="logScope"', html)
        self.assertIn('id="clearTaskLogFilterButton"', html)
        self.assertIn("task-log", app_js)
        self.assertIn("/api/logs?task_id=", app_js)
        self.assertIn("task_id=${encodeURIComponent(state.activeLogTaskId)}", app_js)
        self.assertIn("structured-log-item", app_js)
        self.assertIn("log-line-time", app_js)
        self.assertIn("log-level-chip", app_js)
        self.assertIn("function logLevelClass", app_js)
        self.assertIn("function logLevelLabel", app_js)
        self.assertIn(".log-level-chip.info", styles)
        self.assertIn(".log-level-chip.warning", styles)
        self.assertIn(".log-level-chip.error", styles)
        self.assertIn(".log-level-chip.success", styles)
        self.assertIn("macro-rerun", app_js)
        self.assertIn("function rerunMacroFromReport", app_js)
        self.assertIn("function loadMacroSequenceFromLatestReport()", app_js)
        self.assertIn("setSelectIfOptionExists", app_js)
        self.assertIn("已载入", app_js)
        self.assertIn("state.selectedFiles = new Set(fileIds)", app_js)
        self.assertIn("function exportMacroFailures()", app_js)
        self.assertIn("macro-failures.csv", app_js)
        self.assertIn('task_type: "macro_sequence"', app_js)
        self.assertIn('execute_mode: "local"', app_js)
        self.assertIn("selectedMacros", app_js)
        self.assertIn("image-confirm-type", app_js)
        self.assertIn("function confirmImageType", app_js)
        self.assertIn("类型已确认", app_js)
        self.assertIn("imageKindLabel", app_js)
        self.assertIn("/api/layout-annotations", app_js)
        self.assertIn("function renderLayoutReview()", app_js)
        self.assertIn("function saveLayoutAnnotation()", app_js)
        self.assertIn("layoutStatusClass", app_js)
        self.assertIn("workspaceFailureStrategy", app_js)
        self.assertIn("batchFailureStrategy", app_js)
        self.assertIn("continueOnFailure", app_js)
        self.assertIn("workspaceConcurrency", app_js)
        self.assertIn("maxConcurrentTasks", app_js)
        self.assertIn('colspan="13"', app_js)
        self.assertIn("/api/omml-annotations", app_js)
        self.assertIn("omml-keep", app_js)
        self.assertIn("omml-retry", app_js)
        self.assertIn("omml-fixed", app_js)
        self.assertIn("omml-manual", app_js)
        self.assertIn("manual_omml_path", app_js)
        self.assertIn("手动指定依赖", app_js)
        self.assertIn("window.prompt", app_js)
        self.assertIn("function annotateOmml", app_js)
        self.assertIn("function ommlAnnotation", app_js)
        self.assertIn("function exportOmmlFailures()", app_js)
        self.assertIn("omml-failures.csv", app_js)

    def test_conversion_tab_exposes_prd_conversion_tasks(self) -> None:
        html = (Path(__file__).resolve().parents[1] / "static" / "index.html").read_text(encoding="utf-8")
        for task_type in ["word_to_ppt", "ppt_to_word", "pdf_to_word", "excel_to_pdf", "excel_to_word", "excel_to_ppt"]:
            self.assertIn(f'data-task="{task_type}"', html)
        self.assertIn("使用 Mathpix 识别扫描件、公式和表格线索", html)
        self.assertIn("Word 转 PPT 参数", html)

    def test_ppt_to_word_generates_artifact(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            ppt = processor.create_uploaded_file("slides.pptx", make_pptx_bytes())[0]
            task = processor.create_task({"task_type": "ppt_to_word", "file_ids": [ppt["id"]]})
            report = store.list_reports()[0]
            artifact = report["analysis"]["artifacts"][0]
            self.assertEqual(task["status"], "成功")
            self.assertEqual(artifact["output_type"], "docx")
            self.assertTrue(Path(artifact["path"]).exists())
            with zipfile.ZipFile(artifact["path"]) as archive:
                self.assertIn("word/document.xml", archive.namelist())

    def test_ppt_to_word_uses_notes_mode_and_object_summary(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "slides.pptx"
            path.write_bytes(make_pptx_bytes(make_pptx_object_files(), make_pptx_slide_with_objects_xml()))

            analyzed = DocumentAnalyzer().analyze_file(path)
            slides = extract_pptx_slides(path)

            self.assertEqual(analyzed.content_summary["notes"], 1)
            self.assertEqual(analyzed.content_summary["images"], 1)
            self.assertEqual(analyzed.content_summary["charts"], 1)
            self.assertEqual(analyzed.content_summary["tables"], 1)
            self.assertEqual(analyzed.content_summary["formulas"], 1)
            self.assertTrue(analyzed.has_formula)
            self.assertEqual(analyzed.content_summary["textBoxes"], 2)
            self.assertEqual(slides[0]["notes"], ["讲者备注：先讲目标，再讲例题"])
            self.assertEqual(slides[0]["image_count"], 1)
            self.assertEqual(slides[0]["chart_count"], 1)
            self.assertEqual(slides[0]["table_count"], 1)
            self.assertEqual(slides[0]["formula_count"], 1)

            store = AppStore(tmp)
            store.update_settings({"pptToWordMode": "备注优先模式", "pptToWordTemplate": "K12 Word 讲义模板", "pptToWordGenerateToc": True})
            processor = TaskProcessor(store)
            ppt = processor.create_uploaded_file("slides.pptx", path.read_bytes())[0]
            task = processor.create_task({"task_type": "ppt_to_word", "file_ids": [ppt["id"]]})
            report = store.list_reports()[0]
            artifact = report["analysis"]["artifacts"][0]
            formulas = report["analysis"]["formulas"]
            preview = processor.file_preview(ppt["id"])

            self.assertEqual(task["status"], "成功")
            self.assertEqual(artifact["conversion_settings"]["ppt_to_word_mode"], "备注优先模式")
            self.assertEqual(artifact["conversion_settings"]["ppt_to_word_template"], "K12 Word 讲义模板")
            self.assertTrue(artifact["conversion_settings"]["ppt_to_word_generate_toc"])
            self.assertTrue(formulas)
            self.assertEqual(formulas[0]["file_id"], ppt["id"])
            self.assertTrue(any(item["label"] == "公式" and item["count"] == 1 for item in preview["objects"]))
            with zipfile.ZipFile(artifact["path"]) as archive:
                document_xml = archive.read("word/document.xml").decode("utf-8")
            self.assertIn("PPT 转 Word 模式：备注优先模式", document_xml)
            self.assertIn("Word 模板：K12 Word 讲义模板", document_xml)
            self.assertIn("目录", document_xml)
            self.assertIn("1. 第一课", document_xml)
            self.assertIn("对象摘要：文本框 2 个；图片 1 个；表格 1 个；公式 1 个；图表 1 个；形状 2 个；备注 1 个", document_xml)
            self.assertLess(document_xml.index("讲者备注"), document_xml.index("知识点 A"))
            self.assertIn("Word模板 K12 Word 讲义模板", Path(report["txt_path"]).read_text(encoding="utf-8"))

    def test_ppt_to_word_records_workspace_extract_options(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "slides.pptx"
            source.write_bytes(make_pptx_bytes(make_pptx_object_files(), make_pptx_slide_with_objects_xml()))
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            ppt = processor.create_uploaded_file("slides.pptx", source.read_bytes())[0]
            processor.create_task(
                {
                    "task_type": "ppt_to_word",
                    "file_ids": [ppt["id"]],
                    "options": {
                        "ppt": {
                            "mode": "备注优先模式",
                            "generateToc": False,
                            "extractNotes": False,
                            "retainImages": False,
                            "retainFormulas": False,
                            "templateName": "K12 Word 讲义模板",
                        }
                    },
                }
            )
            report = store.list_reports()[0]
            artifact = report["analysis"]["artifacts"][0]
            settings = artifact["conversion_settings"]
            with zipfile.ZipFile(artifact["path"]) as archive:
                document_xml = archive.read("word/document.xml").decode("utf-8")
            text_report = Path(report["txt_path"]).read_text(encoding="utf-8")

            self.assertFalse(settings["ppt_to_word_generate_toc"])
            self.assertFalse(settings["ppt_extract_notes"])
            self.assertFalse(settings["ppt_retain_images"])
            self.assertFalse(settings["ppt_retain_formulas"])
            self.assertNotIn("讲者备注", document_xml)
            self.assertNotIn("图片 1 个", document_xml)
            self.assertNotIn("公式 1 个", document_xml)
            self.assertIn("表格 1 个", document_xml)
            self.assertIn("PPT备注 否", text_report)
            self.assertIn("PPT图片 否", text_report)
            self.assertIn("PPT公式 否", text_report)

    def test_excel_to_pdf_generates_artifact(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            excel = processor.create_uploaded_file("scores.xlsx", make_xlsx_bytes())[0]
            task = processor.create_task({"task_type": "excel_to_pdf", "file_ids": [excel["id"]]})
            report = store.list_reports()[0]
            artifact = report["analysis"]["artifacts"][0]
            self.assertEqual(task["status"], "成功")
            self.assertEqual(artifact["output_type"], "pdf")
            self.assertTrue(Path(artifact["path"]).exists())
            self.assertTrue(Path(artifact["path"]).read_bytes().startswith(b"%PDF-1.4"))

    def test_excel_to_word_generates_artifact(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            excel = processor.create_uploaded_file("scores.xlsx", make_xlsx_bytes())[0]
            task = processor.create_task({"task_type": "excel_to_word", "file_ids": [excel["id"]]})
            report = store.list_reports()[0]
            artifact = report["analysis"]["artifacts"][0]
            self.assertEqual(task["status"], "成功")
            self.assertEqual(artifact["output_type"], "docx")
            self.assertTrue(Path(artifact["path"]).exists())
            with zipfile.ZipFile(artifact["path"]) as archive:
                self.assertIn("word/document.xml", archive.namelist())

    def test_excel_objects_are_extracted_and_written_to_word_artifact(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "objects.xlsx"
            path.write_bytes(make_xlsx_bytes(make_xlsx_object_files()))

            analyzed = DocumentAnalyzer().analyze_file(path)
            sheets = extract_xlsx_sheets(path)

            self.assertEqual(analyzed.content_summary["charts"], 1)
            self.assertEqual(analyzed.content_summary["drawings"], 1)
            self.assertEqual(analyzed.content_summary["images"], 1)
            self.assertEqual(analyzed.content_summary["comments"], 1)
            self.assertEqual(analyzed.content_summary["tables"], 1)
            self.assertEqual(sheets[0]["chart_count"], 1)
            self.assertEqual(sheets[0]["drawing_count"], 1)
            self.assertEqual(sheets[0]["image_count"], 1)
            self.assertEqual(sheets[0]["comment_count"], 1)
            self.assertEqual(sheets[0]["table_count"], 1)

            store = AppStore(tmp)
            processor = TaskProcessor(store)
            excel = processor.create_uploaded_file("objects.xlsx", path.read_bytes())[0]
            task = processor.create_task({"task_type": "excel_to_word", "file_ids": [excel["id"]]})
            report = store.list_reports()[0]
            artifact = report["analysis"]["artifacts"][0]

            self.assertEqual(task["status"], "成功")
            with zipfile.ZipFile(artifact["path"]) as archive:
                document_xml = archive.read("word/document.xml").decode("utf-8")
            self.assertIn("图表 1 个", document_xml)
            self.assertIn("图片 1 个", document_xml)
            self.assertIn("批注 1 个", document_xml)
            self.assertIn("表格对象 1 个", document_xml)

    def test_excel_conversion_range_filters_current_sheet_selected_sheet_and_cells(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            store.update_settings({"excelConversionRange": "当前工作表"})
            processor = TaskProcessor(store)
            excel = processor.create_uploaded_file("scores.xlsx", make_xlsx_multisheet_bytes())[0]
            processor.create_task({"task_type": "excel_to_word", "file_ids": [excel["id"]]})
            artifact = store.list_reports()[0]["analysis"]["artifacts"][0]
            with zipfile.ZipFile(artifact["path"]) as archive:
                document_xml = archive.read("word/document.xml").decode("utf-8")
            self.assertIn("Sheet 1: 成绩", document_xml)
            self.assertIn("转换范围：当前工作表", document_xml)
            self.assertNotIn("Sheet 2: 汇总", document_xml)

        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            excel = processor.create_uploaded_file("scores.xlsx", make_xlsx_multisheet_bytes())[0]
            processor.create_task({"task_type": "excel_to_ppt", "file_ids": [excel["id"]], "options": {"excel": {"sheetNames": ["汇总"]}}})
            artifact = store.list_reports()[0]["analysis"]["artifacts"][0]
            with zipfile.ZipFile(artifact["path"]) as archive:
                slide_xml = archive.read("ppt/slides/slide1.xml").decode("utf-8")
            self.assertIn("scores.xlsx · 汇总", slide_xml)
            self.assertIn("转换范围 指定工作表", slide_xml)
            self.assertNotIn("scores.xlsx · 成绩", slide_xml)

        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            store.update_settings({"excelConversionRange": "选区"})
            processor = TaskProcessor(store)
            excel = processor.create_uploaded_file("scores.xlsx", make_xlsx_bytes())[0]
            processor.create_task({"task_type": "excel_to_word", "file_ids": [excel["id"]], "options": {"excel": {"cellRefs": ["A2"]}}})
            artifact = store.list_reports()[0]["analysis"]["artifacts"][0]
            with zipfile.ZipFile(artifact["path"]) as archive:
                document_xml = archive.read("word/document.xml").decode("utf-8")
            self.assertIn("转换范围：选区 A2", document_xml)
            self.assertIn("A2: Alice", document_xml)
            self.assertNotIn("B2: 98", document_xml)

    def test_excel_formula_mode_can_keep_formulas_or_result_only(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            store.update_settings({"excelFormulaMode": "仅保留计算结果"})
            processor = TaskProcessor(store)
            excel = processor.create_uploaded_file("scores.xlsx", make_xlsx_bytes())[0]
            processor.create_task({"task_type": "excel_to_word", "file_ids": [excel["id"]]})
            report = store.list_reports()[0]
            artifact = report["analysis"]["artifacts"][0]
            with zipfile.ZipFile(artifact["path"]) as archive:
                document_xml = archive.read("word/document.xml").decode("utf-8")
            self.assertEqual(artifact["conversion_settings"]["excel_formula_mode"], "仅保留计算结果")
            self.assertIn("公式模式：仅保留计算结果", document_xml)
            self.assertIn("B3: 98", document_xml)
            self.assertNotIn("SUM(B2)", document_xml)

        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            store.update_settings({"excelFormulaMode": "仅保留计算结果"})
            processor = TaskProcessor(store)
            excel = processor.create_uploaded_file("scores.xlsx", make_xlsx_bytes())[0]
            processor.create_task({"task_type": "excel_to_word", "file_ids": [excel["id"]], "options": {"excel": {"retainFormulas": True}}})
            artifact = store.list_reports()[0]["analysis"]["artifacts"][0]
            with zipfile.ZipFile(artifact["path"]) as archive:
                document_xml = archive.read("word/document.xml").decode("utf-8")
            self.assertIn("公式模式：保留公式", document_xml)
            self.assertIn("B3: 98 (=SUM(B2))", document_xml)

    def test_excel_split_sheets_generates_one_artifact_per_sheet(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            store.update_settings({"excelSplitSheets": True})
            processor = TaskProcessor(store)
            excel = processor.create_uploaded_file("scores.xlsx", make_xlsx_multisheet_bytes())[0]
            task = processor.create_task({"task_type": "excel_to_word", "file_ids": [excel["id"]]})
            report = store.list_reports()[0]
            artifacts = report["analysis"]["artifacts"]

            self.assertEqual(task["status"], "成功")
            self.assertEqual(len(artifacts), 2)
            self.assertEqual([item["file_name"] for item in artifacts], ["scores-成绩.docx", "scores-汇总.docx"])
            self.assertTrue(all(item["conversion_settings"]["excel_split_sheets"] for item in artifacts))
            with zipfile.ZipFile(artifacts[0]["path"]) as archive:
                first_xml = archive.read("word/document.xml").decode("utf-8")
            with zipfile.ZipFile(artifacts[1]["path"]) as archive:
                second_xml = archive.read("word/document.xml").decode("utf-8")
            self.assertIn("Sheet 1: 成绩", first_xml)
            self.assertNotIn("Sheet 2: 汇总", first_xml)
            self.assertIn("Sheet 2: 汇总", second_xml)
            self.assertNotIn("Sheet 1: 成绩", second_xml)

        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            excel = processor.create_uploaded_file("scores.xlsx", make_xlsx_multisheet_bytes())[0]
            processor.create_task({"task_type": "excel_to_pdf", "file_ids": [excel["id"]], "options": {"excel": {"splitSheets": True}}})
            artifacts = store.list_reports()[0]["analysis"]["artifacts"]
            self.assertEqual([item["output_type"] for item in artifacts], ["pdf", "pdf"])
            self.assertEqual([item["file_name"] for item in artifacts], ["scores-成绩.pdf", "scores-汇总.pdf"])

    def test_excel_to_ppt_generates_artifact(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            excel = processor.create_uploaded_file("scores.xlsx", make_xlsx_bytes())[0]
            task = processor.create_task({"task_type": "excel_to_ppt", "file_ids": [excel["id"]]})
            report = store.list_reports()[0]
            artifact = report["analysis"]["artifacts"][0]
            self.assertEqual(task["status"], "成功")
            self.assertEqual(artifact["output_type"], "pptx")
            self.assertTrue(Path(artifact["path"]).exists())
            with zipfile.ZipFile(artifact["path"]) as archive:
                self.assertIn("ppt/presentation.xml", archive.namelist())

    def test_report_exports_html_json_pdf_xlsx_and_txt(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            word = processor.create_uploaded_file("lesson.docx", make_docx_bytes())[0]
            processor.create_task({"task_type": "word_to_ppt", "file_ids": [word["id"]]})
            report = store.list_reports()[0]
            self.assertGreaterEqual(len(report["qualityChecks"]), 8)
            self.assertIn("quality_issue_count", report)
            quality_summary = {item["label"]: item for item in report["qualitySummary"]}
            self.assertIn("处理文件总数", quality_summary)
            self.assertIn("OMML 公式数量", quality_summary)
            self.assertIn("MathType 公式数量", quality_summary)
            self.assertIn("可能丢失对象", quality_summary)
            self.assertEqual(quality_summary["处理文件总数"]["metric"], "1")
            for key in ["html_path", "json_path", "pdf_path", "xlsx_path", "txt_path"]:
                self.assertTrue(Path(report[key]).exists(), key)
            self.assertTrue(Path(report["pdf_path"]).read_bytes().startswith(b"%PDF-1.4"))
            text_report = Path(report["txt_path"]).read_text(encoding="utf-8")
            html_report = Path(report["html_path"]).read_text(encoding="utf-8")
            json_payload = json.loads(Path(report["json_path"]).read_text(encoding="utf-8"))
            self.assertIn("质量摘要", text_report)
            self.assertIn("转换输出", text_report)
            self.assertIn("转换质量检查", text_report)
            self.assertIn("公式明细", text_report)
            self.assertIn("质量摘要", html_report)
            self.assertIn("公式明细", html_report)
            self.assertIn("qualitySummary", json_payload)
            with zipfile.ZipFile(report["xlsx_path"]) as archive:
                self.assertIn("xl/worksheets/sheet1.xml", archive.namelist())
                sheet_xml = archive.read("xl/worksheets/sheet1.xml").decode("utf-8")
                self.assertIn("质量摘要", sheet_xml)
                self.assertIn("质量检查", sheet_xml)
                self.assertIn("公式明细", sheet_xml)

    def test_task_pause_and_resume(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            word = processor.create_uploaded_file("lesson.docx", make_docx_bytes())[0]
            task = Task(task_type="word_to_ppt", execute_mode="hybrid", file_ids=[word["id"]]).to_dict()
            store.save_task(task)
            paused = processor.pause_task(task["id"])
            self.assertEqual(paused["status"], "已暂停")
            self.assertTrue(any("暂停任务" in log["message"] for log in store.list_logs(task["id"])))
            resumed = processor.resume_task(task["id"])
            self.assertEqual(resumed["status"], "成功")
            self.assertTrue(any("继续任务" in log["message"] for log in store.list_logs(task["id"])))

    def test_interrupted_task_is_marked_recoverable_and_can_resume(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            word = processor.create_uploaded_file("lesson.docx", make_docx_bytes())[0]
            task = Task(task_type="word_to_ppt", execute_mode="hybrid", file_ids=[word["id"]]).to_dict()
            task["status"] = "处理中"
            task["progress"] = 42
            store.save_task(task)

            result = processor.recover_interrupted_tasks()
            interrupted = store.get_task(task["id"])
            resumed = processor.resume_task(task["id"])
            logs = store.list_logs(task["id"], limit=1000)

            self.assertEqual(result["recovered_count"], 1)
            self.assertEqual(interrupted["status"], "已中断")
            self.assertTrue(interrupted["recoverable"])
            self.assertIn("服务重启", interrupted["error_message"])
            self.assertEqual(resumed["status"], "成功")
            self.assertTrue(any("恢复中断任务" in log["message"] for log in logs))

    def test_task_recovery_summary_exposes_resume_contract_without_paths(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            word = processor.create_uploaded_file("lesson.docx", make_docx_bytes())[0]
            interrupted = Task(task_type="word_to_ppt", execute_mode="hybrid", file_ids=[word["id"]]).to_dict()
            interrupted["status"] = "处理中"
            interrupted["progress"] = 58
            interrupted["error_message"] = f"处理中断：{Path(tmp) / 'secret' / 'lesson.docx'}"
            store.save_task(interrupted)
            paused = Task(task_type="formula_precheck", execute_mode="local", file_ids=[word["id"]]).to_dict()
            paused["status"] = "已暂停"
            store.save_task(paused)
            failed = Task(task_type="pdf_to_word", execute_mode="hybrid", file_ids=[word["id"]]).to_dict()
            failed["status"] = "失败"
            failed["error_message"] = f"转换失败：{Path(tmp) / 'secret' / 'source.pdf'}"
            store.save_task(failed)
            completed = Task(task_type="word_to_ppt", execute_mode="hybrid", file_ids=[word["id"]]).to_dict()
            completed["status"] = "成功"
            completed["fail_count"] = 2
            store.save_task(completed)

            processor.recover_interrupted_tasks()
            summary = processor.task_recovery_summary()
            handler = make_handler(store, processor)
            K12RequestHandler._handle_api_get(handler, "/api/tasks/recovery-summary", {})
            response = json.loads(handler.wfile.getvalue().decode("utf-8"))
            body = json.dumps(response, ensure_ascii=False)

            self.assertEqual(summary["schema_version"], "k12.taskRecoverySummary.v1")
            self.assertEqual(summary["summary"]["interrupted_count"], 1)
            self.assertEqual(summary["summary"]["paused_count"], 1)
            self.assertEqual(summary["summary"]["failed_count"], 1)
            self.assertEqual(summary["summary"]["recoverable_count"], 2)
            self.assertNotIn(completed["id"], {item["task_id"] for item in summary["items"]})
            interrupted_item = next(item for item in summary["items"] if item["task_id"] == interrupted["id"])
            self.assertEqual(interrupted_item["recovery_status"], "restart_recoverable")
            self.assertEqual(interrupted_item["resume_endpoint"], f"/api/tasks/{interrupted['id']}/resume")
            self.assertEqual(interrupted_item["retry_endpoint"], f"/api/tasks/{interrupted['id']}/retry")
            self.assertEqual(response["taskRecoverySummary"]["contracts"]["path_policy"], "local_paths_redacted")
            self.assertEqual(handler.status, 200)
            self.assertNotIn(tmp, body)
            self.assertIn("本地路径已隐藏", body)

    def test_file_preview_extracts_structured_content(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            word = processor.create_uploaded_file("lesson.docx", make_docx_bytes())[0]
            excel = processor.create_uploaded_file("scores.xlsx", make_xlsx_bytes())[0]
            word_preview = processor.file_preview(word["id"])
            excel_preview = processor.file_preview(excel["id"])
            self.assertEqual(word_preview["status"], "ready")
            self.assertTrue(any("题目" in page["text"] for page in word_preview["pages"]))
            word_markers = [marker for page in word_preview["pages"] for marker in page.get("markers", [])]
            self.assertTrue(any(marker["label"] == "公式高亮" for marker in word_markers))
            self.assertTrue(any(marker["label"] == "OMML 公式标记" for marker in word_markers))
            self.assertEqual(excel_preview["status"], "ready")
            self.assertTrue(any("Alice" in page["text"] for page in excel_preview["pages"]))
            self.assertTrue(any(item["label"] == "公式" for item in excel_preview["objects"]))

    def test_pdf_preview_marks_scanned_pdf_for_ocr(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            pdf = processor.create_uploaded_file("scan.pdf", b"%PDF-1.4\n1 0 obj<< /Type /Page /Subtype /Image >>Math Formula endobj\n")[0]
            preview = processor.file_preview(pdf["id"])
            self.assertEqual(preview["status"], "ready")
            self.assertEqual(preview["metrics"]["pdf_type"], "扫描型 PDF")
            self.assertTrue(any("Mathpix" in item["status"] for item in preview["objects"]))

    def test_install_profiles_distinguish_windows_and_macos(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            store.update_settings({"localClientPlatform": "Windows"})
            windows = processor.capabilities()
            store.update_settings({"localClientPlatform": "macOS"})
            macos = processor.capabilities()
            self.assertIn("Office COM", windows["installProfile"]["officeAutomation"])
            self.assertIn("受限", macos["installProfile"]["officeAutomation"])
            self.assertIn("不通用", macos["mathType"]["compatibility"])
            self.assertTrue(windows["macroExecution"]["available"])
            self.assertFalse(macos["macroExecution"]["available"])
            self.assertIn("K12 Windows 本地客户端", windows["installProfile"]["recommendedInstaller"])
            self.assertIn("MathType macOS 版本", macos["installProfile"]["requiredComponents"])
            windows_interop = windows["installProfile"]["formulaObjectInterop"]
            macos_interop = macos["installProfile"]["formulaObjectInterop"]
            self.assertEqual(windows_interop["schema_version"], "k12.formulaObjectInterop.v1")
            self.assertFalse(windows_interop["platform_objects_cross_compatible"])
            self.assertFalse(macos_interop["platform_objects_cross_compatible"])
            self.assertTrue(windows_interop["same_platform_required_for_native_objects"])
            self.assertTrue(macos_interop["same_platform_required_for_native_objects"])
            self.assertIn("Windows OLE", windows_interop["native_object_format"])
            self.assertIn("macOS MathType", macos_interop["native_object_format"])
            self.assertEqual(windows_interop["fallback_formats"], ["MathML", "LaTeX", "图片"])

            store.update_settings({"localClientPlatform": "Windows", "mathtypeCompatibilityMode": "mathml-latex"})
            fallback = processor.capabilities()["installProfile"]["formulaObjectInterop"]
            self.assertFalse(fallback["native_mathtype_object_allowed"])
            self.assertTrue(fallback["fallback_required_for_cross_platform"])

    def test_install_plan_distinguishes_platform_packages_and_no_fake_download(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            handler = make_handler(store, processor)
            K12RequestHandler._handle_api_get(handler, "/api/install-plan", {"platform": ["Windows"]})
            response = json.loads(handler.wfile.getvalue().decode("utf-8"))
            windows = response["installPlan"]

            self.assertEqual(handler.status, 200)
            self.assertEqual(windows["schema_version"], "k12.installPlan.v1")
            self.assertEqual(windows["platform"], "Windows")
            self.assertEqual(windows["installer_kind"], "windows-msi")
            self.assertEqual(windows["package"]["file_name"], "K12-Local-Client-Windows-x64.msi")
            self.assertEqual(windows["package"]["status"], "待打包")
            self.assertEqual(windows["package"]["download_url"], "")
            self.assertEqual(windows["package"]["expected_location"], "本地路径已隐藏/K12-Local-Client-Windows-x64.msi")
            self.assertTrue(windows["package"]["expected_location_available"])
            self.assertNotIn(tmp, json.dumps(windows, ensure_ascii=False))
            self.assertIn("Windows 与 macOS", windows["warnings"][1])
            self.assertEqual(windows["formula_compatibility"]["schema_version"], "k12.installFormulaCompatibility.v1")
            self.assertEqual(windows["formula_compatibility"]["target_platform"], "Windows")
            self.assertFalse(windows["formula_compatibility"]["platform_objects_cross_compatible"])
            self.assertTrue(windows["formula_compatibility"]["same_platform_required_for_native_objects"])
            self.assertFalse(windows["formula_compatibility"]["native_handoff_allowed"])
            self.assertIn("client_platform_heartbeat_missing", windows["formula_compatibility"]["native_handoff_blocking_reasons"])
            self.assertTrue(windows["formula_compatibility"]["fallback_required_for_cross_platform"])
            self.assertEqual(windows["formula_compatibility"]["status"], "需同平台")
            self.assertIn("同平台心跳", windows["formula_compatibility"]["recommended_action"])
            self.assertIn("MathML", windows["formula_compatibility"]["fallback_formats"])
            self.assertIn(".msi", windows["formula_compatibility"]["package_boundary"])
            self.assertIn(".pkg", windows["formula_compatibility"]["package_boundary"])
            self.assertTrue(any("pywin32" in step["detail"] for step in windows["steps"]))
            readiness = {item["key"]: item for item in windows["readiness"]}
            self.assertEqual(readiness["client_platform_heartbeat"]["status"], "未连接")

            handler = make_handler(store, processor)
            K12RequestHandler._handle_api_get(handler, "/api/install-plan", {"platform": ["macOS"]})
            macos = json.loads(handler.wfile.getvalue().decode("utf-8"))["installPlan"]
            self.assertEqual(macos["installer_kind"], "macos-pkg")
            self.assertEqual(macos["package"]["file_name"], "K12-Local-Client-macOS-universal.pkg")
            self.assertEqual(macos["formula_compatibility"]["target_platform"], "macOS")
            self.assertFalse(macos["formula_compatibility"]["platform_objects_cross_compatible"])
            self.assertTrue(any("不要把 macOS MathType 对象直接交给 Windows" in step["detail"] for step in macos["steps"]))

            windows_package = store.installers_dir / "K12-Local-Client-Windows-x64.msi"
            windows_package.write_bytes(b"")
            invalid_windows_plan = processor.install_plan("Windows")
            self.assertEqual(invalid_windows_plan["package"]["status"], "安装包无效")
            self.assertEqual(invalid_windows_plan["package"]["download_url"], "")
            self.assertEqual(invalid_windows_plan["package"]["size"], 0)
            self.assertEqual(invalid_windows_plan["package"]["sha256"], "")
            self.assertEqual(invalid_windows_plan["package"]["expected_location"], "本地路径已隐藏/K12-Local-Client-Windows-x64.msi")
            self.assertTrue(invalid_windows_plan["package"]["expected_location_available"])
            self.assertIn("/api/installers/{file_name}?platform=...", invalid_windows_plan["package"]["path_policy"])
            self.assertNotIn(tmp, json.dumps(invalid_windows_plan, ensure_ascii=False))

            windows_package.write_bytes(b"k12 windows installer")
            macos_package = store.installers_dir / "K12-Local-Client-macOS-universal.pkg"
            macos_package.write_bytes(b"k12 mac installer")

            windows_plan = processor.install_plan("Windows")
            macos_plan = processor.install_plan("macOS")
            self.assertIn("?platform=Windows", windows_plan["package"]["download_url"])
            self.assertIn("?platform=macOS", macos_plan["package"]["download_url"])
            self.assertEqual(windows_plan["package"]["expected_location"], "本地路径已隐藏/K12-Local-Client-Windows-x64.msi")
            self.assertEqual(macos_plan["package"]["expected_location"], "本地路径已隐藏/K12-Local-Client-macOS-universal.pkg")
            self.assertNotIn(tmp, json.dumps(windows_plan, ensure_ascii=False))
            self.assertNotIn(tmp, json.dumps(macos_plan, ensure_ascii=False))

    def test_install_plan_compares_target_platform_with_client_heartbeat(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            store.update_settings({"localClientPlatform": "Windows"})
            processor = TaskProcessor(store)
            processor.record_local_client_heartbeat({"client_id": "desktop-win", "platform": "Windows"})

            matched = processor.install_plan("Windows")
            matched_readiness = {item["key"]: item for item in matched["readiness"]}

            self.assertEqual(matched_readiness["client_platform_heartbeat"]["status"], "已匹配")
            self.assertEqual(matched_readiness["client_platform_heartbeat"]["level"], "good")
            self.assertEqual(matched["formula_compatibility"]["status"], "同平台可交接")
            self.assertTrue(matched["formula_compatibility"]["native_handoff_allowed"])
            self.assertEqual(matched["formula_compatibility"]["native_handoff_blocking_reasons"], [])
            self.assertIn("同平台 MathType 对象合同", matched["formula_compatibility"]["message"])

            processor.record_local_client_heartbeat({"client_id": "desktop-mac", "platform": "macOS"})
            mismatched = processor.install_plan("Windows")
            mismatched_readiness = {item["key"]: item for item in mismatched["readiness"]}

            self.assertEqual(mismatched_readiness["client_platform_heartbeat"]["status"], "平台不符")
            self.assertEqual(mismatched_readiness["client_platform_heartbeat"]["level"], "warn")
            self.assertIn("不通用", mismatched_readiness["client_platform_heartbeat"]["detail"])
            self.assertEqual(mismatched["formula_compatibility"]["status"], "平台不符")
            self.assertFalse(mismatched["formula_compatibility"]["native_handoff_allowed"])
            self.assertIn("platform_mismatch", mismatched["formula_compatibility"]["native_handoff_blocking_reasons"])
            self.assertIn("同平台本地客户端", mismatched["formula_compatibility"]["recommended_action"])

    def test_installer_download_only_serves_existing_installer_files(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            with self.assertRaises(FileNotFoundError):
                processor.installer_download_info("K12-Local-Client-Windows-x64.msi")

            installer = store.installers_dir / "K12-Local-Client-Windows-x64.msi"
            installer.write_bytes(b"")
            with self.assertRaises(ValueError) as empty_installer:
                processor.installer_download_info(installer.name)
            self.assertIn("empty", str(empty_installer.exception))

            installer.write_bytes(b"k12 installer placeholder")
            info = processor.installer_download_info(installer.name)
            self.assertEqual(info["file_name"], installer.name)
            self.assertEqual(info["platform"], "Windows")
            self.assertEqual(info["installer_kind"], "windows-msi")
            self.assertIn(".msi", info["package_boundary"])
            self.assertIn(".pkg", info["package_boundary"])
            self.assertIn("MathType 原生对象", info["package_boundary"])
            self.assertIn("不跨平台兼容", info["formula_object_boundary"])
            self.assertEqual(info["formula_object_boundary_header"], "mathtype-native-objects-require-same-platform")
            self.assertEqual(info["fallback_formats"], ["MathML", "LaTeX", "图片"])
            self.assertEqual(info["fallback_formats_header"], "MathML,LaTeX,image")
            self.assertEqual(info["file_size"], len(b"k12 installer placeholder"))
            self.assertEqual(len(info["sha256"]), 64)
            self.assertEqual(info["package_boundary_header"], "windows-msi-and-macos-pkg-are-not-cross-platform")
            self.assertEqual(processor.installer_download_info(installer.name, "Windows")["platform"], "Windows")
            with self.assertRaises(ValueError):
                processor.installer_download_info(installer.name, "macOS")

            mac_installer = store.installers_dir / "K12-Local-Client-macOS-universal.pkg"
            mac_installer.write_bytes(b"k12 mac installer placeholder")
            mac_info = processor.installer_download_info(mac_installer.name)
            self.assertEqual(mac_info["platform"], "macOS")
            self.assertEqual(mac_info["installer_kind"], "macos-pkg")

            unknown = store.installers_dir / "notes.txt"
            unknown.write_text("not an installer", encoding="utf-8")
            with self.assertRaises(ValueError):
                processor.installer_download_info(unknown.name)
            spoofed = store.installers_dir / "K12-Local-Client-Windows-x64.pkg"
            spoofed.write_bytes(b"wrong extension")
            with self.assertRaises(ValueError):
                processor.installer_download_info(spoofed.name)

    def test_installer_download_headers_expose_platform_boundary(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            installer = store.installers_dir / "K12-Local-Client-Windows-x64.msi"
            installer.write_bytes(b"k12 installer placeholder")
            mac_installer = store.installers_dir / "K12-Local-Client-macOS-universal.pkg"
            mac_installer.write_bytes(b"k12 mac installer placeholder")
            handler = make_handler(store, processor)

            K12RequestHandler._send_installer(handler, installer.name, {"platform": ["Windows"]})

            headers = dict(handler.output_headers)
            self.assertEqual(handler.status, 200)
            self.assertEqual(handler.wfile.getvalue(), b"k12 installer placeholder")
            self.assertEqual(headers["X-K12-Platform"], "Windows")
            self.assertEqual(headers["X-K12-Installer-Kind"], "windows-msi")
            self.assertEqual(headers["X-K12-Package-Boundary"], "windows-msi-and-macos-pkg-are-not-cross-platform")
            self.assertEqual(headers["X-K12-Formula-Object-Boundary"], "mathtype-native-objects-require-same-platform")
            self.assertEqual(headers["X-K12-Formula-Fallback-Formats"], "MathML,LaTeX,image")
            self.assertEqual(len(headers["X-K12-SHA256"]), 64)
            self.assertIn(".msi", headers["Content-Disposition"])
            for key, value in handler.output_headers:
                f"{key}: {value}\r\n".encode("latin-1")

            handler = make_handler(store, processor)
            K12RequestHandler._send_installer(handler, mac_installer.name, {"platform": ["macOS"]})
            mac_headers = dict(handler.output_headers)
            self.assertEqual(handler.status, 200)
            self.assertEqual(handler.wfile.getvalue(), b"k12 mac installer placeholder")
            self.assertEqual(mac_headers["X-K12-Platform"], "macOS")
            self.assertEqual(mac_headers["X-K12-Installer-Kind"], "macos-pkg")
            self.assertEqual(mac_headers["X-K12-Package-Boundary"], "windows-msi-and-macos-pkg-are-not-cross-platform")
            self.assertEqual(mac_headers["X-K12-Formula-Object-Boundary"], "mathtype-native-objects-require-same-platform")
            self.assertEqual(mac_headers["X-K12-Formula-Fallback-Formats"], "MathML,LaTeX,image")
            self.assertIn(".pkg", mac_headers["Content-Disposition"])
            for key, value in handler.output_headers:
                f"{key}: {value}\r\n".encode("latin-1")

            handler = make_handler(store, processor)
            with self.assertRaises(JsonError) as raised:
                K12RequestHandler._send_installer(handler, installer.name, {"platform": ["macOS"]})
            self.assertEqual(raised.exception.status, 404)
            self.assertIn("does not match", raised.exception.message)

            handler = make_handler(store, processor)
            with self.assertRaises(JsonError) as raised:
                K12RequestHandler._send_installer(handler, mac_installer.name, {"platform": ["Windows"]})
            self.assertEqual(raised.exception.status, 404)
            self.assertIn("does not match", raised.exception.message)

    def test_mathtype_preflight_warns_for_platform_specific_formula_objects(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            store.update_settings({"localClientPlatform": "Windows", "mathtypeCompatibilityMode": "platform-specific"})
            processor = TaskProcessor(store)
            checks = processor.preflight_checks({"task_type": "omml_to_mathtype"})
            compatibility = next(item for item in checks if item["id"] == "mathtype_compatibility")

            self.assertEqual(compatibility["status"], "需确认")
            self.assertIn("Windows", compatibility["metric"])
            self.assertIn("不通用", compatibility["message"])
            self.assertIn("MathML/LaTeX", compatibility["recommendation"])

    def test_mathtype_preflight_passes_when_cross_platform_fallback_is_selected(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            store.update_settings({"localClientPlatform": "macOS", "mathtypeCompatibilityMode": "mathml-latex"})
            processor = TaskProcessor(store)
            checks = processor.preflight_checks({"task_type": "mathtype_format"})
            compatibility = next(item for item in checks if item["id"] == "mathtype_compatibility")

            self.assertEqual(compatibility["status"], "通过")
            self.assertIn("macOS", compatibility["metric"])
            self.assertIn("MathML/LaTeX", compatibility["metric"])

    def test_local_client_platform_preflight_fails_on_cross_platform_heartbeat(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            store.update_settings({"localClientPlatform": "Windows"})
            processor = TaskProcessor(store)
            processor.record_local_client_heartbeat(
                {
                    "client_id": "desktop-macos",
                    "status": "online",
                    "platform": "macOS",
                    "preflight": {
                        "platform": "macOS",
                        "capabilities": {"ommlDependencySearch": True, "mathTypeAutomation": True},
                    },
                }
            )
            checks = processor.preflight_checks({"task_type": "omml_to_mathtype", "execute_mode": "local"})
            platform = next(item for item in checks if item["id"] == "local_client_platform")

            self.assertEqual(platform["status"], "失败")
            self.assertEqual(platform["severity"], "error")
            self.assertEqual(platform["expected_platform"], "Windows")
            self.assertEqual(platform["heartbeat_platform"], "macOS")
            self.assertFalse(platform["platform_compatible"])
            self.assertIn("不通用", platform["message"])

    def test_global_preflight_reports_cross_platform_heartbeat_mismatch(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            store.update_settings({"localClientPlatform": "Windows"})
            processor = TaskProcessor(store)
            processor.record_local_client_heartbeat(
                {
                    "client_id": "desktop-macos",
                    "status": "online",
                    "platform": "macOS",
                    "preflight": {"platform": "macOS"},
                }
            )
            checks = processor.preflight_checks()
            platform = next(item for item in checks if item["id"] == "local_client_platform")

            self.assertEqual(platform["status"], "失败")
            self.assertEqual(platform["expected_platform"], "Windows")
            self.assertEqual(platform["heartbeat_platform"], "macOS")

    def test_local_client_platform_preflight_passes_when_heartbeat_matches(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            store.update_settings({"localClientPlatform": "macOS"})
            processor = TaskProcessor(store)
            processor.record_local_client_heartbeat(
                {
                    "client_id": "desktop-macos",
                    "status": "online",
                    "platform": "macOS",
                    "preflight": {
                        "platform": "macOS",
                        "capabilities": {"ommlDependencySearch": True, "mathTypeAutomation": True},
                    },
                }
            )
            checks = processor.preflight_checks({"task_type": "mathtype_format", "execute_mode": "local"})
            platform = next(item for item in checks if item["id"] == "local_client_platform")

            self.assertEqual(platform["status"], "通过")
            self.assertEqual(platform["expected_platform"], "macOS")
            self.assertEqual(platform["heartbeat_platform"], "macOS")
            self.assertTrue(platform["platform_compatible"])
            self.assertIn("macOS", platform["metric"])

    def test_local_component_preflight_failure_enters_report_failures(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            store.update_settings({"localClientPlatform": "macOS"})
            processor = TaskProcessor(store)
            processor.record_local_client_heartbeat(
                {
                    "client_id": "desktop-formula",
                    "status": "online",
                    "platform": "macOS",
                    "preflight": {
                        "platform": "macOS",
                        "components": {
                            "omml_dependency": {"label": "OMML 依赖文件", "available": True, "status": "available"},
                            "mathtype": {"label": "MathType 组件", "available": False, "status": "missing", "path": "/Applications/MathType.app"},
                        },
                        "capabilities": {"ommlDependencySearch": True, "mathTypeAutomation": False},
                    },
                }
            )
            word = processor.create_uploaded_file("lesson_omml.docx", make_docx_bytes())[0]
            checks = processor.preflight_checks({"task_type": "omml_to_mathtype", "file_ids": [word["id"]], "execute_mode": "local"})
            component_check = next(item for item in checks if item["id"] == "local_client_components")

            self.assertEqual(component_check["status"], "失败")
            self.assertEqual(component_check["severity"], "error")
            self.assertEqual(component_check["missing_capabilities"], ["mathTypeAutomation"])
            self.assertIn("MathType 自动化", component_check["metric"])
            self.assertNotIn("/Applications", json.dumps(component_check, ensure_ascii=False))

            task = processor.create_task({"task_type": "omml_to_mathtype", "file_ids": [word["id"]]})
            report = store.list_reports()[0]
            quality = {item["id"]: item for item in report["qualityChecks"]}
            preflight_failure = next(row for row in report["failureRows"] if row["source_id"] == "local_client_components")

            self.assertEqual(task["status"], "失败")
            self.assertEqual(quality["local_client_components"]["status"], "失败")
            self.assertEqual(preflight_failure["category"], "系统预检")
            self.assertIn("MathType 自动化", preflight_failure["message"])

    def test_capabilities_expose_local_connection_authorization(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            store.update_settings(
                {
                    "localClientEnabled": True,
                    "allowCloudSync": True,
                    "allowWebLaunchLocalClient": True,
                    "allowTaskStatusCloudSync": True,
                    "sensitiveFilesPreferLocal": False,
                }
            )
            processor = TaskProcessor(store)
            connection = processor.capabilities()["localConnection"]

            self.assertTrue(connection["client_enabled"])
            self.assertTrue(connection["web_launch_allowed"])
            self.assertTrue(connection["cloud_sync_allowed"])
            self.assertTrue(connection["task_status_cloud_sync_allowed"])
            self.assertFalse(connection["sensitive_files_prefer_local"])

    def test_local_task_payload_hands_off_parameters_to_authorized_client(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            store.update_settings(
                {
                    "allowWebLaunchLocalClient": True,
                    "allowCloudSync": True,
                    "allowTaskStatusCloudSync": True,
                    "localSecurityToken": "secret-token",
                }
            )
            processor = TaskProcessor(store)
            word = processor.create_uploaded_file("macro.docm", make_docx_bytes())[0]
            task = processor.create_task(
                {
                    "task_type": "macro_sequence",
                    "file_ids": [word["id"]],
                    "options": {
                        "workflowOrder": ["pdf_to_word", "macro_sequence", "small_image_scan"],
                        "selectedMacros": [{"id": "macro_clean_empty_paragraphs", "execute_order": 1}],
                        "confirmMacroRisk": True,
                        "macroBackup": True,
                    },
                }
            )

            payload = processor.local_task_payload(task["id"])

            self.assertEqual(payload["schema_version"], "k12.localTaskPayload.v1")
            self.assertEqual(payload["handoff"]["status"], "ready")
            self.assertTrue(payload["handoff"]["requires_local_client"])
            self.assertEqual(payload["client_readiness"]["schema_version"], "k12.localClientReadiness.v1")
            self.assertEqual(payload["client_readiness"]["status"], "needs_heartbeat")
            self.assertIn("macroExecution", payload["client_readiness"]["missing_capabilities"])
            self.assertEqual(payload["task"]["execute_mode"], "local")
            self.assertEqual(payload["workflow_plan"]["schema_version"], "k12.workflowPlan.v1")
            self.assertEqual(payload["workflow_plan"]["current_task"], "macro_sequence")
            self.assertEqual(payload["workflow_plan"]["current_index"], 2)
            self.assertEqual(payload["workflow_plan"]["order"], ["pdf_to_word", "macro_sequence", "small_image_scan"])
            self.assertEqual(payload["workflow_plan"]["labels"], ["PDF 转 Word", "Word 宏顺序执行", "微小图片检索"])
            self.assertEqual(payload["files"][0]["input_path"], word["storage_path"])
            self.assertTrue(payload["files"][0]["input_path_exists"])
            self.assertTrue(payload["sync"]["result_upload_allowed"])
            self.assertTrue(payload["sync"]["task_status_cloud_sync_allowed"])
            macro_action = next(action for action in payload["local_actions"] if action["type"] == "macro_sequence")
            self.assertEqual(macro_action["status"], "queued")
            self.assertEqual(macro_action["macros"][0]["macro_name"], "CleanEmptyParagraphs")
            self.assertTrue(macro_action["macros"][0]["backup_path"])
            execution_plan = payload["desktop_execution_plan"]
            self.assertEqual(execution_plan["schema_version"], "k12.desktopExecutionPlan.v1")
            self.assertEqual(execution_plan["status"], "waiting_for_heartbeat")
            self.assertFalse(execution_plan["web_executes_native_documents"])
            self.assertFalse(execution_plan["current_companion_cli_executes_native_documents"])
            self.assertEqual(execution_plan["native_action_count"], 1)
            self.assertEqual(execution_plan["authorization_gates"]["local_security_token_configured"], True)
            plan_action = next(action for action in execution_plan["actions"] if action["type"] == "macro_sequence")
            self.assertEqual(plan_action["gate_status"], "waiting_for_heartbeat")
            self.assertEqual(plan_action["required_capabilities"][0]["key"], "macroExecution")
            self.assertTrue(plan_action["input_paths"][0])
            self.assertTrue(plan_action["backup_paths"][0])
            self.assertTrue(plan_action["output_contract"]["output_directory"])
            self.assertNotIn("secret-token", json.dumps(payload, ensure_ascii=False))

            handler = make_handler(store, processor)
            K12RequestHandler._handle_api_get(handler, f"/api/tasks/{task['id']}/local-payload", {})
            response = json.loads(handler.wfile.getvalue().decode("utf-8"))
            self.assertEqual(handler.status, 200)
            self.assertEqual(response["localPayload"]["task"]["id"], task["id"])
            self.assertEqual(response["localPayload"]["files"][0]["input_path"], word["storage_path"])
            self.assertEqual(response["localPayload"]["workflow_plan"]["current_index"], 2)

            readiness = processor.local_task_readiness(task["id"])
            self.assertEqual(readiness["workflow_plan"]["order"], payload["workflow_plan"]["order"])
            self.assertEqual(readiness["path_policy"], "no local file paths or tokens are returned")

    def test_local_task_payload_includes_platform_formula_delivery_contract(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            store.update_settings(
                {
                    "localClientPlatform": "macOS",
                    "mathtypeCompatibilityMode": "mathml-latex",
                    "keepFormulaImages": True,
                }
            )
            processor = TaskProcessor(store)
            word = processor.create_uploaded_file("lesson.docx", make_docx_bytes())[0]
            task = processor.create_task({"task_type": "omml_to_mathtype", "file_ids": [word["id"]]})

            payload = processor.local_task_payload(task["id"])
            contract = payload["formula_delivery"]
            omml_action = next(action for action in payload["local_actions"] if action["type"] == "omml_mathtype")
            action_contract = omml_action["formula_delivery"]

            self.assertEqual(contract["schema_version"], "k12.formulaDeliveryContract.v1")
            self.assertEqual(contract["platform"], "macOS")
            self.assertEqual(contract["compatibility_mode"], "mathml-latex")
            self.assertTrue(contract["cross_platform_safe"])
            self.assertFalse(contract["platform_objects_cross_compatible"])
            self.assertFalse(contract["native_mathtype_object_allowed"])
            self.assertIn("MathML", contract["output_priority"])
            self.assertIn("LaTeX", contract["fallback_formats"])
            self.assertIn("不通用", contract["message"])
            self.assertEqual(action_contract["compatibility_mode"], contract["compatibility_mode"])
            self.assertFalse(action_contract["native_mathtype_object_allowed"])

    def test_task_options_cannot_override_formula_delivery_platform_contract(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            store.update_settings(
                {
                    "localClientPlatform": "Windows",
                    "mathtypeCompatibilityMode": "platform-specific",
                }
            )
            processor = TaskProcessor(store)
            word = processor.create_uploaded_file("lesson.docx", make_docx_bytes())[0]
            task = processor.create_task(
                {
                    "task_type": "omml_to_mathtype",
                    "file_ids": [word["id"]],
                    "options": {"localClientPlatform": "macOS", "mathtypeCompatibilityMode": "mathml-latex"},
                }
            )

            payload = processor.local_task_payload(task["id"])
            contract = payload["formula_delivery"]
            omml_action = next(action for action in payload["local_actions"] if action["type"] == "omml_mathtype")
            plan_action = next(action for action in payload["desktop_execution_plan"]["actions"] if action["type"] == "omml_mathtype")

            self.assertEqual(task["options"]["localClientPlatform"], "macOS")
            self.assertEqual(contract["platform"], "Windows")
            self.assertEqual(contract["compatibility_mode"], "platform-specific")
            self.assertFalse(contract["platform_objects_cross_compatible"])
            self.assertTrue(contract["native_object_requires_same_platform"])
            self.assertFalse(contract["cross_platform_safe"])
            self.assertEqual(omml_action["formula_delivery"]["platform"], "Windows")
            self.assertEqual(plan_action["formula_delivery"]["platform"], "Windows")
            self.assertTrue(plan_action["formula_delivery"]["native_object_requires_same_platform"])

    def test_local_task_payload_reports_client_readiness_from_heartbeat_preflight(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            store.update_settings(
                {
                    "allowWebLaunchLocalClient": True,
                    "localSecurityToken": "secret-token",
                    "localClientPlatform": "Windows",
                }
            )
            processor = TaskProcessor(store)
            processor.record_local_client_heartbeat(
                {
                    "client_id": "desktop-ready",
                    "status": "online",
                    "platform": "Windows",
                    "capabilities": {"macroExecution": True},
                    "preflight": {
                        "platform": "Windows",
                        "components": {
                            "word": {"label": "Word 桌面组件", "available": True, "status": "available", "path": "C:/Program Files/Microsoft Office/WINWORD.EXE"},
                            "secret": {"label": "token", "available": True, "status": "available"},
                        },
                        "capabilities": {
                            "officeAutomation": True,
                            "mathTypeAutomation": True,
                            "macroExecution": True,
                            "ommlDependencySearch": True,
                        },
                    },
                }
            )
            word = processor.create_uploaded_file("macro.docm", make_docx_bytes())[0]
            task = processor.create_task(
                {
                    "task_type": "macro_sequence",
                    "file_ids": [word["id"]],
                    "options": {
                        "selectedMacros": [{"id": "macro_clean_empty_paragraphs", "execute_order": 1}],
                        "confirmMacroRisk": True,
                    },
                }
            )

            payload = processor.local_task_payload(task["id"])
            readiness = payload["client_readiness"]
            readiness_json = json.dumps(readiness, ensure_ascii=False)

            self.assertEqual(readiness["status"], "ready_for_handoff")
            self.assertEqual(readiness["platform"], "Windows")
            self.assertEqual(readiness["expected_platform"], "Windows")
            self.assertTrue(readiness["platform_compatible"])
            self.assertEqual(readiness["client_id"], "desktop-ready")
            self.assertEqual(readiness["missing_capabilities"], [])
            self.assertFalse(readiness["executes_native_documents"])
            self.assertTrue(readiness["capabilities"]["macroExecution"])
            self.assertIn("word", readiness["components"])
            self.assertNotIn("secret", readiness["components"])
            self.assertNotIn("Program Files", readiness_json)
            self.assertNotIn("secret-token", json.dumps(payload, ensure_ascii=False))

    def test_local_task_payload_flags_missing_formula_client_capability(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            store.update_settings({"localClientPlatform": "macOS", "localSecurityToken": "secret-token"})
            processor = TaskProcessor(store)
            processor.record_local_client_heartbeat(
                {
                    "client_id": "desktop-formula",
                    "status": "online",
                    "platform": "macOS",
                    "preflight": {
                        "platform": "macOS",
                        "components": {
                            "omml_dependency": {"label": "OMML 依赖文件", "available": True, "status": "available"},
                            "mathtype": {"label": "MathType 组件", "available": False, "status": "missing"},
                        },
                        "capabilities": {"ommlDependencySearch": True, "mathTypeAutomation": False},
                    },
                }
            )
            word = processor.create_uploaded_file("lesson_omml.docx", make_docx_bytes())[0]
            task = processor.create_task({"task_type": "omml_to_mathtype", "file_ids": [word["id"]]})

            readiness = processor.local_task_payload(task["id"])["client_readiness"]
            required_keys = [item["key"] for item in readiness["required_capabilities"]]

            self.assertEqual(readiness["status"], "missing_capability")
            self.assertIn("ommlDependencySearch", required_keys)
            self.assertIn("mathTypeAutomation", required_keys)
            self.assertEqual(readiness["missing_capabilities"], ["mathTypeAutomation"])

    def test_local_task_payload_blocks_cross_platform_mathtype_handoff(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            store.update_settings({"localClientPlatform": "Windows", "localSecurityToken": "secret-token"})
            processor = TaskProcessor(store)
            processor.record_local_client_heartbeat(
                {
                    "client_id": "desktop-macos",
                    "status": "online",
                    "platform": "macOS",
                    "preflight": {
                        "platform": "macOS",
                        "components": {
                            "mathtype": {"label": "MathType 组件", "available": True, "status": "available"},
                            "omml_dependency": {"label": "OMML 依赖文件", "available": True, "status": "available"},
                        },
                        "capabilities": {"ommlDependencySearch": True, "mathTypeAutomation": True},
                    },
                }
            )
            word = processor.create_uploaded_file("lesson_omml.docx", make_docx_bytes())[0]
            task = processor.create_task({"task_type": "omml_to_mathtype", "file_ids": [word["id"]]})

            readiness = processor.local_task_payload(task["id"])["client_readiness"]

            self.assertEqual(readiness["status"], "platform_mismatch")
            self.assertEqual(readiness["expected_platform"], "Windows")
            self.assertEqual(readiness["platform"], "macOS")
            self.assertFalse(readiness["platform_compatible"])
            self.assertEqual(readiness["missing_capabilities"], [])
            self.assertIn("不通用", readiness["platform_compatibility_message"])

    def test_local_task_readiness_endpoint_is_path_safe_for_ui(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            store.update_settings({"localSecurityToken": "secret-token", "localClientPlatform": "Windows"})
            processor = TaskProcessor(store)
            processor.record_local_client_heartbeat(
                {
                    "client_id": "desktop-ui",
                    "platform": "Windows",
                    "capabilities": {"macroExecution": True},
                    "preflight": {
                        "platform": "Windows",
                        "components": {
                            "word": {"label": "Word 桌面组件", "available": True, "status": "available", "path": "C:/Program Files/Microsoft Office/WINWORD.EXE"},
                        },
                        "capabilities": {"macroExecution": True},
                    },
                }
            )
            word = processor.create_uploaded_file("macro.docm", make_docx_bytes())[0]
            task = processor.create_task(
                {
                    "task_type": "macro_sequence",
                    "file_ids": [word["id"]],
                    "options": {
                        "selectedMacros": [{"id": "macro_clean_empty_paragraphs", "execute_order": 1}],
                        "confirmMacroRisk": True,
                    },
                }
            )
            handler = make_handler(store, processor)

            K12RequestHandler._handle_api_get(handler, f"/api/tasks/{task['id']}/local-readiness", {})
            response = json.loads(handler.wfile.getvalue().decode("utf-8"))
            payload = response["localReadiness"]
            payload_json = json.dumps(payload, ensure_ascii=False)

            self.assertEqual(handler.status, 200)
            self.assertEqual(payload["schema_version"], "k12.localTaskReadinessPayload.v1")
            self.assertEqual(payload["client_readiness"]["schema_version"], "k12.localClientReadiness.v1")
            self.assertEqual(payload["client_readiness"]["status"], "ready_for_handoff")
            self.assertEqual(payload["local_actions"][0]["required_capabilities"][0]["key"], "macroExecution")
            self.assertEqual(payload["desktop_execution_plan"]["schema_version"], "k12.desktopExecutionPlan.v1")
            self.assertEqual(payload["desktop_execution_plan"]["status"], "ready_for_native_client")
            self.assertFalse(payload["desktop_execution_plan"]["web_executes_native_documents"])
            self.assertEqual(payload["desktop_execution_plan"]["actions"][0]["gate_status"], "ready")
            self.assertEqual(payload["desktop_execution_plan"]["actions"][0]["output_contract"]["output_directory_display"], "本地路径已隐藏")
            self.assertEqual(payload["path_policy"], "no local file paths or tokens are returned")
            self.assertNotIn(word["storage_path"], payload_json)
            self.assertNotIn("Program Files", payload_json)
            self.assertNotIn("secret-token", payload_json)

    def test_local_task_readiness_endpoint_exposes_redacted_formula_delivery_contract(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            store.update_settings(
                {
                    "localSecurityToken": "secret-token",
                    "localClientPlatform": "Windows",
                    "mathtypeCompatibilityMode": "platform-specific",
                }
            )
            processor = TaskProcessor(store)
            processor.record_local_client_heartbeat(
                {
                    "client_id": "formula-desktop",
                    "platform": "Windows",
                    "preflight": {
                        "platform": "Windows",
                        "components": {
                            "mathtype": {"label": "MathType 组件", "available": True, "status": "available", "path": "C:/Program Files/MathType/MathType.exe"},
                            "omml_dependency": {"label": "OMML 依赖文件", "available": True, "status": "available", "path": "C:/Users/a1/.k12/omml2mml.xsl"},
                        },
                        "capabilities": {"ommlDependencySearch": True, "mathTypeAutomation": True},
                    },
                }
            )
            word = processor.create_uploaded_file("lesson_omml.docx", make_docx_bytes())[0]
            task = processor.create_task({"task_type": "omml_to_mathtype", "file_ids": [word["id"]]})
            handler = make_handler(store, processor)

            K12RequestHandler._handle_api_get(handler, f"/api/tasks/{task['id']}/local-readiness", {})
            response = json.loads(handler.wfile.getvalue().decode("utf-8"))
            payload = response["localReadiness"]
            payload_json = json.dumps(payload, ensure_ascii=False)
            plan_contract = payload["desktop_execution_plan"]["formula_delivery"]
            action_contract = next(
                action["formula_delivery"]
                for action in payload["desktop_execution_plan"]["actions"]
                if action["type"] == "omml_mathtype"
            )

            self.assertEqual(handler.status, 200)
            self.assertEqual(payload["client_readiness"]["status"], "ready_for_handoff")
            self.assertEqual(plan_contract["platform"], "Windows")
            self.assertEqual(plan_contract["compatibility_mode"], "platform-specific")
            self.assertFalse(plan_contract["platform_objects_cross_compatible"])
            self.assertTrue(plan_contract["native_object_requires_same_platform"])
            self.assertTrue(plan_contract["native_mathtype_object_allowed"])
            self.assertEqual(action_contract["platform"], "Windows")
            self.assertFalse(action_contract["platform_objects_cross_compatible"])
            self.assertTrue(action_contract["native_object_requires_same_platform"])
            self.assertNotIn(word["storage_path"], payload_json)
            self.assertNotIn("Program Files", payload_json)
            self.assertNotIn("secret-token", payload_json)

    def test_local_task_payload_endpoint_requires_configured_token(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            word = processor.create_uploaded_file("lesson.docx", make_docx_bytes())[0]
            task = processor.create_task({"task_type": "word_to_ppt", "file_ids": [word["id"]]})
            handler = make_handler(store, processor)

            with self.assertRaises(JsonError) as raised:
                K12RequestHandler._handle_api_get(handler, f"/api/tasks/{task['id']}/local-payload", {})

            self.assertEqual(raised.exception.status, 403)
            self.assertIn("安全令牌", raised.exception.message)

    def test_local_task_status_sync_updates_progress_outputs_and_logs(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            store.update_settings({"allowTaskStatusCloudSync": True, "localSecurityToken": "secret-token"})
            processor = TaskProcessor(store)
            word = processor.create_uploaded_file("lesson.docx", make_docx_bytes())[0]
            task = processor.create_task({"task_type": "word_to_ppt", "file_ids": [word["id"]]})
            output_path = str(Path(tmp) / "lesson.pptx")

            synced = processor.sync_local_task_status(
                task["id"],
                {
                    "status": "running",
                    "progress": 63,
                    "message": "本地客户端正在转换",
                    "outputs": [{"name": "lesson.pptx", "path": output_path, "output_type": "pptx", "status": "生成中"}],
                    "dryRunExecution": {
                        "schema_version": "k12.localDryRunExecution.v1",
                        "task": {"id": task["id"], "task_type": "word_to_ppt", "task_label": "Word 转 PPT"},
                        "plan_status": "ready_for_native_client",
                        "native_execution_allowed": True,
                        "companion_cli_native_execution": False,
                        "action_count": 1,
                        "ready_action_count": 1,
                        "blocked_action_count": 0,
                        "waiting_action_count": 0,
                        "actions": [
                            {
                                "action_id": f"{task['id']}:office_conversion",
                                "type": "office_conversion",
                                "label": "Office 转换",
                                "payload_status": "queued",
                                "gate_status": "ready",
                                "dry_run_status": "ready_for_native_executor",
                                "native_execution_performed": False,
                                "required_capabilities": [{"key": "officeAutomation", "label": "Office 自动化", "available": True}],
                                "step_count": 3,
                                "required_step_count": 3,
                                "output_artifact_types": ["pptx"],
                                "input_paths": [str(Path(tmp) / "private.docx")],
                                "backup_paths": [str(Path(tmp) / "backup.docx")],
                            }
                        ],
                    },
                    "resultUploadRequested": True,
                },
            )

            self.assertEqual(synced["status"], "处理中")
            self.assertEqual(synced["progress"], 63)
            self.assertEqual(synced["local_client_outputs"][0]["path"], output_path)
            self.assertEqual(synced["local_execution_summary"]["schema_version"], "k12.localDryRunExecution.v1")
            self.assertEqual(synced["local_execution_summary"]["ready_action_count"], 1)
            self.assertEqual(synced["local_execution_summary"]["actions"][0]["type"], "office_conversion")
            self.assertFalse(synced["local_execution_summary"]["actions"][0]["native_execution_performed"])
            self.assertNotIn(tmp, json.dumps(synced["local_execution_summary"], ensure_ascii=False))
            self.assertEqual(synced["local_sync"]["result_upload_status"], "not_authorized")
            self.assertEqual(synced["local_sync"]["dry_run_execution_status"], "ready_for_native_client")
            self.assertEqual(synced["local_sync"]["dry_run_ready_action_count"], 1)
            self.assertTrue(any("本地客户端同步状态" in log["message"] for log in store.list_logs(task["id"], limit=1000)))

            store.update_settings({"allowCloudSync": True})
            handler = make_handler(store, processor)
            K12RequestHandler._handle_api_post(
                handler,
                f"/api/tasks/{task['id']}/local-sync",
                {
                    "status": "completed",
                    "message": "本地客户端完成转换",
                    "outputs": [{"name": "lesson.pptx", "path": output_path, "output_type": "pptx", "status": "已生成"}],
                    "resultUploadRequested": True,
                },
            )
            response = json.loads(handler.wfile.getvalue().decode("utf-8"))
            saved = store.get_task(task["id"])

            self.assertEqual(handler.status, 200)
            self.assertEqual(saved["status"], "成功")
            self.assertEqual(saved["progress"], 100)
            self.assertEqual(saved["local_sync"]["result_upload_status"], "queued")
            self.assertNotIn(tmp, json.dumps(response, ensure_ascii=False))
            self.assertIn("本地路径已隐藏", json.dumps(response, ensure_ascii=False))
            queue = processor.local_result_upload_queue()
            self.assertEqual(queue["schema_version"], "k12.localResultUploadQueue.v1")
            self.assertTrue(queue["cloud_sync_allowed"])
            self.assertEqual(queue["summary"]["queued"], 1)
            self.assertEqual(queue["items"][0]["upload_status"], "queued")
            self.assertEqual(queue["items"][0]["outputs"][0]["path_display"], "本地路径已隐藏")
            self.assertNotIn(output_path, json.dumps(queue, ensure_ascii=False))

    def test_local_task_status_sync_rejects_mismatched_execution_summary_identity(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            store.update_settings({"allowTaskStatusCloudSync": True, "localSecurityToken": "secret-token"})
            processor = TaskProcessor(store)
            word = processor.create_uploaded_file("lesson.docx", make_docx_bytes())[0]
            task = processor.create_task({"task_type": "word_to_ppt", "file_ids": [word["id"]]})

            with self.assertRaises(ValueError) as wrong_id:
                processor.sync_local_task_status(
                    task["id"],
                    {
                        "status": "running",
                        "progress": 20,
                        "dryRunExecution": {
                            "schema_version": "k12.localDryRunExecution.v1",
                            "task": {"id": "task_other", "task_type": "word_to_ppt", "task_label": "Word 转 PPT"},
                        },
                    },
                )
            self.assertIn("任务 ID 不匹配", str(wrong_id.exception))

            with self.assertRaises(ValueError) as wrong_type:
                processor.sync_local_task_status(
                    task["id"],
                    {
                        "status": "running",
                        "progress": 20,
                        "dryRunExecution": {
                            "schema_version": "k12.localDryRunExecution.v1",
                            "task": {"id": task["id"], "task_type": "pdf_to_word", "task_label": "PDF 转 Word"},
                        },
                    },
                )
            self.assertIn("任务类型不匹配", str(wrong_type.exception))
            saved = store.get_task(task["id"])
            self.assertFalse(saved.get("local_execution_summary"))
            self.assertEqual(saved["status"], "成功")

    def test_local_result_upload_registration_requires_cloud_sync_and_redacts_paths(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            store.update_settings({"allowTaskStatusCloudSync": True, "allowCloudSync": True})
            processor = TaskProcessor(store)
            word = processor.create_uploaded_file("lesson.docx", make_docx_bytes())[0]
            task = processor.create_task({"task_type": "word_to_ppt", "file_ids": [word["id"]]})
            output_path = str(Path(tmp) / "private" / "lesson.pptx")

            with self.assertRaises(ValueError) as missing_token:
                processor.register_local_result_upload(
                    {
                        "task_id": task["id"],
                        "outputs": [{"name": "lesson.pptx", "path": output_path, "size": 128}],
                    }
                )

            self.assertIn("安全令牌", str(missing_token.exception))
            store.update_settings({"allowCloudSync": False, "localSecurityToken": "secret-token"})
            with self.assertRaises(ValueError) as blocked:
                processor.register_local_result_upload(
                    {
                        "task_id": task["id"],
                        "outputs": [{"name": "lesson.pptx", "path": output_path, "size": 128}],
                    }
                )

            self.assertIn("未授权", str(blocked.exception))
            store.update_settings({"allowCloudSync": True})
            handler = make_handler(store, processor)
            K12RequestHandler._handle_api_post(
                handler,
                "/api/local-client/uploads",
                {
                    "task_id": task["id"],
                    "client_id": "desktop-1",
                    "package_sha256": "A" * 64,
                    "outputs": [
                        {
                            "name": "lesson.pptx",
                            "path": output_path,
                            "output_type": "pptx",
                            "size": "128",
                            "sha256": "B" * 64,
                        },
                        {
                            "path": r"C:\Users\a1\secret\answer.docx",
                            "output_type": "docx",
                            "size": 64,
                        }
                    ],
                },
            )
            response = json.loads(handler.wfile.getvalue().decode("utf-8"))
            saved = store.get_task(task["id"])
            queue = response["uploadQueue"]

            self.assertEqual(handler.status, 200)
            self.assertEqual(response["upload"]["schema_version"], "k12.localResultUpload.v1")
            self.assertEqual(response["upload"]["status"], "registered")
            self.assertEqual(response["upload"]["package_sha256"], "a" * 64)
            self.assertEqual(response["upload"]["outputs"][0]["sha256"], "b" * 64)
            self.assertEqual(response["upload"]["outputs"][1]["name"], "answer.docx")
            self.assertEqual(saved["local_sync"]["result_upload_status"], "registered")
            self.assertEqual(queue["summary"]["registered"], 1)
            self.assertEqual(queue["summary"]["output_count"], 2)
            self.assertEqual(queue["items"][0]["upload_status"], "registered")
            self.assertEqual(queue["items"][0]["upload_status_label"], "已登记待云端接收")
            self.assertEqual(queue["items"][0]["outputs"][0]["path_display"], "本地路径已隐藏")
            self.assertTrue(queue["items"][0]["manifest_available"])
            self.assertEqual(queue["items"][0]["manifest_endpoint"], f"/api/local-client/uploads/{response['upload']['upload_id']}/manifest")
            manifest = processor.local_result_upload_manifest(response["upload"]["upload_id"])
            self.assertEqual(manifest["schema_version"], "k12.localResultUploadManifest.v1")
            self.assertEqual(manifest["upload_id"], response["upload"]["upload_id"])
            self.assertEqual(manifest["receive_contract"]["content_transfer"], "not_included")
            self.assertEqual(manifest["package"]["package_sha256"], "a" * 64)
            self.assertEqual(manifest["package"]["output_count"], 2)
            self.assertEqual(manifest["integrity"]["output_sha256_count"], 1)
            self.assertTrue(manifest["integrity"]["all_outputs_named"])
            self.assertEqual(manifest["outputs"][0]["path_display"], "本地路径已隐藏")
            manifest_handler = make_handler(store, processor)
            K12RequestHandler._handle_api_get(manifest_handler, f"/api/local-client/uploads/{response['upload']['upload_id']}/manifest", {})
            manifest_response = json.loads(manifest_handler.wfile.getvalue().decode("utf-8"))
            self.assertEqual(manifest_handler.status, 200)
            self.assertEqual(manifest_response["uploadManifest"]["upload_id"], response["upload"]["upload_id"])
            self.assertNotIn(output_path, json.dumps(response, ensure_ascii=False))
            self.assertNotIn("C:\\Users\\a1\\secret", json.dumps(response, ensure_ascii=False))
            self.assertNotIn(output_path, json.dumps(manifest_response, ensure_ascii=False))
            self.assertNotIn("C:\\Users\\a1\\secret", json.dumps(manifest_response, ensure_ascii=False))
            self.assertTrue(any("本地结果上传登记" in log["message"] for log in store.list_logs(task["id"], limit=1000)))

    def test_local_result_upload_can_receive_content_package_without_path_leak(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            store.update_settings({"allowCloudSync": True, "localSecurityToken": "secret-token"})
            processor = TaskProcessor(store)
            word = processor.create_uploaded_file("lesson.docx", make_docx_bytes())[0]
            task = processor.create_task({"task_type": "word_to_ppt", "file_ids": [word["id"]]})
            content = b"K12 uploaded result package"
            digest = hashlib.sha256(content).hexdigest()
            handler = make_handler(store, processor)

            K12RequestHandler._handle_api_post(
                handler,
                "/api/local-client/uploads",
                {
                    "task_id": task["id"],
                    "client_id": "desktop-content",
                    "outputs": [
                        {
                            "name": "lesson.pptx",
                            "path": str(Path(tmp) / "private" / "lesson.pptx"),
                            "output_type": "pptx",
                            "size": len(content),
                            "sha256": digest,
                        }
                    ],
                    "files": [
                        {
                            "name": "lesson.pptx",
                            "output_type": "pptx",
                            "content_base64": base64.b64encode(content).decode("ascii"),
                            "sha256": digest,
                        }
                    ],
                },
            )
            response = json.loads(handler.wfile.getvalue().decode("utf-8"))
            upload = response["upload"]
            manifest = processor.local_result_upload_manifest(upload["upload_id"])
            serialized = json.dumps({"response": response, "manifest": manifest}, ensure_ascii=False)
            received_dir = store.cloud_uploads_dir / upload["upload_id"]

            self.assertEqual(handler.status, 200)
            self.assertEqual(upload["receive_mode"], "content_received")
            self.assertEqual(upload["received_file_count"], 1)
            self.assertEqual(upload["received_files"][0]["sha256"], digest)
            self.assertEqual(upload["received_files"][0]["path_display"], "云端接收区路径已隐藏")
            self.assertTrue(any(path.read_bytes() == content for path in received_dir.iterdir()))
            self.assertEqual(manifest["receive_state"], "received")
            self.assertEqual(manifest["receive_contract"]["content_transfer"], "included")
            self.assertEqual(manifest["package"]["received_file_count"], 1)
            self.assertTrue(manifest["integrity"]["all_received_files_hashed"])
            self.assertIn("云端接收区路径已隐藏", serialized)
            self.assertNotIn(str(Path(tmp) / "private"), serialized)
            self.assertNotIn(base64.b64encode(content).decode("ascii"), serialized)

    def test_local_result_upload_rejects_hash_mismatch_before_persisting_content(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            store.update_settings({"allowCloudSync": True, "localSecurityToken": "secret-token"})
            processor = TaskProcessor(store)
            word = processor.create_uploaded_file("lesson.docx", make_docx_bytes())[0]
            task = processor.create_task({"task_type": "word_to_ppt", "file_ids": [word["id"]]})
            content = b"K12 tampered upload package"

            with self.assertRaises(ValueError) as raised:
                processor.register_local_result_upload(
                    {
                        "task_id": task["id"],
                        "client_id": "desktop-content",
                        "files": [
                            {
                                "name": "lesson.pptx",
                                "output_type": "pptx",
                                "content_base64": base64.b64encode(content).decode("ascii"),
                                "sha256": "0" * 64,
                            }
                        ],
                    }
            )

            self.assertIn("哈希不匹配", str(raised.exception))
            self.assertFalse(store.get_task(task["id"]).get("local_result_uploads"))
            upload_dirs = list(store.cloud_uploads_dir.glob("local_upload_*"))
            self.assertTrue(all(not any(path.is_file() for path in directory.rglob("*")) for directory in upload_dirs))

    def test_local_companion_cli_helpers_redact_paths_and_build_dry_run_sync(self) -> None:
        payload = {
            "task": {"id": "task_1", "task_type": "macro_sequence", "task_label": "Word 宏顺序执行", "execute_mode": "local", "status": "待处理", "progress": 0},
            "handoff": {"requires_local_client": True, "status": "ready", "message": "本地客户端可接收任务参数"},
            "workflow_plan": {
                "schema_version": "k12.workflowPlan.v1",
                "current_task": "macro_sequence",
                "current_index": 2,
                "order": ["pdf_to_word", "macro_sequence"],
                "labels": ["PDF 转 Word", "Word 宏顺序执行"],
            },
            "files": [
                {
                    "id": "file_1",
                    "file_name": "private.docm",
                    "file_type": "Word",
                    "input_path": "/Users/a1/private/private.docm",
                    "input_path_exists": True,
                    "has_macro": True,
                    "validation_errors": [],
                }
            ],
            "local_actions": [
                {
                    "type": "macro_sequence",
                    "status": "queued",
                    "macros": [{"macro_name": "CleanEmptyParagraphs", "backup_path": "/tmp/private-backup.docm"}],
                }
            ],
            "desktop_execution_plan": {
                "schema_version": "k12.desktopExecutionPlan.v1",
                "status": "waiting_for_heartbeat",
                "native_execution_allowed": False,
                "web_executes_native_documents": False,
                "current_companion_cli_executes_native_documents": False,
                "native_action_count": 1,
                "actions": [
                    {
                        "type": "macro_sequence",
                        "label": "Word 宏顺序执行",
                        "gate_status": "waiting_for_heartbeat",
                        "input_paths": ["/Users/a1/private/private.docm"],
                        "backup_paths": ["/tmp/private-backup.docm"],
                        "required_capabilities": [{"key": "macroExecution", "label": "Word 宏执行"}],
                        "steps": [{"order": 1}, {"order": 2}],
                    }
                ],
            },
            "sync": {"task_status_cloud_sync_allowed": True, "result_upload_allowed": False, "policy": "仅本地保存结果"},
        }

        summary = summarize_payload(payload)
        summary_json = json.dumps(summary, ensure_ascii=False)
        self.assertEqual(summary["schema_version"], "k12.localClientDryRunSummary.v1")
        self.assertEqual(summary["file_count"], 1)
        self.assertEqual(summary["workflow_plan"]["current_index"], 2)
        self.assertEqual(summary["workflow_plan"]["order"], ["pdf_to_word", "macro_sequence"])
        self.assertEqual(summary["actions"][0]["macro_count"], 1)
        self.assertEqual(summary["desktop_execution_plan"]["schema_version"], "k12.desktopExecutionPlan.v1")
        self.assertEqual(summary["desktop_execution_plan"]["actions"][0]["step_count"], 2)
        self.assertEqual(summary["desktop_execution_plan"]["actions"][0]["required_capabilities"], ["macroExecution"])
        self.assertFalse(summary["desktop_execution_plan"]["web_executes_native_documents"])
        self.assertNotIn("/Users/a1/private", summary_json)
        self.assertNotIn("private-backup", summary_json)

        execution = build_dry_run_execution_summary(payload)
        execution_json = json.dumps(execution, ensure_ascii=False)
        self.assertEqual(execution["schema_version"], "k12.localDryRunExecution.v1")
        self.assertEqual(execution["action_count"], 1)
        self.assertEqual(execution["ready_action_count"], 0)
        self.assertEqual(execution["waiting_action_count"], 1)
        self.assertEqual(execution["actions"][0]["dry_run_status"], "waiting_for_heartbeat")
        self.assertEqual(execution["actions"][0]["blockers"], ["local_client_heartbeat_required"])
        self.assertFalse(execution["actions"][0]["native_execution_performed"])
        self.assertNotIn("/Users/a1/private", execution_json)
        self.assertNotIn("private-backup", execution_json)

        heartbeat = build_heartbeat(client_id="cli-1", platform_name="Windows", capabilities={"macroExecution": False, "token": "secret"})
        self.assertEqual(heartbeat["schema_version"], "k12.localClientHeartbeat.v1")
        self.assertEqual(heartbeat["platform"], "Windows")
        self.assertNotIn("token", heartbeat["capabilities"])
        self.assertEqual(heartbeat["preflight"]["schema_version"], "k12.localClientPreflight.v1")
        self.assertFalse(heartbeat["preflight"]["executes_native_documents"])

        sync = build_dry_run_sync_payload(payload)
        self.assertEqual(sync["status"], "running")
        self.assertEqual(sync["progress"], 5)
        self.assertEqual(sync["outputs"], [])
        self.assertEqual(sync["dryRunExecution"]["schema_version"], "k12.localDryRunExecution.v1")
        self.assertEqual(sync["dryRunExecution"]["waiting_action_count"], 1)
        self.assertIn("dry-run", sync["message"])
        self.assertNotIn("/Users/a1/private", json.dumps(sync["dryRunExecution"], ensure_ascii=False))

        native_request = build_native_execution_request(payload, platform_name="Windows")
        native_json = json.dumps(native_request, ensure_ascii=False)
        self.assertEqual(native_request["schema_version"], "k12.localNativeExecutionRequest.v1")
        self.assertFalse(native_request["native_execution_requested"])
        self.assertFalse(native_request["companion_cli_executes_native_documents"])
        self.assertFalse(native_request["platform"]["mathtype_objects_cross_platform_compatible"])
        self.assertTrue(native_request["platform"]["same_platform_required_for_native_mathtype"])
        self.assertEqual(native_request["platform"]["fallback_formula_formats"], ["MathML", "LaTeX", "image"])
        self.assertEqual(native_request["actions"][0]["native_request_status"], "blocked_by_capability")
        self.assertIn("native_execution_not_requested", native_request["actions"][0]["blockers"])
        self.assertIn("missing_capability:macroExecution", native_request["actions"][0]["blockers"])
        self.assertTrue(native_request["actions"][0]["sensitive_path_present"])
        self.assertNotIn("/Users/a1/private", native_json)
        self.assertNotIn("private-backup", native_json)

    def test_local_client_contract_helpers_are_documented_without_divider_comments(self) -> None:
        root = Path(__file__).resolve().parent.parent
        local_client_path = Path(local_client_module.__file__)
        tree = ast.parse(local_client_path.read_text(encoding="utf-8"))
        docstrings = {
            node.name: ast.get_docstring(node)
            for node in ast.walk(tree)
            if isinstance(node, ast.FunctionDef)
        }
        for name in [
            "_component_available",
            "_component_candidates",
            "_dry_run_action_summary",
            "_native_action_request",
            "_native_action_blockers",
            "_native_runner_profile",
            "_action_requires_native_document_runner",
            "_file_action_blocked_status",
        ]:
            self.assertTrue(docstrings.get(name), name)

        divider_comment = re.compile(r"#\s*(?:-{3,}|={3,}|\u2014{2,})")
        scan_paths = list((root / "k12").glob("*.py")) + [
            root / "static" / "app.js",
            root / "static" / "styles.css",
            root / "README.md",
            root / "docs" / "architecture.md",
        ]
        for path in scan_paths:
            text = path.read_text(encoding="utf-8")
            self.assertIsNone(divider_comment.search(text), str(path))

    def test_mathpix_task_option_audit_is_documented(self) -> None:
        root = Path(__file__).resolve().parent.parent
        for path in [root / "README.md", root / "docs" / "architecture.md"]:
            text = path.read_text(encoding="utf-8")
            self.assertIn("task_option_audit", text)
            self.assertIn("k12.mathpixTaskOptionAudit.v1", text)
            self.assertIn("allowExternalMathpixUpload", text)
            self.assertIn("externalUploadAuthorized", text)
            self.assertIn("settings.allowExternalMathpixUpload", text)

    def test_local_installer_manifest_is_documented(self) -> None:
        root = Path(__file__).resolve().parent.parent
        for path in [root / "README.md", root / "docs" / "architecture.md"]:
            text = path.read_text(encoding="utf-8")
            self.assertIn("platform.installer", text)
            self.assertIn("k12.localInstallerManifest.v1", text)
            self.assertIn("/api/installers/{file_name}?platform=", text)
            self.assertIn("SHA256", text)
            self.assertIn("MathType 原生对象边界", text)
            self.assertIn("不暴露 `installers/` 本地目录", text)

    def test_local_companion_dry_run_blocks_cross_platform_formula_contract(self) -> None:
        payload = {
            "task": {"id": "task_formula", "task_type": "omml_to_mathtype", "task_label": "OMML 转 MathType"},
            "desktop_execution_plan": {
                "schema_version": "k12.desktopExecutionPlan.v1",
                "status": "ready_for_native_client",
                "native_execution_allowed": True,
                "web_executes_native_documents": False,
                "current_companion_cli_executes_native_documents": False,
                "native_action_count": 1,
                "platform": {"expected": "Windows", "actual": "Windows", "compatible": True},
                "formula_delivery": {
                    "schema_version": "k12.formulaDeliveryContract.v1",
                    "platform": "macOS",
                    "compatibility_mode": "platform-specific",
                    "platform_object_format": "macOS MathType 对象",
                    "platform_objects_cross_compatible": True,
                    "native_mathtype_object_allowed": True,
                    "native_object_requires_same_platform": False,
                    "fallback_formats": ["MathML", "LaTeX", "图片"],
                    "output_priority": ["macOS MathType 对象"],
                    "message": "macOS 与 Windows 的 MathType 对象不通用",
                },
                "actions": [
                    {
                        "action_id": "task_formula:omml_mathtype",
                        "type": "omml_mathtype",
                        "label": "OMML/MathType 处理",
                        "gate_status": "ready",
                        "required_capabilities": [{"key": "mathTypeAutomation", "label": "MathType 自动化", "available": True}],
                        "input_paths": ["C:/private/lesson.docx"],
                        "steps": [{"operation": "mathtype.convert", "required": True}],
                    }
                ],
            },
            "sync": {"task_status_cloud_sync_allowed": True, "result_upload_allowed": False},
        }

        summary = summarize_payload(payload)
        execution = build_dry_run_execution_summary(payload)
        sync = build_dry_run_sync_payload(payload)
        execution_json = json.dumps(execution, ensure_ascii=False)

        self.assertEqual(summary["desktop_execution_plan"]["formula_delivery"]["platform"], "macOS")
        self.assertFalse(summary["desktop_execution_plan"]["formula_delivery"]["platform_objects_cross_compatible"])
        self.assertTrue(summary["desktop_execution_plan"]["formula_delivery"]["native_object_requires_same_platform"])
        self.assertTrue(summary["desktop_execution_plan"]["actions"][0]["formula_delivery"]["required"])
        self.assertEqual(execution["ready_action_count"], 0)
        self.assertEqual(execution["blocked_action_count"], 1)
        self.assertEqual(execution["formula_delivery"]["platform"], "macOS")
        self.assertFalse(execution["formula_delivery"]["platform_objects_cross_compatible"])
        self.assertTrue(execution["formula_delivery"]["native_object_requires_same_platform"])
        self.assertEqual(execution["actions"][0]["dry_run_status"], "blocked_by_platform")
        self.assertTrue(execution["actions"][0]["formula_delivery"]["required"])
        self.assertIn("formula_delivery_platform_mismatch", execution["actions"][0]["blockers"])
        self.assertIn("formula_contract_expected_platform_mismatch", execution["actions"][0]["blockers"])
        self.assertIn("0/1 个动作可交接", sync["message"])
        self.assertNotIn("C:/private", execution_json)

    def test_local_native_execution_request_marks_ready_windows_actions_without_running_them(self) -> None:
        payload = {
            "task": {"id": "task_2", "task_type": "macro_sequence", "task_label": "Word 宏顺序执行"},
            "desktop_execution_plan": {
                "schema_version": "k12.desktopExecutionPlan.v1",
                "task_id": "task_2",
                "task_type": "macro_sequence",
                "task_label": "Word 宏顺序执行",
                "status": "ready_for_native_client",
                "native_execution_allowed": True,
                "platform": {"expected": "Windows", "actual": "Windows", "compatible": True},
                "formula_delivery": {
                    "schema_version": "k12.formulaDeliveryContract.v1",
                    "platform": "Windows",
                    "compatibility_mode": "platform-specific",
                    "platform_object_format": "Windows OLE / Equation Native",
                    "platform_objects_cross_compatible": False,
                    "native_mathtype_object_allowed": True,
                    "native_object_requires_same_platform": True,
                    "fallback_formats": ["MathML", "LaTeX", "图片"],
                    "output_priority": ["Windows OLE / Equation Native"],
                },
                "actions": [
                    {
                        "action_id": "task_2:macro_sequence",
                        "type": "macro_sequence",
                        "label": "Word 宏顺序执行",
                        "gate_status": "ready",
                        "input_paths": ["C:/private/macro.docm"],
                        "required_capabilities": [{"key": "macroExecution", "label": "Word 宏执行", "available": True}],
                        "output_contract": {"artifact_types": ["macro-report", "backup"], "output_directory": "C:/private/out"},
                        "steps": [
                            {"operation": "macro.backup"},
                            {"operation": "macro.run_ordered"},
                            {"operation": "macro.write_report"},
                        ],
                    }
                ],
            },
        }

        native_request = build_native_execution_request(payload, platform_name="Windows", allow_native_execution=True)
        native_json = json.dumps(native_request, ensure_ascii=False)

        self.assertTrue(native_request["native_execution_requested"])
        self.assertTrue(native_request["native_execution_allowed_by_payload"])
        self.assertFalse(native_request["companion_cli_executes_native_documents"])
        self.assertFalse(native_request["platform"]["mathtype_objects_cross_platform_compatible"])
        self.assertTrue(native_request["platform"]["same_platform_required_for_native_mathtype"])
        self.assertEqual(native_request["platform"]["runner_profile"]["support_level"], "windows_office_com_adapter")
        self.assertTrue(native_request["platform"]["runner_profile"]["native_document_runner_available"])
        self.assertIn("MathType 原生对象必须按同平台交接", native_request["platform"]["message"])
        self.assertEqual(native_request["formula_delivery"]["platform"], "Windows")
        self.assertFalse(native_request["formula_delivery"]["platform_objects_cross_compatible"])
        self.assertTrue(native_request["formula_delivery"]["native_object_requires_same_platform"])
        self.assertIn("图片", native_request["formula_delivery"]["fallback_formats"])
        self.assertEqual(native_request["ready_action_count"], 1)
        self.assertEqual(native_request["actions"][0]["native_request_status"], "ready_for_native_runner")
        self.assertFalse(native_request["actions"][0]["formula_delivery"]["required"])
        self.assertEqual(native_request["actions"][0]["operations"], ["macro.backup", "macro.run_ordered", "macro.write_report"])
        self.assertFalse(native_request["actions"][0]["native_execution_performed"])
        self.assertNotIn("C:/private", native_json)

    def test_local_native_execution_request_marks_macos_formula_actions_as_limited_handoff(self) -> None:
        payload = {
            "task": {"id": "task_macos", "task_type": "omml_to_mathtype", "task_label": "OMML 转 MathType"},
            "desktop_execution_plan": {
                "schema_version": "k12.desktopExecutionPlan.v1",
                "status": "ready_for_native_client",
                "native_execution_allowed": True,
                "platform": {"expected": "macOS", "actual": "macOS", "compatible": True},
                "formula_delivery": {
                    "schema_version": "k12.formulaDeliveryContract.v1",
                    "platform": "macOS",
                    "compatibility_mode": "platform-specific",
                    "platform_object_format": "macOS MathType 对象 / MathML / LaTeX / 图片兜底",
                    "platform_objects_cross_compatible": False,
                    "native_mathtype_object_allowed": True,
                    "native_object_requires_same_platform": True,
                    "fallback_formats": ["MathML", "LaTeX", "图片"],
                    "output_priority": ["macOS MathType 对象", "MathML", "LaTeX", "图片"],
                },
                "actions": [
                    {
                        "action_id": "task_macos:omml_mathtype",
                        "type": "omml_mathtype",
                        "label": "OMML/MathType 处理",
                        "gate_status": "ready",
                        "required_capabilities": [{"key": "mathTypeAutomation", "label": "MathType 自动化", "available": True}],
                        "steps": [{"operation": "mathtype.convert"}],
                    }
                ],
            },
        }

        native_request = build_native_execution_request(payload, platform_name="macOS", allow_native_execution=True)

        self.assertTrue(native_request["platform"]["macos_native_runner"])
        self.assertFalse(native_request["platform"]["windows_native_runner"])
        self.assertEqual(native_request["platform"]["runner_profile"]["support_level"], "macos_limited_handoff")
        self.assertFalse(native_request["platform"]["runner_profile"]["native_document_runner_available"])
        self.assertIn("formula.export_fallbacks", native_request["platform"]["runner_profile"]["supported_operations"])
        self.assertEqual(native_request["formula_delivery"]["platform"], "macOS")
        self.assertFalse(native_request["formula_delivery"]["platform_objects_cross_compatible"])
        self.assertTrue(native_request["formula_delivery"]["native_object_requires_same_platform"])
        self.assertEqual(native_request["ready_action_count"], 0)
        self.assertEqual(native_request["blocked_action_count"], 1)
        action = native_request["actions"][0]
        self.assertEqual(action["native_request_status"], "blocked_by_platform")
        self.assertTrue(action["formula_delivery"]["required"])
        self.assertIn("macos_native_runner_limited", action["blockers"])
        self.assertNotIn("formula_delivery_platform_mismatch", action["blockers"])
        self.assertFalse(action["native_execution_performed"])

    def test_local_native_execution_request_blocks_mismatched_formula_delivery_platform(self) -> None:
        payload = {
            "task": {"id": "task_3", "task_type": "omml_to_mathtype", "task_label": "OMML 转 MathType"},
            "desktop_execution_plan": {
                "schema_version": "k12.desktopExecutionPlan.v1",
                "status": "ready_for_native_client",
                "native_execution_allowed": True,
                "platform": {"expected": "Windows", "actual": "Windows", "compatible": True},
                "formula_delivery": {
                    "schema_version": "k12.formulaDeliveryContract.v1",
                    "platform": "macOS",
                    "compatibility_mode": "platform-specific",
                    "platform_object_format": "macOS MathType 对象",
                    "platform_objects_cross_compatible": True,
                    "native_mathtype_object_allowed": True,
                    "native_object_requires_same_platform": False,
                    "fallback_formats": ["MathML", "LaTeX", "图片"],
                    "output_priority": ["macOS MathType 对象"],
                },
                "actions": [
                    {
                        "action_id": "task_3:omml_mathtype",
                        "type": "omml_mathtype",
                        "label": "OMML/MathType 处理",
                        "gate_status": "ready",
                        "required_capabilities": [{"key": "mathTypeAutomation", "label": "MathType 自动化", "available": True}],
                        "steps": [{"operation": "mathtype.convert"}],
                    }
                ],
            },
        }

        native_request = build_native_execution_request(payload, platform_name="Windows", allow_native_execution=True)

        self.assertEqual(native_request["formula_delivery"]["platform"], "macOS")
        self.assertFalse(native_request["formula_delivery"]["platform_objects_cross_compatible"])
        self.assertTrue(native_request["formula_delivery"]["native_object_requires_same_platform"])
        self.assertEqual(native_request["ready_action_count"], 0)
        self.assertEqual(native_request["blocked_action_count"], 1)
        action = native_request["actions"][0]
        self.assertEqual(action["native_request_status"], "blocked_by_platform")
        self.assertTrue(action["formula_delivery"]["required"])
        self.assertIn("formula_delivery_platform_mismatch", action["blockers"])
        self.assertIn("formula_contract_expected_platform_mismatch", action["blockers"])
        self.assertFalse(action["native_execution_performed"])

    def test_local_native_execution_request_blocks_formula_actions_without_delivery_contract(self) -> None:
        payload = {
            "task": {"id": "task_4", "task_type": "omml_to_mathtype", "task_label": "OMML 转 MathType"},
            "desktop_execution_plan": {
                "schema_version": "k12.desktopExecutionPlan.v1",
                "status": "ready_for_native_client",
                "native_execution_allowed": True,
                "platform": {"expected": "Windows", "actual": "Windows", "compatible": True},
                "actions": [
                    {
                        "action_id": "task_4:omml_mathtype",
                        "type": "omml_mathtype",
                        "label": "OMML/MathType 处理",
                        "gate_status": "ready",
                        "required_capabilities": [{"key": "mathTypeAutomation", "label": "MathType 自动化", "available": True}],
                        "steps": [{"operation": "mathtype.convert"}],
                    }
                ],
            },
        }

        native_request = build_native_execution_request(payload, platform_name="Windows", allow_native_execution=True)

        self.assertFalse(native_request["formula_delivery"]["contract_present"])
        self.assertFalse(native_request["formula_delivery"]["native_mathtype_object_allowed"])
        self.assertEqual(native_request["ready_action_count"], 0)
        self.assertEqual(native_request["blocked_action_count"], 1)
        action = native_request["actions"][0]
        self.assertEqual(action["native_request_status"], "blocked_by_platform")
        self.assertIn("formula_delivery_contract_missing", action["blockers"])
        self.assertTrue(action["formula_delivery"]["required"])
        self.assertFalse(action["native_execution_performed"])

    def test_local_file_action_execution_copies_omml_dependency_only_when_explicitly_allowed(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            document = root / "lesson.docx"
            source = root / "deps" / "OMML2MML.XSL"
            target = root / "OMML2MML.XSL"
            document.write_bytes(make_docx_bytes())
            source.parent.mkdir()
            source.write_text("<xsl:stylesheet>k12</xsl:stylesheet>", encoding="utf-8")
            payload = {
                "local_actions": [
                    {
                        "type": "omml_mathtype",
                        "copy_strategy": "自动重命名",
                        "dependencies": [
                            {
                                "id": "dep_1",
                                "file_id": "file_1",
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
            self.assertEqual(blocked["schema_version"], "k12.localFileActionExecution.v1")
            self.assertEqual(blocked["blocked_count"], 1)
            self.assertEqual(blocked["actions"][0]["status"], "blocked_until_explicit_file_action_request")
            self.assertFalse(target.exists())

            executed = execute_local_file_actions(payload, allow_file_actions=True)
            executed_json = json.dumps(executed, ensure_ascii=False)

            self.assertEqual(executed["performed_count"], 1)
            self.assertEqual(executed["actions"][0]["status"], "copied")
            self.assertTrue(executed["actions"][0]["copy_performed"])
            self.assertTrue(executed["actions"][0]["dependency_extension_allowed"])
            self.assertTrue(executed["actions"][0]["target_directory_allowed"])
            self.assertTrue(executed["actions"][0]["target_directory_writeable"])
            self.assertIn("仅允许 .xsl、.xslt、.xml、.mml 依赖文件", executed["guardrails"])
            self.assertIn("目标文档目录必须具备写入权限", executed["guardrails"])
            self.assertEqual(executed["actions"][0]["source_sha256"], hashlib.sha256(source.read_bytes()).hexdigest())
            self.assertEqual(executed["actions"][0]["target_sha256"], hashlib.sha256(target.read_bytes()).hexdigest())
            self.assertEqual(target.read_text(encoding="utf-8"), source.read_text(encoding="utf-8"))
            self.assertNotIn(str(root), executed_json)

    def test_local_file_action_execution_rejects_omml_copy_outside_document_directory(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            document = root / "lesson.docx"
            source = root / "OMML2MML.XSL"
            outside = root / "outside" / "OMML2MML.XSL"
            document.write_bytes(make_docx_bytes())
            source.write_text("<xsl:stylesheet />", encoding="utf-8")
            payload = {
                "local_actions": [
                    {
                        "type": "omml_mathtype",
                        "dependencies": [
                            {
                                "id": "dep_2",
                                "file_id": "file_2",
                                "document_path": str(document),
                                "omml_source_path": str(source),
                                "omml_target_path": str(outside),
                            }
                        ],
                    }
                ]
            }

            result = execute_local_file_actions(payload, allow_file_actions=True)

            self.assertEqual(result["actions"][0]["status"], "blocked_target_directory")
            self.assertIn("target_not_document_directory", result["actions"][0]["blockers"])
            self.assertFalse(outside.exists())

    def test_local_file_action_execution_rejects_omml_copy_without_write_permission(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            document = root / "lesson.docx"
            source = root / "OMML2MML.XSL"
            target = root / "OMML2MML-copy.XSL"
            document.write_bytes(make_docx_bytes())
            source.write_text("<xsl:stylesheet />", encoding="utf-8")
            payload = {
                "local_actions": [
                    {
                        "type": "omml_mathtype",
                        "dependencies": [
                            {
                                "id": "dep_3",
                                "file_id": "file_3",
                                "document_path": str(document),
                                "omml_source_path": str(source),
                                "omml_target_path": str(target),
                            }
                        ],
                    }
                ]
            }

            with patch("k12.local_client.os.access", return_value=False):
                result = execute_local_file_actions(payload, allow_file_actions=True)

            self.assertEqual(result["actions"][0]["status"], "blocked_target_directory_not_writeable")
            self.assertIn("target_directory_not_writeable", result["actions"][0]["blockers"])
            self.assertFalse(result["actions"][0]["target_directory_writeable"])
            self.assertFalse(target.exists())

    def test_local_file_action_execution_rejects_non_omml_dependency_extension(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            document = root / "lesson.docx"
            source = root / "notes.txt"
            target = root / "notes.txt"
            document.write_bytes(make_docx_bytes())
            source.write_text("not an OMML dependency", encoding="utf-8")
            payload = {
                "local_actions": [
                    {
                        "type": "omml_mathtype",
                        "dependencies": [
                            {
                                "id": "dep_4",
                                "file_id": "file_4",
                                "document_path": str(document),
                                "omml_source_path": str(source),
                                "omml_target_path": str(target),
                            }
                        ],
                    }
                ]
            }

            result = execute_local_file_actions(payload, allow_file_actions=True)

            self.assertEqual(result["actions"][0]["status"], "blocked_unsupported_dependency_extension")
            self.assertIn("unsupported_dependency_extension", result["actions"][0]["blockers"])
            self.assertFalse(result["actions"][0]["dependency_extension_allowed"])
            self.assertFalse(result["actions"][0]["copy_performed"])

    def test_local_companion_cli_component_preflight_is_path_safe(self) -> None:
        preflight = build_component_preflight("Windows")
        sanitized = sanitize_component_preflight(
            {
                **preflight,
                "components": {
                    **preflight["components"],
                    "word": {"label": "Word", "available": True, "status": "available", "path": r"C:\Program Files\Word\WINWORD.EXE"},
                    "token": {"label": "secret", "available": True, "status": "available"},
                },
                "capabilities": {**preflight["capabilities"], "token": "secret"},
            }
        )
        heartbeat = build_heartbeat(client_id="cli-preflight", platform_name="Windows", preflight=sanitized)
        heartbeat_json = json.dumps(heartbeat, ensure_ascii=False)

        self.assertEqual(preflight["schema_version"], "k12.localClientPreflight.v1")
        self.assertIn("word", sanitized["components"])
        self.assertNotIn("token", sanitized["components"])
        self.assertNotIn("token", sanitized["capabilities"])
        self.assertFalse(sanitized["executes_native_documents"])
        self.assertNotIn("Program Files", heartbeat_json)

    def test_local_companion_cli_requires_configured_security_token(self) -> None:
        manifest = {
            "localClientManifest": {
                "security": {"heartbeat_requires_configured_token": True, "token_configured": False},
                "queue": {},
            }
        }
        with patch("k12.local_client.request_json", return_value=manifest):
            with self.assertRaises(LocalClientError) as raised:
                run_once("http://127.0.0.1:8765")
        self.assertIn("安全令牌", str(raised.exception))

    def test_local_companion_cli_run_once_can_include_native_execution_request(self) -> None:
        responses = [
            {
                "localClientManifest": {
                    "security": {"heartbeat_requires_configured_token": True, "token_configured": True},
                    "queue": {"next_task_id": "task_2", "pending_local_task_count": 1},
                    "platform": {"selected": "Windows"},
                }
            },
            {"heartbeat": {"schema_version": "k12.localClientHeartbeat.v1", "client_id": "cli-1"}},
            {
                "localPayload": {
                    "task": {"id": "task_2", "task_type": "macro_sequence", "task_label": "Word 宏顺序执行"},
                    "desktop_execution_plan": {
                        "schema_version": "k12.desktopExecutionPlan.v1",
                        "status": "ready_for_native_client",
                        "native_execution_allowed": True,
                        "platform": {"expected": "Windows", "actual": "Windows", "compatible": True},
                        "actions": [
                            {
                                "action_id": "task_2:macro_sequence",
                                "type": "macro_sequence",
                                "label": "Word 宏顺序执行",
                                "gate_status": "ready",
                                "required_capabilities": [{"key": "macroExecution", "label": "Word 宏执行", "available": True}],
                                "steps": [{"operation": "macro.run_ordered"}],
                            }
                        ],
                    },
                }
            },
        ]

        with patch("k12.local_client.request_json", side_effect=responses):
            result = run_once("http://127.0.0.1:8765", token="secret-token", native_plan=True, allow_native_execution=True)

        native_request = result["native_execution_request"]
        self.assertEqual(native_request["schema_version"], "k12.localNativeExecutionRequest.v1")
        self.assertIsNotNone(native_request["actions"][0]["native_request_status"])
        self.assertFalse(native_request["companion_cli_executes_native_documents"])
        self.assertFalse(native_request["actions"][0]["native_execution_performed"])
        self.assertIsNone(result["sync"])

    def test_local_task_status_sync_requires_authorization(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            store.update_settings({"localSecurityToken": "secret-token"})
            processor = TaskProcessor(store)
            word = processor.create_uploaded_file("lesson.docx", make_docx_bytes())[0]
            task = processor.create_task({"task_type": "word_to_ppt", "file_ids": [word["id"]]})

            with self.assertRaises(ValueError) as raised:
                processor.sync_local_task_status(task["id"], {"status": "running", "progress": 20})

            self.assertIn("未授权", str(raised.exception))

    def test_public_api_payload_redacts_local_paths_by_default(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            store.update_settings({"autoOpenOutputDirectory": True})
            processor = TaskProcessor(store)
            word = processor.create_uploaded_file("macro.docm", make_docx_bytes())[0]
            processor.create_task(
                {
                    "task_type": "macro_sequence",
                    "file_ids": [word["id"]],
                    "options": {
                        "selectedMacros": [{"id": "macro_clean_empty_paragraphs", "execute_order": 1}],
                        "confirmMacroRisk": True,
                        "macroBackup": True,
                    },
                }
            )
            handler = make_handler(store, processor)
            public_file = K12RequestHandler._public_payload(handler, word)
            public_task = K12RequestHandler._public_payload(handler, store.list_tasks()[0])
            public_report = K12RequestHandler._public_payload(handler, store.list_reports()[0])
            public_log = K12RequestHandler._public_log(handler, store.list_logs(limit=1000)[0])
            serialized = json.dumps([public_file, public_task, public_report, public_log], ensure_ascii=False)

            self.assertNotIn(tmp, serialized)
            self.assertIn("本地路径已隐藏", serialized)
            self.assertTrue(public_file["storage_path_available"])
            self.assertTrue(public_task["output_path_available"])
            self.assertTrue(public_report["analysis"]["macros"][0]["backup_available"])

    def test_public_reference_get_routes_do_not_leak_local_token_or_paths(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            store = AppStore(root / "data")
            store.update_settings(
                {
                    "localSecurityToken": "secret-token",
                    "outputDirectory": str(root / "outputs"),
                    "manualOmmlPath": str(root / "OMML2MML.XSL"),
                    "pptToWordTemplatePath": str(root / "template.dotx"),
                    "exposeLocalPaths": False,
                }
            )
            processor = TaskProcessor(store)
            routes = [
                ("/api/settings", {}),
                ("/api/capabilities", {}),
                ("/api/architecture", {}),
                ("/api/api-catalog", {}),
                ("/api/data-dictionary", {}),
                ("/api/acceptance-matrix", {}),
                ("/api/product-summary", {}),
                ("/api/enhancement-plan", {}),
                ("/api/install-profile", {}),
                ("/api/install-plan", {"platform": ["Windows"]}),
                ("/api/local-client/manifest", {}),
                ("/api/mathpix-jobs", {}),
                ("/api/authorizations", {}),
                ("/api/preflight", {}),
            ]

            for path, query in routes:
                with self.subTest(path=path):
                    handler = make_handler(store, processor)
                    K12RequestHandler._handle_api_get(handler, path, query)
                    payload = handler.wfile.getvalue().decode("utf-8")

                    self.assertEqual(handler.status, 200)
                    self.assertNotIn("secret-token", payload)
                    self.assertNotIn(str(root), payload)

    def test_public_api_payload_can_expose_local_paths_when_enabled(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            store.update_settings({"exposeLocalPaths": True})
            processor = TaskProcessor(store)
            word = processor.create_uploaded_file("lesson.docx", make_docx_bytes())[0]
            handler = make_handler(store, processor)
            public_file = K12RequestHandler._public_payload(handler, word)

            self.assertIn(tmp, public_file["storage_path"])

    def test_install_profile_can_be_previewed_without_saving_platform(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            preview = processor.install_profile("macOS")
            settings = store.get_settings()
            self.assertEqual(preview["platform"], "macOS")
            self.assertEqual(settings["localClientPlatform"], "auto")
            self.assertIn("MathML", preview["formulaPortability"])

    def test_architecture_blueprint_api_exposes_prd16_contract(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            handler = make_handler(store, processor)
            K12RequestHandler._handle_api_get(handler, "/api/architecture", {})
            response = json.loads(handler.wfile.getvalue().decode("utf-8"))
            blueprint = response["architecture"]

            self.assertEqual(handler.status, 200)
            self.assertEqual(blueprint["schema_version"], "k12.architectureBlueprint.v1")
            self.assertEqual(blueprint["runtime"]["service"], "Python 标准库本地 HTTP API")
            self.assertTrue(any(layer["key"] == "local_service" for layer in blueprint["layers"]))
            self.assertTrue(any(layer["key"] == "cloud_service" for layer in blueprint["layers"]))
            self.assertTrue(any(contract["endpoint"] == "/api/tasks/{task_id}/local-payload" for contract in blueprint["contracts"]))
            self.assertTrue(any(contract["name"] == "Mathpix PDF" for contract in blueprint["contracts"]))
            self.assertTrue(any(item["recommendation"] == "FastAPI + SQLite" for item in blueprint["recommended_split"]))
            self.assertTrue(any(item["recommendation"] == "Django" for item in blueprint["recommended_split"]))

    def test_api_catalog_exposes_open_interface_boundaries(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            store.update_settings({"localSecurityToken": "secret-token"})
            processor = TaskProcessor(store)
            handler = make_handler(store, processor)
            K12RequestHandler._handle_api_get(handler, "/api/api-catalog", {})
            response = json.loads(handler.wfile.getvalue().decode("utf-8"))
            catalog = response["apiCatalog"]

            self.assertEqual(handler.status, 200)
            self.assertEqual(catalog["schema_version"], "k12.apiCatalog.v1")
            self.assertTrue(catalog["auth"]["token_required"])
            self.assertTrue(catalog["auth"]["local_payload_requires_configured_token"])
            endpoints = {(item["method"], item["path"]): item for item in catalog["endpoints"]}
            self.assertIn(("GET", "/api/tasks/{task_id}/local-payload"), endpoints)
            self.assertEqual(endpoints[("GET", "/api/tasks/{task_id}/local-payload")]["auth"], "configured-token")
            self.assertTrue(endpoints[("GET", "/api/tasks/{task_id}/local-payload")]["sensitive"])
            self.assertIn(("GET", "/api/tasks/{task_id}/local-readiness"), endpoints)
            self.assertEqual(endpoints[("GET", "/api/tasks/{task_id}/local-readiness")]["auth"], "token")
            self.assertFalse(endpoints[("GET", "/api/tasks/{task_id}/local-readiness")]["sensitive"])
            self.assertIn(("GET", "/api/api-catalog"), endpoints)
            self.assertIn(("GET", "/api/data-dictionary"), endpoints)
            self.assertIn(("GET", "/api/enhancement-plan"), endpoints)
            self.assertIn(("GET", "/api/files/{file_id}/reports"), endpoints)
            self.assertIn(("GET", "/api/local-client/manifest"), endpoints)
            self.assertIn(("GET", "/api/local-client/uploads"), endpoints)
            self.assertTrue(endpoints[("GET", "/api/local-client/uploads")]["sensitive"])
            self.assertIn(("GET", "/api/local-client/uploads/{upload_id}/manifest"), endpoints)
            self.assertTrue(endpoints[("GET", "/api/local-client/uploads/{upload_id}/manifest")]["sensitive"])
            self.assertIn(("GET", "/api/tasks/recovery-summary"), endpoints)
            self.assertFalse(endpoints[("GET", "/api/tasks/recovery-summary")]["sensitive"])
            review_endpoints = {
                ("GET", "/api/image-annotations"),
                ("POST", "/api/image-annotations"),
                ("DELETE", "/api/image-annotations/{annotation_id}"),
                ("GET", "/api/formula-annotations"),
                ("POST", "/api/formula-annotations"),
                ("POST", "/api/formula-annotations/bulk-confirm"),
                ("DELETE", "/api/formula-annotations/{annotation_id}"),
                ("GET", "/api/omml-annotations"),
                ("POST", "/api/omml-annotations"),
                ("DELETE", "/api/omml-annotations/{annotation_id}"),
                ("GET", "/api/layout-annotations"),
                ("POST", "/api/layout-annotations"),
                ("DELETE", "/api/layout-annotations/{annotation_id}"),
            }
            for endpoint in review_endpoints:
                self.assertIn(endpoint, endpoints)
                self.assertTrue(endpoints[endpoint]["sensitive"])
            self.assertIn("置信度阈值", endpoints[("POST", "/api/formula-annotations/bulk-confirm")]["description"])
            self.assertIn("手动依赖", endpoints[("GET", "/api/omml-annotations")]["description"])
            self.assertIn(("POST", "/api/image-annotations/replacement"), endpoints)
            self.assertTrue(endpoints[("POST", "/api/image-annotations/replacement")]["sensitive"])
            self.assertIn(("POST", "/api/image-assets/reexport"), endpoints)
            self.assertTrue(endpoints[("POST", "/api/image-assets/reexport")]["sensitive"])
            self.assertIn(("GET", "/api/assets/replacements/{file_name}"), endpoints)
            self.assertTrue(endpoints[("GET", "/api/assets/replacements/{file_name}")]["sensitive"])
            self.assertIn(("GET", "/api/assets/images/{asset_path}"), endpoints)
            self.assertTrue(endpoints[("GET", "/api/assets/images/{asset_path}")]["sensitive"])
            self.assertIn("受控图片缓存目录", endpoints[("GET", "/api/assets/images/{asset_path}")]["description"])
            self.assertIn(("GET", "/api/artifacts/{task_id}/{file_name}"), endpoints)
            self.assertTrue(endpoints[("GET", "/api/artifacts/{task_id}/{file_name}")]["sensitive"])
            self.assertIn("报告登记", endpoints[("GET", "/api/artifacts/{task_id}/{file_name}")]["description"])
            self.assertEqual(endpoints[("POST", "/api/uploads")]["permission"], "files.manage")
            self.assertEqual(endpoints[("POST", "/api/files")]["permission"], "files.manage")
            self.assertEqual(endpoints[("POST", "/api/files/{file_id}/replace")]["permission"], "files.manage")
            self.assertEqual(endpoints[("POST", "/api/files/{file_id}/password")]["permission"], "files.manage")
            self.assertEqual(endpoints[("DELETE", "/api/files/{file_id}")]["permission"], "files.manage")
            self.assertEqual(endpoints[("POST", "/api/tasks")]["permission"], "tasks.create")
            self.assertEqual(endpoints[("POST", "/api/tasks/{task_id}/retry")]["permission"], "tasks.control")
            self.assertEqual(endpoints[("POST", "/api/tasks/{task_id}/pause")]["permission"], "tasks.control")
            self.assertEqual(endpoints[("POST", "/api/tasks/{task_id}/resume")]["permission"], "tasks.control")
            self.assertEqual(endpoints[("POST", "/api/tasks/{task_id}/cancel")]["permission"], "tasks.control")
            self.assertEqual(endpoints[("POST", "/api/tasks/{task_id}/skip-file")]["permission"], "tasks.control")
            self.assertEqual(endpoints[("POST", "/api/tasks/{task_id}/restore-backups")]["permission"], "tasks.control")
            self.assertEqual(endpoints[("PUT", "/api/settings")]["permission"], "settings.manage")
            self.assertEqual(endpoints[("POST", "/api/users")]["permission"], "users.manage")
            self.assertEqual(endpoints[("DELETE", "/api/users/{user_id}")]["permission"], "users.manage")
            self.assertEqual(endpoints[("POST", "/api/templates")]["permission"], "templates.manage")
            self.assertEqual(endpoints[("DELETE", "/api/templates/{template_id}")]["permission"], "templates.manage")
            self.assertEqual(endpoints[("POST", "/api/macro-templates")]["permission"], "templates.manage")
            self.assertEqual(endpoints[("DELETE", "/api/macro-templates/{template_id}")]["permission"], "templates.manage")
            self.assertEqual(endpoints[("POST", "/api/authorizations")]["permission"], "authorizations.manage")
            self.assertEqual(endpoints[("DELETE", "/api/reports/{report_id}")]["permission"], "reports.manage")
            report_export_endpoints = {
                ("GET", "/api/reports/{report_id}/images.zip"),
                ("GET", "/api/reports/{report_id}/images.xlsx"),
                ("GET", "/api/reports/{report_id}/formulas.zip"),
                ("GET", "/api/reports/{report_id}/formulas.xlsx"),
                ("GET", "/api/reports/{report_id}/omml-failures.csv"),
                ("GET", "/api/reports/{report_id}/macro-failures.csv"),
            }
            for endpoint in report_export_endpoints:
                self.assertIn(endpoint, endpoints)
                self.assertTrue(endpoints[endpoint]["sensitive"])
            self.assertIn("manifest.csv", endpoints[("GET", "/api/reports/{report_id}/images.zip")]["description"])
            self.assertIn("TEX", endpoints[("GET", "/api/reports/{report_id}/formulas.zip")]["description"])
            self.assertIn("手动依赖", endpoints[("GET", "/api/reports/{report_id}/omml-failures.csv")]["description"])
            self.assertIn("备份恢复", endpoints[("GET", "/api/reports/{report_id}/macro-failures.csv")]["description"])
            self.assertIn(("POST", "/api/local-client/uploads"), endpoints)
            self.assertEqual(endpoints[("POST", "/api/local-client/uploads")]["auth"], "configured-token + cloud-sync-authorization")
            self.assertTrue(endpoints[("POST", "/api/local-client/uploads")]["sensitive"])
            self.assertIn(("POST", "/api/local-client/heartbeat"), endpoints)
            self.assertEqual(endpoints[("POST", "/api/local-client/heartbeat")]["auth"], "configured-token")
            self.assertIn(("POST", "/api/tasks/{task_id}/local-launch"), endpoints)
            self.assertEqual(endpoints[("POST", "/api/tasks/{task_id}/local-launch")]["auth"], "configured-token + launch-authorization")
            self.assertIn(("GET", "/api/mathpix-jobs"), endpoints)
            self.assertIn("不触发新的外部上传", endpoints[("GET", "/api/mathpix-jobs")]["description"])
            self.assertIn(("GET", "/api/acceptance-matrix"), endpoints)
            self.assertIn("验证上下文", endpoints[("GET", "/api/acceptance-matrix")]["description"])
            self.assertIn("未覆盖风险", endpoints[("GET", "/api/acceptance-matrix")]["description"])
            self.assertIn("pdfToWordEngine 固定归一为 Mathpix", endpoints[("GET", "/api/settings")]["description"])
            self.assertIn("PDF 转 Word 引擎固定 Mathpix", endpoints[("PUT", "/api/settings")]["description"])
            self.assertIn("settings.allowExternalMathpixUpload", endpoints[("PUT", "/api/settings")]["description"])
            self.assertIn(("GET", "/api/install-profile"), endpoints)
            self.assertIn(("GET", "/api/install-plan"), endpoints)
            self.assertIn("心跳平台匹配", endpoints[("GET", "/api/install-plan")]["description"])
            self.assertIn("MathType 同平台", endpoints[("GET", "/api/install-plan")]["description"])
            self.assertIn("platform_mismatch", endpoints[("GET", "/api/tasks/{task_id}/local-readiness")]["description"])
            self.assertIn("同平台 MathType/Office", endpoints[("GET", "/api/tasks/{task_id}/local-readiness")]["description"])
            self.assertIn(("GET", "/api/installers/{file_name}"), endpoints)
            installer_endpoint = endpoints[("GET", "/api/installers/{file_name}")]
            self.assertIn("注册的 Windows .msi 或 macOS .pkg", installer_endpoint["description"])
            self.assertIn("公式对象同平台边界", installer_endpoint["description"])
            response_headers = {item["name"]: item for item in installer_endpoint["response_headers"]}
            self.assertEqual(
                set(response_headers),
                {
                    "X-K12-Platform",
                    "X-K12-Installer-Kind",
                    "X-K12-Package-Boundary",
                    "X-K12-Formula-Object-Boundary",
                    "X-K12-Formula-Fallback-Formats",
                    "X-K12-SHA256",
                },
            )
            self.assertIn("MathType 原生对象必须同平台交接", response_headers["X-K12-Formula-Object-Boundary"]["description"])
            self.assertIn("MathML,LaTeX,image", response_headers["X-K12-Formula-Fallback-Formats"]["description"])

    def test_data_dictionary_exposes_prd10_models_and_counts(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            now = "2026-06-26T00:00:00Z"
            store.save_file(
                {
                    "id": "file_1",
                    "file_name": "讲义.docx",
                    "file_type": "Word",
                    "file_size": 128,
                    "file_path": str(Path(tmp) / "讲义.docx"),
                    "storage_path": str(Path(tmp) / "uploads" / "讲义.docx"),
                    "status": "待处理",
                    "created_at": now,
                }
            )
            store.save_task(
                {
                    "id": "task_1",
                    "task_type": "formula_precheck",
                    "execute_mode": "hybrid",
                    "file_ids": ["file_1"],
                    "status": "完成",
                    "progress": 100,
                    "input_path": str(Path(tmp) / "讲义.docx"),
                    "output_path": str(Path(tmp) / "outputs" / "task_1"),
                    "start_time": now,
                    "end_time": now,
                    "created_at": now,
                }
            )
            store.save_report(
                {
                    "id": "report_1",
                    "task_id": "task_1",
                    "file_id": "file_1",
                    "report_type": "公式预检",
                    "success_count": 1,
                    "fail_count": 0,
                    "formula_count": 3,
                    "omml_count": 2,
                    "omml_converted_count": 1,
                    "macro_count": 4,
                    "small_image_count": 5,
                    "report_path": str(Path(tmp) / "reports" / "report_1.html"),
                    "created_at": now,
                }
            )
            processor = TaskProcessor(store)
            handler = make_handler(store, processor)
            K12RequestHandler._handle_api_get(handler, "/api/data-dictionary", {})
            response = json.loads(handler.wfile.getvalue().decode("utf-8"))
            dictionary = response["dataDictionary"]

            self.assertEqual(handler.status, 200)
            self.assertEqual(dictionary["schema_version"], "k12.dataDictionary.v1")
            self.assertEqual(dictionary["source"], "PRD 第 10 章数据结构设计")
            self.assertEqual(dictionary["summary"]["objects"], 8)
            self.assertGreaterEqual(dictionary["summary"]["fields"], 90)
            self.assertIn("本地路径字段默认脱敏", dictionary["path_policy"])
            objects = {item["name"]: item for item in dictionary["objects"]}
            self.assertEqual(
                list(objects),
                ["FileItem", "Task", "FormulaItem", "FormulaAnnotation", "OmmlDependencyItem", "MacroItem", "SmallImageItem", "ReportItem"],
            )
            self.assertEqual(objects["FileItem"]["count"], 1)
            self.assertEqual(objects["Task"]["count"], 1)
            self.assertEqual(objects["FormulaItem"]["count"], 3)
            self.assertEqual(objects["OmmlDependencyItem"]["count"], 2)
            self.assertEqual(objects["MacroItem"]["count"], 4)
            self.assertEqual(objects["SmallImageItem"]["count"], 5)
            self.assertEqual(objects["ReportItem"]["count"], 1)
            self.assertIn("file_path", objects["FileItem"]["sensitive_fields"])
            self.assertIn("report_path", objects["ReportItem"]["sensitive_fields"])
            file_fields = {field["name"]: field for field in objects["FileItem"]["fields"]}
            self.assertEqual(file_fields["source_relative_path"]["type"], "string")
            self.assertEqual(file_fields["file_path"]["type"], "string")
            task_fields = {field["name"]: field for field in objects["Task"]["fields"]}
            self.assertEqual(task_fields["options"]["type"], "object")
            self.assertIn("workflowPlan", task_fields["options"]["description"])
            formula_fields = {field["name"]: field for field in objects["FormulaItem"]["fields"]}
            self.assertEqual(formula_fields["position_status"]["type"], "string")
            self.assertEqual(formula_fields["fallback_position"]["description"], "位置缺失时的异常位置记录")
            formula_annotation_fields = {field["name"]: field for field in objects["FormulaAnnotation"]["fields"]}
            self.assertEqual(formula_annotation_fields["recognition_request"]["type"], "object")
            self.assertIn("k12.formulaRecognitionRequest.v1", formula_annotation_fields["recognition_request"]["description"])
            image_fields = {field["name"]: field for field in objects["SmallImageItem"]["fields"]}
            self.assertEqual(image_fields["duplicate_check_status"]["type"], "string")
            self.assertEqual(image_fields["duplicate_fallback"]["description"], "重复判断异常兜底")
            self.assertEqual(image_fields["export_status"]["type"], "string")
            self.assertEqual(image_fields["reexported_at"]["type"], "datetime")
            report_fields = {field["name"]: field for field in objects["ReportItem"]["fields"]}
            self.assertIn("report_path", report_fields)

    def test_acceptance_matrix_exposes_prd17_item_evidence(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            handler = make_handler(store, processor)
            K12RequestHandler._handle_api_get(handler, "/api/acceptance-matrix", {})
            response = json.loads(handler.wfile.getvalue().decode("utf-8"))
            matrix = response["acceptanceMatrix"]

            self.assertEqual(handler.status, 200)
            self.assertEqual(matrix["schema_version"], "k12.acceptanceMatrix.v1")
            self.assertEqual(matrix["source"], "PRD 第 17 章验收标准")
            self.assertEqual(matrix["summary"]["groups"], 10)
            self.assertEqual(matrix["summary"]["items"], 96)
            self.assertEqual(matrix["summary"]["uncovered_risks"], 4)
            risks = {item["key"]: item for item in matrix["uncovered_risks"]}
            self.assertEqual(set(risks), {"17.6.2", "17.6.5", "17.7.4", "17.10.7"})
            self.assertIn("未覆盖风险", risks["17.6.2"]["uncovered_risk"])
            self.assertIn("mathpix_credentials_required", risks["17.6.5"]["blocking_reasons"])
            self.assertIn("native_desktop_runner_required", risks["17.10.7"]["blocking_reasons"])
            self.assertIn("同平台 MathType 环境", risks["17.10.7"]["required_environment"])
            groups = {group["section"]: group for group in matrix["groups"]}
            word_items = {item["key"]: item for item in groups["17.2"]["items"]}
            self.assertEqual(word_items["17.2.2"]["status"], "已覆盖")
            self.assertIn("内置转换自检通过", word_items["17.2.2"]["evidence"])
            self.assertIn("无历史任务", word_items["17.2.2"]["current"])
            self.assertEqual(word_items["17.2.4"]["status"], "已覆盖")
            self.assertIn("图片对象进入 PPTX 对象保留清单", word_items["17.2.4"]["evidence"])
            self.assertIn("真实排版和 OLE 写回仍需本地 Office/MathType", word_items["17.2.4"]["current"])
            self.assertEqual(word_items["17.2.5"]["status"], "已覆盖")
            self.assertIn("表格结构进入 PPTX 对象保留清单", word_items["17.2.5"]["evidence"])
            self.assertEqual(word_items["17.2.6"]["status"], "已覆盖")
            self.assertIn("OMML 与 MathType/嵌入对象进入 PPTX 对象保留清单", word_items["17.2.6"]["evidence"])
            self.assertEqual(word_items["17.2.8"]["status"], "已覆盖")
            self.assertIn("公式批量格式化自检通过", word_items["17.2.8"]["evidence"])
            self.assertIn("真实 MathType 对象写回仍需本地客户端", word_items["17.2.8"]["current"])
            ppt_items = {item["key"]: item for item in groups["17.5"]["items"]}
            self.assertEqual(ppt_items["17.5.2"]["status"], "已覆盖")
            self.assertIn("内置转换自检通过", ppt_items["17.5.2"]["evidence"])
            self.assertIn("无历史任务", ppt_items["17.5.2"]["current"])
            self.assertEqual(ppt_items["17.5.6"]["status"], "已覆盖")
            self.assertIn("PPT 公式自检通过", ppt_items["17.5.6"]["evidence"])
            self.assertIn("本地客户端", ppt_items["17.5.6"]["current"])
            self.assertEqual(groups["17.6"]["title"], "PDF 处理验收")
            pdf_items = {item["key"]: item for item in groups["17.6"]["items"]}
            self.assertEqual(pdf_items["17.6.1"]["status"], "合同覆盖")
            self.assertIn("Mathpix 合同自检通过", pdf_items["17.6.1"]["evidence"])
            self.assertIn("无历史任务", pdf_items["17.6.1"]["current"])
            self.assertIn("未发起外部上传", pdf_items["17.6.1"]["current"])
            self.assertEqual(pdf_items["17.6.2"]["status"], "需 Mathpix 实测")
            self.assertIn("扫描 PDF OCR 合同自检通过", pdf_items["17.6.2"]["evidence"])
            self.assertIn("未发起外部上传", pdf_items["17.6.2"]["current"])
            self.assertEqual(pdf_items["17.6.2"]["verification"]["scope"], "external_mathpix")
            self.assertIn("外部上传授权", pdf_items["17.6.2"]["verification"]["required_environment"])
            self.assertEqual(pdf_items["17.6.3"]["status"], "已覆盖")
            self.assertIn("PDF 图片保留自检通过", pdf_items["17.6.3"]["evidence"])
            self.assertIn("未发起外部 Mathpix 上传", pdf_items["17.6.3"]["current"])
            self.assertEqual(pdf_items["17.6.4"]["status"], "已覆盖")
            self.assertIn("PDF 表格保留自检通过", pdf_items["17.6.4"]["evidence"])
            self.assertEqual(pdf_items["17.6.5"]["status"], "需 Mathpix 实测")
            self.assertIn("PDF 公式 OCR 合同自检通过", pdf_items["17.6.5"]["evidence"])
            self.assertIn("tex.zip", pdf_items["17.6.5"]["current"])
            self.assertIn("external_ocr_result_not_verified", pdf_items["17.6.5"]["verification"]["blocking_reasons"])
            self.assertEqual(pdf_items["17.6.6"]["status"], "已覆盖")
            self.assertIn("PDF 公式 MathType 后处理自检通过", pdf_items["17.6.6"]["evidence"])
            self.assertIn("真实 MathType 对象写回仍需本地客户端", pdf_items["17.6.6"]["current"])
            omml_items = {item["key"]: item for item in groups["17.3"]["items"]}
            self.assertEqual(omml_items["17.3.3"]["status"], "合同覆盖")
            self.assertIn("OMML 转 MathType 交接自检通过", omml_items["17.3.3"]["evidence"])
            self.assertIn("真实 MathType 对象写回仍需本地客户端", omml_items["17.3.3"]["current"])
            self.assertEqual(omml_items["17.3.4"]["status"], "已覆盖")
            self.assertIn("OMML 依赖自检通过", omml_items["17.3.4"]["evidence"])
            self.assertIn("未触碰用户文件", omml_items["17.3.4"]["current"])
            self.assertEqual(omml_items["17.3.5"]["status"], "已覆盖")
            self.assertIn("复制到当前 Word 文档目录", omml_items["17.3.5"]["evidence"])
            self.assertIn("未触碰用户文件", omml_items["17.3.5"]["current"])
            self.assertEqual(omml_items["17.3.6"]["status"], "合同覆盖")
            self.assertIn("OMML 重新转换自检通过", omml_items["17.3.6"]["evidence"])
            self.assertIn("retry_queued", omml_items["17.3.6"]["current"])
            macro_items = {item["key"]: item for item in groups["17.4"]["items"]}
            self.assertEqual(macro_items["17.4.1"]["status"], "已覆盖")
            self.assertIn("DOCM 当前 Word 文档会标记 has_macro", macro_items["17.4.1"]["evidence"])
            self.assertIn("真实 VBA 模块枚举仍需本地 Word COM", macro_items["17.4.1"]["current"])
            self.assertEqual(macro_items["17.4.2"]["status"], "已覆盖")
            self.assertIn("DOTM Word 模板会标记 has_macro", macro_items["17.4.2"]["evidence"])
            self.assertEqual(macro_items["17.4.5"]["status"], "合同覆盖")
            self.assertIn("宏顺序执行交接自检通过", macro_items["17.4.5"]["evidence"])
            self.assertIn("dry-run 未执行真实 Word 宏", macro_items["17.4.5"]["current"])
            self.assertEqual(macro_items["17.4.10"]["status"], "已覆盖")
            self.assertIn("宏批量顺序自检通过", macro_items["17.4.10"]["evidence"])
            self.assertIn("真实 VBA 执行仍需本地客户端", macro_items["17.4.10"]["current"])
            mathtype_items = {item["key"]: item for item in groups["17.7"]["items"]}
            self.assertEqual(mathtype_items["17.7.1"]["status"], "已覆盖")
            self.assertIn("已有 MathType 保留自检通过", mathtype_items["17.7.1"]["evidence"])
            self.assertIn("原生 OLE 写回仍需本地客户端", mathtype_items["17.7.1"]["current"])
            self.assertEqual(mathtype_items["17.7.2"]["status"], "合同覆盖")
            self.assertIn("OMML 转 MathType 交接自检通过", mathtype_items["17.7.2"]["evidence"])
            self.assertEqual(mathtype_items["17.7.3"]["status"], "已覆盖")
            self.assertIn("LaTeX 转 MathType 自检通过", mathtype_items["17.7.3"]["evidence"])
            self.assertIn("原生 MathType 对象写回仍需本地客户端", mathtype_items["17.7.3"]["current"])
            self.assertEqual(mathtype_items["17.7.4"]["status"], "需 Mathpix 实测")
            self.assertIn("PDF 公式识别转 MathType 合同自检通过", mathtype_items["17.7.4"]["evidence"])
            self.assertIn("blocked_mathpix", mathtype_items["17.7.4"]["current"])
            self.assertIn("OCR 轮询", mathtype_items["17.7.4"]["verification"]["uncovered_risk"])
            self.assertEqual(mathtype_items["17.7.7"]["status"], "已覆盖")
            self.assertIn("多个 Word 文件的公式共享全文格式化参数", mathtype_items["17.7.7"]["evidence"])
            self.assertEqual(mathtype_items["17.7.8"]["status"], "已覆盖")
            self.assertIn("当前章节、选中区域和单个公式范围", mathtype_items["17.7.8"]["evidence"])
            hybrid_items = {item["key"]: item for item in groups["17.10"]["items"]}
            self.assertEqual(hybrid_items["17.10.1"]["status"], "已覆盖")
            self.assertIn("本地伴随 CLI", hybrid_items["17.10.1"]["evidence"])
            self.assertIn("动作级 dry-run", hybrid_items["17.10.1"]["evidence"])
            self.assertEqual(hybrid_items["17.10.7"]["status"], "需本地客户端实测")
            self.assertIn("动作级 dry-run", hybrid_items["17.10.7"]["evidence"])
            self.assertIn("k12.localNativeExecutionRequest.v1", hybrid_items["17.10.7"]["evidence"])
            self.assertIn("k12.localFileActionExecution.v1", hybrid_items["17.10.7"]["evidence"])
            self.assertIn("本地文件动作执行自检通过", hybrid_items["17.10.7"]["evidence"])
            self.assertIn("依赖扩展名受控", hybrid_items["17.10.7"]["evidence"])
            self.assertIn("目标目录可写", hybrid_items["17.10.7"]["evidence"])
            self.assertIn("blocked_until_explicit_file_action_request", hybrid_items["17.10.7"]["current"])
            self.assertIn("OMML 依赖扩展名", hybrid_items["17.10.7"]["current"])
            self.assertIn("目标目录写入权限已校验", hybrid_items["17.10.7"]["current"])
            self.assertIn("显式授权后复制 1 个 OMML 依赖", hybrid_items["17.10.7"]["current"])
            self.assertIn("不会执行 Office、MathType、OMML 写回或 Word 宏", hybrid_items["17.10.7"]["current"])
            self.assertEqual(hybrid_items["17.10.7"]["verification"]["scope"], "native_desktop_client")
            self.assertIn("native_writeback_not_verified", hybrid_items["17.10.7"]["verification"]["blocking_reasons"])
            self.assertEqual(hybrid_items["17.10.8"]["status"], "已覆盖")
            self.assertIn("本地结果上传自检通过", hybrid_items["17.10.8"]["evidence"])
            self.assertIn("content_transfer=included", hybrid_items["17.10.8"]["current"])
            self.assertEqual(hybrid_items["17.10.10"]["status"], "已覆盖")
            self.assertIn("本地 API 安全自检通过", hybrid_items["17.10.10"]["evidence"])
            self.assertIn("当前令牌未配置", hybrid_items["17.10.10"]["current"])
            self.assertTrue(any("Windows 与 macOS" in item for item in matrix["guardrails"]))

    def test_product_summary_exposes_prd21_core_capabilities(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            handler = make_handler(store, processor)
            K12RequestHandler._handle_api_get(handler, "/api/product-summary", {})
            response = json.loads(handler.wfile.getvalue().decode("utf-8"))
            summary = response["productSummary"]

            self.assertEqual(handler.status, 200)
            self.assertEqual(summary["schema_version"], "k12.productSummary.v1")
            self.assertEqual(len(summary["core_capabilities"]), 13)
            self.assertEqual(len(summary["phase_roadmap"]), 4)
            titles = [item["title"] for item in summary["core_capabilities"]]
            self.assertIn("PDF 可转换为 Word", titles)
            self.assertIn("用户可选择 Word 宏，并按指定顺序执行", titles)
            self.assertTrue(any(item["key"] == "phase_1" for item in summary["phase_roadmap"]))
            self.assertTrue(any(item["key"] == "phase_4" and item["level"] == "blue" for item in summary["phase_roadmap"]))

    def test_enhancement_plan_exposes_v3_boundaries_without_fake_ai_execution(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            handler = make_handler(store, processor)
            K12RequestHandler._handle_api_get(handler, "/api/enhancement-plan", {})
            response = json.loads(handler.wfile.getvalue().decode("utf-8"))
            plan = response["enhancementPlan"]

            self.assertEqual(handler.status, 200)
            self.assertEqual(plan["schema_version"], "k12.enhancementPlan.v1")
            items = {item["key"]: item for item in plan["items"]}
            self.assertIn("ai_layout_repair", items)
            self.assertIn("smart_template_apply", items)
            self.assertIn("private_deployment", items)
            self.assertIn("open_api", items)
            self.assertEqual(items["ai_layout_repair"]["level"], "blue")
            self.assertNotEqual(items["ai_layout_repair"]["status"], "已落地")
            self.assertEqual(items["smart_template_apply"]["level"], "blue")
            self.assertIn("默认不调用外部 AI", items["ai_layout_repair"]["guardrails"])
            self.assertIn("外部模型需单独授权", items["ai_layout_repair"]["guardrails"])
            self.assertTrue(any(item["key"] == "local_security_token" for item in plan["readiness"]))

    def test_local_client_manifest_and_heartbeat_define_desktop_runtime_contract(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            store.update_settings({
                "localSecurityToken": "secret-token",
                "localClientPlatform": "Windows",
                "allowWebLaunchLocalClient": True,
                "allowTaskStatusCloudSync": True,
            })
            processor = TaskProcessor(store)
            windows_package = store.installers_dir / "K12-Local-Client-Windows-x64.msi"
            windows_package.write_bytes(b"k12 windows installer")
            handler = make_handler(store, processor)
            K12RequestHandler._handle_api_get(handler, "/api/local-client/manifest", {})
            response = json.loads(handler.wfile.getvalue().decode("utf-8"))
            manifest = response["localClientManifest"]

            self.assertEqual(handler.status, 200)
            self.assertEqual(manifest["schema_version"], "k12.localClientManifest.v1")
            self.assertEqual(manifest["platform"]["selected"], "Windows")
            self.assertEqual(manifest["platform"]["install_plan_endpoint"], "/api/install-plan")
            self.assertIn(".msi", manifest["platform"]["package_boundary"])
            self.assertIn(".pkg", manifest["platform"]["package_boundary"])
            self.assertIn("MathType", manifest["platform"]["package_boundary"])
            installer = manifest["platform"]["installer"]
            self.assertEqual(installer["schema_version"], "k12.localInstallerManifest.v1")
            self.assertEqual(installer["platform"], "Windows")
            self.assertEqual(installer["installer_kind"], "windows-msi")
            self.assertEqual(installer["file_name"], "K12-Local-Client-Windows-x64.msi")
            self.assertEqual(installer["status"], "可下载")
            self.assertIn("?platform=Windows", installer["download_url"])
            self.assertTrue(installer["download_requires_platform_query"])
            self.assertEqual(installer["size"], len(b"k12 windows installer"))
            self.assertTrue(installer["sha256"])
            self.assertTrue(installer["checksum_required"])
            self.assertTrue(installer["expected_location_available"])
            self.assertIn("/api/installers/{file_name}", installer["path_policy"])
            self.assertIn("MathType 原生对象", installer["formula_object_boundary"])
            self.assertNotIn(tmp, json.dumps(installer, ensure_ascii=False))
            self.assertEqual(manifest["platform"]["formula_object_interop"]["schema_version"], "k12.formulaObjectInterop.v1")
            self.assertFalse(manifest["platform"]["formula_object_interop"]["platform_objects_cross_compatible"])
            self.assertIn("MathML", manifest["platform"]["formula_object_interop"]["fallback_formats"])
            self.assertTrue(manifest["security"]["token_configured"])
            self.assertTrue(manifest["security"]["payload_requires_configured_token"])
            self.assertEqual(manifest["launch"]["protocol"], "k12-local")
            self.assertTrue(manifest["companion_cli"]["available"])
            self.assertTrue(manifest["companion_cli"]["dry_run_supported"])
            self.assertEqual(manifest["companion_cli"]["dry_run_execution_schema"], "k12.localDryRunExecution.v1")
            self.assertTrue(manifest["companion_cli"]["native_plan_supported"])
            self.assertEqual(manifest["companion_cli"]["native_execution_request_schema"], "k12.localNativeExecutionRequest.v1")
            self.assertTrue(manifest["companion_cli"]["file_actions_supported"])
            self.assertEqual(manifest["companion_cli"]["file_action_execution_schema"], "k12.localFileActionExecution.v1")
            self.assertIn("--native-plan", manifest["companion_cli"]["native_plan_command"])
            self.assertIn("--execute-file-actions", manifest["companion_cli"]["file_action_command"])
            self.assertFalse(manifest["companion_cli"]["executes_native_documents"])
            self.assertIn("安全本地文件动作", manifest["companion_cli"]["description"])
            self.assertIn("python3 -m k12.local_client", manifest["companion_cli"]["command"])
            endpoints = {(item["method"], item["path"]): item for item in manifest["endpoints"]}
            self.assertEqual(endpoints[("GET", "/api/tasks/{task_id}/local-payload")]["auth"], "configured-token")
            self.assertEqual(endpoints[("GET", "/api/tasks/{task_id}/local-readiness")]["auth"], "token-or-local")
            self.assertEqual(endpoints[("POST", "/api/tasks/{task_id}/local-launch")]["auth"], "configured-token + launch-authorization")
            self.assertEqual(endpoints[("GET", "/api/local-client/uploads/{upload_id}/manifest")]["auth"], "token-or-local")
            self.assertEqual(endpoints[("POST", "/api/local-client/uploads")]["auth"], "configured-token + cloud-sync-authorization")
            self.assertEqual(endpoints[("POST", "/api/local-client/heartbeat")]["auth"], "configured-token")
            self.assertNotIn("secret-token", json.dumps(manifest, ensure_ascii=False))

            heartbeat = processor.record_local_client_heartbeat({
                "client_id": "desktop-1",
                "status": "online",
                "platform": "macOS",
                "version": "0.2.0",
                "capabilities": {"macroExecution": True, "file_path": "/tmp/should-not-leak"},
                "preflight": {
                    "platform": "macOS",
                    "components": {
                        "mathtype": {"label": "MathType", "available": True, "status": "available", "path": "/Applications/MathType.app"},
                        "secret": {"label": "secret", "available": True, "status": "available"},
                    },
                    "capabilities": {"mathTypeAutomation": False, "token": "secret"},
                },
            })
            heartbeat_json = json.dumps(heartbeat, ensure_ascii=False)
            self.assertEqual(heartbeat["schema_version"], "k12.localClientHeartbeat.v1")
            self.assertEqual(heartbeat["platform"], "macOS")
            self.assertTrue(heartbeat["capabilities"]["macroExecution"])
            self.assertNotIn("file_path", heartbeat["capabilities"])
            self.assertIn("mathtype", heartbeat["preflight"]["components"])
            self.assertNotIn("secret", heartbeat["preflight"]["components"])
            self.assertNotIn("token", heartbeat["preflight"]["capabilities"])
            self.assertNotIn("/Applications", heartbeat_json)
            self.assertEqual(processor.local_client_manifest()["heartbeat"]["client_id"], "desktop-1")

    def test_local_launch_request_records_protocol_url_without_exposing_token(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            word = processor.create_uploaded_file("lesson.docx", make_docx_bytes())[0]
            task = processor.create_task({"task_type": "omml_to_mathtype", "file_ids": [word["id"]]})

            with self.assertRaises(ValueError):
                processor.create_local_launch_request(task["id"], {"api_origin": "http://127.0.0.1:9999"})

            store.update_settings({"localSecurityToken": "secret-token", "allowWebLaunchLocalClient": True})
            launch = processor.create_local_launch_request(task["id"], {"api_origin": "http://127.0.0.1:9999"})

            self.assertEqual(launch["schema_version"], "k12.localLaunchRequest.v1")
            self.assertEqual(launch["task_id"], task["id"])
            self.assertEqual(launch["status"], "requested")
            self.assertIn("k12-local://open", launch["protocol_url"])
            self.assertIn("http://127.0.0.1:9999", launch["payload_url"])
            self.assertTrue(launch["payload_url_requires_token"])
            self.assertNotIn("secret-token", json.dumps(launch, ensure_ascii=False))
            saved_task = store.get_task(task["id"])
            self.assertEqual(saved_task["local_launch_request"]["id"], launch["id"])
            self.assertEqual(saved_task["local_client_status"], "启动已请求")


class LocalApiSecurityTests(unittest.TestCase):
    def test_security_token_accepts_header_bearer_and_query(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            store.update_settings({"localSecurityToken": "secret-token"})
            handler = SimpleNamespace(store=store, headers={})
            with self.assertRaises(JsonError) as raised:
                K12RequestHandler._ensure_authorized(handler, "/api/files", {})
            self.assertEqual(raised.exception.status, 401)

            handler.headers = {"X-K12-Token": "secret-token"}
            K12RequestHandler._ensure_authorized(handler, "/api/files", {})

            handler.headers = {"Authorization": "Bearer secret-token"}
            K12RequestHandler._ensure_authorized(handler, "/api/files", {})

            handler.headers = {}
            K12RequestHandler._ensure_authorized(handler, "/api/files", {"token": ["secret-token"]})

            K12RequestHandler._ensure_authorized(handler, "/api/health", {})

    def test_sensitive_download_routes_require_local_token_when_configured(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            store.update_settings({"localSecurityToken": "secret-token"})
            handler = SimpleNamespace(store=store, headers={})

            for path in ("/api/assets/images/report/icon.png", "/api/artifacts/task_demo/output.docx"):
                with self.assertRaises(JsonError) as raised:
                    K12RequestHandler._ensure_authorized(handler, path, {})
                self.assertEqual(raised.exception.status, 401)

                handler.headers = {"X-K12-Token": "secret-token"}
                K12RequestHandler._ensure_authorized(handler, path, {})
                handler.headers = {}

    def test_image_asset_download_stays_inside_managed_image_cache(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            image = store.images_dir / "report_demo" / "icon.png"
            image.parent.mkdir(parents=True, exist_ok=True)
            image.write_bytes(make_png_bytes(12, 10))

            handler = make_handler(store, processor)
            K12RequestHandler._send_image_asset(handler, "images/report_demo/icon.png")

            self.assertEqual(handler.status, 200)
            self.assertEqual(dict(handler.output_headers)["Content-Type"], "image/png")
            self.assertEqual(handler.wfile.getvalue(), image.read_bytes())

            (store.data_dir / "secret.png").write_bytes(make_png_bytes(8, 8))
            with self.assertRaises(JsonError) as raised:
                K12RequestHandler._send_image_asset(make_handler(store, processor), "images/../secret.png")
            self.assertEqual(raised.exception.status, 404)
            self.assertEqual(raised.exception.message, "Image not found")

    def test_artifact_download_only_serves_successful_report_entries(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            processor = TaskProcessor(store)
            task_id = "task_demo"
            output_dir = store.output_task_dir(task_id)
            registered = output_dir / "lesson.pptx"
            registered.write_bytes(b"registered artifact")
            unregistered = output_dir / "scratch.txt"
            unregistered.write_text("not linked from report", encoding="utf-8")
            failed = output_dir / "failed.docx"
            failed.write_bytes(b"failed artifact")
            store.save_report(
                {
                    "id": "report_demo",
                    "task_id": task_id,
                    "report_type": "处理报告",
                    "created_at": "2026-06-24T00:00:00+00:00",
                    "files": [],
                    "analysis": {
                        "artifacts": [
                            {"task_id": task_id, "file_name": registered.name, "path": str(registered), "status": "成功"},
                            {"task_id": task_id, "file_name": failed.name, "path": str(failed), "status": "失败"},
                        ]
                    },
                }
            )

            handler = make_handler(store, processor)
            K12RequestHandler._send_artifact(handler, f"/api/artifacts/{task_id}/{registered.name}")

            self.assertEqual(handler.status, 200)
            self.assertEqual(handler.wfile.getvalue(), b"registered artifact")
            self.assertIn(("Content-Disposition", f'attachment; filename="{registered.name}"'), handler.output_headers)
            with self.assertRaises(JsonError) as unregistered_error:
                K12RequestHandler._send_artifact(make_handler(store, processor), f"/api/artifacts/{task_id}/{unregistered.name}")
            self.assertEqual(unregistered_error.exception.status, 404)
            with self.assertRaises(JsonError) as failed_error:
                K12RequestHandler._send_artifact(make_handler(store, processor), f"/api/artifacts/{task_id}/{failed.name}")
            self.assertEqual(failed_error.exception.status, 404)


def make_handler(store: AppStore, processor: TaskProcessor):
    handler = K12RequestHandler.__new__(K12RequestHandler)
    handler.server = SimpleNamespace(store=store, processor=processor)
    handler.headers = {}
    handler.wfile = BytesIO()
    handler.status = 0
    handler.output_headers = []
    handler.send_response = lambda status: setattr(handler, "status", status)
    handler.send_header = lambda key, value: handler.output_headers.append((key, value))
    handler.end_headers = lambda: None
    return handler


def make_docx_bytes(extra_files: dict[str, bytes] | None = None, document_xml: str | None = None) -> bytes:
    buffer = BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        archive.writestr("[Content_Types].xml", "<Types></Types>")
        archive.writestr("word/document.xml", document_xml or "<w:document><w:p>题目</w:p><m:oMath>x</m:oMath></w:document>")
        archive.writestr("docProps/app.xml", "<Properties><Pages>3</Pages></Properties>")
        for name, content in (extra_files or {}).items():
            archive.writestr(name, content)
    return buffer.getvalue()


def make_zip_bytes(files: dict[str, bytes]) -> bytes:
    buffer = BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        for name, content in files.items():
            archive.writestr(name, content)
    return buffer.getvalue()


def make_pptx_bytes(extra_files: dict[str, bytes] | None = None, slide_xml: str | None = None) -> bytes:
    buffer = BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        archive.writestr("[Content_Types].xml", "<Types></Types>")
        archive.writestr("ppt/presentation.xml", "<p:presentation></p:presentation>")
        archive.writestr(
            "ppt/slides/slide1.xml",
            slide_xml
            or (
                '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
                'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
                "<a:t>第一课</a:t><a:t>知识点 A</a:t><a:t>知识点 B</a:t>"
                "</p:sld>"
            ),
        )
        for name, content in (extra_files or {}).items():
            archive.writestr(name, content)
    return buffer.getvalue()


def make_pptx_slide_with_objects_xml() -> str:
    return (
        '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">'
        "<p:cSld><p:spTree>"
        "<p:sp><p:txBody><a:p><a:r><a:t>第一课</a:t></a:r></a:p></p:txBody></p:sp>"
        "<p:sp><p:txBody><a:p><a:r><a:t>知识点 A</a:t></a:r></a:p></p:txBody></p:sp>"
        "<a:tbl><a:tr><a:tc><a:txBody><a:p><a:r><a:t>表格单元格</a:t></a:r></a:p></a:txBody></a:tc></a:tr></a:tbl>"
        "<m:oMath><m:r><m:t>x^2+y^2=z^2</m:t></m:r></m:oMath>"
        "</p:spTree></p:cSld>"
        "</p:sld>"
    )


def make_pptx_object_files() -> dict[str, bytes]:
    return {
        "ppt/slides/_rels/slide1.xml.rels": b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>',
        "ppt/notesSlides/notesSlide1.xml": '<p:notes xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:t>讲者备注：先讲目标，再讲例题</a:t></p:notes>'.encode("utf-8"),
        "ppt/media/image1.png": make_png_bytes(32, 24),
        "ppt/charts/chart1.xml": b"<c:chartSpace><c:chart></c:chart></c:chartSpace>",
        "ppt/slideMasters/slideMaster1.xml": b"<p:sldMaster></p:sldMaster>",
        "ppt/slideLayouts/slideLayout1.xml": b"<p:sldLayout></p:sldLayout>",
    }


def make_xlsx_bytes(extra_files: dict[str, bytes] | None = None) -> bytes:
    buffer = BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        archive.writestr(
            "[Content_Types].xml",
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>',
        )
        archive.writestr("_rels/.rels", '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')
        archive.writestr("xl/workbook.xml", '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheets><sheet name="成绩" sheetId="1"/></sheets></workbook>')
        archive.writestr(
            "xl/worksheets/sheet1.xml",
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>姓名</t></is></c><c r="B1" t="inlineStr"><is><t>分数</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>Alice</t></is></c><c r="B2"><v>98</v></c></row><row r="3"><c r="B3"><f>SUM(B2)</f><v>98</v></c></row></sheetData></worksheet>',
        )
        for name, content in (extra_files or {}).items():
            archive.writestr(name, content)
    return buffer.getvalue()


def make_xlsx_multisheet_bytes() -> bytes:
    buffer = BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        archive.writestr(
            "[Content_Types].xml",
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>',
        )
        archive.writestr("_rels/.rels", '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')
        archive.writestr("xl/workbook.xml", '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheets><sheet name="成绩" sheetId="1"/><sheet name="汇总" sheetId="2"/></sheets></workbook>')
        archive.writestr(
            "xl/worksheets/sheet1.xml",
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>姓名</t></is></c><c r="B1" t="inlineStr"><is><t>分数</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>Alice</t></is></c><c r="B2"><v>98</v></c></row></sheetData></worksheet>',
        )
        archive.writestr(
            "xl/worksheets/sheet2.xml",
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>班级</t></is></c><c r="B1" t="inlineStr"><is><t>平均分</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>一班</t></is></c><c r="B2"><v>92</v></c></row></sheetData></worksheet>',
        )
    return buffer.getvalue()


def make_xlsx_object_files() -> dict[str, bytes]:
    return {
        "xl/worksheets/_rels/sheet1.xml.rels": b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments1.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/></Relationships>',
        "xl/drawings/drawing1.xml": b'<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><xdr:twoCellAnchor><xdr:graphicFrame><a:graphic><a:graphicData><c:chart r:id="rId1"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:pic><xdr:nvPicPr/><xdr:blipFill/></xdr:pic></xdr:twoCellAnchor></xdr:wsDr>',
        "xl/drawings/_rels/drawing1.xml.rels": b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/></Relationships>',
        "xl/charts/chart1.xml": b"<c:chartSpace><c:chart></c:chart></c:chartSpace>",
        "xl/media/image1.png": make_png_bytes(32, 24),
        "xl/comments1.xml": '<comments><comment ref="C1"><text><r><t>复核</t></r></text></comment></comments>'.encode("utf-8"),
        "xl/tables/table1.xml": b'<table id="1" name="Scores" displayName="Scores" ref="A1:B3"></table>',
    }


def make_pdf_with_image_bytes(width: int, height: int, stream: bytes | None = None) -> bytes:
    image_stream = stream or b"\xff\xd8K12PDFIMAGE\xff\xd9"
    header = (
        "%PDF-1.4\n"
        "1 0 obj<< /Type /Page >>endobj\n"
        "2 0 obj<< /Type /XObject /Subtype /Image "
        f"/Width {width} /Height {height} "
        "/ColorSpace /DeviceRGB /BitsPerComponent 8 "
        f"/Filter /DCTDecode /Length {len(image_stream)} >>\n"
        "stream\n"
    ).encode("ascii")
    return header + image_stream + b"\nendstream\nendobj\n%%EOF\n"


def make_png_bytes(width: int, height: int, alpha: bool = False) -> bytes:
    pixel = b"\x00\x00\x00\x00" if alpha else b"\x00\x00\x00"
    color_type = 6 if alpha else 2
    raw = b"".join(b"\x00" + pixel * width for _ in range(height))
    chunks = [
        png_chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, color_type, 0, 0, 0)),
        png_chunk(b"IDAT", zlib.compress(raw)),
        png_chunk(b"IEND", b""),
    ]
    return b"\x89PNG\r\n\x1a\n" + b"".join(chunks)


def png_chunk(kind: bytes, data: bytes) -> bytes:
    return len(data).to_bytes(4, "big") + kind + data + zlib.crc32(kind + data).to_bytes(4, "big")


if __name__ == "__main__":
    unittest.main()
