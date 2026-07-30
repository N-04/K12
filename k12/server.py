"""Local HTTP API and static-file server for the K12 workbench.

The server exposes the PRD's browser workflow, report downloads, local-client
handoff APIs, and token-gated sensitive payloads. It also redacts local paths by
default so the web UI can operate without leaking filesystem details.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import mimetypes
import re
import secrets
import zipfile
from io import BytesIO
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse

from .exports import build_formula_xlsx, build_formula_zip, build_image_manifest_xlsx, build_macro_failure_csv, build_omml_failure_csv
from .processor import TaskProcessor
from .reports import ReportBuilder
from .store import AppStore


ROOT = Path(__file__).resolve().parent.parent
STATIC_DIR = ROOT / "static"
LOCAL_PATH_KEYS = {
    "file_path",
    "storage_path",
    "input_path",
    "output_path",
    "report_path",
    "path",
    "html_path",
    "json_path",
    "pdf_path",
    "xlsx_path",
    "txt_path",
    "failure_csv_path",
    "backup_path",
    "expected_location",
    "manual_omml_path",
    "target_path",
    "document_path",
    "omml_source_path",
    "omml_target_path",
    "image_path",
    "original_image_path",
    "template_path",
}


class JsonError(Exception):
    """HTTP-facing error with a status code and user-readable message."""

    def __init__(self, status: int, message: str) -> None:
        """Store the HTTP status and response message for JSON error replies."""
        super().__init__(message)
        self.status = status
        self.message = message


class K12RequestHandler(BaseHTTPRequestHandler):
    """Route local API requests while enforcing token and login boundaries."""

    server_version = "K12LocalAPI/0.1"

    @property
    def store(self) -> AppStore:
        """Return the request-scoped runtime store owned by the server."""
        return self.server.store  # type: ignore[attr-defined]

    @property
    def processor(self) -> TaskProcessor:
        """Return the task processor bound to the current local server."""
        return self.server.processor  # type: ignore[attr-defined]

    def do_GET(self) -> None:
        """Serve static assets or authorized JSON/download GET routes."""
        try:
            parsed = urlparse(self.path)
            if parsed.path.startswith("/api/"):
                query = parse_qs(parsed.query)
                self._ensure_authorized(parsed.path, query)
                self._ensure_logged_in(parsed.path, "GET")
                self._handle_api_get(parsed.path, query)
            else:
                self._serve_static(parsed.path)
        except JsonError as exc:
            self._json({"error": exc.message}, exc.status)
        except Exception as exc:  # pragma: no cover
            self._json({"error": str(exc)}, 500)

    def do_POST(self) -> None:
        """Handle authorized JSON and multipart POST workflow routes."""
        try:
            parsed = urlparse(self.path)
            if not parsed.path.startswith("/api/"):
                raise JsonError(404, "Not found")
            self._ensure_authorized(parsed.path, parse_qs(parsed.query))
            self._ensure_logged_in(parsed.path, "POST")
            if parsed.path == "/api/uploads":
                uploads = self._read_multipart_uploads()
                files = []
                try:
                    for upload in uploads:
                        files.extend(self.processor.create_uploaded_file(upload["file_name"], upload["content"]))
                except ValueError as exc:
                    raise JsonError(400, str(exc)) from exc
                self._json({"files": [self._public_payload(file) for file in files]}, 201)
                return
            replace_match = re.fullmatch(r"/api/files/([^/]+)/replace", parsed.path)
            if replace_match:
                uploads = self._read_multipart_uploads()
                if len(uploads) != 1:
                    raise JsonError(400, "重新上传一次只能选择一个文件")
                upload = uploads[0]
                try:
                    file = self.processor.replace_uploaded_file(replace_match.group(1), upload["file_name"], upload["content"])
                except KeyError as exc:
                    raise JsonError(404, "File not found") from exc
                except ValueError as exc:
                    raise JsonError(400, str(exc)) from exc
                self._json({"file": self._public_payload(file)})
                return
            payload = self._read_json()
            self._handle_api_post(parsed.path, payload)
        except JsonError as exc:
            self._json({"error": exc.message}, exc.status)
        except Exception as exc:  # pragma: no cover
            self._json({"error": str(exc)}, 500)

    def do_PUT(self) -> None:
        """Handle authorized settings updates."""
        try:
            parsed = urlparse(self.path)
            self._ensure_authorized(parsed.path, parse_qs(parsed.query))
            self._ensure_logged_in(parsed.path, "PUT")
            payload = self._read_json()
            if parsed.path == "/api/settings":
                try:
                    self._json({"settings": self._public_settings(self.processor.update_settings(payload))})
                except ValueError as exc:
                    raise JsonError(400, str(exc)) from exc
            else:
                raise JsonError(404, "Not found")
        except JsonError as exc:
            self._json({"error": exc.message}, exc.status)
        except Exception as exc:  # pragma: no cover
            self._json({"error": str(exc)}, 500)

    def do_DELETE(self) -> None:
        """Handle authorized deletion routes for files, reports, and annotations."""
        try:
            parsed = urlparse(self.path)
            self._ensure_authorized(parsed.path, parse_qs(parsed.query))
            self._ensure_logged_in(parsed.path, "DELETE")
            parts = parsed.path.strip("/").split("/")
            if len(parts) == 3 and parts[:2] == ["api", "files"]:
                try:
                    deleted = self.processor.delete_file(parts[2])
                except ValueError as exc:
                    raise JsonError(400, str(exc)) from exc
                self._json({"deleted": deleted})
            elif len(parts) == 3 and parts[:2] == ["api", "reports"]:
                try:
                    self._json(self.processor.delete_report(parts[2]))
                except ValueError as exc:
                    raise JsonError(400, str(exc)) from exc
            elif len(parts) == 3 and parts[:2] == ["api", "macro-templates"]:
                try:
                    deleted = self.processor.delete_macro_template(parts[2])
                except ValueError as exc:
                    raise JsonError(400, str(exc)) from exc
                self._json({"deleted": deleted})
            elif len(parts) == 3 and parts[:2] == ["api", "users"]:
                try:
                    deleted = self.processor.delete_user(parts[2])
                except ValueError as exc:
                    raise JsonError(400, str(exc)) from exc
                self._json({"deleted": deleted})
            elif len(parts) == 3 and parts[:2] == ["api", "templates"]:
                try:
                    deleted = self.processor.delete_template(parts[2])
                except ValueError as exc:
                    raise JsonError(400, str(exc)) from exc
                self._json({"deleted": deleted})
            elif len(parts) == 3 and parts[:2] == ["api", "image-annotations"]:
                deleted = self.processor.delete_image_annotation(parts[2])
                self._json({"deleted": deleted})
            elif len(parts) == 3 and parts[:2] == ["api", "formula-annotations"]:
                deleted = self.processor.delete_formula_annotation(parts[2])
                self._json({"deleted": deleted})
            elif len(parts) == 3 and parts[:2] == ["api", "omml-annotations"]:
                deleted = self.processor.delete_omml_annotation(parts[2])
                self._json({"deleted": deleted})
            elif len(parts) == 3 and parts[:2] == ["api", "layout-annotations"]:
                deleted = self.processor.delete_layout_annotation(parts[2])
                self._json({"deleted": deleted})
            else:
                raise JsonError(404, "Not found")
        except JsonError as exc:
            self._json({"error": exc.message}, exc.status)
        except Exception as exc:  # pragma: no cover
            self._json({"error": str(exc)}, 500)

    def do_OPTIONS(self) -> None:
        """Return CORS preflight headers for local API clients."""
        self.send_response(204)
        self._send_common_headers()
        self.end_headers()

    def _handle_api_get(self, path: str, query: dict[str, list[str]]) -> None:
        """Dispatch authorized GET routes for browser, report, installer, and local-client reads."""
        if path == "/api/health":
            self._json({"status": "ok", "service": "K12", "version": "0.1.0"})
        elif path == "/api/files":
            self._json({"files": [self._public_payload(file) for file in self.store.list_files()]})
        elif path == "/api/files/download":
            self._send_files_bundle(query)
        elif path.startswith("/api/files/") and path.endswith("/reports"):
            file_id = path.strip("/").split("/")[2]
            try:
                reports = self.processor.reports_for_file(file_id)
            except KeyError as exc:
                raise JsonError(404, "File not found") from exc
            self._json({"reports": [self._public_payload(report) for report in reports]})
        elif path.startswith("/api/files/") and path.endswith("/preview"):
            file_id = path.strip("/").split("/")[2]
            self._json({"preview": self.processor.file_preview(file_id)})
        elif path.startswith("/api/files/") and path.endswith("/download"):
            file_id = path.strip("/").split("/")[2]
            self._send_source_file(file_id)
        elif path == "/api/tasks":
            self._json({"tasks": [self._public_payload(task) for task in self.store.list_tasks()]})
        elif path == "/api/tasks/recovery-summary":
            self._json({"taskRecoverySummary": self.processor.task_recovery_summary()})
        elif path.startswith("/api/tasks/") and path.endswith("/local-readiness"):
            task_id = path.strip("/").split("/")[2]
            try:
                self._json({"localReadiness": self.processor.local_task_readiness(task_id)})
            except KeyError as exc:
                raise JsonError(404, "Task not found") from exc
        elif path.startswith("/api/tasks/") and path.endswith("/local-payload"):
            self._ensure_local_payload_token_configured()
            task_id = path.strip("/").split("/")[2]
            try:
                self._json({"localPayload": self.processor.local_task_payload(task_id)})
            except KeyError as exc:
                raise JsonError(404, "Task not found") from exc
        elif path == "/api/reports":
            self._json({"reports": [self._public_payload(report) for report in self.store.list_reports()]})
        elif path == "/api/logs":
            task_id = query.get("task_id", [None])[0]
            self._json({"logs": [self._public_log(log) for log in self.store.list_logs(task_id)]})
        elif path == "/api/logs/download":
            task_id = query.get("task_id", [None])[0]
            self._send_logs_file(task_id, query.get("format", [None])[0])
        elif path == "/api/local-client/manifest":
            self._json({"localClientManifest": self.processor.local_client_manifest()})
        elif path.startswith("/api/local-client/uploads/") and path.endswith("/manifest"):
            upload_id = path.strip("/").split("/")[3]
            try:
                self._json({"uploadManifest": self.processor.local_result_upload_manifest(upload_id)})
            except KeyError as exc:
                raise JsonError(404, "Upload not found") from exc
        elif path == "/api/local-client/uploads":
            self._json({"uploadQueue": self.processor.local_result_upload_queue()})
        elif path == "/api/mathpix-jobs":
            self._json({"mathpixJobs": self.processor.mathpix_job_queue()})
        elif path == "/api/image-annotations":
            report_id = query.get("report_id", [None])[0]
            self._json({"annotations": self.processor.list_image_annotations(report_id)})
        elif path == "/api/formula-annotations":
            report_id = query.get("report_id", [None])[0]
            self._json({"annotations": self.processor.list_formula_annotations(report_id)})
        elif path == "/api/omml-annotations":
            report_id = query.get("report_id", [None])[0]
            self._json({"annotations": self._public_payload(self.processor.list_omml_annotations(report_id))})
        elif path == "/api/layout-annotations":
            report_id = query.get("report_id", [None])[0]
            self._json({"annotations": self.processor.list_layout_annotations(report_id)})
        elif path == "/api/settings":
            self._json({"settings": self._public_settings(self.store.get_settings())})
        elif path == "/api/capabilities":
            self._json({"capabilities": self.processor.capabilities()})
        elif path == "/api/architecture":
            self._json({"architecture": self.processor.architecture_blueprint()})
        elif path == "/api/api-catalog":
            self._json({"apiCatalog": self.processor.api_catalog()})
        elif path == "/api/data-dictionary":
            self._json({"dataDictionary": self.processor.data_dictionary()})
        elif path == "/api/acceptance-matrix":
            self._json({"acceptanceMatrix": self.processor.acceptance_matrix()})
        elif path == "/api/product-summary":
            self._json({"productSummary": self.processor.product_summary()})
        elif path == "/api/enhancement-plan":
            self._json({"enhancementPlan": self.processor.enhancement_plan()})
        elif path == "/api/install-profile":
            platform_name = query.get("platform", [None])[0]
            self._json({"installProfile": self.processor.install_profile(platform_name)})
        elif path == "/api/install-plan":
            platform_name = query.get("platform", [None])[0]
            self._json({"installPlan": self._public_payload(self.processor.install_plan(platform_name))})
        elif path.startswith("/api/installers/"):
            self._send_installer(path.removeprefix("/api/installers/"), query)
        elif path == "/api/macros":
            self._json({"macros": self.processor.macro_library()})
        elif path == "/api/macro-templates":
            self._json({"templates": self.processor.list_macro_templates()})
        elif path == "/api/users":
            self._json({"users": self._public_payload(self.processor.list_users()), "currentUser": self.processor.current_user()})
        elif path == "/api/templates":
            self._json({"templates": self._public_payload(self.processor.list_templates())})
        elif path == "/api/authorizations":
            self._json({"authorizations": self.processor.list_authorizations()})
        elif path == "/api/preflight":
            self._json({"checks": self._public_payload(self.processor.preflight_checks())})
        elif path.startswith("/api/assets/replacements/"):
            self._send_replacement_asset(path.removeprefix("/api/assets/replacements/"))
        elif path.startswith("/api/assets/images/"):
            self._send_image_asset(path.removeprefix("/api/assets/"))
        elif path.startswith("/api/artifacts/"):
            self._send_artifact(path)
        elif path.startswith("/api/tasks/") and path.endswith("/download"):
            task_id = path.strip("/").split("/")[2]
            self._send_task_bundle(task_id)
        elif path.startswith("/api/reports/") and path.endswith("/download"):
            report_id = path.strip("/").split("/")[2]
            report = self.store.get_report(report_id)
            if not report:
                raise JsonError(404, "Report not found")
            self._send_report_file(report, query.get("format", ["html"])[0])
        elif path.startswith("/api/reports/") and path.endswith("/images.zip"):
            report_id = path.strip("/").split("/")[2]
            self._send_report_images(report_id)
        elif path.startswith("/api/reports/") and path.endswith("/images.xlsx"):
            report_id = path.strip("/").split("/")[2]
            self._send_report_image_manifest_xlsx(report_id)
        elif path.startswith("/api/reports/") and path.endswith("/formulas.zip"):
            report_id = path.strip("/").split("/")[2]
            self._send_report_formulas(report_id)
        elif path.startswith("/api/reports/") and path.endswith("/formulas.xlsx"):
            report_id = path.strip("/").split("/")[2]
            self._send_report_formula_xlsx(report_id)
        elif path.startswith("/api/reports/") and path.endswith("/omml-failures.csv"):
            report_id = path.strip("/").split("/")[2]
            self._send_report_omml_failures(report_id)
        elif path.startswith("/api/reports/") and path.endswith("/macro-failures.csv"):
            report_id = path.strip("/").split("/")[2]
            self._send_report_macro_failures(report_id)
        elif path.startswith("/api/reports/") and path.endswith("/failures.csv"):
            report_id = path.strip("/").split("/")[2]
            self._send_report_failures(report_id)
        elif path.startswith("/api/reports/") and path.endswith("/comparison"):
            report_id = path.strip("/").split("/")[2]
            try:
                self._json({"comparison": self._public_payload(self.processor.report_comparison(report_id))})
            except KeyError as exc:
                raise JsonError(404, "Report not found") from exc
        elif path.startswith("/api/reports/"):
            report_id = path.rsplit("/", 1)[-1]
            report = self.store.get_report(report_id)
            if not report:
                raise JsonError(404, "Report not found")
            self._json({"report": self._public_payload(report)})
        else:
            raise JsonError(404, "Not found")

    def _handle_api_post(self, path: str, payload: dict) -> None:
        """Dispatch authorized POST routes for uploads, tasks, annotations, sync, and settings."""
        if path == "/api/files":
            try:
                if "files" in payload:
                    files = [self.processor.create_file(item) for item in payload["files"]]
                    self._json({"files": [self._public_payload(file) for file in files]}, 201)
                else:
                    self._json({"file": self._public_payload(self.processor.create_file(payload))}, 201)
            except ValueError as exc:
                raise JsonError(400, str(exc)) from exc
        elif path.startswith("/api/files/") and path.endswith("/password"):
            file_id = path.split("/")[-2]
            try:
                self._json({"file": self._public_payload(self.processor.set_file_password(file_id, str(payload.get("password") or "")))})
            except ValueError as exc:
                raise JsonError(400, str(exc)) from exc
        elif path == "/api/tasks":
            try:
                self._json({"task": self._public_payload(self.processor.create_task(payload))}, 201)
            except ValueError as exc:
                raise JsonError(400, str(exc)) from exc
        elif path.startswith("/api/tasks/") and path.endswith("/retry"):
            task_id = path.split("/")[-2]
            try:
                self._json({"task": self._public_payload(self.processor.retry_task(task_id))})
            except ValueError as exc:
                raise JsonError(400, str(exc)) from exc
        elif path.startswith("/api/tasks/") and path.endswith("/completion-notice/ack"):
            task_id = path.strip("/").split("/")[2]
            try:
                self._json({"task": self._public_payload(self.processor.acknowledge_task_completion_notice(task_id))})
            except KeyError as exc:
                raise JsonError(404, "Task not found") from exc
        elif path.startswith("/api/tasks/") and path.endswith("/skip-file"):
            task_id = path.split("/")[-2]
            try:
                self._json({"task": self._public_payload(self.processor.skip_batch_file(task_id, str(payload.get("file_id") or payload.get("fileId") or "")))})
            except KeyError as exc:
                raise JsonError(404, "Task not found") from exc
            except ValueError as exc:
                raise JsonError(400, str(exc)) from exc
        elif path.startswith("/api/tasks/") and path.endswith("/cancel"):
            task_id = path.split("/")[-2]
            try:
                self._json({"task": self._public_payload(self.processor.cancel_task(task_id))})
            except ValueError as exc:
                raise JsonError(400, str(exc)) from exc
        elif path.startswith("/api/tasks/") and path.endswith("/pause"):
            task_id = path.split("/")[-2]
            try:
                self._json({"task": self._public_payload(self.processor.pause_task(task_id))})
            except ValueError as exc:
                raise JsonError(400, str(exc)) from exc
        elif path.startswith("/api/tasks/") and path.endswith("/resume"):
            task_id = path.split("/")[-2]
            try:
                self._json({"task": self._public_payload(self.processor.resume_task(task_id))})
            except ValueError as exc:
                raise JsonError(400, str(exc)) from exc
        elif path.startswith("/api/tasks/") and path.endswith("/local-sync"):
            self._ensure_local_payload_token_configured()
            task_id = path.split("/")[-2]
            try:
                self._json({"task": self._public_payload(self.processor.sync_local_task_status(task_id, payload))})
            except KeyError as exc:
                raise JsonError(404, "Task not found") from exc
            except ValueError as exc:
                raise JsonError(400, str(exc)) from exc
        elif path.startswith("/api/tasks/") and path.endswith("/local-launch"):
            self._ensure_local_payload_token_configured()
            task_id = path.split("/")[-2]
            try:
                launch_request = self.processor.create_local_launch_request(task_id, payload)
                self._json({
                    "launchRequest": launch_request,
                    "task": self._public_payload(self.store.get_task(task_id) or {}),
                    "localClientManifest": self.processor.local_client_manifest(),
                })
            except KeyError as exc:
                raise JsonError(404, "Task not found") from exc
            except ValueError as exc:
                raise JsonError(400, str(exc)) from exc
        elif path == "/api/local-client/heartbeat":
            self._ensure_local_payload_token_configured()
            self._json({"heartbeat": self.processor.record_local_client_heartbeat(payload), "localClientManifest": self.processor.local_client_manifest()})
        elif path == "/api/local-client/uploads":
            self._ensure_local_payload_token_configured()
            try:
                upload = self.processor.register_local_result_upload(payload)
                self._json({"upload": upload, "uploadQueue": self.processor.local_result_upload_queue()})
            except KeyError as exc:
                raise JsonError(404, "Task not found") from exc
            except ValueError as exc:
                raise JsonError(400, str(exc)) from exc
        elif path.startswith("/api/tasks/") and path.endswith("/restore-backups"):
            task_id = path.split("/")[-2]
            try:
                self._json(self.processor.restore_task_backups(task_id))
            except ValueError as exc:
                raise JsonError(400, str(exc)) from exc
        elif path == "/api/maintenance/cleanup":
            self._json(self.processor.cleanup_runtime_history(payload))
        elif path == "/api/macro-templates":
            try:
                self._json({"template": self.processor.save_macro_template(payload)}, 201)
            except ValueError as exc:
                raise JsonError(400, str(exc)) from exc
        elif path == "/api/users":
            try:
                self._json({"user": self._public_payload(self.processor.save_user(payload))}, 201)
            except ValueError as exc:
                raise JsonError(400, str(exc)) from exc
        elif path == "/api/session":
            try:
                self._json({"currentUser": self.processor.activate_user(str(payload.get("user_id") or payload.get("userId") or ""))})
            except ValueError as exc:
                raise JsonError(400, str(exc)) from exc
        elif path == "/api/templates":
            try:
                self._json({"template": self._public_payload(self.processor.save_template(payload))}, 201)
            except ValueError as exc:
                raise JsonError(400, str(exc)) from exc
        elif path == "/api/authorizations":
            try:
                self._json({"authorization": self.processor.save_authorization(payload)}, 201)
            except ValueError as exc:
                raise JsonError(400, str(exc)) from exc
        elif path == "/api/image-annotations":
            try:
                self._json({"annotation": self.processor.save_image_annotation(payload)}, 201)
            except ValueError as exc:
                raise JsonError(400, str(exc)) from exc
        elif path == "/api/image-annotations/replacement":
            try:
                self._json({"annotation": self.processor.save_image_replacement(payload)}, 201)
            except ValueError as exc:
                raise JsonError(400, str(exc)) from exc
        elif path == "/api/image-assets/reexport":
            try:
                self._json(self._public_payload(self.processor.reexport_image_asset(payload)))
            except ValueError as exc:
                raise JsonError(400, str(exc)) from exc
        elif path == "/api/formula-annotations/bulk-confirm":
            try:
                self._json(self.processor.bulk_confirm_formulas(payload), 201)
            except ValueError as exc:
                raise JsonError(400, str(exc)) from exc
        elif path == "/api/formula-annotations":
            try:
                self._json({"annotation": self.processor.save_formula_annotation(payload)}, 201)
            except ValueError as exc:
                raise JsonError(400, str(exc)) from exc
        elif path == "/api/omml-annotations":
            try:
                self._json({"annotation": self._public_payload(self.processor.save_omml_annotation(payload))}, 201)
            except ValueError as exc:
                raise JsonError(400, str(exc)) from exc
        elif path == "/api/layout-annotations":
            try:
                self._json({"annotation": self.processor.save_layout_annotation(payload)}, 201)
            except ValueError as exc:
                raise JsonError(400, str(exc)) from exc
        else:
            raise JsonError(404, "Not found")

    def _serve_static(self, path: str) -> None:
        """Serve static UI files while keeping requests inside STATIC_DIR."""
        clean_path = path.strip("/") or "index.html"
        target = (STATIC_DIR / clean_path).resolve()
        if not str(target).startswith(str(STATIC_DIR.resolve())):
            raise JsonError(403, "Forbidden")
        if target.is_dir():
            target = target / "index.html"
        if not target.exists():
            target = STATIC_DIR / "index.html"
        content_type = mimetypes.guess_type(str(target))[0] or "application/octet-stream"
        data = target.read_bytes()
        self.send_response(200)
        self.send_header("Content-Type", content_type)
        self._send_common_headers()
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def _ensure_authorized(self, path: str, query: dict[str, list[str]]) -> None:
        """Require the local security token once the user configures one."""
        if path == "/api/health":
            return
        expected = str(self.store.get_settings().get("localSecurityToken") or "").strip()
        if not expected:
            return
        incoming = self.headers.get("X-K12-Token", "").strip()
        auth_header = self.headers.get("Authorization", "")
        if auth_header.lower().startswith("bearer "):
            incoming = auth_header[7:].strip() or incoming
        if not incoming:
            incoming = query.get("token", [""])[0].strip()
        if not secrets.compare_digest(incoming, expected):
            raise JsonError(401, "本地 API 安全令牌无效或缺失")

    def _ensure_logged_in(self, path: str, method: str) -> None:
        """Require an active local user session when the workspace enables login."""
        if not path.startswith("/api/"):
            return
        if not self.store.get_settings().get("requireLogin", False):
            return
        public = path == "/api/health" or (path == "/api/session" and method == "POST") or (path == "/api/users" and method == "GET")
        if public:
            return
        user = self.processor.current_user()
        if user.get("status") != "启用" or not user.get("login_enabled", True) or not user.get("last_login_at"):
            raise JsonError(401, "请先登录后再访问本地 API")

    def _ensure_local_payload_token_configured(self) -> None:
        """Block local task payload reads until a local security token exists."""
        expected = str(self.store.get_settings().get("localSecurityToken") or "").strip()
        if not expected:
            raise JsonError(403, "本地任务载荷包含本地路径，请先配置本地安全令牌")

    def _read_json(self) -> dict:
        """Read one JSON request body and map parse failures to a JSON API error."""
        length = int(self.headers.get("Content-Length") or 0)
        if length == 0:
            return {}
        raw = self.rfile.read(length).decode("utf-8")
        try:
            return json.loads(raw)
        except json.JSONDecodeError as exc:
            raise JsonError(400, "Invalid JSON") from exc

    def _read_multipart_uploads(self) -> list[dict]:
        """Parse multipart file uploads while enforcing the configured size limit."""
        content_type = self.headers.get("Content-Type") or ""
        match = re.search(r"boundary=([^;]+)", content_type)
        if not match:
            raise JsonError(400, "Missing multipart boundary")
        boundary = match.group(1).strip('"')
        length = int(self.headers.get("Content-Length") or 0)
        limit = int(self.store.get_settings().get("singleFileLimitMb", 500)) * 1024 * 1024
        if length > limit * 2:
            raise JsonError(413, "上传内容超过限制")
        raw = self.rfile.read(length)
        uploads: list[dict] = []
        marker = f"--{boundary}".encode("utf-8")
        for part in raw.split(marker):
            if not part or part in {b"--\r\n", b"--"}:
                continue
            header_blob, _, body = part.partition(b"\r\n\r\n")
            if not body:
                continue
            headers = header_blob.decode("utf-8", errors="ignore")
            name_match = re.search(r'name="([^"]+)"', headers, re.I)
            filename_match = re.search(r'filename="([^"]*)"', headers, re.I)
            if not name_match or name_match.group(1) != "files" or not filename_match:
                continue
            content = body.rsplit(b"\r\n", 1)[0]
            if len(content) > limit:
                raise JsonError(413, f"{filename_match.group(1)} 超过单文件大小限制")
            uploads.append({"file_name": filename_match.group(1), "content": content})
        if not uploads:
            raise JsonError(400, "没有收到文件")
        return uploads

    def _send_report_file(self, report: dict, file_format: str) -> None:
        """Stream one generated report format from the managed report paths."""
        choices = {
            "json": ("json_path", "application/json; charset=utf-8"),
            "html": ("html_path", "text/html; charset=utf-8"),
            "pdf": ("pdf_path", "application/pdf"),
            "xlsx": ("xlsx_path", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
            "txt": ("txt_path", "text/plain; charset=utf-8"),
        }
        key, content_type = choices.get(file_format, choices["html"])
        path, data = self._read_managed_report(report, key)
        self.send_response(200)
        self.send_header("Content-Type", content_type)
        self._send_common_headers()
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Content-Disposition", f'attachment; filename="{self._attachment_name(path.name)}"')
        self.end_headers()
        self.wfile.write(data)

    def _send_source_file(self, file_id: str) -> None:
        """Stream a registered source file through the processor download contract."""
        info, data = self._read_verified_source(file_id)
        path = Path(info["path"])
        content_type = mimetypes.guess_type(str(path))[0] or "application/octet-stream"
        file_name = self._attachment_name(info.get("file_name") or path.name)
        self.send_response(200)
        self.send_header("Content-Type", content_type)
        self._send_common_headers()
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Content-Disposition", f'attachment; filename="{file_name}"')
        self.end_headers()
        self.wfile.write(data)

    def _read_verified_source(self, file_id: str) -> tuple[dict, bytes]:
        """Read a source snapshot and recheck its registered size and hash."""
        try:
            info = self.processor.file_download_info(file_id)
        except KeyError as exc:
            raise JsonError(404, "File not found") from exc
        except ValueError as exc:
            raise JsonError(404, str(exc)) from exc
        path = Path(info["path"])
        try:
            data = path.read_bytes()
        except OSError as exc:
            raise JsonError(404, "文件源不存在，无法下载") from exc
        if len(data) != int(info["file_size"]) or hashlib.sha256(data).hexdigest() != str(info.get("sha256") or ""):
            raise JsonError(404, "文件源内容与登记记录不一致，已拒绝下载")
        return info, data

    def _send_installer(self, file_name: str, query: dict[str, list[str]] | None = None) -> None:
        """Stream a registered Windows or macOS installer with MathType boundary headers."""
        info, data = self._read_verified_installer(file_name, query)
        path = Path(info["path"])
        content_type = mimetypes.guess_type(str(path))[0] or "application/octet-stream"
        safe_name = self._attachment_name(info.get("file_name") or path.name)
        self.send_response(200)
        self.send_header("Content-Type", content_type)
        self._send_common_headers()
        self.send_header("Content-Length", str(len(data)))
        self.send_header("X-K12-SHA256", str(info.get("sha256") or ""))
        self.send_header("X-K12-Platform", str(info.get("platform") or ""))
        self.send_header("X-K12-Installer-Kind", str(info.get("installer_kind") or ""))
        self.send_header("X-K12-Package-Boundary", str(info.get("package_boundary_header") or ""))
        self.send_header("X-K12-Formula-Object-Boundary", str(info.get("formula_object_boundary_header") or ""))
        self.send_header("X-K12-Formula-Fallback-Formats", str(info.get("fallback_formats_header") or "MathML,LaTeX,image"))
        self.send_header("Content-Disposition", f'attachment; filename="{safe_name}"')
        self.end_headers()
        self.wfile.write(data)

    def _read_verified_installer(self, file_name: str, query: dict[str, list[str]] | None = None) -> tuple[dict, bytes]:
        """Read an installer snapshot and recheck its descriptor size and hash."""
        try:
            requested_platform = (query or {}).get("platform", [None])[0]
            info = self.processor.installer_download_info(file_name, requested_platform)
        except FileNotFoundError as exc:
            raise JsonError(404, "Installer not found") from exc
        except ValueError as exc:
            raise JsonError(404, str(exc)) from exc
        path = Path(info["path"])
        try:
            data = path.read_bytes()
        except OSError as exc:
            raise JsonError(404, "Installer not found") from exc
        if len(data) != int(info.get("file_size") or 0) or hashlib.sha256(data).hexdigest() != str(info.get("sha256") or ""):
            raise JsonError(404, "Installer content changed after validation")
        return info, data

    def _send_files_bundle(self, query: dict[str, list[str]]) -> None:
        """Build a ZIP of selected source files plus a manifest of successes and failures."""
        ids = self._query_ids(query)
        file_type = query.get("type", [""])[0]
        if ids:
            files = [self.store.get_file(file_id) for file_id in ids]
            files = [file for file in files if file]
        else:
            files = self.store.list_files()
        if file_type and file_type != "all":
            files = [file for file in files if file.get("file_type") == file_type]
        buffer = BytesIO()
        added: set[str] = set()
        rows = ["file_id,file_name,file_type,status,message,zip_path,size"]
        with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
            for index, file in enumerate(files, start=1):
                try:
                    info, source_data = self._read_verified_source(file["id"])
                    zip_path = self._unique_zip_name(f"files/{index:03d}-{Path(info['file_name']).name}", added)
                    archive.writestr(zip_path, source_data)
                    rows.append(
                        ",".join(
                            self._csv_cell(value)
                            for value in [file["id"], info["file_name"], file.get("file_type", ""), "成功", "", zip_path, info["file_size"]]
                        )
                    )
                except (KeyError, ValueError, OSError, JsonError) as exc:
                    message = exc.message if isinstance(exc, JsonError) else str(exc)
                    rows.append(
                        ",".join(
                            self._csv_cell(value)
                            for value in [file.get("id", ""), file.get("file_name", ""), file.get("file_type", ""), "失败", message, "", 0]
                        )
                    )
            archive.writestr("manifest.csv", "\n".join(rows) + "\n")
        data = buffer.getvalue()
        name = f"k12-{file_type}-files.zip" if file_type and file_type != "all" else "k12-files.zip"
        self.send_response(200)
        self.send_header("Content-Type", "application/zip")
        self._send_common_headers()
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Content-Disposition", f'attachment; filename="{self._attachment_name(name)}"')
        self.end_headers()
        self.wfile.write(data)

    def _send_image_asset(self, asset_path: str) -> None:
        """Serve a registered report-image snapshot from the managed image root."""
        target, data = self._read_verified_image_asset(asset_path)
        content_type = mimetypes.guess_type(str(target))[0] or "application/octet-stream"
        self.send_response(200)
        self.send_header("Content-Type", content_type)
        self._send_common_headers()
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def _read_verified_image_asset(self, asset_path: str) -> tuple[Path, bytes]:
        """Read a report image only when its record size and hash match the snapshot."""
        root = self.store.images_dir.resolve()
        target = (self.store.data_dir / asset_path).resolve()
        try:
            target.relative_to(root)
        except ValueError as exc:
            raise JsonError(404, "Image not found") from exc
        if not target.exists() or not target.is_file():
            raise JsonError(404, "Image not found")
        image = next(
            (
                item
                for report in self.store.list_reports()
                for item in (report.get("analysis") or {}).get("smallImages") or []
                if Path(str(item.get("image_path") or "")).resolve() == target
            ),
            {},
        )
        declared_hash = str(image.get("image_hash") or "")
        if image.get("image_size", "") == "" or not re.fullmatch(r"[0-9a-f]{64}", declared_hash):
            raise JsonError(404, "Image not found")
        try:
            data = target.read_bytes()
        except OSError as exc:
            raise JsonError(404, "Image not found") from exc
        if len(data) != int(image["image_size"]) or hashlib.sha256(data).hexdigest() != declared_hash:
            raise JsonError(404, "Image not found")
        return target, data

    def _send_replacement_asset(self, file_name: str) -> None:
        """Serve one registered replacement-image snapshot from the upload cache."""
        safe_name = Path(file_name).name
        root = self.store.uploads_dir.resolve()
        target = (root / safe_name).resolve()
        try:
            target.relative_to(root)
        except ValueError as exc:
            raise JsonError(404, "Replacement image not found") from exc
        if not target.exists() or not target.is_file():
            raise JsonError(404, "Replacement image not found")
        annotation = next(
            (item for item in self.store.list_image_annotations() if item.get("replacement_storage_name") == safe_name),
            {},
        )
        declared_hash = str(annotation.get("replacement_sha256") or "")
        if annotation.get("replacement_file_size", "") == "" or not re.fullmatch(r"[0-9a-f]{64}", declared_hash):
            raise JsonError(404, "Replacement image not found")
        data = target.read_bytes()
        if len(data) != int(annotation["replacement_file_size"]) or hashlib.sha256(data).hexdigest() != declared_hash:
            raise JsonError(404, "Replacement image not found")
        content_type = mimetypes.guess_type(str(target))[0] or "application/octet-stream"
        self.send_response(200)
        self.send_header("Content-Type", content_type)
        self._send_common_headers()
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def _send_artifact(self, path: str) -> None:
        """Serve successful task artifacts only when a report registered the output."""
        parts = path.strip("/").split("/")
        if len(parts) < 3:
            raise JsonError(404, "Artifact not found")
        task_id = parts[2]
        file_name = Path("/".join(parts[3:])).name
        target, data = self._read_verified_artifact(task_id, file_name)
        content_type = mimetypes.guess_type(str(target))[0] or "application/octet-stream"
        self.send_response(200)
        self.send_header("Content-Type", content_type)
        self._send_common_headers()
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Content-Disposition", f'attachment; filename="{self._attachment_name(target.name)}"')
        self.end_headers()
        self.wfile.write(data)

    def _read_verified_artifact(self, task_id: str, file_name: str) -> tuple[Path, bytes]:
        """Read an artifact snapshot and recheck its registered size and hash."""
        target = self._find_artifact_path(task_id, file_name)
        reports = [report for report in self.store.list_reports() if report.get("task_id") == task_id]
        latest_report = reports[0] if reports else {}
        artifact = next(
            (
                item
                for item in latest_report.get("analysis", {}).get("artifacts", [])
                if item.get("status") == "成功" and item.get("file_name") == file_name
            ),
            {},
        )
        try:
            data = target.read_bytes()
            registered_size = int(artifact.get("size"))
        except (OSError, TypeError, ValueError) as exc:
            raise JsonError(404, "Artifact not found") from exc
        registered_sha256 = str(artifact.get("sha256") or "")
        if len(data) != registered_size or hashlib.sha256(data).hexdigest() != registered_sha256:
            raise JsonError(404, "Artifact not found")
        return target, data

    def _find_artifact_path(self, task_id: str, file_name: str) -> Path:
        """Find a successful artifact from the latest report without trusting URL paths."""
        output_root = self.store.output_task_dir(task_id).resolve()
        reports = [report for report in self.store.list_reports() if report.get("task_id") == task_id]
        latest_report = reports[0] if reports else {}
        for artifact in latest_report.get("analysis", {}).get("artifacts", []):
            candidate = Path(artifact.get("path") or "")
            artifact_task_id = str(artifact.get("task_id") or task_id)
            if (
                artifact_task_id == task_id
                and artifact.get("status") == "成功"
                and artifact.get("file_name") == file_name
                and candidate.exists()
                and candidate.is_file()
            ):
                resolved = candidate.resolve()
                try:
                    resolved.relative_to(output_root)
                except ValueError:
                    continue
                if resolved.name != file_name:
                    continue
                output_type = str(artifact.get("output_type") or resolved.suffix.lstrip("."))
                if self.processor._output_artifact_validation_error(resolved, output_type):
                    continue
                declared_size = artifact.get("size", "")
                declared_sha256 = str(artifact.get("sha256") or "")
                if declared_size == "" or not re.fullmatch(r"[0-9a-f]{64}", declared_sha256):
                    continue
                if int(declared_size) != resolved.stat().st_size:
                    continue
                if self.processor._sha256(resolved) != declared_sha256:
                    continue
                return resolved
        raise JsonError(404, "Artifact not found")

    def _send_task_bundle(self, task_id: str) -> None:
        """Build a task result ZIP containing reports, logs, and registered outputs."""
        task = self.store.get_task(task_id)
        if not task:
            raise JsonError(404, "Task not found")
        reports = [report for report in self.store.list_reports() if report.get("task_id") == task_id]
        buffer = BytesIO()
        added: set[str] = set()
        with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
            archive.writestr("task.json", json.dumps(self._public_payload(task), ensure_ascii=False, indent=2))
            logs = self.store.list_logs(task_id, limit=1000)
            archive.writestr("logs/task.log", "\n".join(f"{item['created_at']} [{item.get('category', 'system')}/{item['level']}] {self._public_log(item)['message']}" for item in reversed(logs)) + "\n")
            for report in reports:
                for key in ("html_path", "json_path", "pdf_path", "xlsx_path", "txt_path"):
                    try:
                        report_path, report_data = self._read_managed_report(report, key, allow_fallback=False)
                    except JsonError:
                        continue
                    archive_name = f"reports/{key.removesuffix('_path')}/{report_path.name}"
                    if archive_name not in added:
                        archive.writestr(archive_name, report_data)
                        added.add(archive_name)
                try:
                    failure_path, failure_data = self._read_managed_report(report, "failure_csv_path", allow_fallback=False)
                except JsonError:
                    failure_path = None
                if failure_path:
                    archive_name = f"reports/failures/{failure_path.name}"
                    if archive_name not in added:
                        archive.writestr(archive_name, failure_data)
                        added.add(archive_name)
            latest_report = reports[0] if reports else {}
            for artifact in latest_report.get("analysis", {}).get("artifacts", []):
                if artifact.get("status") != "成功":
                    continue
                try:
                    artifact_path, artifact_data = self._read_verified_artifact(task_id, str(artifact.get("file_name") or ""))
                except JsonError:
                    continue
                archive_name = f"outputs/{artifact_path.name}"
                if archive_name not in added:
                    archive.writestr(archive_name, artifact_data)
                    added.add(archive_name)
        data = buffer.getvalue()
        self.send_response(200)
        self.send_header("Content-Type", "application/zip")
        self._send_common_headers()
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Content-Disposition", f'attachment; filename="{self._attachment_name(f"{task_id}-results.zip")}"')
        self.end_headers()
        self.wfile.write(data)

    def _managed_report_path(self, report: dict, key: str, allow_fallback: bool = True) -> Path:
        """Resolve one report file only when its path and registered integrity remain valid."""
        raw_path = report.get(key) or (report.get("report_path") if allow_fallback else "") or ""
        candidate = Path(str(raw_path))
        if not candidate.exists() or not candidate.is_file():
            raise JsonError(404, "Report file not found")
        resolved = candidate.resolve()
        try:
            resolved.relative_to(self.store.reports_dir.resolve())
        except ValueError as exc:
            raise JsonError(404, "Report file not found") from exc
        integrity = (report.get("report_file_integrity") or {}).get(key) or {}
        declared_sha256 = str(integrity.get("sha256") or "")
        if integrity.get("size", "") == "" or not re.fullmatch(r"[0-9a-f]{64}", declared_sha256):
            raise JsonError(404, "Report file not found")
        if int(integrity["size"]) != resolved.stat().st_size or self.processor._sha256(resolved) != declared_sha256:
            raise JsonError(404, "Report file not found")
        return resolved

    def _read_managed_report(self, report: dict, key: str, allow_fallback: bool = True) -> tuple[Path, bytes]:
        """Read a report snapshot and recheck its registered size and hash."""
        path = self._managed_report_path(report, key, allow_fallback=allow_fallback)
        integrity = (report.get("report_file_integrity") or {}).get(key) or {}
        try:
            data = path.read_bytes()
            declared_size = int(integrity.get("size"))
        except (OSError, TypeError, ValueError) as exc:
            raise JsonError(404, "Report file not found") from exc
        if len(data) != declared_size or hashlib.sha256(data).hexdigest() != str(integrity.get("sha256") or ""):
            raise JsonError(404, "Report file not found")
        return path, data

    def _send_report_images(self, report_id: str) -> None:
        """Build the small-image ZIP export with annotations and a manifest CSV."""
        report = self.store.get_report(report_id)
        if not report:
            raise JsonError(404, "Report not found")
        images = report.get("analysis", {}).get("smallImages", [])
        annotations = {item["image_id"]: item for item in self.processor.list_image_annotations(report_id)}
        buffer = BytesIO()
        added: set[str] = set()
        with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
            rows = ["image_id,file_id,source_name,location,width,height,area,image_format,suspected_type,confirmed_type,confidence,header_footer,watermark,transparent,duplicate,duplicate_check_status,duplicate_fallback,export_status,export_message,reexported_at,export_format,annotation,note,replacement_file,file"]
            for index, image in enumerate(images, start=1):
                image_path = Path(image.get("image_path") or "")
                file_name = ""
                if image_path.exists() and image_path.is_file():
                    file_name = f"images/{index:03d}-{image_path.name}"
                    if file_name not in added:
                        try:
                            asset_path = image_path.resolve().relative_to(self.store.data_dir.resolve()).as_posix()
                            _verified_path, image_data = self._read_verified_image_asset(asset_path)
                        except (JsonError, ValueError):
                            file_name = ""
                        else:
                            archive.writestr(file_name, image_data)
                            added.add(file_name)
                annotation = annotations.get(image.get("id"), {})
                rows.append(
                    ",".join(
                        self._csv_cell(value)
                        for value in [
                            image.get("id", ""),
                            image.get("file_id", ""),
                            image.get("source_name", ""),
                            image.get("location", ""),
                            image.get("width", ""),
                            image.get("height", ""),
                            image.get("area", ""),
                            image.get("image_type", ""),
                            self._small_image_kind(image),
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
                            file_name,
                        ]
                    )
                )
            archive.writestr("manifest.csv", "\n".join(rows) + "\n")
        data = buffer.getvalue()
        self.send_response(200)
        self.send_header("Content-Type", "application/zip")
        self._send_common_headers()
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Content-Disposition", f'attachment; filename="{self._attachment_name(f"{report_id}-images.zip")}"')
        self.end_headers()
        self.wfile.write(data)

    def _send_report_image_manifest_xlsx(self, report_id: str) -> None:
        """Stream the small-image manifest workbook for one report."""
        report = self.store.get_report(report_id)
        if not report:
            raise JsonError(404, "Report not found")
        annotations = self.processor.list_image_annotations(report_id)
        data = build_image_manifest_xlsx(report, annotations)
        self.send_response(200)
        self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        self._send_common_headers()
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Content-Disposition", f'attachment; filename="{self._attachment_name(f"{report_id}-images.xlsx")}"')
        self.end_headers()
        self.wfile.write(data)

    def _send_report_formulas(self, report_id: str) -> None:
        """Stream the formula ZIP export with MathML, TEX, JSON, and annotations."""
        report = self.store.get_report(report_id)
        if not report:
            raise JsonError(404, "Report not found")
        annotations = self.processor.list_formula_annotations(report_id)
        data = build_formula_zip(report, annotations)
        self.send_response(200)
        self.send_header("Content-Type", "application/zip")
        self._send_common_headers()
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Content-Disposition", f'attachment; filename="{self._attachment_name(f"{report_id}-formulas.zip")}"')
        self.end_headers()
        self.wfile.write(data)

    def _send_report_formula_xlsx(self, report_id: str) -> None:
        """Stream the formula workbook export for one report."""
        report = self.store.get_report(report_id)
        if not report:
            raise JsonError(404, "Report not found")
        annotations = self.processor.list_formula_annotations(report_id)
        data = build_formula_xlsx(report, annotations)
        self.send_response(200)
        self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        self._send_common_headers()
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Content-Disposition", f'attachment; filename="{self._attachment_name(f"{report_id}-formulas.xlsx")}"')
        self.end_headers()
        self.wfile.write(data)

    def _send_report_failures(self, report_id: str) -> None:
        """Stream the standard failure-list CSV for one report."""
        report = self.store.get_report(report_id)
        if not report:
            raise JsonError(404, "Report not found")
        path = Path(report.get("failure_csv_path") or "")
        data = path.read_bytes() if path.exists() else ReportBuilder._failure_csv(report.get("failureRows", [])).encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "text/csv; charset=utf-8")
        self._send_common_headers()
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Content-Disposition", f'attachment; filename="{self._attachment_name(f"{report_id}-failures.csv")}"')
        self.end_headers()
        self.wfile.write(data)

    def _send_report_omml_failures(self, report_id: str) -> None:
        """Stream OMML dependency and conversion failure rows with path redaction settings."""
        report = self.store.get_report(report_id)
        if not report:
            raise JsonError(404, "Report not found")
        annotations = self.processor.list_omml_annotations(report_id)
        expose_paths = bool(self.store.get_settings().get("exposeLocalPaths", False))
        data = build_omml_failure_csv(report, annotations, expose_paths=expose_paths).encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "text/csv; charset=utf-8")
        self._send_common_headers()
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Content-Disposition", f'attachment; filename="{self._attachment_name(f"{report_id}-omml-failures.csv")}"')
        self.end_headers()
        self.wfile.write(data)

    def _send_report_macro_failures(self, report_id: str) -> None:
        """Stream macro failure rows with backup and restore guidance."""
        report = self.store.get_report(report_id)
        if not report:
            raise JsonError(404, "Report not found")
        expose_paths = bool(self.store.get_settings().get("exposeLocalPaths", False))
        data = build_macro_failure_csv(report, expose_paths=expose_paths).encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "text/csv; charset=utf-8")
        self._send_common_headers()
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Content-Disposition", f'attachment; filename="{self._attachment_name(f"{report_id}-macro-failures.csv")}"')
        self.end_headers()
        self.wfile.write(data)

    def _send_logs_file(self, task_id: str | None = None, export_format: str | None = None) -> None:
        """Export redacted task or system logs in the requested safe format."""
        logs = self.store.list_logs(task_id, limit=5000)
        public_logs = [self._public_log(item) for item in reversed(logs)]
        export_format = self._log_export_format(export_format)
        data = self._log_export_bytes(public_logs, export_format)
        name = f"{task_id}-logs.{export_format}" if task_id else f"k12-logs.{export_format}"
        content_type = {
            "txt": "text/plain; charset=utf-8",
            "log": "text/plain; charset=utf-8",
            "csv": "text/csv; charset=utf-8",
            "json": "application/json; charset=utf-8",
        }[export_format]
        self.send_response(200)
        self.send_header("Content-Type", content_type)
        self._send_common_headers()
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Content-Disposition", f'attachment; filename="{self._attachment_name(name)}"')
        self.end_headers()
        self.wfile.write(data)

    def _log_export_format(self, requested: str | None) -> str:
        """Normalize log download formats to the supported extension set."""
        value = str(requested or self.store.get_settings().get("logExportFormat", "txt") or "txt").lower()
        return value if value in {"txt", "log", "csv", "json"} else "txt"

    def _log_export_bytes(self, logs: list[dict], export_format: str) -> bytes:
        """Render public log rows as JSON, CSV, or plain text bytes."""
        if export_format == "json":
            return json.dumps({"logs": logs}, ensure_ascii=False, indent=2).encode("utf-8")
        if export_format == "csv":
            rows = ["created_at,category,level,task_id,message"]
            for item in logs:
                rows.append(
                    ",".join(
                        self._csv_cell(value)
                        for value in [
                            item.get("created_at", ""),
                            item.get("category", "system"),
                            item.get("level", ""),
                            item.get("task_id", ""),
                            item.get("message", ""),
                        ]
                    )
                )
            return ("\n".join(rows) + "\n").encode("utf-8")
        lines = [
            f"{item['created_at']} [{item.get('category', 'system')}/{item['level']}] {item['task_id']} {item['message']}"
            for item in logs
        ]
        return ("\n".join(lines) + ("\n" if lines else "")).encode("utf-8")

    @staticmethod
    def _add_bundle_path(archive: zipfile.ZipFile, path: Path, folder: str, added: set[str]) -> None:
        """Add one existing file to a ZIP folder while avoiding duplicate names."""
        if not path.exists() or not path.is_file():
            return
        name = f"{folder}/{path.name}"
        if name in added:
            return
        archive.write(path, name)
        added.add(name)

    @staticmethod
    def _query_ids(query: dict[str, list[str]]) -> list[str]:
        """Return de-duplicated ids parsed from repeated or comma-separated query values."""
        ids: list[str] = []
        for raw in query.get("ids", []):
            ids.extend(item.strip() for item in raw.split(",") if item.strip())
        return list(dict.fromkeys(ids))

    @staticmethod
    def _unique_zip_name(name: str, added: set[str]) -> str:
        """Reserve a safe unique ZIP entry name under the requested folder."""
        safe = "/".join(Path(part).name for part in name.split("/") if part)
        path = Path(safe)
        candidate = safe
        index = 1
        while candidate in added:
            candidate = f"{path.parent}/{path.stem}-{index}{path.suffix}" if str(path.parent) != "." else f"{path.stem}-{index}{path.suffix}"
            index += 1
        added.add(candidate)
        return candidate

    @staticmethod
    def _csv_cell(value: object) -> str:
        """Quote one CSV cell without relying on platform-specific csv dialect state."""
        text = str(value).replace('"', '""')
        return f'"{text}"'

    @staticmethod
    def _small_image_kind(image: dict) -> str:
        """Classify a small image for exports using the PRD review categories."""
        if image.get("is_formula_like"):
            return "公式"
        if image.get("is_qrcode_like"):
            return "二维码"
        if image.get("is_stamp_like"):
            return "印章"
        if image.get("is_signature_like"):
            return "签名"
        return "图标"

    def _public_payload(self, value: object) -> object:
        """Return API payloads with local filesystem paths redacted by default."""
        if self.store.get_settings().get("exposeLocalPaths", False):
            return value
        if isinstance(value, list):
            return [self._public_payload(item) for item in value]
        if isinstance(value, dict):
            redacted: dict[str, object] = {}
            for key, item in value.items():
                if key in LOCAL_PATH_KEYS:
                    redacted[f"{key}_available"] = bool(item)
                    redacted[key] = self._redacted_path_label(item)
                elif key == "outputs" and isinstance(item, dict):
                    redacted[key] = {name: self._redacted_path_label(path) for name, path in item.items()}
                    redacted["outputs_available"] = {name: bool(path) for name, path in item.items()}
                else:
                    redacted[key] = self._public_payload(item)
            if "backup_path" in value:
                redacted["backup_available"] = bool(value.get("backup_path"))
            return redacted
        if isinstance(value, str):
            return self._redact_text_paths(value)
        return value

    def _public_settings(self, settings: dict) -> dict:
        """Return settings safe for browser reads without exposing local secrets."""
        public = self._public_payload(dict(settings))
        if not isinstance(public, dict):
            public = dict(settings)
        token_configured = bool(str(settings.get("localSecurityToken") or "").strip())
        public["localSecurityTokenConfigured"] = token_configured
        public["localSecurityToken"] = ""
        return public

    def _public_log(self, log: dict) -> dict:
        """Return one log row with local paths redacted unless explicitly allowed."""
        if self.store.get_settings().get("exposeLocalPaths", False):
            return dict(log)
        item = dict(log)
        item["message"] = self._redact_text_paths(str(item.get("message") or ""))
        return item

    @staticmethod
    def _redacted_path_label(value: object) -> str:
        """Replace a local path with a stable filename-only redaction label."""
        if not value:
            return ""
        name = Path(str(value)).name
        return f"本地路径已隐藏/{name}" if name else "本地路径已隐藏"

    @classmethod
    def _redact_text_paths(cls, text: str) -> str:
        """Redact common macOS, Linux, and Windows absolute paths from log text."""
        pattern = r"(/Users/[^\s，,;]+|/private/[^\s，,;]+|/var/folders/[^\s，,;]+|/tmp/[^\s，,;]+|[A-Za-z]:\\[^\s，,;]+)"
        return re.sub(pattern, lambda match: cls._redacted_path_label(match.group(0)), text)

    @staticmethod
    def _attachment_name(value: object) -> str:
        """Keep Content-Disposition filenames on one safe header line."""
        name = Path(str(value or "download")).name or "download"
        return name.replace("\r", "_").replace("\n", "_").replace('"', "_")

    def _json(self, payload: dict, status: int = 200) -> None:
        """Send one JSON response with common CORS headers."""
        data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self._send_common_headers()
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def _send_common_headers(self) -> None:
        """Attach CORS headers used by the local browser UI and companion client."""
        self.send_header("Access-Control-Allow-Origin", self._cors_origin())
        self.send_header("Access-Control-Allow-Methods", "GET, POST, PUT, DELETE, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type, X-K12-Token, Authorization")

    def _cors_origin(self) -> str:
        """Allow only localhost-style origins for browser access to the local API."""
        origin = self.headers.get("Origin") or ""
        if re.fullmatch(r"http://(127\.0\.0\.1|localhost)(:\d+)?", origin):
            return origin
        return "http://127.0.0.1"

    def log_message(self, format: str, *args: object) -> None:
        """Silence default HTTP logging so task logs remain the source of truth."""
        return


class K12Server(ThreadingHTTPServer):
    """Threaded local server that owns one AppStore and TaskProcessor."""

    def __init__(self, server_address: tuple[str, int], handler_class: type[K12RequestHandler], data_dir: Path) -> None:
        """Create the runtime store, processor, recovery pass, and cleanup pass."""
        super().__init__(server_address, handler_class)
        self.store = AppStore(data_dir)
        self.processor = TaskProcessor(self.store)
        self.processor.recover_interrupted_tasks()
        try:
            self.processor.auto_cleanup_runtime_history()
        except Exception as exc:  # pragma: no cover
            self.store.append_log("system", f"自动清理失败：{exc}", "error")


def main() -> None:
    """Start the local K12 HTTP server with a configurable runtime data dir."""
    parser = argparse.ArgumentParser(description="Run the K12 local document processing workbench.")
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", default=8765, type=int)
    parser.add_argument("--data-dir", default=".k12-data")
    args = parser.parse_args()

    server = K12Server((args.host, args.port), K12RequestHandler, Path(args.data_dir))
    print(f"K12 running at http://{args.host}:{args.port}")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nK12 stopped")


if __name__ == "__main__":
    main()
