"""SQLite-backed runtime store for the K12 local-first workbench.

The store keeps uploads, reports, outputs, settings, users, authorizations, and
annotation records under one data directory. Passwords and sensitive local paths
are handled as runtime data; public redaction happens at the API layer.
"""

from __future__ import annotations

import json
import shutil
import sqlite3
import threading
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any

from .models import new_id, utc_now


DEFAULT_SETTINGS: dict[str, Any] = {
    "outputDirectory": "outputs",
    "outputConflictStrategy": "自动重命名",
    "maxConcurrentTasks": 3,
    "singleFileLimitMb": 500,
    "minFreeDiskMb": 512,
    "keepOriginalFile": True,
    "autoOpenOutputDirectory": False,
    "generateReport": True,
    "enableTaskCompletionNotice": True,
    "saveHistory": True,
    "verboseLogging": True,
    "logUploadEvents": True,
    "logConversionEvents": True,
    "logFormulaEvents": True,
    "logOmmlEvents": True,
    "logMacroEvents": True,
    "logImageEvents": True,
    "logErrorEvents": True,
    "logExportFormat": "txt",
    "autoCleanTemp": False,
    "cleanupRetentionDays": 30,
    "logRetentionDays": 30,
    "cleanupIntervalDays": 1,
    "lastCleanupAt": "",
    "duplicateFileStrategy": "自动重命名",
    "wordToPptTemplate": "教学讲义默认模板",
    "pptToWordMode": "逐页讲义模式",
    "pptToWordTemplate": "默认 Word 讲义模板",
    "pptToWordTemplatePath": "",
    "pptToWordGenerateToc": True,
    "pdfPrecisionMode": "平衡",
    "excelConversionRange": "全部工作表",
    "excelFormulaMode": "保留公式",
    "excelSplitSheets": False,
    "retainImages": True,
    "retainTables": True,
    "retainHeadersFooters": False,
    "retainFootnotesEndnotes": True,
    "retainComments": False,
    "retainRevisions": False,
    "pdfToWordEngine": "Mathpix",
    "allowExternalMathpixUpload": False,
    "mathpixAppIdEnv": "MATHPIX_APP_ID",
    "mathpixAppKeyEnv": "MATHPIX_APP_KEY",
    "mathpixPollTimeoutSeconds": 600,
    "waitForMathpix": False,
    "enableFormulaOcr": True,
    "enableMathType": True,
    "enableMathTypeFormatting": True,
    "formulaOutputFormat": "LaTeX+MathML",
    "formulaFont": "Cambria Math",
    "formulaFontSize": 12,
    "formulaFormatScope": "全文",
    "formulaAlignment": "居中",
    "formulaVariableStyle": "斜体",
    "formulaFunctionStyle": "正体",
    "formulaScriptScale": 70,
    "formulaFractionStyle": "标准",
    "formulaRadicalStyle": "标准",
    "formulaMatrixSpacing": "标准",
    "formulaGreekStyle": "标准",
    "formulaInlineBaseline": "跟随正文",
    "formulaDisplaySpacing": "标准",
    "formulaNumbering": "按文档位置",
    "formulaConfidenceThreshold": 80,
    "lowConfidenceFormulaStrategy": "人工确认",
    "keepFormulaImages": True,
    "generateFormulaReport": True,
    "enableOmmlPrecheck": True,
    "autoOmmlToMathType": False,
    "autoSearchOmml": True,
    "allowManualOmml": True,
    "manualOmmlPath": "",
    "ommlSearchPaths": "",
    "ommlSearchMaxFiles": 3000,
    "ommlCopyStrategy": "自动重命名",
    "enableMacroDetection": True,
    "enableMacroExecution": True,
    "allowDocumentMacros": True,
    "allowTemplateMacros": True,
    "allowLocalMacroLibrary": True,
    "macroWhitelistOnly": False,
    "macroWhitelist": "",
    "macroFailureStrategy": "跳过",
    "macroBackup": True,
    "macroTimeoutSeconds": 120,
    "allowBatchMacroExecution": True,
    "smallImageMaxWidth": 96,
    "smallImageMaxHeight": 96,
    "smallImageMaxArea": 9216,
    "includeHeaderFooterImages": True,
    "includeWatermarkImages": True,
    "includeTransparentImages": True,
    "includeDuplicateImages": True,
    "imageExportFormat": "原格式",
    "ocrLanguage": "中文+英文",
    "enableTextOcr": True,
    "enableTableOcr": True,
    "ocrPrecisionMode": "平衡",
    "ocrSpeedMode": "标准",
    "localClientEnabled": True,
    "localClientPlatform": "auto",
    "localApiHost": "127.0.0.1",
    "localApiPort": 8765,
    "localSecurityToken": "",
    "allowWebLaunchLocalClient": False,
    "allowTaskStatusCloudSync": False,
    "exposeLocalPaths": False,
    "mathtypeCompatibilityMode": "platform-specific",
    "localClientHeartbeat": {},
    "allowCloudSync": False,
    "sensitiveFilesPreferLocal": True,
    "activeUserId": "user_admin",
    "requireLogin": False,
}


ROLE_PERMISSIONS: dict[str, list[str]] = {
    "管理员": [
        "files.manage",
        "tasks.create",
        "tasks.control",
        "reports.manage",
        "templates.manage",
        "users.manage",
        "authorizations.manage",
        "settings.manage",
        "macros.execute",
        "cloud.sync",
    ],
    "教师": [
        "files.manage",
        "tasks.create",
        "tasks.control",
        "reports.manage",
        "templates.manage",
        "macros.execute",
    ],
    "学生": [
        "files.manage",
        "tasks.create",
        "reports.manage",
    ],
    "访客": [],
}


DEFAULT_AUTHORIZATIONS: dict[str, dict[str, Any]] = {
    "mathpix_external_upload": {
        "name": "Mathpix 外部上传",
        "scope": "PDF 转 Word 与 PDF 公式识别",
        "setting_key": "allowExternalMathpixUpload",
        "enabled": False,
        "risk_level": "高",
        "description": "允许将 PDF 上传到 Mathpix 进行 OCR、DOCX 和公式识别。",
    },
    "web_launch_local_client": {
        "name": "网页唤起本地客户端",
        "scope": "混合模式任务分流",
        "setting_key": "allowWebLaunchLocalClient",
        "enabled": False,
        "risk_level": "中",
        "description": "允许网页端发起本地客户端唤起请求。",
    },
    "cloud_sync": {
        "name": "结果上传云端",
        "scope": "混合模式结果同步",
        "setting_key": "allowCloudSync",
        "enabled": False,
        "risk_level": "高",
        "description": "允许将处理结果或任务摘要同步到云端。",
    },
    "task_status_cloud_sync": {
        "name": "任务状态同步云端",
        "scope": "任务中心状态同步",
        "setting_key": "allowTaskStatusCloudSync",
        "enabled": False,
        "risk_level": "中",
        "description": "允许网页端同步本地任务状态到云端任务中心。",
    },
    "macro_execution": {
        "name": "宏执行队列",
        "scope": "Word 宏顺序执行",
        "setting_key": "enableMacroExecution",
        "enabled": True,
        "risk_level": "高",
        "description": "允许已确认的宏进入本地客户端执行队列。",
    },
    "expose_local_paths": {
        "name": "显示本地路径",
        "scope": "报告、日志和文件详情",
        "setting_key": "exposeLocalPaths",
        "enabled": False,
        "risk_level": "中",
        "description": "允许 API 响应展示完整本地路径。",
    },
}


class AppStore:
    """Manage local runtime directories and SQLite records for one workspace."""

    def __init__(self, data_dir: Path | str = ".k12-data") -> None:
        self.data_dir = Path(data_dir)
        self.data_dir.mkdir(parents=True, exist_ok=True)
        self.reports_dir = self.data_dir / "reports"
        self.reports_dir.mkdir(parents=True, exist_ok=True)
        self.outputs_dir = self.data_dir / "outputs"
        self.outputs_dir.mkdir(parents=True, exist_ok=True)
        self.uploads_dir = self.data_dir / "uploads"
        self.uploads_dir.mkdir(parents=True, exist_ok=True)
        self.backups_dir = self.data_dir / "backups"
        self.backups_dir.mkdir(parents=True, exist_ok=True)
        self.images_dir = self.data_dir / "images"
        self.images_dir.mkdir(parents=True, exist_ok=True)
        self.installers_dir = self.data_dir / "installers"
        self.installers_dir.mkdir(parents=True, exist_ok=True)
        self.cloud_uploads_dir = self.data_dir / "cloud_uploads"
        self.cloud_uploads_dir.mkdir(parents=True, exist_ok=True)
        self.db_path = self.data_dir / "k12.sqlite3"
        self._file_passwords: dict[str, str] = {}
        self._lock = threading.RLock()
        self._init_db()
        self._reset_password_sessions()

    def connect(self) -> sqlite3.Connection:
        """Open a SQLite connection with row objects for store methods."""
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        return conn

    def _init_db(self) -> None:
        with self._lock, self.connect() as conn:
            conn.executescript(
                """
                create table if not exists files (
                    id text primary key,
                    payload text not null,
                    created_at text not null
                );
                create table if not exists tasks (
                    id text primary key,
                    payload text not null,
                    created_at text not null,
                    updated_at text not null
                );
                create table if not exists reports (
                    id text primary key,
                    task_id text not null,
                    payload text not null,
                    created_at text not null
                );
                create table if not exists logs (
                    id integer primary key autoincrement,
                    task_id text not null,
                    message text not null,
                    level text not null,
                    category text not null default 'system',
                    created_at text not null
                );
                create table if not exists settings (
                    key text primary key,
                    value text not null,
                    updated_at text not null
                );
                create table if not exists macro_templates (
                    id text primary key,
                    payload text not null,
                    created_at text not null,
                    updated_at text not null
                );
                create table if not exists users (
                    id text primary key,
                    payload text not null,
                    created_at text not null,
                    updated_at text not null
                );
                create table if not exists templates (
                    id text primary key,
                    payload text not null,
                    created_at text not null,
                    updated_at text not null
                );
                create table if not exists authorizations (
                    key text primary key,
                    payload text not null,
                    created_at text not null,
                    updated_at text not null
                );
                create table if not exists image_annotations (
                    id text primary key,
                    image_id text not null,
                    report_id text not null,
                    payload text not null,
                    created_at text not null,
                    updated_at text not null
                );
                create table if not exists formula_annotations (
                    id text primary key,
                    formula_id text not null,
                    report_id text not null,
                    payload text not null,
                    created_at text not null,
                    updated_at text not null
                );
                create table if not exists omml_annotations (
                    id text primary key,
                    dependency_id text not null,
                    report_id text not null,
                    payload text not null,
                    created_at text not null,
                    updated_at text not null
                );
                create table if not exists layout_annotations (
                    id text primary key,
                    report_id text not null,
                    location text not null,
                    payload text not null,
                    created_at text not null,
                    updated_at text not null
                );
                """
            )
            self._ensure_column(conn, "logs", "category", "text not null default 'system'")
            for key, value in DEFAULT_SETTINGS.items():
                conn.execute(
                    "insert or ignore into settings (key, value, updated_at) values (?, ?, ?)",
                    (key, json.dumps(value, ensure_ascii=False), utc_now()),
                )
            self._ensure_default_users(conn)
            self._ensure_default_authorizations(conn)

    @staticmethod
    def _ensure_column(conn: sqlite3.Connection, table: str, column: str, definition: str) -> None:
        columns = {row["name"] for row in conn.execute(f"pragma table_info({table})").fetchall()}
        if column not in columns:
            conn.execute(f"alter table {table} add column {column} {definition}")

    @staticmethod
    def _ensure_default_users(conn: sqlite3.Connection) -> None:
        count = conn.execute("select count(*) as count from users").fetchone()["count"]
        if count:
            return
        now = utc_now()
        admin = {
            "id": "user_admin",
            "name": "本地管理员",
            "role": "管理员",
            "status": "启用",
            "permissions": ROLE_PERMISSIONS["管理员"],
            "login_enabled": True,
            "last_login_at": "",
            "created_at": now,
            "updated_at": now,
        }
        conn.execute(
            "insert into users (id, payload, created_at, updated_at) values (?, ?, ?, ?)",
            (admin["id"], json.dumps(admin, ensure_ascii=False), now, now),
        )

    @staticmethod
    def _ensure_default_authorizations(conn: sqlite3.Connection) -> None:
        now = utc_now()
        for key, value in DEFAULT_AUTHORIZATIONS.items():
            payload = {
                "key": key,
                **value,
                "status": "已授权" if value.get("enabled") else "未授权",
                "note": "",
                "updated_by": "system",
                "created_at": now,
                "updated_at": now,
            }
            conn.execute(
                "insert or ignore into authorizations (key, payload, created_at, updated_at) values (?, ?, ?, ?)",
                (key, json.dumps(payload, ensure_ascii=False), now, now),
            )

    def _log_category_enabled(self, category: str) -> bool:
        setting_key = {
            "upload": "logUploadEvents",
            "conversion": "logConversionEvents",
            "formula": "logFormulaEvents",
            "omml": "logOmmlEvents",
            "macro": "logMacroEvents",
            "image": "logImageEvents",
            "error": "logErrorEvents",
        }.get(category)
        if not setting_key:
            return True
        return bool(self.get_settings().get(setting_key, True))

    def save_file(self, payload: dict[str, Any]) -> dict[str, Any]:
        """Persist one file payload exactly as analyzed by the processor."""
        with self._lock, self.connect() as conn:
            conn.execute(
                "insert or replace into files (id, payload, created_at) values (?, ?, ?)",
                (payload["id"], json.dumps(payload, ensure_ascii=False), payload["created_at"]),
            )
        return payload

    def replace_file(self, file_id: str, payload: dict[str, Any]) -> dict[str, Any]:
        """Replace one uploaded file record and remove old cached upload entries."""
        with self._lock, self.connect() as conn:
            rows = conn.execute("select payload from files").fetchall()
            all_payloads = [json.loads(row["payload"]) for row in rows]
            old_payloads = [
                item
                for item in all_payloads
                if item.get("id") == file_id or item.get("archive_parent_id") == file_id
            ]
            if not old_payloads:
                raise KeyError(f"File not found: {file_id}")
            previous = next(item for item in old_payloads if item.get("id") == file_id)
            for item in old_payloads:
                conn.execute("delete from files where id = ?", (item["id"],))
                self._file_passwords.pop(item["id"], None)
            payload["id"] = file_id
            payload["created_at"] = previous.get("created_at", payload.get("created_at") or utc_now())
            payload["replaced_at"] = utc_now()
            payload["archive_parent_id"] = ""
            payload["source_kind"] = "upload"
            conn.execute(
                "insert or replace into files (id, payload, created_at) values (?, ?, ?)",
                (payload["id"], json.dumps(payload, ensure_ascii=False), payload["created_at"]),
            )
        self._cleanup_uploaded_payloads(old_payloads)
        return payload

    def get_file(self, file_id: str) -> dict[str, Any] | None:
        """Return one stored file payload by id."""
        with self._lock, self.connect() as conn:
            row = conn.execute("select payload from files where id = ?", (file_id,)).fetchone()
        return json.loads(row["payload"]) if row else None

    def list_files(self) -> list[dict[str, Any]]:
        """Return stored files newest first."""
        with self._lock, self.connect() as conn:
            rows = conn.execute("select payload from files order by created_at desc").fetchall()
        return [json.loads(row["payload"]) for row in rows]

    def delete_file(self, file_id: str) -> bool:
        """Delete a file record plus archive children and managed upload cache."""
        payloads: list[dict[str, Any]] = []
        with self._lock, self.connect() as conn:
            rows = conn.execute("select payload from files").fetchall()
            all_payloads = [json.loads(row["payload"]) for row in rows]
            payloads = [
                payload
                for payload in all_payloads
                if payload.get("id") == file_id or payload.get("archive_parent_id") == file_id
            ]
            if not payloads:
                return False
            ids = [payload["id"] for payload in payloads]
            for item_id in ids:
                conn.execute("delete from files where id = ?", (item_id,))
                self._file_passwords.pop(item_id, None)
        self._cleanup_uploaded_payloads(payloads)
        return True

    def set_file_password(self, file_id: str, password: str) -> dict[str, Any]:
        """Store an encrypted-file password only for the current local process."""
        if not password:
            raise ValueError("密码不能为空")
        payload = self.get_file(file_id)
        if not payload:
            raise ValueError("文件不存在")
        self._file_passwords[file_id] = password
        payload["password_provided"] = True
        payload["password_session_active"] = True
        payload["password_updated_at"] = utc_now()
        payload["validation_errors"] = [
            error
            for error in list(payload.get("validation_errors") or [])
            if "加密" not in str(error) and "密码" not in str(error)
        ]
        summary = dict(payload.get("content_summary") or {})
        summary["passwordStatus"] = "已输入，仅当前本地会话有效"
        payload["content_summary"] = summary
        if not payload["validation_errors"] and payload.get("status") == "校验失败":
            payload["status"] = "待处理"
        return self.save_file(payload)

    def has_file_password(self, file_id: str) -> bool:
        """Return whether a runtime-only password exists for the file."""
        return file_id in self._file_passwords

    def save_task(self, payload: dict[str, Any]) -> dict[str, Any]:
        """Persist a task payload while refreshing duration metadata."""
        now = utc_now()
        payload = self._task_with_duration(payload, now)
        with self._lock, self.connect() as conn:
            conn.execute(
                """
                insert or replace into tasks (id, payload, created_at, updated_at)
                values (?, ?, coalesce((select created_at from tasks where id = ?), ?), ?)
                """,
                (payload["id"], json.dumps(payload, ensure_ascii=False), payload["id"], payload.get("created_at", now), now),
            )
        return payload

    @staticmethod
    def _task_with_duration(payload: dict[str, Any], now: str) -> dict[str, Any]:
        enriched = dict(payload)
        start = AppStore._parse_timestamp(enriched.get("start_time"))
        end = AppStore._parse_timestamp(enriched.get("end_time")) or AppStore._parse_timestamp(now)
        if start and end and end >= start:
            seconds = int((end - start).total_seconds())
        else:
            seconds = 0
        enriched["duration_seconds"] = seconds
        enriched["duration_label"] = AppStore._duration_label(seconds) if start else "-"
        return enriched

    @staticmethod
    def _parse_timestamp(value: Any) -> datetime | None:
        if not value:
            return None
        try:
            parsed = datetime.fromisoformat(str(value).replace("Z", "+00:00"))
        except ValueError:
            return None
        if parsed.tzinfo is None:
            return parsed.replace(tzinfo=timezone.utc)
        return parsed

    @staticmethod
    def _duration_label(seconds: int) -> str:
        seconds = max(0, int(seconds or 0))
        hours, remainder = divmod(seconds, 3600)
        minutes, secs = divmod(remainder, 60)
        if hours:
            return f"{hours}小时{minutes}分{secs}秒"
        if minutes:
            return f"{minutes}分{secs}秒"
        return f"{secs}秒"

    def get_task(self, task_id: str) -> dict[str, Any] | None:
        """Return one task payload with current duration metadata."""
        with self._lock, self.connect() as conn:
            row = conn.execute("select payload from tasks where id = ?", (task_id,)).fetchone()
        return self._task_with_duration(json.loads(row["payload"]), utc_now()) if row else None

    def list_tasks(self) -> list[dict[str, Any]]:
        """Return all tasks newest-updated first with current duration metadata."""
        with self._lock, self.connect() as conn:
            rows = conn.execute("select payload from tasks order by updated_at desc").fetchall()
        now = utc_now()
        return [self._task_with_duration(json.loads(row["payload"]), now) for row in rows]

    def delete_task_history(self, task_id: str, delete_outputs: bool = True) -> dict[str, Any]:
        """Delete a task, its reports, annotations, logs, and managed runtime files."""
        reports: list[dict[str, Any]] = []
        with self._lock, self.connect() as conn:
            rows = conn.execute("select payload from reports where task_id = ?", (task_id,)).fetchall()
            reports = [json.loads(row["payload"]) for row in rows]
            report_ids = [report["id"] for report in reports]
            if report_ids:
                placeholders = self._sql_placeholders(report_ids)
                conn.execute(f"delete from image_annotations where report_id in ({placeholders})", tuple(report_ids))
                conn.execute(f"delete from formula_annotations where report_id in ({placeholders})", tuple(report_ids))
                conn.execute(f"delete from omml_annotations where report_id in ({placeholders})", tuple(report_ids))
                conn.execute(f"delete from layout_annotations where report_id in ({placeholders})", tuple(report_ids))
                conn.execute(f"delete from reports where id in ({placeholders})", tuple(report_ids))
            conn.execute("delete from logs where task_id = ?", (task_id,))
            cur = conn.execute("delete from tasks where id = ?", (task_id,))
        report_files_deleted = self._cleanup_report_files(reports)
        image_files_deleted = sum(self._cleanup_report_image_files(report) for report in reports)
        dirs_deleted = self._cleanup_runtime_dirs({task_id}) if delete_outputs else {"output_dirs_deleted": 0, "backup_dirs_deleted": 0, "image_dirs_deleted": 0}
        return {
            "deleted": cur.rowcount > 0,
            "reports_deleted": len(reports),
            "report_files_deleted": report_files_deleted,
            "image_files_deleted": image_files_deleted,
            **dirs_deleted,
        }

    def save_report(self, payload: dict[str, Any]) -> dict[str, Any]:
        """Persist one report payload and its task association."""
        with self._lock, self.connect() as conn:
            conn.execute(
                "insert or replace into reports (id, task_id, payload, created_at) values (?, ?, ?, ?)",
                (payload["id"], payload["task_id"], json.dumps(payload, ensure_ascii=False), payload["created_at"]),
            )
        return payload

    def get_report(self, report_id: str) -> dict[str, Any] | None:
        """Return one report payload by id."""
        with self._lock, self.connect() as conn:
            row = conn.execute("select payload from reports where id = ?", (report_id,)).fetchone()
        return json.loads(row["payload"]) if row else None

    def list_reports(self) -> list[dict[str, Any]]:
        """Return stored reports newest first."""
        with self._lock, self.connect() as conn:
            rows = conn.execute("select payload from reports order by created_at desc, rowid desc").fetchall()
        return [json.loads(row["payload"]) for row in rows]

    def delete_report(self, report_id: str) -> dict[str, Any]:
        """Delete one report, report artifacts, image cache, and annotations."""
        with self._lock, self.connect() as conn:
            row = conn.execute("select payload from reports where id = ?", (report_id,)).fetchone()
            if not row:
                return {"deleted": False, "report_files_deleted": 0, "image_files_deleted": 0, "annotations_deleted": 0}
            report = json.loads(row["payload"])
            image_cur = conn.execute("delete from image_annotations where report_id = ?", (report_id,))
            formula_cur = conn.execute("delete from formula_annotations where report_id = ?", (report_id,))
            omml_cur = conn.execute("delete from omml_annotations where report_id = ?", (report_id,))
            layout_cur = conn.execute("delete from layout_annotations where report_id = ?", (report_id,))
            conn.execute("delete from reports where id = ?", (report_id,))
        report_files_deleted = self._cleanup_report_files([report])
        image_files_deleted = self._cleanup_report_image_files(report)
        return {
            "deleted": True,
            "report_files_deleted": report_files_deleted,
            "image_files_deleted": image_files_deleted,
            "annotations_deleted": image_cur.rowcount + formula_cur.rowcount + omml_cur.rowcount + layout_cur.rowcount,
        }

    def append_log(self, task_id: str, message: str, level: str = "info", category: str = "system") -> None:
        """Append a log entry when history and category settings allow it."""
        category = str(category or "system")
        if level == "error" and category == "system":
            category = "error"
        if level == "error" and not self.get_settings().get("logErrorEvents", True):
            return
        if not self._log_category_enabled(category):
            return
        if not self.get_settings().get("saveHistory", True) and category in {"upload", "conversion", "formula", "omml", "macro", "image", "error"}:
            return
        with self._lock, self.connect() as conn:
            conn.execute(
                "insert into logs (task_id, message, level, category, created_at) values (?, ?, ?, ?, ?)",
                (task_id, message, level, category, utc_now()),
            )

    def list_logs(self, task_id: str | None = None, limit: int = 200) -> list[dict[str, Any]]:
        """Return global or task-specific logs newest first."""
        with self._lock, self.connect() as conn:
            if task_id:
                rows = conn.execute(
                    "select task_id, message, level, category, created_at from logs where task_id = ? order by id desc limit ?",
                    (task_id, limit),
                ).fetchall()
            else:
                rows = conn.execute(
                    "select task_id, message, level, category, created_at from logs order by id desc limit ?",
                    (limit,),
                ).fetchall()
        return [dict(row) for row in rows]

    def get_settings(self) -> dict[str, Any]:
        """Return settings with PDF-to-Word normalized to Mathpix."""
        with self._lock, self.connect() as conn:
            rows = conn.execute("select key, value from settings").fetchall()
        settings = {row["key"]: json.loads(row["value"]) for row in rows}
        settings["pdfToWordEngine"] = "Mathpix"
        return settings

    def output_base_dir(self) -> Path:
        """Return and create the configured output base directory."""
        raw = str(self.get_settings().get("outputDirectory") or "outputs").strip() or "outputs"
        path = Path(raw).expanduser()
        if not path.is_absolute():
            path = self.data_dir / path
        path.mkdir(parents=True, exist_ok=True)
        return path

    def output_task_dir(self, task_id: str) -> Path:
        """Return and create the output directory for one task."""
        path = self.output_base_dir() / task_id
        path.mkdir(parents=True, exist_ok=True)
        return path

    def update_settings(self, patch: dict[str, Any]) -> dict[str, Any]:
        """Persist valid settings while forcing the PDF engine to Mathpix."""
        valid_keys = set(DEFAULT_SETTINGS)
        with self._lock, self.connect() as conn:
            for key, value in patch.items():
                if key in valid_keys:
                    if key == "pdfToWordEngine":
                        value = "Mathpix"
                    conn.execute(
                        "insert or replace into settings (key, value, updated_at) values (?, ?, ?)",
                        (key, json.dumps(value, ensure_ascii=False), utc_now()),
                    )
        return self.get_settings()

    def cleanup_runtime_history(
        self,
        retention_days: int = 30,
        log_retention_days: int | None = None,
    ) -> dict[str, Any]:
        """Delete expired task/report/log history and managed runtime artifacts."""
        retention_days = max(0, int(retention_days))
        log_retention_days = retention_days if log_retention_days is None else max(0, int(log_retention_days))
        now = datetime.now(timezone.utc)
        task_cutoff = now - timedelta(days=retention_days)
        log_cutoff = now - timedelta(days=log_retention_days)
        report_payloads: list[dict[str, Any]] = []
        task_ids: set[str] = set()
        report_ids: set[str] = set()
        deleted = {
            "retention_days": retention_days,
            "log_retention_days": log_retention_days,
            "cutoff": task_cutoff.isoformat(timespec="seconds"),
            "log_cutoff": log_cutoff.isoformat(timespec="seconds"),
            "tasks_deleted": 0,
            "reports_deleted": 0,
            "logs_deleted": 0,
            "image_annotations_deleted": 0,
            "formula_annotations_deleted": 0,
            "omml_annotations_deleted": 0,
            "layout_annotations_deleted": 0,
            "report_files_deleted": 0,
            "output_dirs_deleted": 0,
            "backup_dirs_deleted": 0,
            "image_dirs_deleted": 0,
        }
        with self._lock, self.connect() as conn:
            task_rows = conn.execute("select id, created_at, updated_at from tasks").fetchall()
            for row in task_rows:
                if self._timestamp_at_or_before(str(row["updated_at"] or row["created_at"]), task_cutoff):
                    task_ids.add(str(row["id"]))

            report_rows = conn.execute("select id, task_id, payload, created_at from reports").fetchall()
            for row in report_rows:
                if str(row["task_id"]) in task_ids or self._timestamp_at_or_before(str(row["created_at"]), task_cutoff):
                    report_ids.add(str(row["id"]))
                    report_payloads.append(json.loads(row["payload"]))

            if report_ids:
                placeholders = self._sql_placeholders(report_ids)
                params = tuple(report_ids)
                cur = conn.execute(f"delete from image_annotations where report_id in ({placeholders})", params)
                deleted["image_annotations_deleted"] = cur.rowcount
                cur = conn.execute(f"delete from formula_annotations where report_id in ({placeholders})", params)
                deleted["formula_annotations_deleted"] = cur.rowcount
                cur = conn.execute(f"delete from omml_annotations where report_id in ({placeholders})", params)
                deleted["omml_annotations_deleted"] = cur.rowcount
                cur = conn.execute(f"delete from layout_annotations where report_id in ({placeholders})", params)
                deleted["layout_annotations_deleted"] = cur.rowcount
                cur = conn.execute(f"delete from reports where id in ({placeholders})", params)
                deleted["reports_deleted"] = cur.rowcount

            if task_ids:
                placeholders = self._sql_placeholders(task_ids)
                params = tuple(task_ids)
                cur = conn.execute(f"delete from tasks where id in ({placeholders})", params)
                deleted["tasks_deleted"] = cur.rowcount

            log_rows = conn.execute("select id, task_id, created_at from logs").fetchall()
            log_ids = [
                int(row["id"])
                for row in log_rows
                if str(row["task_id"]) in task_ids or self._timestamp_at_or_before(str(row["created_at"]), log_cutoff)
            ]
            if log_ids:
                placeholders = self._sql_placeholders(log_ids)
                cur = conn.execute(f"delete from logs where id in ({placeholders})", tuple(log_ids))
                deleted["logs_deleted"] = cur.rowcount

        deleted["report_files_deleted"] = self._cleanup_report_files(report_payloads)
        dirs_deleted = self._cleanup_runtime_dirs(task_ids)
        deleted.update(dirs_deleted)
        return deleted

    def save_macro_template(self, payload: dict[str, Any]) -> dict[str, Any]:
        """Persist a reusable Word macro execution-order template."""
        now = utc_now()
        template = {
            "id": str(payload.get("id") or new_id("macro_tpl")),
            "name": str(payload.get("name") or "未命名宏模板").strip() or "未命名宏模板",
            "description": str(payload.get("description") or ""),
            "macro_sequence": list(payload.get("macro_sequence") or payload.get("macroSequence") or []),
            "defaults": dict(payload.get("defaults") or {}),
            "created_at": str(payload.get("created_at") or now),
            "updated_at": now,
        }
        with self._lock, self.connect() as conn:
            conn.execute(
                """
                insert or replace into macro_templates (id, payload, created_at, updated_at)
                values (?, ?, coalesce((select created_at from macro_templates where id = ?), ?), ?)
                """,
                (template["id"], json.dumps(template, ensure_ascii=False), template["id"], template["created_at"], now),
            )
        return template

    def get_macro_template(self, template_id: str) -> dict[str, Any] | None:
        """Return one macro template by id."""
        with self._lock, self.connect() as conn:
            row = conn.execute("select payload from macro_templates where id = ?", (template_id,)).fetchone()
        return json.loads(row["payload"]) if row else None

    def list_macro_templates(self) -> list[dict[str, Any]]:
        """Return macro templates newest-updated first."""
        with self._lock, self.connect() as conn:
            rows = conn.execute("select payload from macro_templates order by updated_at desc").fetchall()
        return [json.loads(row["payload"]) for row in rows]

    def delete_macro_template(self, template_id: str) -> bool:
        """Delete one macro sequence template by id."""
        with self._lock, self.connect() as conn:
            cur = conn.execute("delete from macro_templates where id = ?", (template_id,))
        return cur.rowcount > 0

    def save_user(self, payload: dict[str, Any]) -> dict[str, Any]:
        """Persist a local user with role-derived permissions."""
        now = utc_now()
        role = str(payload.get("role") or "学生")
        if role not in ROLE_PERMISSIONS:
            role = "学生"
        user_id = str(payload.get("id") or new_id("user"))
        with self._lock, self.connect() as conn:
            existing = conn.execute("select payload from users where id = ?", (user_id,)).fetchone()
            previous = json.loads(existing["payload"]) if existing else {}
            permissions = payload.get("permissions")
            if not isinstance(permissions, list):
                permissions = previous.get("permissions") or ROLE_PERMISSIONS[role]
            user = {
                "id": user_id,
                "name": str(payload.get("name") or previous.get("name") or "未命名用户").strip() or "未命名用户",
                "role": role,
                "status": str(payload.get("status") or previous.get("status") or "启用"),
                "permissions": list(dict.fromkeys(str(item) for item in permissions)),
                "login_enabled": bool(payload.get("login_enabled", previous.get("login_enabled", True))),
                "last_login_at": str(payload.get("last_login_at") or previous.get("last_login_at") or ""),
                "created_at": str(previous.get("created_at") or payload.get("created_at") or now),
                "updated_at": now,
            }
            conn.execute(
                """
                insert or replace into users (id, payload, created_at, updated_at)
                values (?, ?, coalesce((select created_at from users where id = ?), ?), ?)
                """,
                (user["id"], json.dumps(user, ensure_ascii=False), user["id"], user["created_at"], now),
            )
        return user

    def get_user(self, user_id: str) -> dict[str, Any] | None:
        """Return one local user by id."""
        with self._lock, self.connect() as conn:
            row = conn.execute("select payload from users where id = ?", (user_id,)).fetchone()
        return json.loads(row["payload"]) if row else None

    def list_users(self) -> list[dict[str, Any]]:
        """Return local users newest-updated first."""
        with self._lock, self.connect() as conn:
            rows = conn.execute("select payload from users order by updated_at desc").fetchall()
        return [json.loads(row["payload"]) for row in rows]

    def delete_user(self, user_id: str) -> bool:
        """Delete one local user by id."""
        with self._lock, self.connect() as conn:
            cur = conn.execute("delete from users where id = ?", (user_id,))
        return cur.rowcount > 0

    def save_template(self, payload: dict[str, Any]) -> dict[str, Any]:
        """Persist a conversion or OCR template descriptor."""
        now = utc_now()
        template_id = str(payload.get("id") or new_id("tpl"))
        with self._lock, self.connect() as conn:
            existing = conn.execute("select payload from templates where id = ?", (template_id,)).fetchone()
            previous = json.loads(existing["payload"]) if existing else {}
            template = {
                "id": template_id,
                "name": str(payload.get("name") or previous.get("name") or "未命名模板").strip() or "未命名模板",
                "template_type": str(payload.get("template_type") or payload.get("templateType") or previous.get("template_type") or "PPT 模板"),
                "applies_to": str(payload.get("applies_to") or payload.get("appliesTo") or previous.get("applies_to") or "word_to_ppt"),
                "description": str(payload.get("description") or previous.get("description") or ""),
                "template_path": str(payload.get("template_path") or payload.get("templatePath") or previous.get("template_path") or ""),
                "settings": dict(payload.get("settings") or previous.get("settings") or {}),
                "tags": list(payload.get("tags") or previous.get("tags") or []),
                "status": str(payload.get("status") or previous.get("status") or "启用"),
                "created_at": str(previous.get("created_at") or payload.get("created_at") or now),
                "updated_at": now,
            }
            conn.execute(
                """
                insert or replace into templates (id, payload, created_at, updated_at)
                values (?, ?, coalesce((select created_at from templates where id = ?), ?), ?)
                """,
                (template["id"], json.dumps(template, ensure_ascii=False), template["id"], template["created_at"], now),
            )
        return template

    def list_templates(self) -> list[dict[str, Any]]:
        """Return general templates newest-updated first."""
        with self._lock, self.connect() as conn:
            rows = conn.execute("select payload from templates order by updated_at desc").fetchall()
        return [json.loads(row["payload"]) for row in rows]

    def delete_template(self, template_id: str) -> bool:
        """Delete one general template by id."""
        with self._lock, self.connect() as conn:
            cur = conn.execute("delete from templates where id = ?", (template_id,))
        return cur.rowcount > 0

    def save_authorization(self, payload: dict[str, Any]) -> dict[str, Any]:
        """Persist one explicit authorization toggle such as Mathpix upload."""
        key = str(payload.get("key") or "")
        if not key:
            raise ValueError("授权项 key 不能为空")
        now = utc_now()
        with self._lock, self.connect() as conn:
            existing = conn.execute("select payload from authorizations where key = ?", (key,)).fetchone()
            previous = json.loads(existing["payload"]) if existing else {}
            base = dict(DEFAULT_AUTHORIZATIONS.get(key) or {})
            enabled = bool(payload.get("enabled", previous.get("enabled", base.get("enabled", False))))
            authorization = {
                "key": key,
                "name": str(payload.get("name") or previous.get("name") or base.get("name") or key),
                "scope": str(payload.get("scope") or previous.get("scope") or base.get("scope") or ""),
                "setting_key": str(payload.get("setting_key") or previous.get("setting_key") or base.get("setting_key") or ""),
                "enabled": enabled,
                "status": "已授权" if enabled else "未授权",
                "risk_level": str(payload.get("risk_level") or previous.get("risk_level") or base.get("risk_level") or "中"),
                "description": str(payload.get("description") or previous.get("description") or base.get("description") or ""),
                "note": str(payload.get("note") or ""),
                "updated_by": str(payload.get("updated_by") or payload.get("updatedBy") or "本地管理员"),
                "created_at": str(previous.get("created_at") or payload.get("created_at") or now),
                "updated_at": now,
            }
            conn.execute(
                """
                insert or replace into authorizations (key, payload, created_at, updated_at)
                values (?, ?, coalesce((select created_at from authorizations where key = ?), ?), ?)
                """,
                (key, json.dumps(authorization, ensure_ascii=False), key, authorization["created_at"], now),
            )
        return authorization

    def list_authorizations(self) -> list[dict[str, Any]]:
        """Return authorization records, creating defaults when missing."""
        with self._lock, self.connect() as conn:
            rows = conn.execute("select payload from authorizations order by key").fetchall()
        existing = {item["key"]: item for item in (json.loads(row["payload"]) for row in rows)}
        for key, value in DEFAULT_AUTHORIZATIONS.items():
            if key not in existing:
                existing[key] = self.save_authorization({"key": key, **value})
        return [existing[key] for key in sorted(existing)]

    def save_image_annotation(self, payload: dict[str, Any]) -> dict[str, Any]:
        """Persist manual review or replacement metadata for one extracted image."""
        now = utc_now()
        image_id = str(payload.get("image_id") or "")
        report_id = str(payload.get("report_id") or "")
        annotation = {
            "id": str(payload.get("id") or new_id("img_note")),
            "image_id": image_id,
            "report_id": report_id,
            "status": str(payload.get("status") or "误判"),
            "label": str(payload.get("label") or ""),
            "note": str(payload.get("note") or ""),
            "replacement_file_name": str(payload.get("replacement_file_name") or payload.get("replacementFileName") or ""),
            "replacement_storage_name": str(payload.get("replacement_storage_name") or payload.get("replacementStorageName") or ""),
            "replacement_asset_url": str(payload.get("replacement_asset_url") or payload.get("replacementAssetUrl") or ""),
            "replacement_file_size": int(payload.get("replacement_file_size") or payload.get("replacementFileSize") or 0),
            "replacement_sha256": str(payload.get("replacement_sha256") or payload.get("replacementSha256") or ""),
            "replacement_mime_type": str(payload.get("replacement_mime_type") or payload.get("replacementMimeType") or ""),
            "created_at": str(payload.get("created_at") or now),
            "updated_at": now,
        }
        with self._lock, self.connect() as conn:
            existing = conn.execute(
                "select payload from image_annotations where image_id = ? and report_id = ?",
                (image_id, report_id),
            ).fetchone()
            if existing:
                previous = json.loads(existing["payload"])
                annotation["id"] = previous.get("id", annotation["id"])
                annotation["created_at"] = previous.get("created_at", annotation["created_at"])
            conn.execute(
                """
                insert or replace into image_annotations (id, image_id, report_id, payload, created_at, updated_at)
                values (?, ?, ?, ?, ?, ?)
                """,
                (annotation["id"], image_id, report_id, json.dumps(annotation, ensure_ascii=False), annotation["created_at"], now),
            )
        return annotation

    def list_image_annotations(self, report_id: str | None = None) -> list[dict[str, Any]]:
        """Return image annotations globally or for one report."""
        with self._lock, self.connect() as conn:
            if report_id:
                rows = conn.execute(
                    "select payload from image_annotations where report_id = ? order by updated_at desc",
                    (report_id,),
                ).fetchall()
            else:
                rows = conn.execute("select payload from image_annotations order by updated_at desc").fetchall()
        return [json.loads(row["payload"]) for row in rows]

    def delete_image_annotation(self, annotation_id: str) -> bool:
        """Delete one image annotation by id."""
        with self._lock, self.connect() as conn:
            cur = conn.execute("delete from image_annotations where id = ?", (annotation_id,))
        return cur.rowcount > 0

    def save_formula_annotation(self, payload: dict[str, Any]) -> dict[str, Any]:
        """Persist manual confirmation or correction for one formula."""
        now = utc_now()
        formula_id = str(payload.get("formula_id") or "")
        report_id = str(payload.get("report_id") or "")
        annotation = {
            "id": str(payload.get("id") or new_id("formula_note")),
            "formula_id": formula_id,
            "report_id": report_id,
            "status": str(payload.get("status") or "已确认"),
            "latex": str(payload.get("latex") or ""),
            "mathml": str(payload.get("mathml") or ""),
            "note": str(payload.get("note") or ""),
            "created_at": str(payload.get("created_at") or now),
            "updated_at": now,
        }
        with self._lock, self.connect() as conn:
            existing = conn.execute(
                "select payload from formula_annotations where formula_id = ? and report_id = ?",
                (formula_id, report_id),
            ).fetchone()
            if existing:
                previous = json.loads(existing["payload"])
                annotation["id"] = previous.get("id", annotation["id"])
                annotation["created_at"] = previous.get("created_at", annotation["created_at"])
            conn.execute(
                """
                insert or replace into formula_annotations (id, formula_id, report_id, payload, created_at, updated_at)
                values (?, ?, ?, ?, ?, ?)
                """,
                (annotation["id"], formula_id, report_id, json.dumps(annotation, ensure_ascii=False), annotation["created_at"], now),
            )
        return annotation

    def list_formula_annotations(self, report_id: str | None = None) -> list[dict[str, Any]]:
        """Return formula annotations globally or for one report."""
        with self._lock, self.connect() as conn:
            if report_id:
                rows = conn.execute(
                    "select payload from formula_annotations where report_id = ? order by updated_at desc",
                    (report_id,),
                ).fetchall()
            else:
                rows = conn.execute("select payload from formula_annotations order by updated_at desc").fetchall()
        return [json.loads(row["payload"]) for row in rows]

    def delete_formula_annotation(self, annotation_id: str) -> bool:
        """Delete one formula annotation by id."""
        with self._lock, self.connect() as conn:
            cur = conn.execute("delete from formula_annotations where id = ?", (annotation_id,))
        return cur.rowcount > 0

    def save_omml_annotation(self, payload: dict[str, Any]) -> dict[str, Any]:
        """Persist OMML keep/retry/manual-dependency correction metadata."""
        now = utc_now()
        dependency_id = str(payload.get("dependency_id") or payload.get("dependencyId") or "")
        report_id = str(payload.get("report_id") or payload.get("reportId") or "")
        annotation = {
            "id": str(payload.get("id") or new_id("omml_note")),
            "dependency_id": dependency_id,
            "report_id": report_id,
            "status": str(payload.get("status") or "保留OMML"),
            "keep_omml": bool(payload.get("keep_omml", payload.get("keepOmml", True))),
            "retry_conversion": bool(payload.get("retry_conversion", payload.get("retryConversion", False))),
            "manual_omml_path": str(payload.get("manual_omml_path") or payload.get("manualOmmlPath") or ""),
            "note": str(payload.get("note") or ""),
            "created_at": str(payload.get("created_at") or now),
            "updated_at": now,
        }
        with self._lock, self.connect() as conn:
            existing = conn.execute(
                "select payload from omml_annotations where dependency_id = ? and report_id = ?",
                (dependency_id, report_id),
            ).fetchone()
            if existing:
                previous = json.loads(existing["payload"])
                annotation["id"] = previous.get("id", annotation["id"])
                annotation["created_at"] = previous.get("created_at", annotation["created_at"])
            conn.execute(
                """
                insert or replace into omml_annotations (id, dependency_id, report_id, payload, created_at, updated_at)
                values (?, ?, ?, ?, ?, ?)
                """,
                (annotation["id"], dependency_id, report_id, json.dumps(annotation, ensure_ascii=False), annotation["created_at"], now),
            )
        return annotation

    def list_omml_annotations(self, report_id: str | None = None) -> list[dict[str, Any]]:
        """Return OMML annotations globally or for one report."""
        with self._lock, self.connect() as conn:
            if report_id:
                rows = conn.execute(
                    "select payload from omml_annotations where report_id = ? order by updated_at desc",
                    (report_id,),
                ).fetchall()
            else:
                rows = conn.execute("select payload from omml_annotations order by updated_at desc").fetchall()
        return [json.loads(row["payload"]) for row in rows]

    def delete_omml_annotation(self, annotation_id: str) -> bool:
        """Delete one OMML annotation by id."""
        with self._lock, self.connect() as conn:
            cur = conn.execute("delete from omml_annotations where id = ?", (annotation_id,))
        return cur.rowcount > 0

    def save_layout_annotation(self, payload: dict[str, Any]) -> dict[str, Any]:
        """Persist a report layout-quality correction note."""
        now = utc_now()
        report_id = str(payload.get("report_id") or payload.get("reportId") or "")
        location = str(payload.get("location") or payload.get("position") or "").strip()
        annotation = {
            "id": str(payload.get("id") or new_id("layout_note")),
            "report_id": report_id,
            "location": location,
            "issue_type": str(payload.get("issue_type") or payload.get("issueType") or "标题层级"),
            "status": str(payload.get("status") or "待校正"),
            "page_index": payload.get("page_index", payload.get("pageIndex", "")),
            "before": str(payload.get("before") or ""),
            "after": str(payload.get("after") or ""),
            "recommendation": str(payload.get("recommendation") or ""),
            "note": str(payload.get("note") or ""),
            "created_at": str(payload.get("created_at") or now),
            "updated_at": now,
        }
        with self._lock, self.connect() as conn:
            existing = conn.execute(
                "select payload from layout_annotations where report_id = ? and location = ?",
                (report_id, location),
            ).fetchone()
            if existing:
                previous = json.loads(existing["payload"])
                annotation["id"] = previous.get("id", annotation["id"])
                annotation["created_at"] = previous.get("created_at", annotation["created_at"])
            conn.execute(
                """
                insert or replace into layout_annotations (id, report_id, location, payload, created_at, updated_at)
                values (?, ?, ?, ?, ?, ?)
                """,
                (annotation["id"], report_id, location, json.dumps(annotation, ensure_ascii=False), annotation["created_at"], now),
            )
        return annotation

    def list_layout_annotations(self, report_id: str | None = None) -> list[dict[str, Any]]:
        """Return layout annotations globally or for one report."""
        with self._lock, self.connect() as conn:
            if report_id:
                rows = conn.execute(
                    "select payload from layout_annotations where report_id = ? order by updated_at desc",
                    (report_id,),
                ).fetchall()
            else:
                rows = conn.execute("select payload from layout_annotations order by updated_at desc").fetchall()
        return [json.loads(row["payload"]) for row in rows]

    def delete_layout_annotation(self, annotation_id: str) -> bool:
        """Delete one layout annotation by id."""
        with self._lock, self.connect() as conn:
            cur = conn.execute("delete from layout_annotations where id = ?", (annotation_id,))
        return cur.rowcount > 0

    def _cleanup_uploaded_payloads(self, payloads: list[dict[str, Any]]) -> None:
        upload_root = self.uploads_dir.resolve()
        archive_dirs: set[Path] = set()
        for payload in payloads:
            source_kind = str(payload.get("source_kind") or "")
            if source_kind not in {"upload", "archive_entry"}:
                continue
            storage_path = str(payload.get("storage_path") or "")
            if storage_path:
                path = Path(storage_path)
                try:
                    resolved = path.resolve()
                    resolved.relative_to(upload_root)
                except (OSError, ValueError):
                    resolved = None
                if resolved and resolved.is_file():
                    try:
                        resolved.unlink()
                    except OSError:
                        pass
            if source_kind == "upload":
                archive_dirs.add(self.uploads_dir / str(payload.get("id") or ""))
            parent_id = str(payload.get("archive_parent_id") or "")
            if parent_id:
                archive_dirs.add(self.uploads_dir / parent_id)
        for directory in archive_dirs:
            try:
                resolved_dir = directory.resolve()
                resolved_dir.relative_to(upload_root)
            except (OSError, ValueError):
                continue
            if resolved_dir.exists() and resolved_dir.is_dir():
                shutil.rmtree(resolved_dir, ignore_errors=True)

    @staticmethod
    def _parse_timestamp(value: str) -> datetime | None:
        try:
            parsed = datetime.fromisoformat(value)
        except ValueError:
            return None
        if parsed.tzinfo is None:
            parsed = parsed.replace(tzinfo=timezone.utc)
        return parsed.astimezone(timezone.utc)

    @classmethod
    def _timestamp_at_or_before(cls, value: str, cutoff: datetime) -> bool:
        parsed = cls._parse_timestamp(value)
        return bool(parsed and parsed <= cutoff)

    @staticmethod
    def _sql_placeholders(values: set[str] | list[int]) -> str:
        return ",".join("?" for _ in values)

    def _cleanup_report_files(self, reports: list[dict[str, Any]]) -> int:
        report_root = self.reports_dir.resolve()
        deleted = 0
        for report in reports:
            for key in ("html_path", "json_path", "pdf_path", "xlsx_path", "txt_path", "failure_csv_path"):
                path = Path(str(report.get(key) or ""))
                if not path:
                    continue
                try:
                    resolved = path.resolve()
                    resolved.relative_to(report_root)
                except (OSError, ValueError):
                    continue
                if resolved.is_file():
                    try:
                        resolved.unlink()
                        deleted += 1
                    except OSError:
                        pass
        return deleted

    def _cleanup_report_image_files(self, report: dict[str, Any]) -> int:
        image_root = self.images_dir.resolve()
        deleted = 0
        for image in report.get("analysis", {}).get("smallImages", []):
            path = Path(str(image.get("image_path") or ""))
            if not path:
                continue
            try:
                resolved = path.resolve()
                resolved.relative_to(image_root)
            except (OSError, ValueError):
                continue
            if resolved.is_file():
                try:
                    resolved.unlink()
                    deleted += 1
                except OSError:
                    pass
        return deleted

    def _cleanup_runtime_dirs(self, task_ids: set[str]) -> dict[str, int]:
        counts = {"output_dirs_deleted": 0, "backup_dirs_deleted": 0, "image_dirs_deleted": 0}
        output_roots = [self._configured_output_base_dir(), self.outputs_dir]
        backup_roots = [self.backups_dir]
        image_roots = [self.images_dir]
        for task_id in task_ids:
            if not task_id:
                continue
            for root in output_roots:
                counts["output_dirs_deleted"] += self._remove_child_dir(root, task_id)
            for root in backup_roots:
                counts["backup_dirs_deleted"] += self._remove_child_dir(root, task_id)
            for root in image_roots:
                counts["image_dirs_deleted"] += self._remove_child_dir(root, task_id)
        return counts

    def _configured_output_base_dir(self) -> Path:
        raw = str(self.get_settings().get("outputDirectory") or "outputs").strip() or "outputs"
        path = Path(raw).expanduser()
        return path if path.is_absolute() else self.data_dir / path

    @staticmethod
    def _remove_child_dir(root: Path, child_name: str) -> int:
        try:
            resolved_root = root.resolve()
            target = (root / child_name).resolve()
            target.relative_to(resolved_root)
        except (OSError, ValueError):
            return 0
        if not target.exists() or not target.is_dir():
            return 0
        shutil.rmtree(target, ignore_errors=True)
        return 0 if target.exists() else 1

    def _reset_password_sessions(self) -> None:
        for payload in self.list_files():
            if not payload.get("password_session_active"):
                continue
            payload["password_provided"] = False
            payload["password_session_active"] = False
            summary = dict(payload.get("content_summary") or {})
            summary["passwordStatus"] = "未输入"
            payload["content_summary"] = summary
            errors = list(payload.get("validation_errors") or [])
            if payload.get("encrypted") and not any("密码" in str(error) or "加密" in str(error) for error in errors):
                errors.append("文件已加密，需要密码")
            payload["validation_errors"] = errors
            if errors:
                payload["status"] = "校验失败"
            self.save_file(payload)
