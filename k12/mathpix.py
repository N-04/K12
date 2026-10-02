"""封装 Mathpix PDF OCR 的 HTTP 接口。构建客户端前，由 TaskProcessor 检查外部上传授权、凭据及文档是否允许离开本机。"""

from __future__ import annotations

import json
import os
import time
import urllib.error
import urllib.request
from pathlib import Path
from typing import Any

MATHPIX_DOWNLOAD_EXTENSIONS = {"docx", "tex.zip"}
MATHPIX_CONVERSION_FORMATS = ("docx", "tex.zip", "html", "md", "mmd", "lines.json")
MATHPIX_ID_CHARS = set("abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_-")
MATHPIX_ERROR_DETAIL_LIMIT = 600


class MathpixConfigError(RuntimeError):
    """已授权的 Mathpix 凭据环境不完整时抛出。"""

    code = "mathpix_config_error"


class MathpixApiError(RuntimeError):
    """Mathpix HTTP 工作流返回不安全或失败结果时抛出。"""

    code = "mathpix_api_error"


class MathpixClient:
    """使用标准库实现 Mathpix /v3/pdf 工作流客户端。"""

    base_url = "https://api.mathpix.com"

    def __init__(self, app_id: str, app_key: str, timeout: int = 30) -> None:
        """初始化当前对象所需的配置、依赖与运行状态。"""
        if not app_id or not app_key:
            raise MathpixConfigError("Mathpix APP ID 或 APP KEY 未配置")
        self.app_id = app_id
        self.app_key = app_key
        self.timeout = timeout

    @classmethod
    def from_environment(cls, app_id_env: str, app_key_env: str) -> "MathpixClient":
        """读取指定环境变量中的凭据，不记录其值。"""
        return cls(os.getenv(app_id_env, ""), os.getenv(app_key_env, ""))

    def submit_pdf(self, pdf_path: Path, options: dict[str, Any] | None = None) -> dict[str, Any]:
        """向 Mathpix 提交本地 PDF 并返回 JSON 响应。"""
        fields = {"options_json": json.dumps(self._pdf_options(options), ensure_ascii=False)}
        files = {"file": (pdf_path.name, pdf_path.read_bytes(), "application/pdf")}
        return self._request_json("POST", "/v3/pdf", body=self._multipart_body(fields, files))

    def get_pdf_status(self, pdf_id: str) -> dict[str, Any]:
        """读取已提交 PDF 的 Mathpix 处理状态。"""
        safe_pdf_id = self._pdf_result_id(pdf_id)
        return self._request_json("GET", f"/v3/pdf/{safe_pdf_id}")

    def wait_for_pdf(self, pdf_id: str, timeout_seconds: int = 600, poll_seconds: int = 5) -> dict[str, Any]:
        """轮询至 Mathpix 终态或本地超时。"""
        deadline = time.time() + timeout_seconds
        latest: dict[str, Any] = {}
        while time.time() < deadline:
            latest = self.get_pdf_status(pdf_id)
            if latest.get("status") in {"completed", "error"}:
                return latest
            time.sleep(poll_seconds)
        latest["status"] = latest.get("status") or "timeout"
        return latest

    def download_pdf_result(self, pdf_id: str, extension: str) -> bytes:
        """下载 Mathpix 生成的 docx 或 tex.zip 等产物。"""
        safe_pdf_id = self._pdf_result_id(pdf_id)
        safe_extension = self._pdf_result_extension(extension)
        return self._request_bytes("GET", f"/v3/pdf/{safe_pdf_id}.{safe_extension}")

    def _request_json(self, method: str, path: str, body: tuple[bytes, str] | None = None) -> dict[str, Any]:
        """发起 Mathpix 请求并解码 JSON 响应。"""
        data = self._request_bytes(method, path, body)
        try:
            payload = json.loads(data.decode("utf-8"))
        except (UnicodeDecodeError, json.JSONDecodeError) as exc:
            raise MathpixApiError("Mathpix API 返回非 JSON 响应") from exc
        if not isinstance(payload, dict):
            raise MathpixApiError("Mathpix API JSON 响应格式异常")
        return payload

    def _request_bytes(self, method: str, path: str, body: tuple[bytes, str] | None = None) -> bytes:
        """发起 Mathpix 请求并转换传输错误供调用者处理。"""
        payload = body[0] if body else None
        request = urllib.request.Request(f"{self.base_url}{path}", data=payload, method=method)
        request.add_header("app_id", self.app_id)
        request.add_header("app_key", self.app_key)
        if body:
            request.add_header("Content-Type", body[1])
        try:
            with urllib.request.urlopen(request, timeout=self.timeout) as response:
                return response.read()
        except urllib.error.HTTPError as exc:
            detail = self._error_detail(exc.read(MATHPIX_ERROR_DETAIL_LIMIT + 1))
            raise MathpixApiError(f"Mathpix API 返回 {exc.code}: {detail}") from exc
        except urllib.error.URLError as exc:
            raise MathpixApiError(f"Mathpix API 请求失败: {exc.reason}") from exc

    def _error_detail(self, data: bytes) -> str:
        """精简传输错误供任务日志和报告使用。"""
        text = data.decode("utf-8", errors="replace").replace("\r", " ").replace("\n", " ").strip()
        for secret in {self.app_id, self.app_key}:
            if secret:
                text = text.replace(secret, "[redacted]")
        if len(text) > MATHPIX_ERROR_DETAIL_LIMIT:
            return f"{text[:MATHPIX_ERROR_DETAIL_LIMIT]}..."
        return text or "无响应正文"

    @staticmethod
    def _multipart_body(fields: dict[str, str], files: dict[str, tuple[str, bytes, str]]) -> tuple[bytes, str]:
        """使用标准库构建多部分请求体。"""
        boundary = f"k12-mathpix-{os.urandom(12).hex()}"
        chunks: list[bytes] = []
        for name, value in fields.items():
            safe_name = MathpixClient._multipart_header_value(name)
            chunks.extend(
                [
                    f"--{boundary}\r\n".encode("utf-8"),
                    f'Content-Disposition: form-data; name="{safe_name}"\r\n\r\n'.encode("utf-8"),
                    value.encode("utf-8"),
                    b"\r\n",
                ]
            )
        for name, (filename, data, content_type) in files.items():
            safe_name = MathpixClient._multipart_header_value(name)
            safe_filename = MathpixClient._multipart_header_value(filename)
            chunks.extend(
                [
                    f"--{boundary}\r\n".encode("utf-8"),
                    f'Content-Disposition: form-data; name="{safe_name}"; filename="{safe_filename}"\r\n'.encode("utf-8"),
                    f"Content-Type: {content_type}\r\n\r\n".encode("utf-8"),
                    data,
                    b"\r\n",
                ]
            )
        chunks.append(f"--{boundary}--\r\n".encode("utf-8"))
        return b"".join(chunks), f"multipart/form-data; boundary={boundary}"

    @staticmethod
    def _multipart_header_value(value: str) -> str:
        """确保多部分请求参数仅占一行安全头字段。"""
        return str(value).replace("\r", "_").replace("\n", "_").replace('"', "_")

    @staticmethod
    def _pdf_result_extension(extension: str) -> str:
        """仅允许 K12 PDF 转 Word 使用的 Mathpix 产物类型。"""
        value = str(extension or "").strip().lower()
        if value not in MATHPIX_DOWNLOAD_EXTENSIONS:
            raise MathpixApiError(f"Unsupported Mathpix PDF result extension: {extension}")
        return value

    @staticmethod
    def _pdf_result_id(pdf_id: str) -> str:
        """限制 Mathpix PDF 请求在预期接口路径内。"""
        value = str(pdf_id or "").strip()
        if not value or any(char not in MATHPIX_ID_CHARS for char in value):
            raise MathpixApiError("Invalid Mathpix PDF id")
        return value

    @staticmethod
    def _pdf_options(options: dict[str, Any] | None = None) -> dict[str, Any]:
        """分离 Mathpix 参数与 K12 专用 OCR 元数据。"""
        raw = dict(options or {})
        raw_formats = raw.get("conversion_formats") if isinstance(raw.get("conversion_formats"), dict) else {}
        if not raw_formats:
            raw_formats = {
                key: raw.get(key)
                for key in MATHPIX_CONVERSION_FORMATS
                if key in raw
            }
        formats = {
            str(key): True
            for key, value in raw_formats.items()
            if key in MATHPIX_CONVERSION_FORMATS and bool(value)
        }
        if not formats:
            formats = {"docx": True}
        # PDF 转 Word 必须包含 DOCX，即使同时请求 tex.zip、HTML 或行数据等可选产物。
        formats["docx"] = True
        payload: dict[str, Any] = {"conversion_formats": formats}
        allowed_keys = {
            "metadata",
            "conversion_options",
            "enable_tables_fallback",
            "fullwidth_punctuation",
            "include_equation_tags",
            "include_line_data",
            "include_page_info",
            "include_smiles",
            "idiomatic_eqn_arrays",
            "math_display_delimiters",
            "math_inline_delimiters",
            "numbers_default_to_math",
            "page_ranges",
            "rm_spaces",
        }
        for key in allowed_keys:
            if key in raw:
                payload[key] = raw[key]
        return payload
