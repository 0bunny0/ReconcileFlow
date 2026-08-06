from __future__ import annotations

import argparse
from email.parser import BytesParser
from email.policy import default
import json
import mimetypes
from pathlib import Path
import secrets
import threading
import time
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import quote, unquote, urlparse

from reconciliation import Result, ReconciliationError, basis_config, export_excel, module_config, payload, reconcile_module_all


ROOT = Path(__file__).resolve().parent
STATIC = ROOT / "static"
MAX_REQUEST_BYTES = 150 * 1024 * 1024
DOWNLOADS: dict[str, tuple[float, bytes | Result, str]] = {}


def _cleanup_downloads() -> None:
    cutoff = time.time() - 60 * 60
    for token, (created, _, _) in list(DOWNLOADS.items()):
        if created < cutoff:
            DOWNLOADS.pop(token, None)


def _multipart(body: bytes, content_type: str) -> dict[str, tuple[str, bytes]]:
    message = BytesParser(policy=default).parsebytes(
        b"Content-Type: " + content_type.encode("utf-8") + b"\r\nMIME-Version: 1.0\r\n\r\n" + body
    )
    if not message.is_multipart():
        raise ReconciliationError("上传格式不正确，请重新选择当前模块的三份 Excel 文件。")
    fields: dict[str, tuple[str, bytes]] = {}
    for part in message.iter_parts():
        name = part.get_param("name", header="content-disposition")
        if not name:
            continue
        filename = part.get_filename() or ""
        fields[name] = (filename, part.get_payload(decode=True) or b"")
    return fields


class Handler(BaseHTTPRequestHandler):
    server_version = "CurrencyReconciliation/2.0"

    def log_message(self, format: str, *args: object) -> None:
        print(f"[{self.log_date_time_string()}] {format % args}")

    def _bytes(self, status: int, body: bytes, content_type: str, **headers: str) -> None:
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("X-Content-Type-Options", "nosniff")
        for key, value in headers.items():
            self.send_header(key.replace("_", "-"), value)
        self.end_headers()
        self.wfile.write(body)

    def _json(self, status: int, value: object) -> None:
        body = json.dumps(value, ensure_ascii=False, allow_nan=False).encode("utf-8")
        self._bytes(status, body, "application/json; charset=utf-8", Cache_Control="no-store")

    def do_GET(self) -> None:
        path = unquote(urlparse(self.path).path)
        if path == "/api/health":
            self._json(200, {"ok": True})
            return
        if path.startswith("/api/download/"):
            _cleanup_downloads()
            token = path.rsplit("/", 1)[-1]
            item = DOWNLOADS.get(token)
            if not item:
                self._json(404, {"error": "导出文件已过期，请重新核对。"})
                return
            content = item[1]
            if isinstance(content, Result):
                content = export_excel(content)
                DOWNLOADS[token] = (item[0], content, item[2])
            self._bytes(
                200,
                content,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                Content_Disposition=f'attachment; filename="reconciliation_result.xlsx"; filename*=UTF-8\'\'{quote(item[2])}',
                Cache_Control="no-store",
            )
            return

        requested = "index.html" if path in {"", "/"} else path.lstrip("/")
        target = (STATIC / requested).resolve()
        if STATIC.resolve() not in target.parents and target != STATIC.resolve():
            self._json(404, {"error": "页面不存在。"})
            return
        if not target.is_file():
            self._json(404, {"error": "页面不存在。"})
            return
        content_type = mimetypes.guess_type(target.name)[0] or "application/octet-stream"
        if content_type.startswith("text/") or content_type in {"application/javascript", "application/json"}:
            content_type += "; charset=utf-8"
        self._bytes(200, target.read_bytes(), content_type, Cache_Control="no-cache")

    def do_POST(self) -> None:
        if urlparse(self.path).path != "/api/analyze":
            self._json(404, {"error": "接口不存在。"})
            return
        try:
            length = int(self.headers.get("Content-Length", "0"))
            if length <= 0:
                raise ReconciliationError("没有收到上传文件。")
            if length > MAX_REQUEST_BYTES:
                self._json(413, {"error": "三份文件合计不能超过 150 MB。"})
                return
            fields = _multipart(self.rfile.read(length), self.headers.get("Content-Type", ""))
            module = fields.get("module", ("", b""))[1].decode("utf-8")
            module_info = module_config(module)
            amount_key = f"{module}_amount"
            freight_key = f"{module}_freight"
            required_fields = ("document", amount_key, freight_key)
            missing = [name for name in required_fields if name not in fields or not fields[name][1]]
            if missing:
                raise ReconciliationError(f"请完整上传文控登记表、{module_info['amount_source_label']}和{module_info['freight_source_label']}。")
            try:
                tolerance = float(fields.get("tolerance", ("", b"1.00"))[1].decode("utf-8") or "1.00")
            except ValueError as exc:
                raise ReconciliationError("差异容差必须是数字。") from exc
            names = {
                "document": fields["document"][0] or "文控登记表.xlsx",
                "amount": fields[amount_key][0] or f"{module_info['amount_source_label']}.xlsx",
                "freight": fields[freight_key][0] or f"{module_info['freight_source_label']}.xlsx",
            }
            for key, filename in names.items():
                if Path(filename).suffix.lower() not in {".xlsx", ".xlsm"}:
                    raise ReconciliationError(f"{key} 不是 .xlsx 或 .xlsm 文件。")
            results = reconcile_module_all(
                fields["document"][1], fields[amount_key][1], fields[freight_key][1],
                module=module, tolerance=tolerance, source_names=names,
            )
            response: dict[str, object] = {"modules": {}, "module": module, "tolerance": tolerance, "source_names": names}
            module_payloads: dict[str, dict[str, object]] = {module: {}}
            for basis, result in results.items():
                token = secrets.token_urlsafe(24)
                filename = f"{module_config(result.module)['label']}_{basis_config(result.basis)['label']}_核对结果.xlsx"
                DOWNLOADS[token] = (time.time(), result, filename)
                result_payload = payload(result)
                result_payload["download_url"] = f"/api/download/{token}"
                module_payloads[module][basis] = result_payload
            response["modules"] = module_payloads
            self._json(200, response)
        except ReconciliationError as exc:
            self._json(400, {"error": str(exc)})
        except Exception as exc:
            print(f"Unexpected error: {exc!r}")
            self._json(500, {"error": "处理失败。请确认当前模块的三份 Excel 格式未改变，然后重试。"})


def main() -> None:
    parser = argparse.ArgumentParser(description="接单/出货/文控双模块双口径差异核对工具")
    parser.add_argument("--host", default="127.0.0.1", help="监听地址；局域网共享可用 0.0.0.0")
    parser.add_argument("--port", type=int, default=8765, help="监听端口，默认 8765")
    parser.add_argument("--no-browser", action="store_true", help="启动后不自动打开浏览器")
    args = parser.parse_args()
    server = ThreadingHTTPServer((args.host, args.port), Handler)
    visible_host = "127.0.0.1" if args.host == "0.0.0.0" else args.host
    url = f"http://{visible_host}:{args.port}"
    print(f"接单与出货差异核对工具已启动：{url}")
    print("按 Ctrl+C 停止。上传文件仅在本机内存中处理。")
    if not args.no_browser:
        threading.Timer(0.7, lambda: webbrowser.open(url)).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n服务已停止。")
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
