from __future__ import annotations

import argparse
import json
import os
from http import HTTPStatus
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse

from crm_data import ReminderStore

BASE_DIR = Path(__file__).resolve().parent
STATIC_DIR = BASE_DIR / "static"
DEFAULT_EXCEL_PATH = "/Volumes/D/wps同步/AI项目信息/浙江.xlsx"


def safe_int(value: str | None, default: int) -> int:
    try:
        return int(value or default)
    except (TypeError, ValueError):
        return default


class CRMHandler(SimpleHTTPRequestHandler):
    store: ReminderStore

    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(STATIC_DIR), **kwargs)

    def do_GET(self) -> None:
        parsed = urlparse(self.path)

        if parsed.path == "/api/meta":
            self.send_json(self.store.meta())
            return

        if parsed.path == "/api/reminders":
            params = parse_qs(parsed.query)
            payload = self.store.query(
                search=params.get("search", [""])[0],
                recipient=params.get("recipient", [""])[0],
                status=params.get("status", ["all"])[0],
                page=safe_int(params.get("page", ["1"])[0], 1),
                page_size=safe_int(params.get("page_size", ["50"])[0], 50),
            )
            self.send_json(payload)
            return

        if parsed.path.startswith("/api/customers/"):
            record_id = parsed.path.rsplit("/", 1)[-1]
            record = self.store.get_record(record_id)
            if record is None:
                self.send_error(HTTPStatus.NOT_FOUND, "未找到客户详情")
                return
            self.send_json(record)
            return

        if parsed.path in ("", "/"):
            self.path = "/index.html"
        else:
            requested = (STATIC_DIR / parsed.path.lstrip("/")).resolve()
            if not str(requested).startswith(str(STATIC_DIR.resolve())) or not requested.exists():
                self.path = "/index.html"

        super().do_GET()

    def do_POST(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path != "/api/reload":
            self.send_error(HTTPStatus.NOT_FOUND, "不支持的接口")
            return

        self.store.reload()
        self.send_json({"ok": True, "message": "Excel 数据已重新加载。"})

    def send_json(self, payload: dict) -> None:
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def log_message(self, fmt: str, *args) -> None:
        print(f"[CRM] {self.address_string()} - {fmt % args}")


def build_server(port: int, excel_path: str) -> ThreadingHTTPServer:
    store = ReminderStore(excel_path)
    CRMHandler.store = store
    return ThreadingHTTPServer(("127.0.0.1", port), CRMHandler)


def main() -> None:
    parser = argparse.ArgumentParser(description="CRM 续期提醒管理程序")
    parser.add_argument("--port", type=int, default=int(os.environ.get("CRM_PORT", "8765")))
    parser.add_argument("--excel", default=os.environ.get("CRM_EXCEL_PATH", DEFAULT_EXCEL_PATH))
    args = parser.parse_args()

    server = build_server(args.port, args.excel)
    print(f"CRM 提醒系统已启动: http://127.0.0.1:{args.port}")
    print(f"Excel 数据源: {args.excel}")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n服务已停止。")
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
