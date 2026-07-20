from __future__ import annotations

import logging
from logging.handlers import RotatingFileHandler
from multiprocessing import freeze_support
from pathlib import Path
from threading import Thread
from typing import Any, Callable

from .desktop_runtime import (
    SingleInstance,
    log_file_path,
    user_output_dir,
    wait_until_ready,
)
from .web.app import create_app


HOST = "127.0.0.1"
MUTEX_NAME = r"Local\ExcelSplitterTool"
WINDOW_TITLE = "Excel 拆分工具"


def _create_waitress_server(app, **kwargs):
    from waitress.server import create_server

    return create_server(app, **kwargs)


class DesktopServer:
    def __init__(
        self,
        app,
        *,
        server_factory: Callable[..., Any] = _create_waitress_server,
        readiness_probe: Callable[..., None] = wait_until_ready,
    ):
        self._app = app
        self._server_factory = server_factory
        self._readiness_probe = readiness_probe
        self._server = None
        self._thread: Thread | None = None

    def start(self) -> str:
        if self._server is not None:
            return self.url
        self._server = self._server_factory(
            self._app,
            host=HOST,
            port=0,
            threads=4,
            ident="ExcelSplitter",
        )
        self.url = f"http://{HOST}:{int(self._server.effective_port)}/"
        self._thread = Thread(
            target=self._server.run,
            name="excel-splitter-http",
            daemon=True,
        )
        self._thread.start()
        try:
            self._readiness_probe(self.url, timeout=15)
        except Exception:
            self.stop()
            raise
        return self.url

    def stop(self) -> None:
        if self._server is None:
            return
        server = self._server
        thread = self._thread
        dispatcher = getattr(server, "task_dispatcher", None)
        if dispatcher is not None:
            dispatcher.shutdown(cancel_pending=False, timeout=30)
        server.close()
        socket_map = getattr(server, "_map", None)
        asyncore = getattr(server, "asyncore", None)
        if socket_map is not None and asyncore is not None:
            asyncore.close_all(socket_map)
        if thread is not None and thread.is_alive():
            thread.join(timeout=5)
        if thread is not None and thread.is_alive():
            raise RuntimeError("本地服务未能正常停止")
        self._server = None
        self._thread = None

    def has_active_jobs(self) -> bool:
        jobs = self._app.extensions.get("excel_splitter_jobs", {})
        return any(
            record.get("execution", {}).get("status") in {"queued", "running"}
            for record in jobs.values()
        )


def run_window(server: DesktopServer, *, webview_module=None) -> None:
    if webview_module is None:
        import webview as webview_module

        _require_webview2()

    url = server.start()
    try:
        webview_module.settings["ALLOW_DOWNLOADS"] = True
        window = webview_module.create_window(
            WINDOW_TITLE,
            url,
            width=1180,
            height=800,
            min_size=(900, 640),
            resizable=True,
            background_color="#f4f6f8",
            text_select=True,
        )

        def guard_active_job() -> bool | None:
            if server.has_active_jobs():
                _show_message(
                    WINDOW_TITLE,
                    "拆分任务正在处理，请等待任务完成后再关闭程序。",
                )
                return False
            return None

        window.events.closing += guard_active_job
        webview_module.start(
            gui="edgechromium",
            debug=False,
            private_mode=True,
        )
    finally:
        server.stop()


def configure_logging(path: Path | None = None) -> Path:
    target = path or log_file_path()
    target.parent.mkdir(parents=True, exist_ok=True)
    handler = RotatingFileHandler(
        target,
        maxBytes=2 * 1024 * 1024,
        backupCount=3,
        encoding="utf-8",
    )
    handler.setFormatter(
        logging.Formatter("%(asctime)s %(levelname)s %(name)s %(message)s")
    )
    root_logger = logging.getLogger()
    root_logger.setLevel(logging.INFO)
    root_logger.addHandler(handler)
    return target


def _require_webview2() -> None:
    from webview.platforms import winforms

    if winforms.renderer != "edgechromium":
        raise RuntimeError(
            "未检测到 Microsoft Edge WebView2 Runtime，请安装后重新启动。"
        )


def _show_message(title: str, message: str, *, error: bool = False) -> None:
    import ctypes

    flags = 0x10 if error else 0x40
    ctypes.windll.user32.MessageBoxW(None, message, title, flags)


def main() -> int:
    freeze_support()
    instance = SingleInstance(MUTEX_NAME)
    try:
        configure_logging()
        if not instance.acquire():
            _show_message(WINDOW_TITLE, "Excel 拆分工具已经在运行。")
            return 0
        output_dir = user_output_dir()
        output_dir.mkdir(parents=True, exist_ok=True)
        app = create_app({"DEFAULT_OUTPUT_DIR": output_dir})
        run_window(DesktopServer(app))
        return 0
    except Exception as exc:
        logging.getLogger(__name__).exception("Desktop application failed")
        _show_message(WINDOW_TITLE, f"程序启动失败：\n{exc}", error=True)
        return 1
    finally:
        instance.release()


if __name__ == "__main__":
    raise SystemExit(main())
