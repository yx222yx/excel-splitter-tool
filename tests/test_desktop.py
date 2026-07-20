from pathlib import Path
from types import SimpleNamespace

import pytest


class FakeServer:
    effective_port = 43123

    def __init__(self):
        self.ran = False
        self.closed = False
        self._map = {}
        self.close_all_calls = []
        self.asyncore = SimpleNamespace(
            close_all=lambda socket_map: self.close_all_calls.append(socket_map)
        )
        self.task_dispatcher = SimpleNamespace(shutdown_calls=[])

        def shutdown(**kwargs):
            self.task_dispatcher.shutdown_calls.append(kwargs)

        self.task_dispatcher.shutdown = shutdown

    def run(self):
        self.ran = True

    def close(self):
        self.closed = True


def test_desktop_server_starts_on_effective_port_and_stops(tmp_path: Path):
    from excel_splitter.desktop import DesktopServer

    fake_server = FakeServer()
    readiness_urls = []
    desktop_server = DesktopServer(
        app=object(),
        server_factory=lambda _app, **_kwargs: fake_server,
        readiness_probe=lambda url, **_kwargs: readiness_urls.append(url),
    )

    url = desktop_server.start()
    desktop_server.stop()
    desktop_server.stop()

    assert url == "http://127.0.0.1:43123/"
    assert readiness_urls == [url]
    assert fake_server.ran is True
    assert fake_server.closed is True
    assert fake_server.close_all_calls == [fake_server._map]
    assert fake_server.task_dispatcher.shutdown_calls == [
        {"cancel_pending": False, "timeout": 30}
    ]


class FakeEvent:
    def __init__(self):
        self.handlers = []

    def __iadd__(self, handler):
        self.handlers.append(handler)
        return self


class FakeWindow:
    def __init__(self):
        self.events = SimpleNamespace(closing=FakeEvent())


class FakeWebview:
    def __init__(self):
        self.settings = {}
        self.window_args = None
        self.start_args = None
        self.window = FakeWindow()

    def create_window(self, *args, **kwargs):
        self.window_args = (args, kwargs)
        return self.window

    def start(self, **kwargs):
        self.start_args = kwargs


class FakeDesktopServer:
    def __init__(self):
        self.stopped = False
        self.active_jobs = False

    def start(self):
        return "http://127.0.0.1:43123/"

    def stop(self):
        self.stopped = True

    def has_active_jobs(self):
        return self.active_jobs


def test_run_window_enables_downloads_and_always_stops_server():
    from excel_splitter.desktop import run_window

    webview = FakeWebview()
    server = FakeDesktopServer()

    run_window(server, webview_module=webview)

    args, kwargs = webview.window_args
    assert args == ("Excel 拆分工具", "http://127.0.0.1:43123/")
    assert kwargs["width"] == 1180
    assert kwargs["height"] == 800
    assert kwargs["min_size"] == (900, 640)
    assert webview.settings["ALLOW_DOWNLOADS"] is True
    assert webview.start_args == {
        "gui": "edgechromium",
        "debug": False,
        "private_mode": True,
    }
    assert server.stopped is True


def test_window_close_is_blocked_while_split_job_is_active(monkeypatch):
    from excel_splitter import desktop

    messages = []
    monkeypatch.setattr(
        desktop,
        "_show_message",
        lambda title, message, **_kwargs: messages.append((title, message)),
    )
    webview = FakeWebview()
    server = FakeDesktopServer()
    server.active_jobs = True

    desktop.run_window(server, webview_module=webview)
    close_handler = webview.window.events.closing.handlers[0]

    assert close_handler() is False
    assert messages and "正在处理" in messages[0][1]


def test_desktop_server_reports_active_background_jobs():
    from excel_splitter.desktop import DesktopServer

    app = SimpleNamespace(
        extensions={
            "excel_splitter_jobs": {
                "a": {"execution": {"status": "running"}},
                "b": {"execution": {"status": "complete"}},
            }
        }
    )

    server = DesktopServer(app, server_factory=lambda *_args, **_kwargs: None)

    assert server.has_active_jobs() is True
    app.extensions["excel_splitter_jobs"]["a"]["execution"]["status"] = "complete"
    assert server.has_active_jobs() is False


def test_main_shows_error_when_logging_initialization_fails(monkeypatch):
    from excel_splitter import desktop

    messages = []
    released = []

    class FakeInstance:
        def acquire(self):
            return True

        def release(self):
            released.append(True)

    monkeypatch.setattr(desktop, "SingleInstance", lambda _name: FakeInstance())
    monkeypatch.setattr(
        desktop,
        "configure_logging",
        lambda: (_ for _ in ()).throw(OSError("log denied")),
    )
    monkeypatch.setattr(
        desktop,
        "_show_message",
        lambda title, message, **kwargs: messages.append((title, message, kwargs)),
    )

    assert desktop.main() == 1
    assert messages and "log denied" in messages[0][1]
    assert released == [True]


def test_run_window_stops_server_when_window_creation_fails():
    from excel_splitter.desktop import run_window

    class FailingWebview(FakeWebview):
        def create_window(self, *args, **kwargs):
            raise RuntimeError("window failed")

    server = FakeDesktopServer()

    with pytest.raises(RuntimeError, match="window failed"):
        run_window(server, webview_module=FailingWebview())

    assert server.stopped is True

def test_main_calls_freeze_support_before_desktop_initialization(monkeypatch):
    from excel_splitter import desktop

    calls = []

    class FakeInstance:
        def acquire(self):
            calls.append("acquire")
            return False

        def release(self):
            calls.append("release")

    monkeypatch.setattr(
        desktop,
        "freeze_support",
        lambda: calls.append("freeze_support"),
        raising=False,
    )
    monkeypatch.setattr(desktop, "SingleInstance", lambda _name: FakeInstance())
    monkeypatch.setattr(
        desktop,
        "configure_logging",
        lambda: calls.append("configure_logging"),
    )
    monkeypatch.setattr(desktop, "_show_message", lambda *_args, **_kwargs: None)

    assert desktop.main() == 0
    assert calls[0] == "freeze_support"
