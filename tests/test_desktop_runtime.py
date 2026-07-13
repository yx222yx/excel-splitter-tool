from pathlib import Path
from ctypes import wintypes

import pytest


def test_user_paths_use_writable_windows_locations(monkeypatch, tmp_path: Path):
    from excel_splitter.desktop_runtime import log_file_path, user_output_dir

    profile = tmp_path / "profile"
    local_app_data = tmp_path / "local"
    monkeypatch.setenv("USERPROFILE", str(profile))
    monkeypatch.setenv("LOCALAPPDATA", str(local_app_data))

    assert user_output_dir() == profile / "Documents" / "Excel拆分工具输出"
    assert log_file_path() == local_app_data / "ExcelSplitter" / "logs" / "excel-splitter.log"


class FakeKernel32:
    def __init__(self, last_error: int = 0):
        self.last_error = last_error
        self.released = []
        self.closed = []

    def CreateMutexW(self, _security, _owner, _name):
        return 42

    def GetLastError(self):
        return self.last_error

    def ReleaseMutex(self, handle):
        self.released.append(handle)
        return True

    def CloseHandle(self, handle):
        self.closed.append(handle)
        return True


class FakeFunction:
    def __init__(self, result):
        self.result = result
        self.argtypes = None
        self.restype = None

    def __call__(self, *_args):
        return self.result


class SignatureKernel32:
    def __init__(self):
        self.CreateMutexW = FakeFunction(42)
        self.GetLastError = FakeFunction(0)
        self.ReleaseMutex = FakeFunction(True)
        self.CloseHandle = FakeFunction(True)


def test_single_instance_acquires_and_releases_mutex():
    from excel_splitter.desktop_runtime import SingleInstance

    kernel32 = FakeKernel32()
    instance = SingleInstance("Local\\ExcelSplitterTool", kernel32=kernel32)

    assert instance.acquire() is True
    instance.release()

    assert kernel32.released == [42]
    assert kernel32.closed == [42]


def test_single_instance_rejects_existing_mutex():
    from excel_splitter.desktop_runtime import ERROR_ALREADY_EXISTS, SingleInstance

    kernel32 = FakeKernel32(last_error=ERROR_ALREADY_EXISTS)
    instance = SingleInstance("Local\\ExcelSplitterTool", kernel32=kernel32)

    assert instance.acquire() is False
    assert kernel32.closed == [42]


def test_single_instance_configures_windows_handle_signatures():
    from excel_splitter.desktop_runtime import SingleInstance

    kernel32 = SignatureKernel32()
    instance = SingleInstance("Local\\ExcelSplitterTool", kernel32=kernel32)

    assert instance.acquire() is True

    assert kernel32.CreateMutexW.restype is wintypes.HANDLE
    assert kernel32.ReleaseMutex.argtypes == [wintypes.HANDLE]
    assert kernel32.CloseHandle.argtypes == [wintypes.HANDLE]
    instance.release()


def test_wait_until_ready_retries_until_probe_succeeds():
    from excel_splitter.desktop_runtime import wait_until_ready

    attempts = []

    def probe(_url):
        attempts.append(True)
        if len(attempts) < 3:
            raise OSError("not ready")

    now = iter([0.0, 0.1, 0.2, 0.3])
    wait_until_ready(
        "http://127.0.0.1:1234/",
        timeout=1,
        interval=0,
        probe=probe,
        clock=lambda: next(now),
        sleeper=lambda _seconds: None,
    )

    assert len(attempts) == 3


def test_wait_until_ready_raises_after_timeout():
    from excel_splitter.desktop_runtime import wait_until_ready

    now = iter([0.0, 0.1, 1.1])

    with pytest.raises(TimeoutError, match="本地服务启动超时"):
        wait_until_ready(
            "http://127.0.0.1:1234/",
            timeout=1,
            interval=0,
            probe=lambda _url: (_ for _ in ()).throw(OSError("not ready")),
            clock=lambda: next(now),
            sleeper=lambda _seconds: None,
        )
