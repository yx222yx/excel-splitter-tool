from __future__ import annotations

import ctypes
from ctypes import wintypes
import os
from pathlib import Path
import time
from typing import Callable
from urllib.request import urlopen


ERROR_ALREADY_EXISTS = 183


def user_output_dir() -> Path:
    return Path.home() / "Desktop"


def log_file_path() -> Path:
    local_app_data = Path(
        os.environ.get("LOCALAPPDATA", Path.home() / "AppData" / "Local")
    )
    return local_app_data / "ExcelSplitter" / "logs" / "excel-splitter.log"


class SingleInstance:
    def __init__(self, name: str, *, kernel32=None):
        self.name = name
        self._kernel32 = kernel32
        self._handle = None
        self._owns_mutex = False

    def acquire(self) -> bool:
        if self._handle is not None:
            return self._owns_mutex
        kernel32 = self._kernel32 or ctypes.windll.kernel32
        _configure_kernel32(kernel32)
        handle = kernel32.CreateMutexW(None, True, self.name)
        if not handle:
            raise ctypes.WinError()
        self._kernel32 = kernel32
        self._handle = handle
        if kernel32.GetLastError() == ERROR_ALREADY_EXISTS:
            kernel32.CloseHandle(handle)
            self._handle = None
            return False
        self._owns_mutex = True
        return True

    def release(self) -> None:
        if self._handle is None:
            return
        if self._owns_mutex:
            self._kernel32.ReleaseMutex(self._handle)
        self._kernel32.CloseHandle(self._handle)
        self._handle = None
        self._owns_mutex = False

    def __enter__(self):
        if not self.acquire():
            raise RuntimeError("Excel 拆分工具已经在运行")
        return self

    def __exit__(self, _exc_type, _exc, _traceback):
        self.release()


def _configure_kernel32(kernel32) -> None:
    signatures = (
        (
            kernel32.CreateMutexW,
            [wintypes.LPVOID, wintypes.BOOL, wintypes.LPCWSTR],
            wintypes.HANDLE,
        ),
        (kernel32.GetLastError, [], wintypes.DWORD),
        (kernel32.ReleaseMutex, [wintypes.HANDLE], wintypes.BOOL),
        (kernel32.CloseHandle, [wintypes.HANDLE], wintypes.BOOL),
    )
    for function, argtypes, restype in signatures:
        try:
            function.argtypes = argtypes
            function.restype = restype
        except AttributeError:
            # Lightweight injected test doubles can expose bound methods.
            pass

def wait_until_ready(
    url: str,
    *,
    timeout: float = 15,
    interval: float = 0.1,
    probe: Callable[[str], None] | None = None,
    clock: Callable[[], float] = time.monotonic,
    sleeper: Callable[[float], None] = time.sleep,
) -> None:
    probe = probe or _probe_url
    deadline = clock() + timeout
    last_error: Exception | None = None
    while True:
        try:
            probe(url)
            return
        except Exception as exc:
            last_error = exc
        if clock() >= deadline:
            raise TimeoutError("本地服务启动超时") from last_error
        sleeper(interval)


def _probe_url(url: str) -> None:
    with urlopen(url, timeout=1) as response:
        if response.status >= 500:
            raise OSError(f"本地服务返回状态码 {response.status}")
