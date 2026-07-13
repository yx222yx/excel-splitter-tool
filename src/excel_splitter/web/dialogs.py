from __future__ import annotations

from pathlib import Path


def choose_directory(initial_dir: Path) -> Path | None:
    import tkinter as tk
    from tkinter import filedialog

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    root.update()
    try:
        selected = filedialog.askdirectory(
            parent=root,
            initialdir=str(initial_dir),
            title="选择拆分结果输出目录",
            mustexist=True,
        )
    finally:
        root.destroy()
    return Path(selected) if selected else None
