from __future__ import annotations

from io import BytesIO
from pathlib import Path

from msoffcrypto import OfficeFile
from msoffcrypto.format.ooxml import OOXMLFile


def is_encrypted(path: Path) -> bool:
    """快速检查文件是否已加密。"""
    try:
        with open(path, "rb") as f:
            of = OfficeFile(f)
            return of.is_encrypted()
    except Exception:
        return False


def decrypt_file(path: Path, password: str) -> BytesIO:
    """解密加密的 xlsx 文件，返回解密后的字节流。"""
    output = BytesIO()
    with open(path, "rb") as f:
        of = OfficeFile(f)
        of.load_key(password=password)
        of.decrypt(output)
    output.seek(0)
    return output


def encrypt_file(source: Path, password: str) -> None:
    """将未加密的 xlsx 文件加密，原地替换为加密版本。"""
    temp = source.with_name(source.name + ".encrypted")
    try:
        with open(source, "rb") as f:
            of = OOXMLFile(f)
            with open(temp, "wb") as out:
                of.encrypt(password, out)
        temp.replace(source)
    except Exception:
        temp.unlink(missing_ok=True)
        raise
