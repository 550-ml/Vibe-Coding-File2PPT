from __future__ import annotations

import ctypes
import os
import platform
import sys
from dataclasses import dataclass
from pathlib import Path


CONTROL_CODE = b"Test"
CONTROL_DIR_NAME = "control"
DLL_NAME = "DRMSRelClient4Python-x64.dll"
REL_FILE_NAME = "WH-OFDMaker-Rel.xml"


@dataclass
class ControlCheckResult:
    ok: bool
    message: str = ""


def check_software_control() -> ControlCheckResult:
    if platform.system().lower() != "windows":
        return ControlCheckResult(True, "非 Windows 环境，跳过授权 DLL 校验。")

    if os.environ.get("FILE2PPT_SKIP_LICENSE") == "1":
        return ControlCheckResult(True, "已通过环境变量跳过授权校验。")

    control_dir = _find_control_dir()
    if control_dir is None:
        return ControlCheckResult(False, "未找到授权控制文件目录 control。")

    dll_path = control_dir / DLL_NAME
    rel_path = control_dir / REL_FILE_NAME
    if not dll_path.exists():
        return ControlCheckResult(False, f"未找到授权 DLL：{dll_path}")
    if not rel_path.exists():
        return ControlCheckResult(False, f"未找到授权文件：{rel_path}")

    try:
        if hasattr(os, "add_dll_directory"):
            os.add_dll_directory(str(control_dir))
        lib = ctypes.CDLL(str(dll_path))
        lib.RelChecker_Create.argtypes = [ctypes.c_char_p]
        lib.RelChecker_Create.restype = ctypes.c_void_p
        lib.RelChecker_Check.argtypes = [ctypes.c_void_p, ctypes.c_char_p]
        lib.RelChecker_Check.restype = ctypes.c_bool
        lib.RelChecker_Destroy.argtypes = [ctypes.c_void_p]
        lib.RelChecker_Destroy.restype = None

        obj = lib.RelChecker_Create(str(rel_path).encode("utf-8"))
        if not obj:
            return ControlCheckResult(False, "授权控制对象创建失败。")

        try:
            ok = bool(lib.RelChecker_Check(obj, CONTROL_CODE))
        finally:
            lib.RelChecker_Destroy(obj)

        if not ok:
            return ControlCheckResult(False, "当前软件不可使用，请重新确认授权。")
        return ControlCheckResult(True, "授权校验通过。")
    except OSError as exc:
        return ControlCheckResult(False, f"DLL 加载失败：{exc}")
    except Exception as exc:
        return ControlCheckResult(False, f"授权校验失败：{exc}")


def _find_control_dir() -> Path | None:
    candidates = []
    if hasattr(sys, "_MEIPASS"):
        candidates.append(Path(sys._MEIPASS) / CONTROL_DIR_NAME)

    executable_dir = Path(sys.executable).resolve().parent
    candidates.append(executable_dir / CONTROL_DIR_NAME)
    candidates.append(executable_dir / "_internal" / CONTROL_DIR_NAME)

    project_root = Path(__file__).resolve().parent.parent
    candidates.append(project_root / CONTROL_DIR_NAME)
    candidates.append(project_root / "contrl" / "软件控制所需文件")

    for candidate in candidates:
        if candidate.exists():
            return candidate
    return None
