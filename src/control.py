from __future__ import annotations

import ctypes
import os
import platform
import shutil
import sys
from dataclasses import dataclass
from pathlib import Path


DEFAULT_CONTROL_CODE = "Test"
CONTROL_DIR_NAME = "control"
DLL_NAME = "DRMSRelClient4Python-x64.dll"
DEFAULT_REL_FILE_NAME = "WH-OFDMaker-Rel.xml"
USER_REL_FILE_NAME = "User-Rel.xml"


@dataclass
class ControlCheckResult:
    ok: bool
    message: str = ""


def check_software_control(
    control_code: str = DEFAULT_CONTROL_CODE,
    rel_path: str | Path | None = None,
) -> ControlCheckResult:
    if platform.system().lower() != "windows":
        return ControlCheckResult(True, "非 Windows 环境，跳过授权 DLL 校验。")

    if os.environ.get("FILE2PPT_SKIP_LICENSE") == "1":
        return ControlCheckResult(True, "已通过环境变量跳过授权校验。")

    code = control_code.strip()
    if not code:
        return ControlCheckResult(False, "授权控制码为空，请联系管理员。")

    control_dir = _find_control_dir()
    if control_dir is None:
        return ControlCheckResult(False, "未找到授权控制文件目录 control。")

    dll_path = control_dir / DLL_NAME
    rel_file_path = Path(rel_path).expanduser().resolve() if rel_path else _find_rel_file(control_dir)
    if not dll_path.exists():
        return ControlCheckResult(False, f"未找到授权 DLL：{dll_path}")
    if rel_file_path is None or not rel_file_path.exists():
        return ControlCheckResult(False, "未找到授权控制文件 XML，请联系管理员。")

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

        obj = lib.RelChecker_Create(str(rel_file_path).encode("utf-8"))
        if not obj:
            return ControlCheckResult(False, "授权控制对象创建失败。")

        try:
            ok = bool(lib.RelChecker_Check(obj, code.encode("utf-8")))
        finally:
            lib.RelChecker_Destroy(obj)

        if not ok:
            return ControlCheckResult(False, "当前软件不可使用，请重新确认授权。")
        return ControlCheckResult(True, "授权校验通过。")
    except OSError as exc:
        return ControlCheckResult(False, f"DLL 加载失败：{exc}")
    except Exception as exc:
        return ControlCheckResult(False, f"授权校验失败：{exc}")


def get_default_rel_file() -> Path | None:
    control_dir = _find_control_dir()
    if control_dir is None:
        return None
    return _find_rel_file(control_dir)


def install_rel_file(source_path: str | Path) -> Path:
    source = Path(source_path).expanduser().resolve()
    if not source.exists():
        raise FileNotFoundError(f"授权文件不存在：{source}")

    control_dir = _find_control_dir()
    if control_dir is None:
        executable_dir = Path(sys.executable).resolve().parent
        control_dir = executable_dir / CONTROL_DIR_NAME
        control_dir.mkdir(parents=True, exist_ok=True)

    target = control_dir / USER_REL_FILE_NAME
    shutil.copy2(source, target)
    return target


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


def _find_rel_file(control_dir: Path) -> Path | None:
    user_rel = control_dir / USER_REL_FILE_NAME
    if user_rel.exists():
        return user_rel

    default_rel = control_dir / DEFAULT_REL_FILE_NAME
    if default_rel.exists():
        return default_rel

    return None
