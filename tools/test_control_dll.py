from __future__ import annotations

import ctypes
import platform
from pathlib import Path


CONTROL_CODE = b"Test"
CONTROL_DIR = Path(__file__).resolve().parent.parent / "contrl" / "软件控制所需文件"
DLL_PATH = CONTROL_DIR / "DRMSRelClient4Python-x64.dll"
REL_PATH = CONTROL_DIR / "WH-OFDMaker-Rel.xml"


def main() -> int:
    print(f"system: {platform.system()}")
    print(f"dll: {DLL_PATH}")
    print(f"rel: {REL_PATH}")
    print(f"control_code: {CONTROL_CODE!r}")

    if not DLL_PATH.exists():
        print("FAIL: DLL 文件不存在")
        return 1
    if not REL_PATH.exists():
        print("FAIL: 授权 XML 文件不存在")
        return 1
    if platform.system().lower() != "windows":
        print("SKIP: 当前不是 Windows，无法实际加载 PE 格式 DLL。请在 Windows 上运行本脚本做真实校验。")
        return 0

    try:
        if hasattr(__import__("os"), "add_dll_directory"):
            __import__("os").add_dll_directory(str(CONTROL_DIR))
        lib = ctypes.CDLL(str(DLL_PATH))
    except OSError as exc:
        print(f"FAIL: DLL 加载失败: {exc}")
        return 1

    lib.RelChecker_Create.argtypes = [ctypes.c_char_p]
    lib.RelChecker_Create.restype = ctypes.c_void_p
    lib.RelChecker_Check.argtypes = [ctypes.c_void_p, ctypes.c_char_p]
    lib.RelChecker_Check.restype = ctypes.c_bool
    lib.RelChecker_Destroy.argtypes = [ctypes.c_void_p]
    lib.RelChecker_Destroy.restype = None

    obj = lib.RelChecker_Create(str(REL_PATH).encode("utf-8"))
    if not obj:
        print("FAIL: RelChecker_Create 返回空对象")
        return 1

    try:
        ok = bool(lib.RelChecker_Check(obj, CONTROL_CODE))
    finally:
        lib.RelChecker_Destroy(obj)

    print(f"RelChecker_Check: {ok}")
    if ok:
        print("PASS: 授权校验通过")
        return 0

    print("FAIL: 授权校验未通过")
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
