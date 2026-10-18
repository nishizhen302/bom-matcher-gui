# -*- coding: utf-8 -*-
"""问正在运行的 Excel：某个文件开着没有、有没有没保存的改动，必要时替用户按一次保存。

只在系统「运行对象表」(ROT) 里找已打开的工作簿，绝不启动 Excel、也不替用户打开文件。
Excel 打开工作簿时会把完整路径登记进 ROT；WPS 等找不到时返回 'unknown'，由界面兜底。
"""
import os

CLOSED, SAVED, DIRTY, UNKNOWN = 'closed', 'saved', 'dirty', 'unknown'


def lock_file_exists(path):
    """Excel/WPS 打开文件时会在同目录放 ~$文件名 锁文件。"""
    d, base = os.path.split(os.path.abspath(path))
    try:
        return any(n.startswith('~$') and n[2:] in (base, base[2:]) for n in os.listdir(d))
    except OSError:
        return False


def _find_workbook(path):
    import pythoncom
    import win32com.client
    target = os.path.normcase(os.path.abspath(path))
    rot = pythoncom.GetRunningObjectTable()
    ctx = pythoncom.CreateBindCtx(0)
    for mk in rot:
        try:
            name = mk.GetDisplayName(ctx, None)
        except pythoncom.com_error:
            continue
        if os.path.normcase(name) == target:
            obj = rot.GetObject(mk)
            return win32com.client.Dispatch(obj.QueryInterface(pythoncom.IID_IDispatch))
    return None


def _com(fn):
    """在当前线程初始化 COM 再调用（pywebview 的 js_api 跑在工作线程里）。"""
    try:
        import pythoncom
    except ImportError:
        return None
    pythoncom.CoInitialize()
    try:
        return fn()
    finally:
        pythoncom.CoUninitialize()


def status(path):
    """closed=没开着 / saved=开着但已保存 / dirty=有没保存的改动 / unknown=有锁文件但问不到 Excel"""
    if not lock_file_exists(path):
        return CLOSED

    def ask():
        wb = _find_workbook(path)
        if wb is None:
            return UNKNOWN
        return SAVED if wb.Saved else DIRTY
    try:
        return _com(ask) or UNKNOWN
    except Exception:
        return UNKNOWN


def save(path):
    """替用户在 Excel 里保存该工作簿。返回 None 表示成功，否则返回给用户看的原因。"""
    def do():
        wb = _find_workbook(path)
        if wb is None:
            return '在 Excel 里没找到这个文件，请手动保存'
        wb.Save()
        return None if wb.Saved else 'Excel 没有保存成功，请手动保存'
    try:
        import pythoncom  # noqa: F401
    except ImportError:
        return '缺少 Excel 通信组件，请手动保存'
    try:
        return _com(do)
    except Exception as e:
        if 'rejected' in str(e).lower() or '-2147418111' in str(e):
            return 'Excel 正在编辑单元格，先按回车或 Esc 退出编辑再点'
        return f'保存失败：{e}'
