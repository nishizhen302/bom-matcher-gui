# -*- coding: utf-8 -*-
"""OrCAD 原理图 (.DSN) → BOM，免去在 OrCAD 里手动导出 Excel。

借用本机 OrCAD 自带的 tclsh + DBO 接口（orDb_Dll_Tcl64.dll）在界面外读 DSN，
产出与 OrCAD「Bill of Materials」默认导出一致的行：
按 (Value, PCB Footprint) 分组，位号自然排序，组按首位号排序。
2026-10-18 用 VBO_EDP.DSN 对照 OrCAD 原始导出 VBO_EDP.xlsx 逐格一致（32 行含顺序）。
"""
import glob
import os
import re
import shutil
import subprocess
import sys
import tempfile

import openpyxl
from openpyxl.styles import Font

TCL_SCRIPT = 'orcad_dsn.tcl'
CADENCE_GLOBS = [r'C:\Cadence\*\tools\bin', r'D:\Cadence\*\tools\bin']


class DsnError(Exception):
    """读 DSN 失败，消息直接给用户看。"""


def _resource(rel):
    base = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, rel)


def find_orcad():
    """返回 (tclsh.exe, orDb_Dll_Tcl64.dll)；找不到返回 None。"""
    dirs = []
    root = os.environ.get('CDSROOT')
    if root:
        dirs.append(os.path.join(root, 'tools', 'bin'))
    for g in CADENCE_GLOBS:
        dirs += sorted(glob.glob(g), reverse=True)
    for d in dirs:
        tclsh = os.path.join(d, 'tclsh.exe')
        dll = os.path.join(d, 'orDb_Dll_Tcl64.dll')
        if os.path.isfile(tclsh) and os.path.isfile(dll):
            return tclsh, dll
    return None


def _natkey(s):
    return [int(t) if t.isdigit() else t.upper() for t in re.split(r'(\d+)', s)]


def read_parts(dsn_path):
    """读出所有元件 → [(位号, 值, 封装)]，多 section 器件（U14 拆成多片）已按位号去重。"""
    if not os.path.isfile(dsn_path):
        raise DsnError('原理图文件不存在')
    orcad = find_orcad()
    if not orcad:
        raise DsnError('没找到 OrCAD（需要 C:\\Cadence\\SPB_xx\\tools\\bin 下的 tclsh.exe 和 '
                       'orDb_Dll_Tcl64.dll）。没装 OrCAD 的电脑请改拖 BOM Excel')
    tclsh, dll = orcad
    work = tempfile.mkdtemp(prefix='bomdsn_')
    try:
        # 拷到纯英文临时路径再读：避开中文路径，也不碰 OrCAD 正开着的原文件
        src = os.path.join(work, 'design.DSN')
        shutil.copyfile(dsn_path, src)
        out = os.path.join(work, 'parts.tsv')
        flags = getattr(subprocess, 'CREATE_NO_WINDOW', 0)
        try:
            p = subprocess.run([tclsh, _resource(TCL_SCRIPT), dll, src, out],
                               capture_output=True, timeout=120, creationflags=flags)
        except subprocess.TimeoutExpired:
            raise DsnError('读原理图超时（120 秒）')
        if p.returncode != 0 or not os.path.isfile(out):
            err = (p.stderr or b'').decode('gbk', 'replace').strip()
            raise DsnError('OrCAD 打不开这个原理图（可能是更高版本 OrCAD 保存的，或文件已损坏）'
                           + (f'：{err[-200:]}' if err else ''))
        with open(out, encoding='utf-8', errors='replace') as f:
            lines = f.read().splitlines()
    finally:
        shutil.rmtree(work, ignore_errors=True)

    parts = {}
    conflicts = []
    for line in lines:
        cols = line.split('\t')
        if cols[0] != 'P' or len(cols) < 4:
            continue
        ref, val, fp = cols[1].strip(), cols[2], cols[3]
        if ref in parts:
            if parts[ref] != (val, fp):
                conflicts.append(ref)
            continue
        parts[ref] = (val, fp)

    # 取不到 ≠ 取到空：读出 0 个元件当失败处理，不出一份空 BOM
    if not parts:
        raise DsnError('原理图里没读到任何元件')
    bad = sorted((r for r in parts if not r or '?' in r), key=_natkey)
    if bad:
        raise DsnError(f'原理图还没编号（有 {len(bad)} 个元件位号带 ?，如 {bad[0] or "空"}），'
                       '请先在 OrCAD 里 Annotate 并保存')
    if conflicts:
        raise DsnError('同一位号的几片值或封装不一致：' + ','.join(sorted(set(conflicts), key=_natkey)[:10]))
    return [(r, v, fp) for r, (v, fp) in parts.items()]


def group_rows(parts):
    """→ [(位号串, 值, 封装, 数量)]，与 OrCAD 默认 BOM 导出同序。"""
    groups = {}
    for ref, val, fp in parts:
        groups.setdefault((val, fp), []).append(ref)
    rows = [(','.join(sorted(refs, key=_natkey)), val, fp, len(refs))
            for (val, fp), refs in groups.items()]
    rows.sort(key=lambda r: _natkey(r[0].split(',')[0]))
    return rows


def write_bom_xlsx(rows, out_path):
    """写成 OrCAD 导出的列布局：Item / Quantity / Reference / Part / PCB Footprint。"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'BOM'
    ws.append(['Item', 'Quantity', 'Reference', 'Part', 'PCB Footprint'])
    for c in ws[1]:
        c.font = Font(bold=True)
    for i, (refs, val, fp, qty) in enumerate(rows, 1):
        ws.append([i, qty, refs, val, fp])
    for col, w in zip('ABCDE', (6, 9, 60, 24, 24)):
        ws.column_dimensions[col].width = w
    wb.save(out_path)


def bom_path_for(dsn_path):
    stem = os.path.splitext(os.path.abspath(dsn_path))[0]
    return f'{stem}_BOM(由原理图生成).xlsx'


def dsn_to_bom(dsn_path, out_path=None):
    """读 DSN → 在原理图旁写 BOM Excel。返回 (xlsx路径, 元件数, 行数)。"""
    parts = read_parts(dsn_path)
    rows = group_rows(parts)
    out_path = out_path or bom_path_for(dsn_path)
    try:
        write_bom_xlsx(rows, out_path)
    except PermissionError:
        raise DsnError(f'写不进 {os.path.basename(out_path)}，它是不是正被 Excel 打开着？关掉再试')
    return out_path, len(parts), len(rows)
