# -*- mode: python ; coding: utf-8 -*-

# 构建前把当前 git 版本焊进 _build_info.py（版本号 = 提交数），供"检查更新"读取
import datetime as _dt
import subprocess as _sp


def _git(args):
    try:
        return _sp.run(['git'] + args, capture_output=True, text=True).stdout.strip()
    except Exception:
        return ''


with open('_build_info.py', 'w', encoding='utf-8') as _f:
    _f.write(
        f'VERSION = "{_git(["rev-list", "--count", "HEAD"]) or "?"}"\n'
        f'SHA = "{_git(["rev-parse", "--short", "HEAD"])}"\n'
        f'SHA_FULL = "{_git(["rev-parse", "HEAD"])}"\n'
        f'DATE = "{_dt.date.today().isoformat()}"\n')


a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=[],
    datas=[('gui', 'gui'), ('orcad_dsn.tcl', '.')],
    hiddenimports=['_build_info'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='BomMatcher',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['app.ico'],
)
