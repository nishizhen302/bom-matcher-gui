# -*- coding: utf-8 -*-
"""编码表记忆：拖一次存副本→下次启动自动载入；原文件改过→重新载入；原文件删了→用副本。"""
import os
import sys
import time

import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import app
import engine


@pytest.fixture
def env(tmp_path, monkeypatch):
    cfg = tmp_path / 'appdata'
    monkeypatch.setattr(app, 'CONFIG_DIR', str(cfg))
    monkeypatch.setattr(app, 'CONFIG_FILE', str(cfg / 'config.json'))
    monkeypatch.setattr(app, 'CODE_DIR', str(cfg / 'code_table'))
    monkeypatch.setattr(engine, 'load_code_lookup',
                        lambda p: ({'Y1': ('n', 's')} if 'bad' not in p else 1 / 0, set()))
    src = tmp_path / '编码表F27.xlsx'
    src.write_bytes(b'v1')
    return tmp_path, src


def test_remember_and_restore(env):
    tmp, src = env
    info = app.Api().set_file('code', str(src))
    assert '已记住' in info['meta']
    api2 = app.Api()                       # 模拟重启
    r = api2.restore_code()
    assert r['name'] == '编码表F27.xlsx' and '上次的编码表' in r['meta']
    assert open(api2.code_file, 'rb').read() == b'v1'
    assert os.path.dirname(api2.code_file) == app.CODE_DIR


def test_source_updated_reloads(env):
    tmp, src = env
    app.Api().set_file('code', str(src))
    src.write_bytes(b'v2')
    os.utime(src, (time.time() + 5, time.time() + 5))
    api2 = app.Api()
    assert '原文件有更新' in api2.restore_code()['meta']
    assert open(api2.code_file, 'rb').read() == b'v2'


def test_source_deleted_uses_copy(env):
    tmp, src = env
    app.Api().set_file('code', str(src))
    src.unlink()
    api2 = app.Api()
    assert api2.restore_code() is not None and os.path.isfile(api2.code_file)


def test_bad_code_table_not_remembered(env):
    tmp, src = env
    bad = tmp / 'bad.xlsx'
    bad.write_bytes(b'x')
    assert 'error' in app.Api().set_file('code', str(bad))
    assert app.Api().restore_code() is None


def test_excel_lock_file_detected(tmp_path):
    import excel_bridge
    f = tmp_path / 'BOM_匹配结果_1.xlsx'
    f.write_bytes(b'')
    assert not excel_bridge.lock_file_exists(str(f))
    assert excel_bridge.status(str(f)) == excel_bridge.CLOSED
    (tmp_path / '~$BOM_匹配结果_1.xlsx').write_bytes(b'')
    assert excel_bridge.lock_file_exists(str(f))
    # 有锁文件但 Excel 里并没打开（残留锁/WPS）→ 问不到，不能当成已保存
    assert excel_bridge.status(str(f)) == excel_bridge.UNKNOWN


def test_make_final_never_trusts_unsaved(tmp_path, monkeypatch):
    import excel_bridge
    a = app.Api()
    a.final_file = str(tmp_path / 'r.xlsx')
    monkeypatch.setattr(excel_bridge, 'status', lambda p: excel_bridge.DIRTY)
    saved = []
    monkeypatch.setattr(excel_bridge, 'save', lambda p: saved.append(p) or 'Excel 正在编辑单元格')
    assert a.make_final('check')['need'] == 'dirty'
    assert a.make_final('force')['need'] == 'dirty'      # 有改动时 force 也不放行
    assert a.make_final('save') == {'error': 'Excel 正在编辑单元格'} and saved
