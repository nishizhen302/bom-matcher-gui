# -*- coding: utf-8 -*-
"""DSN → BOM：分组规则单测 + 真实原理图对照 OrCAD 原始导出（样本不进仓库，缺了就跳过）。"""
import os
import sys
from pathlib import Path

import openpyxl
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import dsn_reader as D

# 样本：同一块板的 DSN + 在 OrCAD 里手动导出的原始 BOM（板子资料，不提交到公开仓库）
SAMPLE_DIR = Path(os.environ.get('BOM_DSN_SAMPLE_DIR',
                                 Path.home() / 'Desktop' / 'VBO_EDP'))
SAMPLE_DSN = SAMPLE_DIR / 'VBO_EDP.DSN'
SAMPLE_ORCAD_BOM = SAMPLE_DIR / 'VBO_EDP.xlsx'


def test_group_rows_orcad_order():
    parts = [('R10', '10K', 'R0402'), ('C2', '100nF', 'C0402'), ('R2', '10K', 'R0402'),
             ('C10', '100nF', 'C0402'), ('C1', '1uF', 'C0402'), ('R1', '10K', 'R0603')]
    assert D.group_rows(parts) == [
        ('C1', '1uF', 'C0402', 1),
        ('C2,C10', '100nF', 'C0402', 2),     # 自然排序：C2 在 C10 前
        ('R1', '10K', 'R0603', 1),           # 同值不同封装分两行
        ('R2,R10', '10K', 'R0402', 2),
    ]


def test_write_bom_loadable_by_engine(tmp_path):
    from bom_matcher.loader import load_bom
    out = tmp_path / 'b.xlsx'
    D.write_bom_xlsx([('R1,R2', '10K', 'R0402', 2)], out)
    rows, cols = load_bom(str(out))
    r = rows[0]
    assert (r[cols['ref']], r[cols['part']], r[cols['fp']], r[cols['qty']]) == \
        ('R1,R2', '10K', 'R0402', 2)


needs_sample = pytest.mark.skipif(
    not (SAMPLE_DSN.is_file() and SAMPLE_ORCAD_BOM.is_file() and D.find_orcad()),
    reason='缺 VBO_EDP 样本或本机没装 OrCAD')


@needs_sample
def test_vbo_edp_matches_orcad_export():
    ws = openpyxl.load_workbook(SAMPLE_ORCAD_BOM).active
    orcad = [(str(r[2]), str(r[3]), str(r[4]), int(r[1]))
             for r in ws.iter_rows(values_only=True) if isinstance(r[0], int)]
    mine = D.group_rows(D.read_parts(str(SAMPLE_DSN)))
    assert len(orcad) == 32
    assert mine == orcad


def test_missing_file_is_error(tmp_path):
    with pytest.raises(D.DsnError):
        D.read_parts(str(tmp_path / 'nope.DSN'))


@needs_sample
def test_garbage_dsn_is_error_not_empty_bom(tmp_path):
    bad = tmp_path / 'bad.DSN'
    bad.write_bytes(b'not an orcad design')
    with pytest.raises(D.DsnError):
        D.read_parts(str(bad))
