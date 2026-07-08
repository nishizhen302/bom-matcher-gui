# -*- coding: utf-8 -*-
"""生产版输出（finalize）测试：模拟"导出→人工审核→生成生产版"全流程。"""
import os
import sys

import openpyxl
import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from bom_matcher import report
from bom_matcher.finalize import finalize, FINAL_HEADER


@pytest.fixture
def reviewed_file(tmp_path):
    """先用 report.write_excel 生成真实格式的结果文件，再模拟人工审核编辑。"""
    rows = [
        ['C1', '0.1uF', 'C0603', '1', '100-000123', '贴片电容', '0.1uF 50V X7R 0603', '', ''],
        ['C2', '0.1uF', 'C0603', '2', '100-000123', '贴片电容', '0.1uF 50V X7R 0603', '', ''],
        ['R1', '10K', 'R0603', '1', '100-000200', '贴片电阻', '10K 1% 0603', '', ''],
        ['H1', '安装孔', 'HOLE', '1', report.MECH_PN, '安装孔/测试点', '', '', '非采购物料，不参与匹配'],
        ['D5', 'NC', 'C0603', '1', report.UNMATCHED, report.UNMATCHED, report.UNMATCHED, '', '编码表无对应料'],
        ['C9', '0.1uF', 'C0805', '1', report.UNMATCHED, report.UNMATCHED, report.UNMATCHED, '', '编码表无对应料'],
        ['X1', '22pF', 'C0402', '1', report.UNMATCHED, report.UNMATCHED, report.UNMATCHED, '', '编码表无对应料'],
        ['AUTO', '锡膏', '', '3.2g', '100-000900', '锡膏', '有铅锡膏', '', '联动规则自动追加'],
    ]
    report.sort_rows(rows)
    result = str(tmp_path / '结果.xlsx')
    report.write_excel(rows, result)

    # 模拟人工审核：C9 手填料号（黄色底色还在）；追加一行 PCB 料号（无位号）
    wb = openpyxl.load_workbook(result)
    ws = wb.active
    for row in ws.iter_rows(min_row=2):
        if str(row[0].value) == 'C9':
            row[4].value, row[5].value, row[6].value = '100-000123', '贴片电容', '0.1uF 50V X7R 0603'
    pcb = ws.max_row + 1
    for col, val in ((4, '1'), (5, '100-000PCB'), (6, 'PCB'), (7, '4层 FR4 1.6mm')):
        ws.cell(row=pcb, column=col).value = val
    wb.save(result)
    return result


def _load(path):
    ws = openpyxl.load_workbook(path).active
    data = [[c.value for c in r] for r in ws.iter_rows(min_row=2)]
    return ws, data


def test_finalize(reviewed_file):
    out = finalize(reviewed_file, log=lambda m: None)
    assert out.endswith('_生产版.xlsx') and os.path.exists(out)
    ws, data = _load(out)

    assert [c.value for c in ws[1]] == FINAL_HEADER
    assert ws.freeze_panes == 'A2'
    assert ws.max_column == len(FINAL_HEADER)

    # 手动追加的 PCB 行排第一，序号从 1 连续编号
    assert data[0][1] == '100-000PCB' and data[0][0] == 1
    assert [r[0] for r in data] == list(range(1, len(data) + 1))

    by_pn = {r[1]: r for r in data}
    # C9 手填料号后并入同料号汇总：数量 1+2+1，位号合并
    agg = by_pn['100-000123']
    assert agg[4] == '4'
    assert sorted(agg[5].split(',')) == ['C1', 'C2', 'C9']
    assert by_pn['100-000200'][4] == '1'
    # 锡膏行不聚合，原样保留
    assert by_pn['100-000900'][4] == '3.2g'

    # NC 行、非采购项行、旧汇总行都不出现
    assert report.MECH_PN not in by_pn
    refs = ','.join(str(r[5]) for r in data)
    assert 'D5' not in refs and 'H1' not in refs

    # 漏填的 X1 保留在表尾且为黄色
    assert data[-1][1] == report.UNMATCHED and data[-1][5] == 'X1'
    last = list(ws.iter_rows(min_row=ws.max_row))[0]
    assert str(last[0].fill.start_color.rgb).endswith('FFFF00')
    # 其余行无底色
    first = list(ws.iter_rows(min_row=2, max_row=2))[0]
    assert first[0].fill.fill_type != 'solid'


def test_finalize_warns_leftover(reviewed_file, capsys):
    logs = []
    finalize(reviewed_file, log=logs.append)
    assert any('X1' in m and '未填料号' in m for m in logs)


def test_finalize_rejects_foreign_file(tmp_path):
    wb = openpyxl.Workbook()
    wb.active.append(['随便', '什么', '表'])
    path = str(tmp_path / 'foreign.xlsx')
    wb.save(path)
    with pytest.raises(ValueError):
        finalize(path, log=lambda m: None)
