# -*- coding: utf-8 -*-
"""输出层：排序、同料号聚合、汇总行、着色、列宽行高。格式与旧版一致，新增第9列"匹配说明"。"""
import math
import re
import openpyxl

from .parsers import NC_RE

UNMATCHED = '未匹配'
MECH_PN = '非采购项'

HEADER = ['位号', 'Part Value', 'PCB Footprint', 'Quantity', '料号', '名称', '规格', '候选料号', '匹配说明']
N_COLS = len(HEADER)


def _is_nc(part):
    return bool(NC_RE.search(str(part).upper()))


def sort_rows(rows):
    """稳定排序：正常(按料号) → 非采购项 → 未匹配 → NC"""
    rows.sort(key=lambda r: (
        _is_nc(r[1]),
        r[4] == UNMATCHED,
        r[4],
    ))
    return rows


def _parse_qty(value):
    m = re.search(r'\d+(?:\.\d+)?', str(value or ''))
    return float(m.group(0)) if m else 0.0


def add_companions(rows, companions, by_pn, log):
    for trigger_pn, (add_pn, per_qty, label) in companions.items():
        total = sum(_parse_qty(r[3]) for r in rows if r[4] == trigger_pn)
        if total <= 0 or any(r[4] == add_pn for r in rows):
            continue
        e = by_pn.get(add_pn)
        if not e:
            log(f"Warning: 匹配到 {trigger_pn}，但编码表未找到可用的 {add_pn}")
            continue
        qty_text = f"{total * per_qty:.3f}".rstrip('0').rstrip('.') + 'g'
        rows.append(['AUTO', label, '', qty_text, e.pn, e.name, e.spec, '', '联动规则自动追加'])
        log(f"自动添加 {add_pn}: {total:g} * {per_qty}g = {qty_text}")
    return rows


def write_excel(rows, output_file):
    aggregatable = lambda r: (not _is_nc(r[1])) and r[4] not in (UNMATCHED, MECH_PN)
    pn_counts = {}
    for r in rows:
        if aggregatable(r):
            pn_counts[r[4]] = pn_counts.get(r[4], 0) + 1
    repeated = {pn for pn, c in pn_counts.items() if c > 1}

    write_rows = []  # (type, data, is_repeated)
    group = []

    def flush(g):
        if len(g) > 1 and aggregatable(g[0]):
            total = sum(int(str(x[3])) for x in g if str(x[3]).isdigit())
            refs = ','.join(str(x[0]) for x in g)
            write_rows.append(('summary',
                               [refs, '', '', str(total), g[0][4], g[0][5], g[0][6], '', ''], True))

    for r in rows:
        if not group or group[-1][4] != r[4]:
            flush(group)
            group = [r]
        else:
            group.append(r)
        write_rows.append(('data', r, r[4] in repeated))
    flush(group)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(HEADER)
    thin = openpyxl.styles.Side(style='thin')
    border = openpyxl.styles.Border(left=thin, right=thin, top=thin, bottom=thin)
    fills = {
        'yellow': openpyxl.styles.PatternFill('solid', start_color='FFFF00'),
        'blue': openpyxl.styles.PatternFill('solid', start_color='0099CC'),
        'gray': openpyxl.styles.PatternFill('solid', start_color='D9D9D9'),
        'green': openpyxl.styles.PatternFill('solid', start_color='C6EFCE'),
        'mech': openpyxl.styles.PatternFill('solid', start_color='FFF2CC'),
        'review': openpyxl.styles.PatternFill('solid', start_color='F8CBAD'),
    }
    for col in range(1, N_COLS + 1):
        ws.cell(row=1, column=col).border = border

    for row_type, data, is_rep in write_rows:
        row_num = ws.max_row + 1
        ws.append(data[:N_COLS] + [''] * (N_COLS - len(data)))
        fill = None
        if row_type == 'summary':
            fill = fills['green']
            for col in range(1, N_COLS + 1):
                ws.cell(row=row_num, column=col).font = openpyxl.styles.Font(bold=True)
        elif _is_nc(data[1]):
            fill = fills['blue']
        elif data[4] == UNMATCHED:
            fill = fills['yellow']
        elif len(data) > 8 and '[需确认]' in str(data[8]):
            fill = fills['review']
        elif data[4] == MECH_PN:
            fill = fills['mech']
        elif is_rep:
            fill = fills['gray']
        for col in range(1, N_COLS + 1):
            cell = ws.cell(row=row_num, column=col)
            cell.border = border
            if fill:
                cell.fill = fill

    # 对齐
    center = openpyxl.styles.Alignment(horizontal='center', vertical='center', wrap_text=True)
    left = openpyxl.styles.Alignment(horizontal='left', vertical='center', wrap_text=True)
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
        for cell in row:
            cell.alignment = center if (cell.row == 1 or cell.column in (4, 5, 6)) else left

    widths = {1: 44, 2: 16.5, 3: 16.5, 4: 10, 5: 16, 6: 14, 7: 44, 8: 11, 9: 22}
    for col, w in widths.items():
        ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = w

    ws.row_dimensions[1].height = 22
    col_eff = {1: 36, 2: 13, 3: 13, 7: 36, 8: 9, 9: 18}
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        max_lines = 1
        for cell in row:
            if cell.column in col_eff and cell.value:
                w = sum(2 if ord(c) > 127 else 1 for c in str(cell.value))
                max_lines = max(max_lines, math.ceil(w / col_eff[cell.column]))
        ws.row_dimensions[row[0].row].height = max(30, max_lines * 15)

    wb.save(output_file)
