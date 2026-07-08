# -*- coding: utf-8 -*-
"""生成/对比金标准：把匹配结果xlsx转成规范化文本（一行一条：位号|Part|FP|Qty|料号）。

用法:
  python tests/make_golden.py <result.xlsx> <out.txt>
颜色标记: U=未匹配(黄) N=NC(蓝) S=汇总(绿) D=重复(灰) .=普通
"""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
import openpyxl

FILLS = {'00FFFF00': 'U', '000099CC': 'N', '00C6EFCE': 'S', '00D9D9D9': 'D'}


def canon(path):
    wb = openpyxl.load_workbook(path)
    ws = wb.active
    lines = []
    for row in list(ws.iter_rows())[1:]:
        f = row[0].fill
        tag = '.'
        if f and f.fill_type == 'solid' and isinstance(f.start_color.rgb, str):
            tag = FILLS.get(f.start_color.rgb, '.')
        vals = [str(c.value).strip() if c.value is not None else '' for c in row[:5]]
        lines.append(tag + ' ' + '|'.join(vals))
    return lines


if __name__ == '__main__':
    lines = canon(sys.argv[1])
    with open(sys.argv[2], 'w', encoding='utf-8') as fh:
        fh.write('\n'.join(lines) + '\n')
    print(f'{sys.argv[2]}: {len(lines)} rows')
