# -*- coding: utf-8 -*-
"""BOM匹配引擎（新架构适配层）— BOM匹配助手 GUI 专用入口。

内核在 bom_matcher/ 包（解析→引擎→输出，2026-07-08 重写），
本文件负责维持 app.py 依赖的接口契约：
  - _set_log_callback / _set_progress_callback / _set_row_callback
  - process_bom(...) -> dict(output_file/results/unmatched/code_index/candidate_cache/export)
  - results 为可变行列表，人工复核直接改 r[4..6]，export 原地重排后重写 Excel
旧引擎完整实现见 git 历史（首个提交）。
"""
import argparse
import io
import os
import sys
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from bom_matcher.loader import load_bom, load_code
from bom_matcher.rules import load_user_rules, COMPANIONS
from bom_matcher.engine import Engine
from bom_matcher.model import (KIND_BLANK, KIND_MECH,
                               CAT_R, CAT_C, CAT_L, CAT_CONN)
from bom_matcher import report


def _ensure_utf8_stream(stream):
    if stream is None:
        return stream
    try:
        stream.reconfigure(encoding='utf-8', errors='replace')
        return stream
    except Exception:
        if hasattr(stream, 'buffer'):
            return io.TextIOWrapper(stream.buffer, encoding='utf-8', errors='replace')
        return stream


if sys.platform == 'win32':
    sys.stdout = _ensure_utf8_stream(sys.stdout)
    sys.stderr = _ensure_utf8_stream(sys.stderr)

_log_callback = None
_progress_callback = None
_row_callback = None

DTYPE_CN = {CAT_R: '电阻', CAT_C: '电容', CAT_L: '电感', CAT_CONN: '连接器'}


def _set_log_callback(cb):
    global _log_callback
    _log_callback = cb


def _set_progress_callback(cb):
    global _progress_callback
    _progress_callback = cb


def _set_row_callback(cb):
    global _row_callback
    _row_callback = cb


def _log(msg):
    print(msg)
    if _log_callback:
        _log_callback(msg)


def _progress(pct, msg=''):
    if _progress_callback:
        _progress_callback(pct, msg)


def _emit_row(refdes, part, fp, dtype, pn):
    if _row_callback:
        try:
            _row_callback({'ref': refdes, 'part': part, 'fp': fp,
                           'dtype': dtype, 'pn': pn, 'matched': pn != report.UNMATCHED})
        except Exception:
            pass


def process_bom(bom_file, code_file, output_file=None, log_callback=None, progress_callback=None):
    """供GUI直接调用的入口函数"""
    _set_log_callback(log_callback)
    _set_progress_callback(progress_callback)
    return main(bom=bom_file, code=code_file, output=output_file)


def main(bom=None, code=None, output=None, verbose=False, debug=False):
    cli = bom is None and code is None
    if cli:
        parser = argparse.ArgumentParser(description='BOM智能匹配工具')
        parser.add_argument('--bom', required=True)
        parser.add_argument('--code', required=True)
        parser.add_argument('--output')
        args = parser.parse_args()
        bom, code, output = args.bom, args.code, args.output

    output = output or os.path.join(
        os.path.dirname(os.path.abspath(bom)),
        f"BOM_匹配结果_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")

    bom_rows, cols = load_bom(bom)
    entries, discontinued, sheets = load_code(code)
    _log(f"BOM: {len(bom_rows)} 行, 编码表: {', '.join(sheets)}")
    _log(f"编码表条目: {len(entries)} (另有 {len(discontinued)} 条停用料已剔除)")
    _progress(15, '加载文件完成')

    extra_alias, banned = load_user_rules(os.path.dirname(os.path.abspath(__file__)))
    eng = Engine(entries, extra_alias, banned)

    items = []
    need = [c for c in (cols['ref'], cols['part'], cols['fp'], cols['qty']) if c is not None]
    for row in bom_rows:
        if len(row) <= max(need):
            continue

        def cell(i):
            return '' if i is None or row[i] is None else str(row[i]).strip()

        it = eng.build_item(cell(cols['ref']), cell(cols['part']),
                            cell(cols['fp']), cell(cols['qty']))
        if it.kind != KIND_BLANK:
            items.append(it)
    eng.prescan(items)
    _progress(35, '预扫描完成，开始匹配...')

    results = []
    unmatched = []
    for it in items:
        dtype = DTYPE_CN.get(it.category, '其他')
        if it.kind == KIND_MECH:
            results.append([it.refdes, it.part, it.fp, it.qty,
                            report.MECH_PN, '安装孔/测试点', '', '', '非采购物料，不参与匹配'])
            _emit_row(it.refdes, it.part, it.fp, '其他', report.MECH_PN)
            continue
        m = eng.match(it)
        if m:
            cand = '; '.join(f'{pn}: {spec[:40]}' for pn, _n, spec in m.candidates)
            note = m.reason + (' [需确认]' if m.confidence == 'review' else '')
            results.append([it.refdes, it.part, it.fp, it.qty,
                            m.pn, m.name, m.spec, cand, note])
            _emit_row(it.refdes, it.part, it.fp, dtype, m.pn)
        else:
            near = eng.nearest(it)
            cand = '; '.join(f'{e.pn}: {e.spec[:40]}' for e in near)
            results.append([it.refdes, it.part, it.fp, it.qty,
                            report.UNMATCHED, report.UNMATCHED, report.UNMATCHED,
                            cand, '编码表无对应料' + ('，右侧列出近似候选' if cand else '')])
            unmatched.append((it.refdes, it.part, it.fp))
            _emit_row(it.refdes, it.part, it.fp, dtype, report.UNMATCHED)

    _progress(75, '匹配完成，正在生成输出...')

    def export_results(res_list, out_path=None):
        """复核后重导出：原地重排+聚合，人工指定的行更新说明列"""
        for r in res_list:
            if r[4] != report.UNMATCHED and len(r) > 8 and str(r[8]).startswith('编码表无对应料'):
                r[7], r[8] = '', '人工指定'
        report.add_companions(res_list, COMPANIONS, eng.by_pn, _log)
        report.sort_rows(res_list)
        report.write_excel(res_list, out_path or output)

    export_results(results)
    _progress(100, '完成')
    _log(f"\n完成。输出: {output}, 总{len(results)}行, 未匹配{len(unmatched)}条")

    # 复核界面用的编码表索引（中文类别 -> (料号,名称,规格)）
    code_index = {'电容': [], '电阻': [], '电感': [], '连接器': [], '其他': []}
    for e in eng.entries:
        code_index[DTYPE_CN.get(e.category, '其他')].append((e.pn, e.name, e.spec))

    return {
        'output_file': output,
        'results': results,
        'unmatched': unmatched,
        'code_index': code_index,
        'candidate_cache': {},
        'export': export_results,
    }


if __name__ == '__main__':
    main()
