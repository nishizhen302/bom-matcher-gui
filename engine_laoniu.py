# -*- coding: utf-8 -*-
"""
BOM智能匹配 v2 - 适配物料长代码体系
编码表列: 物料长代码、物料名称、规格型号、分类
"""
import argparse
import openpyxl
import re
import sys
import io
from collections import defaultdict, Counter
from datetime import datetime

if sys.platform == 'win32' and sys.stdout is not None and hasattr(sys.stdout, 'buffer'):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
elif sys.stdout is None:
    sys.stdout = io.StringIO()
    sys.stderr = io.StringIO()

DEBUG_MODE = False

def debug_print(*args, **kwargs):
    if DEBUG_MODE:
        print(*args, **kwargs)

def run_match(bom_file, code_file, output_file=None, progress=None, debug=False):
    """执行 BOM 匹配。progress 为事件回调 progress(dict)，事件类型见底部 emit 调用。
    返回 dict: results/unmatched/output_file/bom_count/categories/export/find_candidates
    匹配逻辑与 bom-matcher-laoniu run.py v7 保持一致，仅加了回调与导出参数化。"""
    global DEBUG_MODE
    DEBUG_MODE = debug

    def emit(event):
        if progress is not None:
            try:
                progress(event)
            except Exception:
                pass

    BOM_FILE = bom_file
    CODE_FILE = code_file
    OUTPUT_FILE = output_file or f"BOM_匹配结果_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

    # ========== 工具函数 ==========
    def normalize(s):
        return re.sub(r'[^a-z0-9]', '', str(s).lower())

    def extract_footprint(fp):
        if not fp:
            return ''
        fp = str(fp).upper()
        if 'AE' in fp or 'CAPAE' in fp:
            return fp
        m = re.search(r'(\d{4})', fp)
        return m.group(1) if m else re.search(r'(\d{3})', fp).group(1) if re.search(r'(\d{3})', fp) else ''

    def parse_cap_value(s):
        """解析电容值(返回pF)和电压"""
        s = str(s).strip().upper()
        if s.startswith('NC/'): s = s[3:]
        m = re.search(r'([\d\.]+)\s*(UF|U|NF|N|PF|P)(?:\s*[,/，]\s*(\d+(?:\.\d+)?)V)?', s)
        if not m: return None, None
        val = float(m.group(1)); unit = m.group(2)
        if unit in ('UF','U'): pF = int(round(val*1000000))
        elif unit in ('NF','N'): pF = int(round(val*1000))
        elif unit in ('PF','P'): pF = int(round(val))
        else: return None, None
        if m.group(3): return pF, float(m.group(3))
        mV = re.search(r'\s+(\d+(?:\.\d+)?)V', s)
        return pF, float(mV.group(1)) if mV else None

    def parse_spec_cap(s):
        """从规格中解析电容值、电压、封装"""
        if not s: return None, None, ''
        s = str(s)
        fp_m = re.search(r'(\d{4})', s)
        fp = fp_m.group(1) if fp_m else ''
        mV = re.search(r'(\d+(?:\.\d+)?)V', s, re.I)
        volt = float(mV.group(1)) if mV else None
        # 用 (?=\b|V|\d) 替代 \b，解决 "100uF50V" 中 F 后紧跟数字导致 \b 不匹配的问题
        m = re.search(r'(?<![a-zA-Z])(\d+(?:\.\d+)?)\s*(UF|U|NF|N|PF|P)(?=\b|V|\d|$)', s, re.I)
        if not m: return None, volt, fp
        val = float(m.group(1)); unit = m.group(2).upper()
        if unit in ('UF','U'): pF = int(round(val*1000000))
        elif unit in ('NF','N'): pF = int(round(val*1000))
        elif unit in ('PF','P'): pF = int(round(val))
        else: return None, volt, fp
        return pF, volt, fp

    def parse_res_value(s):
        """解析电阻值"""
        s = str(s).strip().upper().replace('NC','').replace(' ','')
        # 先去掉下划线（如 10K_1% → 10K1%）
        s = re.sub(r'_', '', s)
        # 去掉百分比和精度信息
        s = re.sub(r'\d*\.?\d*%', '', s)
        # 三位一体: 4K7, 2K2, 1M5, 14K3
        m = re.search(r'(\d+(?:\.\d+)?)([KMG])(\d)', s)
        if m:
            val = float(m.group(1) + '.' + m.group(3))
            mult = m.group(2)
            if mult == 'K': val *= 1000
            elif mult == 'M': val *= 1000000
            return int(round(val))
        # 标准格式: 10K, 100R, 330K, 22R
        m = re.search(r'(\d+(?:\.\d+)?)([KMG])?', s)
        if not m: return None
        val = float(m.group(1))
        mult = m.group(2)
        if mult == 'K': val *= 1000
        elif mult == 'M': val *= 1000000
        return int(round(val))

    def parse_ind_value(s):
        """解析电感值(返回uH)"""
        s = str(s).strip().upper()
        if s.startswith('NC/'): s = s[3:]
        m = re.search(r'(\d+(?:\.\d+)?)\s*(UH|U|MH|M|NH|N)\b', s)
        if m:
            val = float(m.group(1)); u = m.group(2)
            if u in ('UH','U'): return val
            if u in ('MH','M'): return val*1000
            if u in ('NH','N'): return val/1000
        m2 = re.search(r'(\d+)R(\d+)', s)
        return float(f"{m2.group(1)}.{m2.group(2)}") if m2 else None

    def parse_spec_res(s):
        """从规格中解析电阻值、封装、精度"""
        s = str(s).strip().upper().replace('，',',').replace('Ω','')
        resistance = None

        # 三位一体: 4K7, 2K2, 1M5
        m = re.search(r'(\d)([KMG])(\d)', s)
        if m:
            val = float(m.group(1) + '.' + m.group(3))
            mult = m.group(2)
            if mult=='K': val*=1000
            elif mult=='M': val*=1000000
            resistance = int(round(val))
        else:
            # 标准格式: 10R, 22R, 100K, 100 OHM
            # 优先匹配带单位的（OHM/R/Ω），避免匹配到物料编号中的数字
            m2 = re.search(r'([\d\.]+)\s*([KMG]?)\s*(?:OHM|OHMS)\b', s)
            if not m2:
                m2 = re.search(r'([\d\.]+)\s*([KMG]?)R(?![A-Z0-9])', s)
            if not m2:
                # 回退到裸数字+可选KMG
                m2 = re.search(r'([\d\.]+)\s*([KMG]?)\b', s)
            if not m2: return None, None, None
            try:
                val = float(m2.group(1))
            except: return None, None, None
            mult = m2.group(2)
            if mult == 'K': val *= 1000
            elif mult == 'M': val *= 1000000
            # 0R/0R1 等接近0的值
            if val < 1: resistance = 0
            else: resistance = int(round(val))

        # 封装：在规格中找标准封装 (0402, 0603, 0805, 1206)
        fp = re.search(r'(?:0402|0603|0805|1206|1210|2010|2512)', s)
        fp_str = fp.group(0) if fp else ''
        # 如果没找到标准封装，再尝试通用4位数字
        if not fp_str:
            fp = re.search(r'(\d{4})', s)
            fp_str = fp.group(1) if fp else ''

        # 精度
        tol = re.search(r'[\xb1]?\s*(\d+)%', s)
        tol_val = int(tol.group(1)) if tol else None
        return resistance, fp_str, tol_val

    def is_nc(part):
        return str(part).strip().upper().startswith('NC/')

    def strip_nc(part):
        s = str(part).strip()
        if s.upper().startswith('NC/'): return s[3:]
        if s.upper() in ('NC', 'NC.', 'N/C'): return ''
        return s

    # ========== 加载数据 ==========
    # 加载BOM
    wb_bom = openpyxl.load_workbook(BOM_FILE, data_only=True, read_only=True)
    ws_bom = wb_bom.active
    bom_rows = list(ws_bom.iter_rows(values_only=True))
    wb_bom.close()

    # 找表头行
    header_row = None
    for i, row in enumerate(bom_rows):
        if row and 'Item' in str(row[0] if row[0] else ''):
            header_row = i
            break

    if header_row is None:
        print("ERROR: 无法找到BOM表头行")
        return

    header = [str(c).strip().lower() if c else '' for c in bom_rows[header_row]]
    col_item = header.index('item')
    col_qty = header.index('quantity')
    col_ref = header.index('reference')
    col_part = header.index('part')
    col_fp = header.index('pcb footprint') if 'pcb footprint' in header else header.index('footprint')

    bom_data = []
    for row in bom_rows[header_row+1:]:
        if len(row) <= max(col_item, col_part, col_fp): continue
        item = row[col_item]
        if item is None or not isinstance(item, (int, float)): continue
        ref = str(row[col_ref] if row[col_ref] is not None else '').strip()
        part = str(row[col_part] if row[col_part] is not None else '').strip()
        fp = str(row[col_fp] if row[col_fp] is not None else '').strip()
        qty = str(row[col_qty] if row[col_qty] is not None else '').strip()
        if not ref and not part and not fp: continue
        bom_data.append((ref, part, fp, qty))

    # 加载编码表（去重后，从各分类sheet加载）
    wb_code = openpyxl.load_workbook(CODE_FILE, data_only=True, read_only=True)
    code_items = []  # (物料长代码, 物料名称, 规格型号, 分类)
    skip_sheets = ['分类汇总', '重复物料汇总', '全部物料(含重复)']
    for name in wb_code.sheetnames:
        if name in skip_sheets: continue
        ws = wb_code[name]
        rows = list(ws.iter_rows(values_only=True))
        if len(rows) < 2: continue
        # 第一行是表头
        for row in rows[1:]:
            if len(row) < 3: continue
            code = str(row[0]).strip() if row[0] else ''
            mat_name = str(row[1]).strip() if row[1] else ''
            spec = str(row[2]).strip() if row[2] else ''
            if not code or not spec: continue
            code_items.append((code, mat_name, spec, name))
    wb_code.close()

    # 去重（按物料长代码+规格型号）
    seen = set()
    code_unique = []
    for item in code_items:
        key = (item[0], item[2])
        if key not in seen:
            seen.add(key)
            code_unique.append(item)

    # 按分类建索引
    code_index = defaultdict(list)
    for item in code_unique:
        code_index[item[3]].append(item)

    all_categories = list(code_index.keys())
    print(f"BOM: {len(bom_data)} 行")
    print(f"编码表: {len(code_unique)} 条物料, 分类: {all_categories}")
    emit({'type': 'loaded', 'bom_count': len(bom_data), 'code_count': len(code_unique),
          'categories': all_categories, 'bom_name': str(BOM_FILE), 'code_name': str(CODE_FILE)})

    # ========== 匹配逻辑 ==========

    def normalize_ae_footprint(fp_str):
        """统一铝电解电容封装为4位尺寸代码 (D×H, 单位mm, 截断取整)
        CAPAE840X1000N: 8.4x10.0mm → int → 8x10 → "0810"
        CAPAE660X770N:  6.6x7.7mm  → int → 6x7  → "0607"
        标准4位码如 0810 → "0810"
        注意：用int()而非round()，避免6.6→7但6.3→6导致的误不匹配
        """
        s = str(fp_str).upper()
        # CAPAE 风格: CAPAE840X1000N (直径x高度, 单位0.1mm)
        m = re.search(r'(\d{3})X(\d{3,4})', s)
        if m:
            d = int(m.group(1)) / 100.0  # 840→8.4, 660→6.6
            h = int(m.group(2)) / 100.0  # 1000→10.0, 770→7.7
            return f"{int(d):02d}{int(h):02d}"
        # 标准4位码如 0810
        m4 = re.search(r'(\d{4})', s)
        return m4.group(1) if m4 else s

    def extract_ae_size_from_spec(spec_str):
        """从代码表规格文本中提取铝电解电容的DxH尺寸(mm)
        如 '8*10mm' → '0810', '6.3*5.7mm' → '0605', '10x10' → '1010'
        注意：用int()而非round()，与normalize_ae_footprint保持一致
        """
        s = str(spec_str)
        # 匹配 D*D 或 DxH 或 D*Dmm 等格式
        m = re.search(r'(\d+(?:\.\d+)?)\s*[*xX×]\s*(\d+(?:\.\d+)?)\s*(?:mm)?', s)
        if m:
            d = int(float(m.group(1)))
            h = int(float(m.group(2)))
            return f"{d:02d}{h:02d}"
        return None

    def match_capacitor(part, fp):
        """匹配电容（严格模式：容值+封装必须匹配）"""
        part_clean = strip_nc(part)
        cap_pF, volt = parse_cap_value(part_clean)
        fp_digits = extract_footprint(fp)
        if cap_pF is None: return None

        # 必须同时匹配容值和封装
        if not fp_digits: return None

        # 判断是否为铝电解电容（AE/CAPAE封装）
        is_ae = 'AE' in str(fp).upper() or 'CAPAE' in str(fp).upper()

        candidates = []
        for code, name, spec, cat in code_index.get('电容', []):
            c_cap, c_volt, c_fp = parse_spec_cap(spec)
            if c_cap != cap_pF: continue

            # 铝电解电容：按尺寸匹配，BOM的CAPAE编码和代码表规格文本都转为 D×H mm
            if is_ae or '铝电解' in str(name):
                ae_bom = normalize_ae_footprint(fp_digits)
                # 优先从规格文本提取尺寸(如 8*10mm)，其次用型号中的4位码
                ae_code = extract_ae_size_from_spec(spec)
                if not ae_code:
                    ae_code = normalize_ae_footprint(c_fp)
                if not ae_code or ae_bom != ae_code:
                    continue
            else:
                if c_fp != fp_digits: continue  # 封装必须一致，否则SMT焊不上

            score = 0
            if volt and c_volt and c_volt == volt:
                score += 5
            elif volt and c_volt and c_volt > volt:
                score += 3
            elif volt and c_volt and c_volt < volt:
                score += 1
            candidates.append((code, name, spec, score))

        if candidates:
            candidates.sort(key=lambda x: -x[3])
            return (candidates[0][0], candidates[0][1], candidates[0][2])

        return None

    def match_resistor(part, fp):
        """匹配电阻（严格模式：阻值+封装必须匹配，精度不低于BOM要求）"""
        part_clean = strip_nc(part)
        bom_val = parse_res_value(part_clean)
        fp_digits = extract_footprint(fp)
        part_norm = normalize(part_clean)
        part_upper = str(part_clean).upper()

        # 提取BOM精度要求（如 200K 1% → 1）
        bom_tol = None
        tol_m = re.search(r'(\d+(?:\.\d+)?)\s*%', part_upper)
        if tol_m:
            bom_tol = float(tol_m.group(1))

        # NC 或空值
        if bom_val is None:
            return None

        if not fp_digits: return None  # 封装未知则不匹配

        # 0R电阻
        if bom_val == 0 or part_norm in ('0r', '0'):
            matches = []
            for code, name, spec, cat in code_index.get('电阻', []):
                r_val, r_fp, r_tol = parse_spec_res(spec)
                if r_val == 0 and r_fp == fp_digits:
                    matches.append((code, name, spec, r_tol))
            if matches:
                matches.sort(key=lambda x: x[3] if x[3] else 999)
                return (matches[0][0], matches[0][1], matches[0][2])
            return None

        # 阻值+封装严格匹配，精度过滤
        matches = []
        for code, name, spec, cat in code_index.get('电阻', []):
            r_val, r_fp, r_tol = parse_spec_res(spec)
            if r_val == bom_val and r_fp == fp_digits:
                # BOM要求精度时，编码表精度必须不高于BOM（更精确）
                if bom_tol is not None:
                    if r_tol is not None and r_tol > bom_tol:
                        continue  # 编码表精度太差，跳过
                matches.append((code, name, spec, r_tol))
        if matches:
            # 按精度升序排列（tol_val越小精度越高），优先选最高精度的电阻
            matches.sort(key=lambda x: x[3] if x[3] else 999)
            return (matches[0][0], matches[0][1], matches[0][2])

        return None

    def match_inductor(part, fp):
        """匹配电感/磁珠"""
        part_clean = strip_nc(part)
        fp_digits = extract_footprint(fp)
        part_lower = str(part_clean).lower()
        part_upper = str(part_clean).upper()

        # === 磁珠匹配 ===
        is_bead = False
        part_and_fp = part_lower + ' ' + fp_digits  # 合并part和fp用于检查
        if 'mpz' in part_and_fp or 'bead' in part_and_fp or 'feb' in part_and_fp:
            is_bead = True
        elif '/100mhz' in part_and_fp or '@100mhz' in part_and_fp:
            is_bead = True
        elif 'fb' == str(ref)[:2] if ref else False:  # FB开头位号
            is_bead = True
        elif 'lw' == str(ref)[:2] if ref else False:  # LW开头位号通常是磁珠
            is_bead = True
        elif '100r/0805' in part_and_fp or '100r/1206' in part_and_fp or '100r/0603' in part_and_fp:
            is_bead = True
        elif re.match(r'bead[_/]', part_and_fp):
            is_bead = True

        # CM_Filter 是共模电感，不是普通磁珠，单独匹配
        is_cm_filter = 'cm_filter' in str(fp).lower() or 'cmfilter' in str(fp).lower()

        if is_cm_filter:
            # 从编码表电感分类中找共模电感（名称含"共模"或CM/ACM型号）
            for code, name, spec, cat in code_index.get('电感', []):
                if '共模' in str(name):
                    # 检查阻抗是否匹配（90R → 90Ω）
                    if part_clean == '90R' or '90' in part_upper:
                        if '90' in str(spec) or '900' in str(spec):
                            return (code, name, spec)
                    else:
                        return (code, name, spec)
            # 没找到共模电感，标记为未匹配
            return None

        if is_bead:
            # 从编码表电感分类中找磁珠
            for code, name, spec, cat in code_index.get('电感', []):
                spec_upper = str(spec).upper()
                spec_lower = str(spec).lower()
                # 磁珠特征：含@100MHz、Ω@100MHz、或名称含磁珠
                is_bead_spec = ('@100MHZ' in spec_upper or '@100MHz' in spec or
                               'Ω@100MHz' in spec or '磁珠' in str(name) or
                               'MPZ' in spec_upper or 'BLM' in spec_upper or
                               'BEAD' in spec_upper or 'FEB' in spec_upper or
                               '100MHZ' in spec_upper)
                if not is_bead_spec:
                    continue
                # 封装匹配
                if fp_digits and fp_digits not in spec_upper:
                    continue
                # 如果BOM指定了MPZ型号，优先精确匹配
                if 'MPZ' in part_upper and 'MPZ' in spec_upper:
                    return (code, name, spec)
            # 如果没精确匹配到，放宽封装限制
            for code, name, spec, cat in code_index.get('电感', []):
                spec_upper = str(spec).upper()
                is_bead_spec = ('@100MHZ' in spec_upper or '磁珠' in str(name) or
                               'MPZ' in spec_upper or 'BLM' in spec_upper or
                               'BEAD' in spec_upper)
                if not is_bead_spec:
                    continue
                if fp_digits and fp_digits not in spec_upper:
                    continue
                return (code, name, spec)

        # === 普通电感匹配 ===
        val = parse_ind_value(part_clean)
        if val is None:
            # 尝试从COIL封装判断
            if 'COIL' in str(fp).upper():
                # 尝试从BOM提取电感值
                m = re.search(r'(\d+(?:\.\d+)?)\s*UH', part_upper)
                if m:
                    val = float(m.group(1))
            if val is None:
                return None

        for code, name, spec, cat in code_index.get('电感', []):
            spec_str = str(spec)
            # 跳过磁珠
            if '磁珠' in name or '@100MHz' in spec_str or '100MHZ' in spec_str.upper():
                continue
            m = re.search(r'(\d+(?:\.\d+)?)\s*(UH|U|MH|M|NH|N|H)\b', spec_str, re.I)
            if not m: continue
            try:
                sv = float(m.group(1)); u = m.group(2).upper()
                if u in ('UH','U'): code_uH = sv
                elif u in ('MH','M'): code_uH = sv*1000
                elif u in ('NH','N'): code_uH = sv/1000
                elif u == 'H': code_uH = sv*1000000
                else: continue
            except: continue
            if abs(val - code_uH) > 0.01: continue
            if fp_digits and fp_digits not in spec_str: continue
            return (code, name, spec)

        # 也在其他分类中查找电感
        for code, name, spec, cat in code_index.get('其他', []):
            spec_str = str(spec)
            m = re.search(r'(\d+(?:\.\d+)?)\s*(UH|U|MH|M|NH|N|H)\b', spec_str, re.I)
            if not m: continue
            try:
                sv = float(m.group(1)); u = m.group(2).upper()
                if u in ('UH','U'): code_uH = sv
                elif u in ('MH','M'): code_uH = sv*1000
                elif u in ('NH','N'): code_uH = sv/1000
                elif u == 'H': code_uH = sv*1000000
                else: continue
            except: continue
            if val is None or abs(val - code_uH) > 0.01: continue
            if fp_digits and fp_digits not in spec_str: continue
            return (code, name, spec)

        return None

    def match_connector(part, fp):
        """匹配连接器/插座"""
        part_clean = strip_nc(part)
        part_norm = normalize(part_clean)
        fp_digits = extract_footprint(fp)

        # 精确型号匹配
        for code, name, spec, cat in code_index.get('插座', []):
            spec_norm = normalize(spec)
            if part_norm in spec_norm:
                return (code, name, spec)

        # 子串匹配（型号精确包含）
        if len(part_norm) >= 5:
            for code, name, spec, cat in code_index.get('插座', []):
                spec_norm = normalize(spec)
                if part_norm in spec_norm:
                    return (code, name, spec)

        # JST B系列前缀匹配（B4B、B5B、B10B等，区分引脚数）
        jst_prefix = re.match(r'(B\d+B)', part_clean, re.I)
        if jst_prefix:
            model_prefix = normalize(jst_prefix.group(1))
            for code, name, spec, cat in code_index.get('插座', []):
                spec_norm = normalize(spec)
                if model_prefix in spec_norm:
                    return (code, name, spec)

        # 特殊连接器类型关键词匹配（需要引脚数一致或BOM不含引脚数信息）
        keywords = ['USB', 'DVI', 'HDMI', 'DP', 'RJ45', 'VGA', 'HEADER']
        part_upper = str(part).upper()
        # DISPLAY_PORT 映射为 DP
        if 'DISPLAY_PORT' in part_upper:
            part_upper_for_match = part_upper.replace('DISPLAY_PORT', 'DP')
        else:
            part_upper_for_match = part_upper
        for kw in keywords:
            if kw in part_upper_for_match:
                # USB类型精确匹配：BOM含USB-A/USB-B/USB-C/MICRO-USB时，优先精确匹配类型
                # 注意：USB-A/USB-B 等是通用接口类型描述，不是具体型号
                # 只有BOM包含具体型号时才匹配，否则标记未匹配（避免单层匹配到双层等问题）
                if kw == 'USB':
                    # 先尝试精确型号匹配（BOM part 包含具体型号）
                    for code, name, spec, cat in code_index.get('插座', []):
                        spec_upper = str(spec).upper() + ' ' + str(name).upper()
                        if 'USB' not in spec_upper:
                            continue
                        if part_norm in normalize(spec):
                            return (code, name, spec)
                    # BOM只有通用USB类型描述，代码表候选可能不兼容 → 不匹配
                    return None
                # 通用关键词匹配（无精确类型匹配时的fallback）
                for code, name, spec, cat in code_index.get('插座', []):
                    if kw in str(spec).upper() or kw in str(name).upper():
                        return (code, name, spec)

        # B4PS-VH 匹配 VH 系列
        if 'B4PS' in part_upper or 'VH' in part_upper:
            for code, name, spec, cat in code_index.get('插座', []):
                if 'VH' in str(spec).upper() or 'VH' in str(name).upper():
                    return (code, name, spec)

        # DDR3 匹配内存插座
        if 'DDR3' in part_upper:
            for code, name, spec, cat in code_index.get('插座', []):
                if 'DDR' in str(spec).upper() or 'DDR' in str(name).upper():
                    return (code, name, spec)

        return None

    def match_ic(part, fp):
        """匹配集成电路"""
        part_clean = strip_nc(part)
        part_upper = str(part_clean).upper()

        # 精确型号匹配 - 遍历所有分类（有些IC可能在开关稳压器等分类里）
        for cat_name in ['集成电路', '开关稳压器', '驱动IC', '半导体', '电感', '其他']:
            for code, name, spec, cat in code_index.get(cat_name, []):
                if part_upper in str(spec).upper():
                    return (code, name, spec)

        # 型号前缀匹配（至少4个字符）
        part_norm = normalize(part_clean)
        for cat_name in ['集成电路', '开关稳压器', '半导体']:
            for code, name, spec, cat in code_index.get(cat_name, []):
                spec_norm = normalize(spec)
                if len(part_norm) >= 4 and part_norm in spec_norm:
                    return (code, name, spec)

        # 特殊：X24C16 匹配 EEPROM
        if 'X24C16' in part_upper or '24C02' in part_upper:
            for code, name, spec, cat in code_index.get('集成电路', []):
                if 'EEPROM' in str(name).upper() or 'X24C16' in str(spec).upper() or '24C02' in str(spec).upper():
                    return (code, name, spec)

        return None

    def match_esd(part, fp):
        """匹配静电保护器"""
        part_clean = strip_nc(part)
        part_upper = str(part_clean).upper()
        part_norm = normalize(part_clean)

        # 提取BOM封装尺寸数字，用于封装兼容性检查
        fp_size = ''
        m_fp = re.search(r'(\d{4})', fp)
        if m_fp:
            fp_size = m_fp.group(1)  # 如 ESD0603 → 0603, ESD0402 → 0402

        # 精确匹配
        for code, name, spec, cat in code_index.get('静电保护器', []):
            if part_upper in str(spec).upper():
                # 封装兼容性检查：BOM封装尺寸不能与编码表标注冲突
                if fp_size:
                    spec_upper = str(spec).upper()
                    # 从规格中提取封装尺寸标注（如0402,0603,0201等）
                    m_spec_size = re.search(r'\b(0201|0402|0603|0805|1008|1206|1210|1806|1812|2010|2512)\b', spec_upper)
                    if m_spec_size and m_spec_size.group(1) != fp_size:
                        continue
                return (code, name, spec)
        # 子串匹配
        for code, name, spec, cat in code_index.get('静电保护器', []):
            spec_norm = normalize(spec)
            if len(part_norm) >= 4 and part_norm in spec_norm:
                # 封装兼容性检查
                if fp_size:
                    spec_upper = str(spec).upper()
                    m_spec_size = re.search(r'\b(0201|0402|0603|0805|1008|1206|1210|1806|1812|2010|2512)\b', spec_upper)
                    if m_spec_size and m_spec_size.group(1) != fp_size:
                        continue
                return (code, name, spec)
        # RCLAMP 系列：精确匹配完整型号
        if 'RCLAMP' in part_upper:
            for code, name, spec, cat in code_index.get('静电保护器', []):
                spec_upper = str(spec).upper()
                if 'RCLAMP' in spec_upper:
                    m = re.search(r'RCLAMP[A-Z0-9]+', spec_upper)
                    if m and m.group() == part_upper:
                        return (code, name, spec)
        # PESD 系列：精确匹配型号前缀（至少到电压位，如PESD5V0）
        if 'PESD' in part_upper:
            for code, name, spec, cat in code_index.get('静电保护器', []):
                spec_upper = str(spec).upper()
                if 'PESD' in spec_upper:
                    # 提取编码表中的PESD型号
                    m = re.search(r'PESD[A-Z0-9]+', spec_upper)
                    if m and m.group() == part_upper:
                        return (code, name, spec)
        return None

    def match_fuse(part, fp):
        """匹配保险丝"""
        part_clean = strip_nc(part)
        part_norm = normalize(part_clean)
        part_upper = str(part_clean).upper()

        # 提取BOM参数：电流(A)、电压(V)、封装
        bom_amps = re.findall(r'([\d\.]+)\s*A', part_upper)
        bom_volts = re.findall(r'([\d\.]+)\s*V', part_upper)
        bom_fp = re.search(r'(0402|0603|0805|1206|1210|1812|2010|2410)', part_upper)

        # 参数匹配：电流+电压+封装全部匹配
        if bom_amps and bom_volts:
            candidates = []
            for code, name, spec, cat in code_index.get('保险丝', []):
                spec_upper = str(spec).upper()
                spec_amps = re.findall(r'([\d\.]+)\s*A', spec_upper)
                spec_volts = re.findall(r'([\d\.]+)\s*V', spec_upper)
                spec_fp = re.search(r'(0402|0603|0805|1206|1210|1812|2010|2410)', spec_upper)

                if not spec_amps or not spec_volts:
                    continue
                # 电流和电压都要匹配
                if any(abs(float(a) - float(b)) < 0.1 for a in bom_amps for b in spec_amps) and \
                   any(abs(float(a) - float(b)) < 0.1 for a in bom_volts for b in spec_volts):
                    score = 0
                    if bom_fp and spec_fp and bom_fp.group(1) == spec_fp.group(1):
                        score += 10
                    candidates.append((code, name, spec, score))
            if candidates:
                candidates.sort(key=lambda x: -x[3])
                return (candidates[0][0], candidates[0][1], candidates[0][2])

        # 精确字符串匹配（兜底）
        for code, name, spec, cat in code_index.get('保险丝', []):
            if part_norm in normalize(spec):
                return (code, name, spec)

        # FUSE/RE-FUSE 关键词
        if 'FUSE' in part_upper or 'RE-FUSE' in part_upper:
            for code, name, spec, cat in code_index.get('保险丝', []):
                if 'FUSE' in str(spec).upper() or '保险' in str(spec):
                    return (code, name, spec)

        return None

    def match_crystal(part, fp):
        """匹配晶振"""
        part_clean = strip_nc(part)
        part_norm = normalize(part_clean)
        for code, name, spec, cat in code_index.get('晶振', []):
            if part_norm in normalize(spec):
                return (code, name, spec)
        return None

    def match_semiconductor(part, fp):
        """匹配半导体（二极管、三极管、MOS管）"""
        part_clean = strip_nc(part)
        part_norm = normalize(part_clean)
        part_upper = str(part_clean).upper()
        fp_upper = str(fp).upper() if fp else ''
        # 提取封装中的数字部分和关键词部分
        fp_digits = extract_footprint(fp)

        # 同类替代映射（封装必须兼容，否则SMT焊不上）
        equiv_map = {
            '1N4148': 'LL4148',
            '2N3904': 'PMBT3904',
            'MMBT3904': 'PMBT3904',
            # 2N7002 移除：小信号MOS不应映射到功率MOS SI2328
        }
        equiv = equiv_map.get(part_upper)
        if equiv:
            equiv_candidates = []
            for code, name, spec, cat in code_index.get('半导体', []):
                if equiv in str(spec).upper() or equiv in str(name).upper():
                    equiv_candidates.append((code, name, spec))
            # 封装兼容性检查：用完整fp字符串去匹配规格中的封装信息
            if fp_upper and equiv_candidates:
                for code, name, spec in equiv_candidates:
                    spec_upper = str(spec).upper()
                    # 用去除连字符的fp匹配（SOD323匹配SOD-323，SOT23匹配SOT-23等）
                    fp_no_hyphen = fp_upper.replace('-', '').replace('_', '')
                    spec_no_hyphen = spec_upper.replace('-', '').replace('_', ' ').replace('(', '').replace(')', '')
                    if fp_no_hyphen in spec_no_hyphen or fp_upper in spec_upper:
                        return (code, name, spec)
                # 有equiv候选但封装不兼容 → 不匹配
                return None

        # 精确子串匹配
        for code, name, spec, cat in code_index.get('半导体', []):
            spec_norm = normalize(spec)
            name_upper = str(name).upper()
            if part_norm in spec_norm or part_norm in name_upper:
                return (code, name, spec)

        # 子串匹配（至少4字符）
        if len(part_norm) >= 4:
            for code, name, spec, cat in code_index.get('半导体', []):
                spec_norm = normalize(spec)
                if part_norm[:6] in spec_norm:
                    return (code, name, spec)

        # UDZ3.6B 特殊：匹配含 3.6B 的
        if 'UDZ' in part_upper or '3.6B' in part_upper:
            for code, name, spec, cat in code_index.get('半导体', []):
                if '3.6B' in str(spec) or '3.6V' in str(spec):
                    return (code, name, spec)

        # SMDJ 系列：匹配代码表中含 SMDJ 的TVS（需精确型号匹配）
        if 'SMDJ' in part_upper:
            for code, name, spec, cat in code_index.get('半导体', []):
                if 'SMDJ' in str(spec).upper() or 'SMDJ' in str(name).upper():
                    return (code, name, spec)
            return None

        # PESD 系列：匹配 ESD 保护（精确型号匹配）
        if 'PESD' in part_upper:
            for code, name, spec, cat in code_index.get('半导体', []):
                spec_upper = str(spec).upper()
                m = re.search(r'PESD[A-Z0-9]+', spec_upper)
                if m and m.group() == part_upper:
                    return (code, name, spec)

        # Green LED 匹配
        if 'GREEN' in part_upper:
            for code, name, spec, cat in code_index.get('半导体', []):
                if '绿色' in str(name) or 'GREEN' in str(spec).upper() or 'GHC' in str(spec).upper():
                    return (code, name, spec)

        return None

    def classify_device(ref, part, fp):
        """根据Part和Footprint判断器件类型"""
        part_upper = str(part).upper()
        fp_lower = str(fp).lower() if fp else ''

        # NC/ 不改变器件类型
        real_part = strip_nc(part)
        real_part_upper = real_part.upper()

        # 连接器关键词
        conn_kw = ['USB', 'DVI', 'HDMI', 'DP', 'RJ45', 'VGA', 'HEADER',
                    'CNF', 'PH-K', 'XH-A', 'XH', '-PH-', 'B8B', 'B4B', 'B3B',
                    'B10B', 'B2B', 'MOLEX', '1251S', 'CON-', 'DDR3',
                    'DISPLAY_PORT', 'B4PS', 'NC/MPC']
        for kw in conn_kw:
            if kw in real_part_upper or kw in fp_lower:
                return '插座'

        # ESD/TVS
        if 'CLAMP' in real_part_upper or 'RCLAMP' in real_part_upper or 'TVS' in real_part_upper:
            return '静电保护器'
        if 'PESD' in real_part_upper or 'LXES' in real_part_upper:
            return '静电保护器'

        # IC关键词
        ic_kw = ['MP9943', 'MP2307', 'SGM61430', 'RT9026', 'PI3HD',
                 'EEPROM', 'FLASH', '24C02', 'X24C16', 'SFLASH',
                 'RTD2797', 'RTD2795', 'DDR3']
        for kw in ic_kw:
            if kw in real_part_upper:
                return '集成电路'
        # U开头的位号通常是IC
        if 'U' == str(ref)[0] if ref else False:
            return '集成电路'
        if fp_lower.startswith('soic') or fp_lower.startswith('qfn') or 'bga' in fp_lower or 'tssop' in fp_lower:
            return '集成电路'
        # Q开头的位号通常是三极管/MOS
        if 'Q' == str(ref)[0] if ref else False:
            return '半导体'

        # 晶振
        if '14.318' in real_part or 'MHZ' in real_part_upper or 'Y' == str(fp)[0] if fp else False:
            return '晶振'
        if '3225' in fp_lower or 'HC-49' in real_part_upper:
            return '晶振'

        # 保险丝（排除FB磁珠）
        if 'FUSE' in real_part_upper or '保险' in real_part or 'RE-FUSE' in real_part_upper:
            return '保险丝'
        if 'F' == str(ref)[0] if ref else False:
            # F开头但FB开头是磁珠，不是保险丝
            if str(ref)[:2] == 'FB':
                return '电感'
            # F开头位号+含电压(A/V)的通常是保险丝
            if re.search(r'\d*A', real_part_upper) or re.search(r'\d*V', real_part_upper):
                return '保险丝'

        # 电感/磁珠
        if 'UH' in real_part_upper:
            return '电感'
        if 'BEAD' in real_part_upper or '/100MHZ' in real_part_upper:
            return '电感'
        if (fp_lower.startswith('l') or 'cm_' in fp_lower or 'cmfilter' in fp_lower) and 'led' not in fp_lower:
            return '电感'
        if '90R' == real_part:
            return '电感'
        if 'CR32' in real_part_upper or 'COIL' in real_part_upper:
            return '电感'

        # 电容
        if fp_lower.startswith('c') and 'cm' not in fp_lower:
            return '电容'
        if 'cap' in fp_lower:
            return '电容'

        # 电阻
        if fp_lower.startswith('r') and 'cm' not in fp_lower:
            return '电阻'

        # 二极管/MOS/三极管
        diode_kw = ['FDS4435', 'SI2328', 'B340', 'BAV70', 'PMBT3904', '1N4148',
                     'UDZ', 'PESD', 'SMDJ', 'SK54', 'GREEN', 'LED',
                     'SCHOTTKY', 'DIODE', 'NPN', 'MOSFET',
                     '2N3904', 'MMBT3904', '2N7002', 'RE-FUSE', 'FUSE']
        for kw in diode_kw:
            if kw in real_part_upper:
                if kw in ('RE-FUSE', 'FUSE'):
                    return '保险丝'
                return '半导体'

        # 磁珠
        if 'MPZ' in real_part_upper or 'BEAD' in real_part_upper or 'BLM' in real_part_upper:
            return '电感'
        if 'FB' == str(ref)[:2] if ref else False:
            return '电感'

        # 蜂鸣器、开关
        if 'BUZZER' in real_part_upper or '蜂鸣' in real_part:
            return '其他'
        if 'SW' in real_part_upper or 'HEATSINK' in real_part_upper or 'TP_' in fp_lower:
            return '其他'
        if 'RCLAMP' in real_part_upper:
            return '静电保护器'

        return '其他'

    # ========== 电容全局复用优化（严格模式）==========
    cap_groups = defaultdict(list)  # (cap_pF, fp_digits, volt) -> [(idx, ref, part, fp)]
    for i, (ref, part, fp, qty) in enumerate(bom_data):
        dtype = classify_device(ref, part, fp)
        if dtype == '电容':
            pc = strip_nc(part)
            cp, v = parse_cap_value(pc)
            fpd = extract_footprint(fp)
            if cp is not None and fpd:
                # 铝电解电容统一为尺寸码
                is_ae = 'AE' in str(fp).upper() or 'CAPAE' in str(fp).upper()
                if is_ae:
                    fpd = normalize_ae_footprint(fpd)
                cap_groups[(cp, fpd, v or 0)].append((i, ref, part, fp))

    cap_group_selection = {}
    for key, items in cap_groups.items():
        cp, fpd, volt = key
        candidates = []
        for code, name, spec, cat in code_index.get('电容', []):
            c_cap, c_volt, c_fp = parse_spec_cap(spec)
            if c_cap != cp: continue
            # 铝电解电容按尺寸匹配
            is_ae_item = '铝电解' in str(name)
            if is_ae_item:
                ae_code = extract_ae_size_from_spec(spec)
                if not ae_code:
                    ae_code = normalize_ae_footprint(c_fp)
                if not ae_code or ae_code != fpd:
                    continue
            else:
                if c_fp != fpd: continue  # 封装必须一致
            score = 0
            if volt > 0 and c_volt and c_volt == volt:
                score += 5
            elif volt > 0 and c_volt and c_volt >= volt:
                score += 3
            candidates.append((code, name, spec, score))

        if candidates:
            candidates.sort(key=lambda x: -x[3])
            cap_group_selection[key] = candidates[0][:3]

    # ========== 主匹配循环 ==========
    results = []
    unmatched = []

    for i, (ref, part, fp, qty) in enumerate(bom_data):
        dtype = classify_device(ref, part, fp)
        match = None

        if dtype == '电容':
            pc = strip_nc(part)
            cp, v = parse_cap_value(pc)
            fpd = extract_footprint(fp)
            # 铝电解电容统一为尺寸码以匹配去重key
            is_ae_loop = 'AE' in str(fp).upper() or 'CAPAE' in str(fp).upper()
            if is_ae_loop and fpd:
                fpd = normalize_ae_footprint(fpd)
            key = (cp, fpd, v or 0) if cp and fpd else None
            if key and key in cap_group_selection:
                match = cap_group_selection[key]
            else:
                match = match_capacitor(part, fp)
        elif dtype == '电阻':
            match = match_resistor(part, fp)
        elif dtype == '电感':
            match = match_inductor(part, fp)
        elif dtype == '插座':
            match = match_connector(part, fp)
        elif dtype == '集成电路':
            match = match_ic(part, fp)
        elif dtype == '静电保护器':
            match = match_esd(part, fp)
        elif dtype == '保险丝':
            match = match_fuse(part, fp)
        elif dtype == '晶振':
            match = match_crystal(part, fp)
        elif dtype == '半导体':
            match = match_semiconductor(part, fp)
        else:
            # 尝试所有分类（仅当part非空时）
            part_norm_fb = normalize(strip_nc(part))
            if part_norm_fb:
                for cat in all_categories:
                    for code, name, spec, cat_name in code_index[cat]:
                        if part_norm_fb in normalize(spec):
                            match = (code, name, spec)
                            dtype = cat
                            break
                    if match: break

        if match:
            code, name, spec = match
            results.append([ref, part, fp, qty, code, name, spec, dtype, ''])
            emit({'type': 'row', 'index': i, 'total': len(bom_data), 'ref': ref,
                  'part': part, 'fp': fp, 'dtype': dtype, 'matched': True, 'code': code})
        else:
            results.append([ref, part, fp, qty, '未匹配', '', '', dtype, ''])
            unmatched.append((ref, part, fp, dtype))
            emit({'type': 'row', 'index': i, 'total': len(bom_data), 'ref': ref,
                  'part': part, 'fp': fp, 'dtype': dtype, 'matched': False, 'code': ''})

    # ========== 候选搜索（供 GUI 人工复核）==========
    def find_candidates(row_index, limit=8):
        """给未匹配行找同分类近似候选料号，返回 [{code,name,spec,hits:[命中点]}]"""
        ref, part, fp, qty = bom_data[row_index]
        dtype = results[row_index][7]
        part_clean = strip_nc(part)
        part_norm = normalize(part_clean)
        fp_digits = extract_footprint(fp)
        part_upper = str(part_clean).upper()
        # 同分类优先，其次相邻分类
        related = {'电阻': ['电阻'], '电容': ['电容'],
                   '电感': ['电感', '其他'], '半导体': ['半导体', '静电保护器'],
                   '静电保护器': ['静电保护器', '半导体'],
                   '集成电路': ['集成电路', '开关稳压器', '半导体', '驱动IC'],
                   '插座': ['插座'], '保险丝': ['保险丝'], '晶振': ['晶振']}
        cats = related.get(dtype, list(code_index.keys()))
        # BOM 侧参数
        bom_res = parse_res_value(part_clean) if dtype == '电阻' else None
        bom_cap, bom_volt = (parse_cap_value(part_clean) if dtype == '电容' else (None, None))
        tokens = [t for t in re.findall(r'[A-Z0-9\.]{2,}', part_upper) if not t.isdigit() or len(t) >= 3]

        scored = []
        for cat in cats:
            for code, name, spec, cat_name in code_index.get(cat, []):
                spec_upper = str(spec).upper()
                spec_norm = normalize(spec)
                score = 0
                hits = []
                if dtype == '电阻' and bom_res is not None:
                    r_val, r_fp, r_tol = parse_spec_res(spec)
                    if r_val == bom_res:
                        score += 6; hits.append('阻值一致')
                    elif r_val and bom_res and 0.9 <= (r_val / max(bom_res, 1)) <= 1.1:
                        score += 2; hits.append('阻值相近')
                    if r_fp and r_fp == fp_digits:
                        score += 3; hits.append('封装一致')
                if dtype == '电容' and bom_cap is not None:
                    c_cap, c_volt, c_fp = parse_spec_cap(spec)
                    if c_cap == bom_cap:
                        score += 6; hits.append('容值一致')
                    if c_fp and c_fp == fp_digits:
                        score += 3; hits.append('封装一致')
                    if bom_volt and c_volt and c_volt >= bom_volt:
                        score += 1; hits.append('耐压满足')
                if part_norm and len(part_norm) >= 4 and part_norm[:6] in spec_norm:
                    score += 5; hits.append('型号相符')
                tok_hit = sum(1 for t in tokens if t in spec_upper)
                if tok_hit:
                    score += tok_hit
                    hits.append(f'{tok_hit}处关键词')
                if fp_digits and dtype not in ('电阻', '电容') and fp_digits in spec_upper:
                    score += 2; hits.append('封装出现')
                if score > 0:
                    scored.append((score, code, name, spec, hits))
        scored.sort(key=lambda x: -x[0])
        return [{'code': c, 'name': n, 'spec': s, 'hits': h, 'score': sc}
                for sc, c, n, s, h in scored[:limit]]

    # ========== 输出（参数化，人工改判后可重导出）==========
    def export_excel(res_list, out_path=None):
        path = out_path or OUTPUT_FILE
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = '匹配结果'

        headers = ['位号', 'Part Value', 'PCB Footprint', '数量', '物料长代码', '物料名称', '规格型号', '器件分类', '备注']
        header_font = openpyxl.styles.Font(bold=True, color='FFFFFF', size=11)
        header_fill = openpyxl.styles.PatternFill('solid', fgColor='4472C4')
        yellow_fill = openpyxl.styles.PatternFill('solid', fgColor='FFFF00')
        green_fill = openpyxl.styles.PatternFill('solid', fgColor='C6EFCE')
        thin_border = openpyxl.styles.Border(
            left=openpyxl.styles.Side(style='thin'), right=openpyxl.styles.Side(style='thin'),
            top=openpyxl.styles.Side(style='thin'), bottom=openpyxl.styles.Side(style='thin')
        )

        for col, h in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=h)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = openpyxl.styles.Alignment(horizontal='center')
            cell.border = thin_border

        for r_idx, row_data in enumerate(res_list, 2):
            for c_idx, val in enumerate(row_data, 1):
                cell = ws.cell(row=r_idx, column=c_idx, value=val)
                cell.border = thin_border
            if row_data[4] == '未匹配':
                for col in range(1, len(headers)+1):
                    ws.cell(row=r_idx, column=col).fill = yellow_fill
            elif row_data[8] == '人工指定':
                for col in range(1, len(headers)+1):
                    ws.cell(row=r_idx, column=col).fill = green_fill

        ws.column_dimensions['A'].width = 15
        ws.column_dimensions['B'].width = 25
        ws.column_dimensions['C'].width = 20
        ws.column_dimensions['D'].width = 8
        ws.column_dimensions['E'].width = 15
        ws.column_dimensions['F'].width = 18
        ws.column_dimensions['G'].width = 70
        ws.column_dimensions['H'].width = 12
        ws.column_dimensions['I'].width = 20

        # 未匹配/统计均从 res_list 重算，保证人工改判后一致
        cur_unmatched = [(r[0], r[1], r[2], r[7]) for r in res_list if r[4] == '未匹配']
        if cur_unmatched:
            ws2 = wb.create_sheet('未匹配物料')
            ws2.append(['位号', 'Part Value', 'PCB Footprint', '器件分类'])
            for ref, part, fp, dtype in cur_unmatched:
                ws2.append([ref, part, fp, dtype])

        ws3 = wb.create_sheet('匹配统计')
        ws3['A1'] = '匹配统计'
        ws3['A1'].font = openpyxl.styles.Font(bold=True, size=14)
        ws3['A3'] = 'BOM总行数:'
        ws3['B3'] = len(res_list)
        ws3['A4'] = '已匹配:'
        ws3['B4'] = len(res_list) - len(cur_unmatched)
        ws3['A5'] = '未匹配:'
        ws3['B5'] = len(cur_unmatched)
        ws3['A6'] = '匹配率:'
        ws3['B6'] = '%.1f%%' % ((len(res_list) - len(cur_unmatched)) / max(len(res_list), 1) * 100)

        ws3['A8'] = '按器件分类:'
        ws3['A8'].font = openpyxl.styles.Font(bold=True)
        ws3.append([])  # skip row 9
        ws3.append(['器件分类', 'BOM数', '已匹配', '未匹配', '匹配率'])
        type_stats = defaultdict(lambda: {'total': 0, 'matched': 0})
        for r in res_list:
            dtype = r[7]
            type_stats[dtype]['total'] += 1
            if r[4] != '未匹配':
                type_stats[dtype]['matched'] += 1
        for dtype in sorted(type_stats.keys()):
            s = type_stats[dtype]
            ws3.append([dtype, s['total'], s['matched'], s['total'] - s['matched'],
                         '%.1f%%' % (s['matched'] / max(s['total'], 1) * 100)])

        wb.save(path)
        return path

    export_excel(results)

    # 分类统计（给 GUI 用）
    cat_stats = defaultdict(lambda: {'total': 0, 'matched': 0})
    for r in results:
        cat_stats[r[7]]['total'] += 1
        if r[4] != '未匹配':
            cat_stats[r[7]]['matched'] += 1

    print(f"\n{'='*50}")
    print(f"匹配完成！输出: {OUTPUT_FILE}")
    print(f"BOM总行数: {len(bom_data)}")
    print(f"已匹配: {len(bom_data) - len(unmatched)}")
    print(f"未匹配: {len(unmatched)}")
    print(f"匹配率: {(len(bom_data) - len(unmatched)) / max(len(bom_data), 1) * 100:.1f}%")

    if unmatched:
        print(f"\n未匹配物料 ({len(unmatched)}条):")
        for ref, part, fp, dtype in unmatched:
            print(f"  {ref:15s} | {part:25s} | {fp:20s} | {dtype}")

    emit({'type': 'done', 'output_file': str(OUTPUT_FILE), 'total': len(bom_data),
          'matched': len(bom_data) - len(unmatched), 'unmatched': len(unmatched),
          'cat_stats': {k: dict(v) for k, v in cat_stats.items()}})

    return {
        'results': results,
        'unmatched': unmatched,
        'output_file': str(OUTPUT_FILE),
        'bom_count': len(bom_data),
        'categories': all_categories,
        'cat_stats': {k: dict(v) for k, v in cat_stats.items()},
        'export': export_excel,
        'find_candidates': find_candidates,
    }

def main():
    parser = argparse.ArgumentParser(description='BOM智能匹配 v2')
    parser.add_argument('--bom', required=True, help='BOM文件路径')
    parser.add_argument('--code', required=True, help='物料编码表路径')
    parser.add_argument('--output', help='输出文件路径')
    parser.add_argument('--debug', '-d', action='store_true')
    args = parser.parse_args()
    run_match(args.bom, args.code, args.output, debug=args.debug)

if __name__ == '__main__':
    main()
