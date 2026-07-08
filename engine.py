# -*- coding: utf-8 -*-
"""
BOM智能匹配 Skill - Production
"""
import argparse
import openpyxl
import re
import sys
import io
from datetime import datetime

# 强制使用UTF-8编码，解决Windows控制台编码问题；优先 reconfigure，避免重复 import 时反复包裹标准流。
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

# 全局调试标志（在main函数中设置）
DEBUG_MODE = False
VERBOSE_MODE = False

# 调试输出函数（全局可访问）
def debug_print(*args, **kwargs):
    if DEBUG_MODE:
        print(*args, **kwargs)

def verbose_print(*args, **kwargs):
    if VERBOSE_MODE:
        print(*args, **kwargs)

# GUI集成回调（由process_bom设置）
_log_callback = None
_progress_callback = None

def _set_log_callback(cb):
    global _log_callback
    _log_callback = cb

def _set_progress_callback(cb):
    global _progress_callback
    _progress_callback = cb

def _log(msg):
    """同时输出到控制台和GUI日志回调"""
    print(msg)
    if _log_callback:
        _log_callback(msg)

def _progress(pct, msg=''):
    """更新进度（0-100）"""
    if _progress_callback:
        _progress_callback(pct, msg)

# 逐行匹配回调（GUI 流式日志用）
_row_callback = None

def _set_row_callback(cb):
    global _row_callback
    _row_callback = cb

def _emit_row(refdes, part, fp, dtype, pn):
    if _row_callback:
        try:
            _row_callback({'ref': refdes, 'part': part, 'fp': fp,
                           'dtype': dtype, 'pn': pn, 'matched': pn != '未匹配'})
        except Exception:
            pass

def process_bom(bom_file, code_file, output_file=None, log_callback=None, progress_callback=None):
    """供GUI直接调用的入口函数"""
    _set_log_callback(log_callback)
    _set_progress_callback(progress_callback)
    return main(bom=bom_file, code=code_file, output=output_file)

def main(bom=None, code=None, output=None, verbose=False, debug=False):
    # CLI模式：解析命令行参数
    if bom is None and code is None:
        parser = argparse.ArgumentParser(description='BOM智能匹配工具')
        parser.add_argument('--bom', required=True, help='BOM文件路径')
        parser.add_argument('--code', required=True, help='物料编码表路径')
        parser.add_argument('--output', help='输出文件路径，默认自动生成')
        parser.add_argument('--verbose', '-v', action='store_true', help='启用详细输出')
        parser.add_argument('--debug', '-d', action='store_true', help='启用调试输出（包含verbose）')
        args = parser.parse_args()
        bom = args.bom
        code = args.code
        output = args.output
        verbose = args.verbose
        debug = args.debug

    # 设置全局调试模式
    global DEBUG_MODE, VERBOSE_MODE
    DEBUG_MODE = debug
    VERBOSE_MODE = verbose or debug

    BOM_FILE = bom
    CODE_FILE = code
    import os as _os
    OUTPUT_FILE = output or _os.path.join(
        _os.path.dirname(_os.path.abspath(bom)),
        f"BOM_匹配结果_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    )

    SPECIAL_RULES = [
        ('90R', 'CM_FILTER', 'Y0604-00030'),
        ('DVI', None, 'Y0614-00013'),
        ('DISPLAY_PORT-CON', None, 'Y0614-00018'),
        ('DP', 'DP_MOLEX_47272A', 'Y0614-00018'),
        ('GREEN', 'LED0603', 'Y0608-00007'),
        ('24C02', None, 'Y0612-00014'),
        ('24LC02', None, 'Y0612-00014'),
        ('X24C16', None, 'Y0612-00018'),
        ('RTD2785', None, 'Y0612-00106'),

        ('RTD2795XX-CG', None, 'Y0612-00108'),
        ('HEATSINK', None, 'Y1209-00005'),
        ('RE-FUSE/49FKA-002', 'F1812', 'Y0607-00008'),
        # 通用LDO规则：1117M-X.XV -> 查找对应电压的1117 LDO
        # 这个规则会在match_special_rules中特殊处理

        # 铝电解电容特殊规则（根据用户要求修改）
        # 47uF/35V capae660x770n -> Y0602-00066 (用户指定)
#         ('47UF/35V', 'CAPAE660X770N', 'Y0602-00066'),
        ('47UF/35V', 'AE660', 'Y0602-00066'),
        # 100uF/25V capae660x770n -> Y0602-00007 (100uF/25V, 6.3×7.7mm, 村田)
#         ('100UF/25V', 'CAPAE660X770N', 'Y0602-00007'),
        ('100UF/25V', 'AE660', 'Y0602-00007'),
        # 100uF/50V 铝电解，CAPAE840X1000N -> Y0602-00033 (8*10mm)
        ('100UF/50V', 'AE840', 'Y0602-00033'),
#         ('100UF/50V', 'CAPAE840X1000N', 'Y0602-00033'),
        # 100nF/0402电容特殊规则（用户要求匹配到Y0602-00068，禁止Y0602-00004）
        ('100NF/0402', 'C0402', 'Y0602-00068'),
        ('100NF', 'C0402', 'Y0602-00068'),

        # 磁珠 (Ferrite Bead) 规则：基于封装匹配 Y0606 系列
        # 简化逻辑：识别为磁珠后，直接根据封装匹配，不关心具体阻抗值
        # 0603 -> Y0606-00006 (100Ω@100MHz, 2A, 0603)
        ('BEAD', 'L0603', 'Y0606-00006'),
        ('FEB', 'L0603', 'Y0606-00006'),
        ('MPZ', 'L0603', 'Y0606-00006'),  # MPZ系列磁珠
        # 0805 -> Y0606-00003 (100Ω@100MHz, 4A, 0805)
        ('BEAD', 'L0805', 'Y0606-00003'),
        ('FEB', 'L0805', 'Y0606-00003'),
        ('MPZ', 'L0805', 'Y0606-00003'),  # MPZ系列磁珠
        # 1206 -> Y0606-00012 (33Ω@100MHz, 6A, 1206)
        ('BEAD', 'L1206', 'Y0606-00012'),
        ('FEB', 'L1206', 'Y0606-00012'),
        ('MPZ', 'L1206', 'Y0606-00012'),  # MPZ系列磁珠
        
        # xxxR/100MHz 格式磁珠（根据封装匹配）
        ('/100MHZ', 'L0603', 'Y0606-00006'),
        ('/100MHZ', 'L0805', 'Y0606-00003'),
        ('/100MHZ', 'L1206', 'Y0606-00012'),
        # 三极管 2N3904 / MMBT3904 -> Y0609-00001 (PMBT3904, NPN, 60V, 0.1A, SOT-23)
        ('2N3904', None, 'Y0609-00001'),
        ('MMBT3904', None, 'Y0609-00001'),

        ('3904', None, 'Y0609-00001'),
        # C4201WR系列连接器
        ('C4201WR', None, 'Y0614-00092'),
        # 稳压管特殊规则：UDZ3.6B -> Y0608-00008，若停用则自动切换到Y0608-00019
        ('UDZ3.6B', None, 'Y0608-00008'),
        # RTD2797显示控制器
        ('RTD2797', None, 'Y0612-00118'),
        ('BGA_32X32', 'RTD2797', 'Y0612-00118'),  # 处理Part和封装颠倒的情况
        # RTD2796显示控制器（应该匹配到Y0612-00027）
        ('RTD2796', None, 'Y0612-00027'),
    ]

    def match_special_rules(part, fp):
        part_clean = str(part).strip().upper()
        if part_clean.startswith('NC/'): part_clean = part_clean[3:]
        fp_clean = str(fp).strip().upper()
        # DEBUG for HEATSINK and RTD2797
        if 'HEAT' in part_clean or 'RTD' in part_clean or 'BGA_32X32' in part_clean:
            debug_print(f'match_special_rules: part_clean={part_clean}, fp_clean={fp_clean}')
            debug_print(f'检查RTD2796规则: rule_part=RTD2796, rule_fp=None, pn=Y0612-00027')
            if 'RTD2796' in part_clean:
                debug_print(f'RTD2796应该匹配到Y0612-00027')
        for rule_part, rule_fp, pn in SPECIAL_RULES:
            if rule_part and rule_part not in part_clean: continue
            if rule_fp and rule_fp not in fp_clean: continue
            
            # 特殊处理：UDZ3.6B -> Y0608-00008，若停用则自动切换到Y0608-00019
            if rule_part == 'UDZ3.6B' and pn == 'Y0608-00008':
                n, s = get_name_spec_by_pn(pn)
                if n and '停用' in str(s):  # Y0608-00008停用
                    debug_print(f'UDZ3.6B 匹配到 Y0608-00008 已停用，自动切换到 Y0608-00019')
                    # 检查Y0608-00019是否可用
                    n2, s2 = get_name_spec_by_pn('Y0608-00019')
                    if n2 and '停用' not in str(s2):
                        return 'Y0608-00019'
                    else:
                        debug_print(f'Y0608-00019 也停用或不存在，继续查找其他规则')
                        continue
                elif n:  # Y0608-00008未停用
                    return pn
                else:  # Y0608-00008不存在
                    continue
            
            # 检查匹配到的料号是否停用
            n, s = get_name_spec_by_pn(pn)
            if n is None:  # 料号停用或不存在
                if 'HEAT' in part_clean:
                    verbose_print(f'  -> 规则匹配但料号停用: {rule_part}, {rule_fp} -> {pn}')
                continue  # 继续查找其他规则
            
            if 'HEAT' in part_clean:
                verbose_print(f'  -> matched rule: {rule_part}, {rule_fp} -> {pn}')
            return pn
        
        return None

    # 工具
    def parse_capacitor_value(s):
        s = str(s).strip().upper()
        if s.startswith('NC/'): s = s[3:]
        
        # 改进的正则：匹配电容值和电压（支持/、空格、,、，分隔）
        # 匹配格式：47UF/35V, 47UF 35V, 47UF,35V, 100nF/0402, 0.1uF,C0402, 0.1uF，C0402
        # 注意：去掉了\b，因为后面可能紧跟/或,等非字母字符
        # 支持中文逗号（，）和英文逗号（,）
        m = re.search(r'([\d\.]+)\s*(UF|U|NF|N|PF|P)(?:\s*[,\/，]\s*(\d+(?:\.\d+)?)V)?', s)
        if not m: return None, None
        val = float(m.group(1)); unit = m.group(2)
        if unit in ('UF','U'): pF = int(round(val*1000000))
        elif unit in ('NF','N'): pF = int(round(val*1000))
        elif unit in ('PF','P'): pF = int(round(val))
        else: return None, None
        
        # 如果正则中已经捕获了电压（对于/或,分隔的情况）
        if m.group(3):
            return pF, float(m.group(3))
        
        # 如果没有捕获到，尝试匹配空格后的电压
        mV = re.search(r'\s+(\d+(?:\.\d+)?)V', s)
        if mV:
            return pF, float(mV.group(1))
        
        return pF, None

    def parse_spec_capacitor(s):
        if not s: return None, None, ''
        fp = (re.search(r'(\d{4})', s) or ['',''])[1]
        mV = re.search(r'(\d+(?:\.\d+)?)V', s, re.I)
        volt = float(mV.group(1)) if mV else None
        m = re.search(r'(?<![a-zA-Z])(\d+(?:\.\d+)?)\s*(UF|U|NF|N|PF|P)\b', s, re.I)
        if not m: return None, volt, fp
        val = float(m.group(1)); unit = m.group(2).upper()
        if unit in ('UF','U'): pF = int(round(val*1000000))
        elif unit in ('NF','N'): pF = int(round(val*1000))
        elif unit in ('PF','P'): pF = int(round(val))
        else: return None, volt, fp
        return pF, volt, fp

    def extract_footprint_digits(fp):
        '''提取封装数字
        1. 铝电解电容：capae660x770n -> 660 (提取直径数字)
        2. 铝电解电容：CAPAE840X1000N -> 840 (提取直径数字)
        3. 铝电解电容简写：AE660 -> 660
        4. 标准封装：C0603 -> 0603, L0805 -> 0805
        '''
        if not fp: return ''
        fp_str = str(fp).upper()
        
        # 先检查是否是铝电解电容（优先处理）
        if 'AE' in fp_str or 'CAPAE' in fp_str:
            # 铝电解电容格式：CAPAE840X1000N 或 AE660
            # 提取直径数字（第一个X前的数字）
            if 'X' in fp_str:
                # 找到第一个X的位置
                x_pos = fp_str.find('X')
                # 提取X之前的数字
                before_x = fp_str[:x_pos]
                digits_match = re.search(r'(\d+)', before_x)
                if digits_match:
                    return digits_match.group(1)
            else:
                # 简写格式：AE660
                digits_match = re.search(r'(\d+)', fp_str)
                if digits_match:
                    return digits_match.group(1)
        
        # 对于标准封装：C0603, L0805, R1206, SOD-323等，提取数字部分
        
        # 优先检查常见封装格式
        # 1. SOD-323, SOT-23, SOIC-8 等格式
        for prefix in ['SOD-', 'SOT-', 'SOIC-', 'DO-', 'LL-']:
            if prefix in fp_str:
                # 提取前缀后的数字
                pattern = prefix + r'(\d+)'
                match = re.search(pattern, fp_str)
                if match:
                    return match.group(1)
        
        # 2. 匹配连续的数字（至少2位）
        match = re.search(r'(\d{2,})', fp_str)
        if match:
            digits = match.group(1)
            # 如果是4位数字（如0805），直接返回
            if len(digits) == 4:
                return digits
            # 如果是3位数字
            elif len(digits) == 3:
                # 060 -> 0603（简写）
                if digits.startswith('060'):
                    return '0603'
                # 其他3位数字如323、120等，直接返回
                return digits
            # 其他情况返回提取的数字
            return digits
        
        return ''
    def normalize_string(s): return re.sub(r'[^a-z0-9\u4e00-\u9fff]', '', str(s).lower())
    def extract_connector_type(s):
        s = str(s).strip().upper()
        if 'C4201WR' in s: return 'C4201WR'
        if 'HEADER' in s: return 'HEADER'
        if 'DVI' in s: return 'DVI'
        if 'USB' in s: return 'USB'
        if 'HDMI' in s: return 'HDMI'
        if 'DP' in s or 'DISPLAY' in s or 'PORT' in s: return 'DP'
        if 'RJ45' in s: return 'RJ45'
        if 'VGA' in s: return 'VGA'
        if 'ETHERNET' in s: return 'ETHERNET'
        if re.search(r'B\d+B-PH-K-S', s) or 'PH-K-S' in s: return 'JST_PH'
        if re.search(r'B\d+B-XH', s) or 'XH' in s: return 'JST_XH'
        # DC JACK / power connector patterns (MPC, DC, DC JACK)
        if re.match(r'^MPC[-\s]', s) or 'DC JACK' in s or 'DC_RA' in s:
            return 'DC_JACK'
        return ''

    def load_bom():
        wb = openpyxl.load_workbook(BOM_FILE, data_only=True, read_only=True)
        ws = wb.active
        rows = list(ws.iter_rows(values_only=True))
        wb.close()
        for idx, row in enumerate(rows):
            header = [re.sub(r'\s+', ' ', str(c).strip().lower()) if c else '' for c in row]
            has_ref = 'reference' in header
            has_part = 'part' in header
            has_fp = 'pcb footprint' in header or 'footprint' in header
            if has_ref and has_part and has_fp:
                return header, rows[idx + 1:]
        raise ValueError("BOM表头未找到：需要包含 Reference、Part、PCB Footprint 或 Footprint")

    def load_code():
        wb = openpyxl.load_workbook(CODE_FILE, data_only=True, read_only=True)
        all_rows = []
        target_names = []
        for name in wb.sheetnames:
            ws = wb[name]
            normalized_rows = []
            for row in ws.iter_rows(values_only=True):
                cells = list(row)
                pn_idx = None
                for i, cell in enumerate(cells):
                    if cell and re.match(r'^Y(?:06|12)\d{2}-\d+', str(cell).strip(), re.I):
                        pn_idx = i
                        break
                if pn_idx is None:
                    continue
                pn = str(cells[pn_idx]).strip()
                item_name = str(cells[pn_idx + 1]).strip() if pn_idx + 1 < len(cells) and cells[pn_idx + 1] is not None else ''
                spec = str(cells[pn_idx + 2]).strip() if pn_idx + 2 < len(cells) and cells[pn_idx + 2] is not None else ''
                normalized_rows.append(('', pn, item_name, spec))
            if normalized_rows:
                all_rows.extend(normalized_rows)
                target_names.append(name)
        wb.close()
        if not all_rows:
            raise ValueError("编码表未找到 Y06/Y12 料号行")
        return ', '.join(target_names), all_rows

    bom_header, bom_data = load_bom()
    code_sheet, code_data = load_code()
    _log(f"BOM: {len(bom_data)} 行, 编码表: {code_sheet}")
    _progress(15, '加载文件完成')

    col_ref = bom_header.index('reference')
    col_part = bom_header.index('part')
    col_fp = bom_header.index('pcb footprint') if 'pcb footprint' in bom_header else bom_header.index('footprint')
    col_qty = bom_header.index('quantity') if 'quantity' in bom_header else (bom_header.index('qty') if 'qty' in bom_header else None)
    col_pn, col_name, col_spec = 1, 2, 3
    _log(f"Quantity列: {'找到' if col_qty is not None else '未找到'}" + (f", 索引={col_qty}" if col_qty is not None else ""))

    # 分类索引
    code_index = {'电容':[], '电阻':[], '电感':[], '连接器':[], '其他':[]}
    for row in code_data:
        if len(row)<4: continue
        pn = str(row[col_pn]).strip()
        name = str(row[col_name]).strip()
        spec = str(row[col_spec]).strip()
        cat = '其他'
        if pn.startswith('Y0601'): cat='电阻'
        elif pn.startswith('Y0602') or '电容' in name or 'capacitor' in name.lower(): cat='电容'
        elif pn.startswith('Y0603') or pn.startswith('Y0604') or '电感' in name or 'inductor' in name.lower(): cat='电感'
        elif pn.startswith('Y0614') or '连接器' in name or 'connector' in name.lower(): cat='连接器'
        elif pn.startswith('Y0612') or '存储' in name or 'eeprom' in name.lower() or 'eeprom' in spec.lower(): cat='其他'
        code_index[cat].append((pn, name, spec))

    _log(f"索引: 电容{len(code_index['电容'])} 电阻{len(code_index['电阻'])} 电感{len(code_index['电感'])} 连接器{len(code_index['连接器'])} 其他{len(code_index['其他'])}")
    _progress(25, '编码表索引建立完成')

    def get_name_spec_by_pn(pn):
        '''获取料号规格，并检查是否停用'''
        for cat in ('电容','电阻','电感','连接器','其他'):
            for p, n, s in code_index[cat]:
                if p == pn: 
                    # 全局停用过滤：如果规格中包含"停用"，返回None表示不可用
                    spec_str = str(s).strip()
                    if '停用' in spec_str:
                        debug_print(f'停用过滤: 料号 {pn} 已停用，规格: {spec_str[:60]}...')
                        return None, None
                    return n, s
        return None, None

    # 电阻辅助
    def extract_resistor_components(spec_str):
        s = str(spec_str).strip().upper()
        # 替换全角逗号为英文逗号，以便分割；如果还是无法分割，则使用字符集搜索
        s = s.replace('，', ',').replace('、', ',')
        # 同时移除常见干扰字符（乱码等），但不影响数字和字母
        # 先尝试按逗号或分号分割
        if ',' in s or ';' in s:
            parts = re.split(r'[,;]', s)
        else:
            # 如果没有明显分隔符，则逐字符扫描寻找模式
            parts = [s]
        resistance = None
        # 在第一轮中，尝试从每个 part 提取阻值
        for p in parts:
            p = p.strip()
            if not p: continue
            # 三位一体: 8K2, 4K7, 2R2
            m = re.match(r'^(\d+)([KMG])(\d+)$', p)
            if m:
                val = float(m.group(1)+'.'+m.group(3))
                mult = m.group(2)
                if mult=='K': val*=1000
                elif mult=='M': val*=1000000
                elif mult=='G': val*=1000000000
                resistance = int(round(val)); break
            # 标准: 10K, 2.2K, 100R
            m = re.match(r'^([\d\.]+)([KMG]?)(R?)$', p)
            if m:
                vs, mult, r_flag = m.groups()
                if vs.isdigit() and len(vs)==4: continue  # skip 4-digit footprint like 0402
                try: val = float(vs)
                except: continue
                if mult=='K': val*=1000
                elif mult=='M': val*=1000000
                elif mult=='G': val*=1000000000
                resistance = int(round(val)); break
        if resistance is None:
            # 尝试在整个字符串中搜索阻值模式
            # 搜索 数字+K/M/G 或 数字+数字+K/M/G 形式
            m = re.search(r'(\d+)([KMG])(\d+)', s)
            if m:
                val = float(m.group(1) + '.' + m.group(3))
                mult = m.group(2)
                if mult=='K': val *= 1000
                elif mult=='M': val *= 1000000
                elif mult=='G': val *= 1000000000
                resistance = int(round(val))
            else:
                # 搜索 数字+(K/M/G)? 在单词边界
                m = re.search(r'(\d+(?:\.\d+)?)([KMG]?)(?=[^0-9]|$)', s)
                if m:
                    try:
                        val = float(m.group(1))
                    except:
                        val = None
                    if val is not None:
                        mult = m.group(2)
                        if mult=='K': val *= 1000
                        elif mult=='M': val *= 1000000
                        elif mult=='G': val *= 1000000000
                        resistance = int(round(val))
        # 提取封装: 搜索 4-digit pattern
        footprint = None
        for p in parts:
            fp_match = re.search(r'(\d{4})', p)
            if fp_match and not p[fp_match.start()-1].isalpha():  # simple heuristic
                footprint = fp_match.group(1)
                break
        if not footprint:
            # 在整个字符串中搜索 4-digit 数字，常见封装 0402, 0603, 0805, 1206, 1210, etc.
            m = re.search(r'(\d{4})', s)
            if m:
                footprint = m.group(1)
        # 容差
        tolerance = None
        mt = re.search(r'(\d+)%', s)
        if mt:
            try:
                tolerance = int(mt.group(1))
            except:
                pass
        return resistance, footprint, tolerance

    def parse_resistor_value(part_str):
        s = str(part_str).strip().upper()
        # Remove all NC markers
        s = s.replace('NC', '')
        
        # 先检查是否有*N模式（表示N个相同电阻）
        array_count = 1
        array_match = re.search(r'\*(\d+)', s)
        if array_match:
            array_count = int(array_match.group(1))
            # 移除*N部分以便解析电阻值
            s = re.sub(r'\*\d+', '', s)
        
        # Split tokens on common separators: whitespace, %, /, _
        tokens = re.split(r'[\s%/_]', s)
        for tok in tokens:
            tok = tok.strip()
            if not tok:
                continue
            # Three-part: 8K2, 4K7, 2R2, 2M5
            if re.fullmatch(r'\d+[KMGR]\d+', tok):
                m = re.fullmatch(r'(\d+)([KMGR])(\d+)', tok)
                val = float(m.group(1) + '.' + m.group(3))
                mult = m.group(2)
                if mult in ('K', 'R'):
                    val *= 1000
                elif mult == 'M':
                    val *= 1000000
                elif mult == 'G':
                    val *= 1000000000
                return int(round(val)), array_count
            # Standard: 10K, 2.2K, 100R, or pure number
            m = re.fullmatch(r'(\d+(?:\.\d+)?)([KMGR]?)', tok)
            if m:
                try:
                    val = float(m.group(1))
                except ValueError:
                    continue
                mult = m.group(2)
                if mult == 'K':
                    val *= 1000
                elif mult == 'M':
                    val *= 1000000
                elif mult == 'G':
                    val *= 1000000000
                # R or empty: no change
                return int(round(val)), array_count
        return None, array_count
    def parse_inductor_value(part_str):
        s = str(part_str).strip().upper()
        if s.startswith('NC/'): s = s[3:]
        m = re.search(r'(\d+(?:\.\d+)?)\s*(UH|U|MH|M|NH|N|H)\b', s, re.I)
        if m:
            val = float(m.group(1)); unit = m.group(2).upper()
            if unit in ('UH','U'): return val
            if unit in ('MH','M'): return val*1000
            if unit in ('NH','N'): return val/1000
            if unit=='H': return val*1000000
            return val
        m2 = re.search(r'(\d+)R(\d+)', s)
        if m2: return float(f"{m2.group(1)}.{m2.group(2)}")
        m3 = re.search(r'([\d\.]+)', s)
        if m3: return float(m3.group(1))
        return None

    # 匹配函数 - 基础版本（不包含优化）
    def match_capacitor_basic(part, fp):
        # 先检查特殊规则
        special_pn = match_special_rules(part, fp)
        if special_pn:
            n, s = get_name_spec_by_pn(special_pn)
            if n: return special_pn, n, s
        
        # 特殊处理：铝电解电容通用匹配
        part_upper = str(part).upper()
        fp_upper = str(fp).upper() if fp else ''
        
        if 'AE' in fp_upper or 'CAPAE' in fp_upper:
            # 1. 解析电容值和电压
            cap_pF, volt = parse_capacitor_value(part)
            if cap_pF and volt:
                # 2. 在编码表中查找匹配的铝电解电容
                best_match = None
                best_score = float('inf')
                
                for pn, name, spec in code_index.get('电容', []):
                    spec_upper = str(spec).upper()
                    
                    # 检查是否是铝电解电容
                    is_aluminum = (
                        '电解' in str(name) or 
                        'ALUMINUM' in spec_upper or 
                        'VEU' in spec_upper or 
                        'VZH' in spec_upper or
                        '×' in spec_upper or 'X' in spec_upper  # 包含尺寸
                    )
                    
                    if not is_aluminum:
                        continue
                    
                    # 解析规格
                    c_cap, c_volt, c_fp = parse_spec_capacitor(spec)
                    if not c_cap or not c_volt:
                        continue
                    
                    # 检查电容值匹配（允许一定误差）
                    cap_diff = abs(c_cap - cap_pF) / cap_pF
                    if cap_diff > 0.5:  # 容量差异超过50%，不匹配
                        continue
                    
                    # 检查电压兼容：编码表电压 ≥ 要求电压
                    if c_volt < volt:
                        continue
                    
                    # 计算匹配分数：容量差异 + 电压余量
                    volt_margin = (c_volt - volt) / volt  # 电压余量比例
                    score = cap_diff + volt_margin * 0.1  # 容量差异权重更高
                    
                    if score < best_score:
                        best_score = score
                        best_match = pn
                
                if best_match:
                    n, s = get_name_spec_by_pn(best_match)
                    if n:
                        debug_print(f'铝电解电容通用匹配: {part} ({fp}) -> {best_match} (分数: {best_score:.3f})')
                        return best_match, n, s
        
        cap_pF, volt = parse_capacitor_value(part)
        fp_digits = extract_footprint_digits(fp)
        if cap_pF is None: return None
        
        # 收集所有可能的匹配
        all_matches = []
        
        # 如果有明确电压，优先严格匹配，如果没有则允许向上兼容（更高电压）
        if volt is not None:
            # 首先尝试严格匹配
            for pn, name, spec in code_index['电容']:
                c_cap, c_volt, c_fp = parse_spec_capacitor(spec)
                if c_cap == cap_pF and c_volt == volt and (fp_digits == '' or c_fp == fp_digits):
                    all_matches.append((pn, name, spec, c_volt, 0))  # 优先级0：严格匹配
            
            # 如果没有严格匹配，寻找更高电压的兼容料号
            if not all_matches:
                for pn, name, spec in code_index['电容']:
                    c_cap, c_volt, c_fp = parse_spec_capacitor(spec)
                    # 电容值相同，封装匹配，且电压更高（兼容）
                    if c_cap == cap_pF and (fp_digits == '' or c_fp == fp_digits) and c_volt is not None and c_volt > volt:
                        all_matches.append((pn, name, spec, c_volt, 1))  # 优先级1：电压兼容
        else:
            # 没有明确电压要求，匹配所有电容值相同的料号
            for pn, name, spec in code_index['电容']:
                c_cap, c_volt, c_fp = parse_spec_capacitor(spec)
                if c_cap == cap_pF and (fp_digits == '' or c_fp == fp_digits):
                    all_matches.append((pn, name, spec, c_volt if c_volt else 0, 2))  # 优先级2：无电压要求
        
        if not all_matches:
            return None
        
        # 默认选择电压最高的兼容料号（安全原则：高电压可以替代低电压）
        all_matches.sort(key=lambda x: (x[3] if x[3] is not None else 0), reverse=True)
        return all_matches[0][:3]
    
    # 匹配函数 - 带物料种类优化
    def match_capacitor_optimized(part, fp, capacitor_group_selection):
        # 先检查特殊规则
        special_pn = match_special_rules(part, fp)
        if special_pn:
            n, s = get_name_spec_by_pn(special_pn)
            if n: return special_pn, n, s
        
        cap_pF, volt = parse_capacitor_value(part)
        fp_digits = extract_footprint_digits(fp)
        if cap_pF is None: return None
        
        # 构建分组键
        volt_key = volt if volt is not None else 0
        group_key = (cap_pF, fp_digits, volt_key)
        
        # 如果优化选择中有这个分组，使用优化后的选择
        if group_key in capacitor_group_selection:
            pn = capacitor_group_selection[group_key]
            n, s = get_name_spec_by_pn(pn)
            if n:
                return pn, n, s
        
        # 回退到基础匹配
        return match_capacitor_basic(part, fp)

    def match_resistor(part, fp):
        part_clean = str(part).strip()
        if part_clean.upper().startswith('NC/'): part_clean = part_clean[3:]
        fp_digits = extract_footprint_digits(fp)
        part_norm = normalize_string(part_clean)
        bom_val_result = parse_resistor_value(part_clean)
        if bom_val_result is None:
            bom_val = None
            array_count = 1
        else:
            bom_val, array_count = bom_val_result
        
        if part_norm in ('0r','0'):
            cands = []
            for pn, name, spec in code_index['电阻']:
                r_val, r_fp, r_tol = extract_resistor_components(spec)
                if r_val is not None and r_val == 0 and (not fp_digits or fp_digits == r_fp):
                    cands.append((pn, name, spec, r_tol))
            if cands:
                cands.sort(key=lambda x: x[3] if x[3] else 999)
                return cands[0][:3]
            return None
        
        if bom_val is None:
            for pn, name, spec in code_index['电阻']:
                spec_norm = normalize_string(spec)
                if part_norm in spec_norm and (not fp_digits or fp_digits in spec_norm):
                    return pn, name, spec
            return None
        matches = []
        for pn, name, spec in code_index['电阻']:
            r_val, r_fp, r_tol = extract_resistor_components(spec)
            if r_val != bom_val:
                continue
            # 如果BOM有封装要求，则必须匹配封装
            if fp_digits:
                if r_fp and r_fp == fp_digits:
                    matches.append((pn, name, spec, r_tol))
            else:
                # BOM无封装要求，接受任何封装
                matches.append((pn, name, spec, r_tol))
        if matches:
            matches.sort(key=lambda x: x[3] if x[3] else 999)
            
            # 检查是否是电阻阵列（*N模式）
            if array_count > 1:
                # 尝试查找电阻阵列器件
                array_keywords = ['排阻', '电阻排', '电阻网络', 'ARRAY', 'NETWORK', f'×{array_count}', f'*{array_count}', f'{array_count}P', f'{array_count}IN']
                array_matches = []
                
                for pn, name, spec in code_index['电阻']:
                    spec_upper = str(spec).upper()
                    name_upper = str(name).upper()
                    
                    # 检查是否是阵列器件
                    is_array = any(keyword in name_upper or keyword in spec_upper for keyword in array_keywords)
                    
                    if is_array:
                        r_val, r_fp, r_tol = extract_resistor_components(spec)
                        # 检查电阻值是否匹配（阵列通常有多个电阻，这里简化处理）
                        if r_val == bom_val or (fp_digits and r_fp == fp_digits):
                            array_matches.append((pn, name, spec, r_tol))
                
                if array_matches:
                    array_matches.sort(key=lambda x: x[3] if x[3] else 999)
                    pn, name, spec, _ = array_matches[0]
                    # 在规格中添加阵列信息
                    spec_with_array = f"{spec} (阵列: {array_count}个{bom_val}Ω电阻)"
                    return pn, name, spec_with_array
            
            # 如果不是阵列或没找到阵列器件，返回单个电阻
            pn, name, spec, _ = matches[0]
            
            # 如果是阵列需求但没找到阵列器件，在规格中注明需要多个
            if array_count > 1:
                spec = f"{spec} (需用{array_count}个)"
        if fp_digits:
            # BOM有明确封装，必须找到封装匹配的候选，否则不应匹配到封装不匹配的料号
            alt = []
            for pn, name, spec in code_index['电阻']:
                r_val, r_fp, r_tol = extract_resistor_components(spec)
                if r_val == bom_val and r_fp == fp_digits:
                    alt.append((pn, name, spec, r_tol))
            if alt:
                alt.sort(key=lambda x: x[3] if x[3] else 999)
                return alt[0][:3]
            # 没有封装匹配，直接返回未匹配
            return None
        else:
            # BOM无封装，可放宽匹配
            alt = []
            for pn, name, spec in code_index['电阻']:
                r_val, r_fp, r_tol = extract_resistor_components(spec)
                if r_val == bom_val:
                    alt.append((pn, name, spec, r_tol))
            if alt:
                alt.sort(key=lambda x: x[3] if x[3] else 999)
                return alt[0][:3]
            # 最后尝试字符串模糊匹配
            for pn, name, spec in code_index['电阻']:
                spec_norm = normalize_string(spec)
                if part_norm in spec_norm:
                    return pn, name, spec
            return None

    def match_inductor_with_candidates(part, fp):
        special_pn = match_special_rules(part, fp)
        if special_pn:
            n, s = get_name_spec_by_pn(special_pn)
            if n: return (special_pn, n, s), []
        part_clean = str(part).strip()
        if part_clean.upper().startswith('NC/'): part_clean = part_clean[3:]
        val_part = part_clean.split('/')[0].split()[0].strip()
        bom_val = parse_inductor_value(val_part)
        fp_digits = extract_footprint_digits(fp)
        candidates = []
        
        # 简化磁珠匹配逻辑：识别磁珠后直接根据封装匹配
        # 磁珠识别条件（任一即可）：
        # 1. 器件名包含 BEAD、FEB
        # 2. 器件名以 MPZ 开头
        # 3. 器件名包含 /100MHz（如 220R/100MHz）
        # 匹配规则：只根据封装，不关心具体阻抗值
        val_part_lower = val_part.lower()
        is_ferrite_bead = (
            'bead' in val_part_lower or 
            'feb' in val_part_lower or
            val_part_lower.startswith('mpz') or
            '/100mhz' in val_part_lower
        )
        
        if is_ferrite_bead:
            # 磁珠封装映射：根据封装直接匹配料号
            ferrite_bead_mapping = {
                '0603': 'Y0606-00006',  # 100Ω@100MHz, 2A, 0603
                '0805': 'Y0606-00003',  # 100Ω@100MHz, 4A, 0805
                '1206': 'Y0606-00012',  # 33Ω@100MHz, 6A, 1206
            }
            

            
            if fp_digits in ferrite_bead_mapping:
                pn = ferrite_bead_mapping[fp_digits]
                n, s = get_name_spec_by_pn(pn)
                if n:
        
                    return (pn, n, s), []

            
            # 如果封装映射失败，尝试在Y0606系列中查找
            for pn, n, s in code_index['电感']:
                if not pn.startswith('Y0606'): continue
                if fp_digits and fp_digits not in s.replace(',', ''): continue
                candidates.append((pn, n, s))
            
            if candidates:
                best = candidates[0]
                debug_print(f'磁珠备选匹配: {part_clean} ({fp}) -> {best[0]}')
                return best, candidates
            
            return None, candidates
        
        # 原有的FEB匹配逻辑（保持兼容）
        if 'feb' in normalize_string(val_part):
            mapping = {'0805':'Y0606-00003','0603':'Y0606-00006','1206':'Y0606-00012'}
            if fp_digits in mapping:
                pn = mapping[fp_digits]; n,s = get_name_spec_by_pn(pn)
                if n: return (pn,n,s), []
            for pn,n,s in code_index['电感']:
                if not pn.startswith('Y0606'): continue
                if fp_digits and fp_digits not in s.replace(',',''): continue
                candidates.append((pn,n,s))
            if candidates: return candidates[0], candidates
            return None, candidates
        if val_part in ('90R','100R'):
            for pn,n,s in code_index['电感']:
                if val_part in normalize_string(s): return (pn,n,s), []
        # 普通电感: 先按感值收集候选，封装匹配的优先，但不硬过滤
        fp_all_nums = re.findall(r'\d+', str(fp)) if fp else []
        exact_fp_cands = []
        other_cands = []
        for pn,n,s in code_index['电感']:
            code_uH = None
            m = re.search(r'(\d+(?:\.\d+)?)\s*(UH|U|MH|M|NH|N|H)\b', s, re.I)
            if m:
                try:
                    val = float(m.group(1)); u = m.group(2).upper()
                    if u in ('UH','U'): code_uH = val
                    elif u in ('MH','M'): code_uH = val*1000
                    elif u in ('NH','N'): code_uH = val/1000
                    elif u=='H': code_uH = val*1000000
                except: code_uH = None
            if code_uH is None:
                m2 = re.search(r'(\d+)R(\d+)', s)
                if m2:
                    try: code_uH = float(f"{m2.group(1)}.{m2.group(2)}")
                    except: code_uH = None
            if code_uH is None or bom_val is None: continue
            if abs(bom_val - code_uH) > 0.01: continue
            spec_norm = normalize_string(s)
            # 用封装全部数字组做匹配(>=2位)，至少一组命中则视为封装匹配
            fp_match = any(len(num) >= 2 and num in spec_norm for num in fp_all_nums) if fp_all_nums else False
            if fp_match:
                exact_fp_cands.append((pn,n,s))
            else:
                other_cands.append((pn,n,s))
        candidates = exact_fp_cands + other_cands
        if not candidates:
            for pn,n,s in code_index['其他']:
                code_uH = None
                m = re.search(r'(\d+(?:\.\d+)?)\s*(UH|U|MH|M|NH|N|H)\b', s, re.I)
                if m:
                    try:
                        val = float(m.group(1)); u = m.group(2).upper()
                        if u in ('UH','U'): code_uH = val
                        elif u in ('MH','M'): code_uH = val*1000
                        elif u in ('NH','N'): code_uH = val/1000
                        elif u=='H': code_uH = val*1000000
                    except: code_uH = None
                if code_uH is None:
                    m2 = re.search(r'(\d+)R(\d+)', s)
                    if m2:
                        try: code_uH = float(f"{m2.group(1)}.{m2.group(2)}")
                        except: code_uH = None
                if code_uH is None or bom_val is None: continue
                if abs(bom_val - code_uH) > 0.01: continue
                spec_norm = normalize_string(s)
                fp_match = any(len(num) >= 2 and num in spec_norm for num in fp_all_nums) if fp_all_nums else False
                if fp_match:
                    exact_fp_cands.append((pn,n,s))
                else:
                    other_cands.append((pn,n,s))
            candidates = exact_fp_cands + other_cands
        if candidates:
            return candidates[0], candidates
        return None, []

    def match_connector(part, fp):
        special_pn = match_special_rules(part, fp)
        if special_pn:
            n, s = get_name_spec_by_pn(special_pn)
            if n: return special_pn, n, s
        part_clean = str(part).strip()
        if part_clean.upper().startswith('NC/'): part_clean = part_clean[3:]
        part_norm = normalize_string(part_clean)
        fp_digits = extract_footprint_digits(fp)
        conn_type = extract_connector_type(part_clean)
        

        
        # 专门处理C4201WR系列连接器
        if conn_type == 'C4201WR':
            # 提取针数信息：C4201WR-F1-2x3P -> 2x3P
            import re
            cjt_match = re.search(r'C4201WR(?:-F1)?-(\d+)[xX](\d+)P?', part_clean)
            if cjt_match:
                rows, pins = cjt_match.group(1), cjt_match.group(2)
                total_pins = str(int(rows) * int(pins))
                # 在连接器中查找匹配的料号
                for pn, n, s in code_index['连接器']:
                    spec_upper = s.upper()
                    # 检查规格中是否包含相同的针数信息
                    if total_pins in spec_upper or f'{pins}P' in spec_upper:
                        return pn, n, s
                # 也在其他中查找
                for pn, n, s in code_index['其他']:
                    spec_upper = s.upper()
                    if total_pins in spec_upper or f'{pins}P' in spec_upper:
                        return pn, n, s
        
        # 专门处理HEADER系列连接器
        if conn_type == 'HEADER':
            # 提取针数信息：HEADER_15*2W -> 15行×2列=30针
            # 或者 con-2_0-2x15p-v -> 2x15p
            import re
            # 尝试从part中提取针数
            header_match = re.search(r'HEADER[_-](\d+)[*Xx](\d+)', part_clean)
            if header_match:
                rows, cols = header_match.group(1), header_match.group(2)
                total_pins = str(int(rows) * int(cols))
                
                # 在连接器中查找匹配的料号
                for pn, n, s in code_index['连接器']:
                    spec_upper = s.upper()
                    # 更精确的针数匹配：检查是否包含"2*15"、"15*2"、"30P"或"30PIN"等
                    # 避免误匹配（如"1300"包含"30"）
                    has_2x15 = f'{rows}*{cols}' in spec_upper or f'{cols}*{rows}' in spec_upper
                    has_2x15_alt = f'{rows}X{cols}' in spec_upper or f'{cols}X{rows}' in spec_upper
                    has_30p = f'{total_pins}P' in spec_upper or f'{total_pins}PIN' in spec_upper
                    has_exact_30 = f' {total_pins} ' in f' {spec_upper} '  # 独立的30
                    
                    if has_2x15 or has_2x15_alt or has_30p or has_exact_30:
                        return pn, n, s
                
                # 也在其他中查找
                for pn, n, s in code_index['其他']:
                    spec_upper = s.upper()
                    has_2x15 = f'{rows}*{cols}' in spec_upper or f'{cols}*{rows}' in spec_upper
                    has_2x15_alt = f'{rows}X{cols}' in spec_upper or f'{cols}X{rows}' in spec_upper
                    has_30p = f'{total_pins}P' in spec_upper or f'{total_pins}PIN' in spec_upper
                    has_exact_30 = f' {total_pins} ' in f' {spec_upper} '  # 独立的30
                    
                    if has_2x15 or has_2x15_alt or has_30p or has_exact_30:
                        return pn, n, s
            
            # 如果从part中没提取到，尝试从footprint中提取
            fp_clean = str(fp).strip().upper()
            fp_match = re.search(r'(\d+)[Xx](\d+)P', fp_clean)
            if fp_match:
                rows, cols = fp_match.group(1), fp_match.group(2)
                total_pins = str(int(rows) * int(cols))
                
                # 在连接器中查找匹配的料号
                for pn, n, s in code_index['连接器']:
                    spec_upper = s.upper()
                    # 使用相同的精确匹配逻辑
                    has_2x15 = f'{rows}*{cols}' in spec_upper or f'{cols}*{rows}' in spec_upper
                    has_2x15_alt = f'{rows}X{cols}' in spec_upper or f'{cols}X{rows}' in spec_upper
                    has_30p = f'{total_pins}P' in spec_upper or f'{total_pins}PIN' in spec_upper
                    has_exact_30 = f' {total_pins} ' in f' {spec_upper} '  # 独立的30
                    
                    if has_2x15 or has_2x15_alt or has_30p or has_exact_30:
                        return pn, n, s
                
                # 也在其他中查找
                for pn, n, s in code_index['其他']:
                    spec_upper = s.upper()
                    has_2x15 = f'{rows}*{cols}' in spec_upper or f'{cols}*{rows}' in spec_upper
                    has_2x15_alt = f'{rows}X{cols}' in spec_upper or f'{cols}X{rows}' in spec_upper
                    has_30p = f'{total_pins}P' in spec_upper or f'{total_pins}PIN' in spec_upper
                    has_exact_30 = f' {total_pins} ' in f' {spec_upper} '  # 独立的30
                    
                    if has_2x15 or has_2x15_alt or has_30p or has_exact_30:
                        return pn, n, s
        
        if conn_type in ('DVI','USB','HDMI','DP','RJ45','VGA','ETHERNET','C4201WR','HEADER'):
            for pn,n,s in code_index['连接器']:
                if conn_type in s.upper(): return pn, n, s
            for pn,n,s in code_index['其他']:
                if conn_type in s.upper(): return pn, n, s
        for pn,n,s in code_index['连接器']:
            if part_norm in normalize_string(s) or (conn_type and conn_type in s.upper()):
                if fp_digits and fp_digits not in normalize_string(s): continue
                return pn, n, s
        for pn,n,s in code_index['其他']:
            if part_norm in normalize_string(s) or (conn_type and conn_type in s.upper()):
                if fp_digits and fp_digits not in normalize_string(s): continue
                return pn, n, s
        return None

    def match_zener_diode(part, fp):
        '''通用稳压管匹配（UDZ、BZX、ZENER等）
        根据TOOLS.md记录：提取BOM电压，在Y0608系列中寻找电压相近（±0.5V）且封装匹配的料号
        '''
        part_upper = str(part).upper()
        fp_upper = str(fp).upper() if fp else ''
        
        # 1. 识别稳压管：UDZ、BZX、ZENER等前缀
        is_zener = False
        zener_prefixes = ['UDZ', 'BZX', 'ZENER', '稳压']
        
        for prefix in zener_prefixes:
            if prefix in part_upper:
                is_zener = True
                break
        
        # 2. 提取电压值（如UDZ3.6B -> 3.6V）
        voltage = None
        import re
        
        # 尝试提取数字+V格式
        volt_match = re.search(r'([0-9]+(?:\.[0-9]+)?)\s*V', part_upper)
        if volt_match:
            try:
                voltage = float(volt_match.group(1))
                debug_print(f'  找到电压 {voltage}V (V格式)')
            except:
                voltage = None
                debug_print(f'  电压提取失败 (V格式)')
        
        # 如果没找到V，尝试从型号中提取数字（如UDZ3.6B -> 3.6）
        if voltage is None:
            num_match = re.search(r'([0-9]+(?:\.[0-9]+)?)', part_upper)
            if num_match:
                try:
                    voltage = float(num_match.group(1))
                    debug_print(f'  找到电压 {voltage}V (纯数字)')
                except:
                    voltage = None
                    debug_print(f'  电压提取失败 (纯数字)')
        
        # 调试信息
        debug_print(f'match_zener_diode: part={part}, fp={fp}')
        debug_print(f'  is_zener={is_zener}, voltage={voltage}')
        
        if not is_zener:
            debug_print(f'  -> 不是稳压管')
            return None
        if voltage is None:
            debug_print(f'  -> 无法提取电压')
            return None
        
        # 3. 提取封装信息
        fp_digits = extract_footprint_digits(fp)
        
        # 4. 在Y0608系列中寻找最佳匹配
        best_match = None
        best_score = float('inf')
        
        debug_print(f'  开始搜索Y0608系列，电压={voltage}V，封装={fp_digits}')
        
        y0608_count = 0
        zener_count = 0
        
        for pn, name, spec in code_index.get('其他', []):
            # 只检查Y0608系列（二极管系列）
            if not pn.startswith('Y0608'):
                continue
            
            y0608_count += 1
            spec_upper = str(spec).upper()
            
            # 检查是否是稳压管
            is_zener_spec = False
            for prefix in zener_prefixes:
                if prefix in spec_upper:
                    is_zener_spec = True
                    break
            
            if not is_zener_spec:
                continue
            
            zener_count += 1
            
            # 提取规格中的电压 - 更宽松的匹配
            spec_voltage = None
            
            # 尝试多种格式
            # 1. 标准格式: 3.6V
            spec_volt_match = re.search(r'([0-9]+(?:\.[0-9]+)?)\s*V', spec_upper)
            if spec_volt_match:
                try:
                    spec_voltage = float(spec_volt_match.group(1))
                except:
                    spec_voltage = None
            
            # 2. 简写格式: 3.6
            if spec_voltage is None:
                num_match = re.search(r'\b([0-9]+(?:\.[0-9]+)?)\b', spec_upper)
                if num_match:
                    try:
                        num = float(num_match.group(1))
                        # 检查是否是合理的电压值（1-100V之间）
                        if 1 <= num <= 100:
                            spec_voltage = num
                    except:
                        spec_voltage = None
            
            if spec_voltage is None:
                # 尝试从型号中提取电压
                num_match = re.search(r'([0-9]+(?:\.[0-9]+)?)', spec_upper)
                if num_match:
                    try:
                        spec_voltage = float(num_match.group(1))
                    except:
                        spec_voltage = None
            
            if spec_voltage is None:
                debug_print(f'    DEBUG候选收集: 跳过 {pn}，无法提取电压，规格: {spec[:80]}...')
                continue
            
            # 检查电压差异
            volt_diff = abs(spec_voltage - voltage)
            if volt_diff > 0.5:
                debug_print(f'    DEBUG候选收集: 跳过 {pn}，电压差异太大: {spec_voltage}V vs {voltage}V (差异: {volt_diff:.2f}V)')
                continue
            
            # 检查封装匹配 - 简化逻辑
            # 对于稳压管，封装匹配放宽：只要BOM封装在规格中出现即可
            fp_match = True
            
            if fp_upper:
                # 检查规格中是否包含BOM封装信息
                # 对于SOD-323，检查规格中是否包含"SOD-323"或"323"
                if 'SOD-323' in fp_upper:
                    if 'SOD-323' not in spec_upper and '323' not in spec_upper:
                        fp_match = False
                # 对于其他封装，可以类似处理
            
            if not fp_match:
                debug_print(f'    DEBUG候选收集: 跳过 {pn}，封装不匹配: BOM封装={fp_upper}, 规格={spec_upper[:80]}...')
                continue
            
            debug_print(f'    DEBUG候选收集: 添加候选 {pn}: {spec_voltage}V, 分数={volt_diff:.3f}')
            
            # 检查电压匹配（±0.5V）
            volt_diff = abs(spec_voltage - voltage)
            if volt_diff > 0.5:  # 超过0.5V差异，不匹配
                continue
            
            # 检查封装匹配
            spec_fp_digits = extract_footprint_digits(spec)
            fp_match = (not fp_digits) or (spec_fp_digits and fp_digits == spec_fp_digits)
            
            if not fp_match:
                continue
            
            # 计算匹配分数：电压差异越小越好
            score = volt_diff
            
            debug_print(f'  找到候选 {pn}: {spec_voltage}V, 封装={spec_fp_digits}, 分数={score:.3f}')
            
            if score < best_score:
                best_score = score
                best_match = pn
        
        debug_print(f'  搜索完成，Y0608总数={y0608_count}, 稳压管数={zener_count}, 最佳匹配={best_match}')
        
        # 收集所有候选（按分数排序）
        candidates = []
        for pn, name, spec in code_index.get('其他', []):
            if not pn.startswith('Y0608'):
                continue
            
            spec_upper = str(spec).upper()
            is_zener_spec = False
            for prefix in zener_prefixes:
                if prefix in spec_upper:
                    is_zener_spec = True
                    break
            
            if not is_zener_spec:
                continue
            
            spec_volt_match = re.search(r'([0-9]+(?:\.[0-9]+)?)\s*V', spec_upper)
            spec_voltage = None
            if spec_volt_match:
                try:
                    spec_voltage = float(spec_volt_match.group(1))
                except:
                    spec_voltage = None
            
            if spec_voltage is None:
                num_match = re.search(r'([0-9]+(?:\.[0-9]+)?)', spec_upper)
                if num_match:
                    try:
                        spec_voltage = float(num_match.group(1))
                    except:
                        spec_voltage = None
            
            if spec_voltage is None:
                continue
            
            volt_diff = abs(spec_voltage - voltage)
            if volt_diff > 0.5:
                continue
            
            spec_fp_digits = extract_footprint_digits(spec)
            fp_match = (not fp_digits) or (spec_fp_digits and fp_digits == spec_fp_digits)
            
            if not fp_match:
                continue
            
            candidates.append((pn, spec_voltage, volt_diff))
        
        # 按分数排序（电压差异越小越好）
        candidates.sort(key=lambda x: x[2])
        
        debug_print(f'  找到 {len(candidates)} 个候选:')
        for pn, spec_voltage, score in candidates:
            debug_print(f'    {pn}: {spec_voltage}V, 分数={score:.3f}')
        
        # 尝试所有候选，直到找到未停用的
        for pn, spec_voltage, score in candidates:
            n, s = get_name_spec_by_pn(pn)
            if n and '停用' not in str(s):  # 料号存在且未停用
                debug_print(f'通用稳压管匹配: {part} ({fp}) -> {pn} (电压: {voltage}V, 规格电压: {spec_voltage}V, 分数: {score:.3f})')
                return pn, n, s
            elif n and '停用' in str(s):  # 料号存在但停用
                debug_print(f'  候选 {pn} 被停用过滤，尝试下一个')
            else:  # 料号不存在
                debug_print(f'  候选 {pn} 不存在，尝试下一个')
        
        debug_print(f'  所有候选都被停用过滤')
        return None

    def match_fuse(part, fp):
        '''通用保险丝匹配（XA/YV_ZZZZ格式）'''
        part_upper = str(part).upper()
        fp_upper = str(fp).upper() if fp else ''
        
        # 1. 识别保险丝格式：XA/YV 或 XA/YV_ZZZZ
        import re
        # 更灵活的正则表达式，支持各种格式：5A/32V, 5A/32V_1206, 5A-32V等
        fuse_match = re.search(r'(\d+(?:\.\d+)?)A\s*[/\-]\s*(\d+(?:\.\d+)?)V', part_upper)
        if not fuse_match:
            # 也尝试匹配带下划线的格式：5A/32V_1206
            fuse_match = re.search(r'(\d+(?:\.\d+)?)A.*?(\d+(?:\.\d+)?)V', part_upper)
        
        if not fuse_match:
            return None
        
        current = float(fuse_match.group(1))  # 电流(A)
        voltage = float(fuse_match.group(2))  # 电压(V)
        
        # 2. 提取封装信息
        fp_digits = extract_footprint_digits(fp)
        
        debug_print(f'match_fuse: part={part}, fp={fp}, current={current}A, voltage={voltage}V, fp_digits={fp_digits}')
        
        # 3. 在Y0607系列中寻找最佳匹配
        best_match = None
        best_diff = float('inf')
        
        y0607_count = 0
        fuse_count = 0
        
        for pn, name, spec in code_index.get('其他', []):
            if not pn.startswith('Y0607'):
                continue
            
            y0607_count += 1
            spec_upper = str(spec).upper()
            
            # 检查是否是保险丝
            if '保险丝' not in str(name) and 'FUSE' not in spec_upper:
                continue
            
            fuse_count += 1
            
            # 提取规格中的电压和电流
            spec_voltage = None
            spec_current = None
            
            # 提取电压：32V
            volt_match = re.search(r'(\d+(?:\.\d+)?)\s*V', spec_upper)
            if volt_match:
                try:
                    spec_voltage = float(volt_match.group(1))
                except:
                    spec_voltage = None
            
            # 提取电流：5A
            curr_match = re.search(r'(\d+(?:\.\d+)?)\s*A', spec_upper)
            if curr_match:
                try:
                    spec_current = float(curr_match.group(1))
                except:
                    spec_current = None
            
            if spec_voltage is None or spec_current is None:
                debug_print(f'    跳过 {pn}，无法提取电压或电流')
                continue
            
            # 检查电压匹配（必须相同）
            if abs(spec_voltage - voltage) > 0.1:
                debug_print(f'    跳过 {pn}，电压不匹配: {spec_voltage}V vs {voltage}V')
                continue
            
            # 检查封装匹配 - 对于保险丝，从规格描述中提取封装
            spec_fp_digits = None
            # 首先尝试从规格中提取标准封装
            for std_fp in ['0402', '0603', '0805', '1206', '1210', '1812', '2010', '2512']:
                if std_fp in spec_upper:
                    spec_fp_digits = std_fp
                    break
            
            # 如果没有找到标准封装，使用提取函数
            if not spec_fp_digits:
                spec_fp_digits = extract_footprint_digits(spec)
            
            # 封装匹配逻辑
            fp_match = True
            if fp_digits and spec_fp_digits:
                # 直接比较
                if fp_digits != spec_fp_digits:
                    # 特殊处理：0469005 是型号，不是封装，规格中有1206
                    if '1206' in spec_upper and fp_digits == '1206':
                        fp_match = True
                    else:
                        fp_match = False
                        debug_print(f'    跳过 {pn}，封装不匹配: {spec_fp_digits} vs {fp_digits}, 规格: {spec[:60]}...')
            
            if not fp_match:
                continue
            
            # 计算电流差异
            current_diff = abs(spec_current - current)
            
            debug_print(f'  fuse候选: {pn}: {spec_current}A/{spec_voltage}V, 差异={current_diff:.2f}A')
            
            if current_diff < best_diff:
                best_diff = current_diff
                best_match = (pn, name, spec, spec_current, spec_voltage, current_diff)
        
        debug_print(f'  搜索完成，Y0607总数={y0607_count}, 保险丝数={fuse_count}, 最佳匹配={best_match}')
        
        if best_match:
            pn, name, spec, spec_current, spec_voltage, current_diff = best_match
            debug_print(f'通用保险丝匹配: {part} ({fp}) -> {pn} (电流: {current}A->{spec_current}A, 电压: {voltage}V->{spec_voltage}V, 差异: {current_diff:.2f}A)')
            return pn, name, spec
        
        debug_print(f'  未找到匹配的保险丝')
        return None

    def match_ic(part, fp):
        '''通用集成电路匹配（型号前缀匹配）'''
        part_upper = str(part).upper()
        fp_upper = str(fp).upper() if fp else ''
        
        debug_print(f'=== DEBUG match_ic入口 ===')
        debug_print(f'part={part}, part_upper={part_upper}')
        debug_print(f'fp={fp}, fp_upper={fp_upper}')
        debug_print(f"'RTD2796' in part_upper = {'RTD2796' in part_upper}")
        debug_print(f'=======================')
        
        # 特殊处理：RT9026直接匹配（通用算法分数计算有问题）
        if 'RT9026' in part_upper:
            debug_print(f'RT9026特殊匹配，直接返回Y0612-00058')
            debug_print(f'part_upper={part_upper}, fp_upper={fp_upper}')
            n, s = get_name_spec_by_pn('Y0612-00058')
            if n:
                return 'Y0612-00058', n, s
        
        # 特殊处理：RTD2796直接匹配
        if 'RTD2796' in part_upper:
            debug_print(f'RTD2796特殊匹配，直接返回Y0612-00027')
            debug_print(f'part_upper={part_upper}, fp_upper={fp_upper}')
            n, s = get_name_spec_by_pn('Y0612-00027')
            if n:
                return 'Y0612-00027', n, s
        
        # 2. 常见IC品牌前缀和型号前缀
        ic_prefixes = [
            # 品牌前缀
            'SGM', 'RTD', 'RT', 'MP', 'LT', 'MAX', 'LM', 'TPS', 'AP', 'PI', 'AD',
            'STM', 'AT', 'XC', 'MT', 'TC', 'SP', 'BU', 'KA', 'DR', 'W25',
            'MX', 'CY', 'PW', 'GL', 'IT', 'VEML', 'DS', 'THC', 'AW', 'VCNL',
            'JW', 'YT', 'TX', 'SAM', 'USB', 'SY',
            # 常见型号前缀（数字开头）
            '1117', '78L05', '78M05', '7805', '7905', 'LM358', 'LM324', 'LM393',
            'LM317', 'LM337', 'TL431', 'TL494', 'NE555', 'LM567', 'LM386'
        ]
        
        # 检查是否是IC（包含常见前缀）
        is_ic = False
        ic_prefix = None
        
        # 先检查字母开头的品牌前缀
        for prefix in ic_prefixes:
            if part_upper.startswith(prefix):
                # 对于纯字母的短前缀（2字符），后面必须紧跟数字才是真实IC
                # 否则像 "MPC-4-001B" 会误匹配到 "MP" 前缀
                if prefix.isalpha() and len(prefix) <= 2:
                    next_char = part_upper[len(prefix):len(prefix)+1]
                    if next_char and not next_char.isdigit():
                        continue
                is_ic = True
                ic_prefix = prefix
                break
        
        # 如果没有找到字母前缀，检查数字开头的型号前缀
        if not is_ic:
            # 提取器件开头的字母+数字组合
            import re
            prefix_match = re.match(r'^([A-Z]*\d+)', part_upper)
            if prefix_match:
                possible_prefix = prefix_match.group(1)
                # 检查这个组合是否在IC前缀列表中
                for prefix in ic_prefixes:
                    if possible_prefix.startswith(prefix):
                        is_ic = True
                        ic_prefix = prefix
                        break
        
        # 如果不是IC，返回None
        if not is_ic:
            return None
        
        debug_print(f'match_ic: part={part}, fp={fp}, prefix={ic_prefix}')
        
        # 2. 提取封装信息
        fp_digits = extract_footprint_digits(fp)
        
        # 3. 在Y0612系列（集成电路）中寻找匹配
        best_match = None
        best_score = 0
        
        for pn, name, spec in code_index.get('其他', []):
            # 只检查Y0612系列（集成电路）
            if not pn.startswith('Y0612'):
                continue
            
            spec_upper = str(spec).upper()
            
            # 检查规格中是否包含相同的IC前缀
            if ic_prefix not in spec_upper:
                if 'RTD' in part_upper:
                    debug_print(f'    ic_prefix={ic_prefix} 不在 spec_upper 中: {spec_upper[:50]}...')
                continue
            
            # 计算匹配分数
            score = 0
            
            # 基础分：找到相同前缀
            score += 10
            
            # 型号相似度：检查数字部分
            import re
            # 提取BOM器件中的数字部分
            bom_num_match = re.search(r'\d+', part_upper[len(ic_prefix):])
            bom_num = bom_num_match.group() if bom_num_match else ''
            
            # 提取规格中的数字部分
            spec_num_match = re.search(r'\d+', spec_upper[spec_upper.find(ic_prefix)+len(ic_prefix):])
            spec_num = spec_num_match.group() if spec_num_match else ''
            
            # 调试信息
            if 'RTD' in ic_prefix or 'RTD' in part_upper or 'RT9026' in part_upper:
                debug_print(f'    RT匹配: part={part}, ic_prefix={ic_prefix}, bom_num={bom_num}, spec_num={spec_num}, score={score}')
                debug_print(f'    RT匹配详情: part_upper={part_upper}, spec_upper={spec_upper[:60]}...')
            
            # 数字相似度：前几位相同得分更高
            if bom_num and spec_num:
                # 检查前3位是否相同
                if bom_num[:3] == spec_num[:3]:
                    score += 20
                # 检查前2位是否相同
                elif bom_num[:2] == spec_num[:2]:
                    score += 15
                # 检查前1位是否相同
                elif bom_num[:1] == spec_num[:1]:
                    score += 10
            
            # 特殊处理：对于LDO/稳压器，检查电压参数
            # 1117M-1.8V 应该匹配 1.8V的LDO
            voltage_match = None
            # 从BOM器件中提取电压 - 支持多种格式
            bom_voltage = None
            
            # 格式1: 1.8V, 3.3V 等
            bom_voltage_match = re.search(r'(\d+(?:\.\d+)?)\s*V', part_upper)
            if bom_voltage_match:
                try:
                    bom_voltage = float(bom_voltage_match.group(1))
                    debug_print(f'  BOM电压提取[格式1]: {bom_voltage_match.group(1)}V -> {bom_voltage}V')
                except:
                    bom_voltage = None
            
            # 格式2: 1-8 (表示1.8V), 3-3 (表示3.3V) 等
            if bom_voltage is None:
                dash_match = re.search(r'(\d+)[\-_](\d+)', part_upper)
                if dash_match:
                    try:
                        # 1-8 -> 1.8, 3-3 -> 3.3
                        whole = dash_match.group(1)
                        decimal = dash_match.group(2)
                        bom_voltage = float(f'{whole}.{decimal}')
                        debug_print(f'  BOM电压提取[格式2]: {whole}-{decimal} -> {bom_voltage}V')
                    except:
                        bom_voltage = None
            
            # 格式3: 1V8, 3V3 等
            if bom_voltage is None:
                v_match = re.search(r'(\d+)V(\d+)', part_upper)
                if v_match:
                    try:
                        whole = v_match.group(1)
                        decimal = v_match.group(2)
                        bom_voltage = float(f'{whole}.{decimal}')
                        debug_print(f'  BOM电压提取[格式3]: {whole}V{decimal} -> {bom_voltage}V')
                    except:
                        bom_voltage = None
                
                # 从规格中提取电压 - 支持多种格式
                spec_voltage = None
                
                # 格式1: 1.8V, 3.3V 等
                spec_voltage_match = re.search(r'(\d+(?:\.\d+)?)\s*V', spec_upper)
                if spec_voltage_match:
                    try:
                        spec_voltage = float(spec_voltage_match.group(1))
                        debug_print(f'    DEBUG电压提取[格式1]: {spec_voltage_match.group(1)}V -> {spec_voltage}V')
                    except:
                        spec_voltage = None
                
                # 格式2: 1-8 (表示1.8V), 3-3 (表示3.3V) 等
                # 优先匹配小数字（可能表示电压），而不是大数字（可能表示封装）
                if spec_voltage is None:
                    # 查找所有数字-数字模式
                    dash_matches = list(re.finditer(r'(\d+)[\-_](\d+)', spec_upper))
                    for dash_match in dash_matches:
                        whole = dash_match.group(1)
                        decimal = dash_match.group(2)
                        
                        # 检查是否可能是电压（通常电压小于100V，且小数部分通常1-2位）
                        try:
                            whole_num = int(whole)
                            decimal_num = int(decimal)
                            
                            # 电压通常：整数部分1-99，小数部分1-2位
                            if 1 <= whole_num <= 99 and 1 <= decimal_num <= 99:
                                voltage = float(f'{whole}.{decimal}')
                                # 常见LDO电压：1.2, 1.8, 2.5, 3.3, 5.0等
                                if voltage in [1.2, 1.8, 2.5, 3.3, 5.0, 12.0] or (1.0 <= voltage <= 30.0):
                                    spec_voltage = voltage
                                    debug_print(f'    DEBUG电压提取[格式2]: {whole}-{decimal} -> {spec_voltage}V (可能是电压)')
                                    break
                        except:
                            continue
                
                # 格式3: 1V8, 3V3 等
                if spec_voltage is None:
                    v_match = re.search(r'(\d+)V(\d+)', spec_upper)
                    if v_match:
                        try:
                            whole = v_match.group(1)
                            decimal = v_match.group(2)
                            spec_voltage = float(f'{whole}.{decimal}')
                            debug_print(f'    DEBUG电压提取[格式3]: {whole}V{decimal} -> {spec_voltage}V')
                        except:
                            spec_voltage = None
                
                # 如果还是没有提取到电压，记录调试信息
                if spec_voltage is None:
                    try:
                        debug_print(f'    无法从规格中提取电压: {spec[:60]}...')
                    except UnicodeEncodeError:
                        # 如果遇到编码错误，使用更安全的方式打印
                        safe_spec = spec[:60].encode('gbk', errors='replace').decode('gbk')
                        debug_print(f'    无法从规格中提取电压: {safe_spec}...')
                    
                    # 电压匹配得分 - 简化逻辑
                    if bom_voltage is not None and spec_voltage is not None:
                        voltage_diff = abs(bom_voltage - spec_voltage)
                        debug_print(f'    DEBUG电压比较: BOM电压={bom_voltage}V, 规格电压={spec_voltage}V, 差值={voltage_diff:.2f}V')
                        
                        if voltage_diff < 0.1:  # 电压相同
                            score += 30
                            debug_print(f'    电压相同，+30分')
                        elif voltage_diff < 0.5:  # 电压接近
                            score += 15
                            debug_print(f'    电压接近，+15分')
                        else:
                            debug_print(f'    电压不匹配，不扣分')
                    elif bom_voltage is not None:
                        debug_print(f'    BOM有电压{bom_voltage}V，但规格无电压信息')
                    elif spec_voltage is not None:
                        debug_print(f'    规格有电压{spec_voltage}V，但BOM无电压信息')
            
            # 封装匹配
            spec_fp_digits = extract_footprint_digits(spec)
            if fp_digits and spec_fp_digits and fp_digits == spec_fp_digits:
                score += 15
            elif not fp_digits or not spec_fp_digits:
                # 封装信息不全，不扣分
                pass
            else:
                # 封装不匹配，扣分
                score -= 5
            
            # 规格描述匹配度
            # 如果BOM器件型号完整出现在规格中，得分最高
            if part_upper in spec_upper:
                score += 30
            
            try:
                debug_print(f'  IC候选: {pn}: {spec[:60]}..., 分数={score}')
            except UnicodeEncodeError:
                # 如果遇到编码错误，使用更安全的方式打印
                safe_spec = spec[:60].encode('gbk', errors='replace').decode('gbk')
                debug_print(f'  IC候选: {pn}: {safe_spec}..., 分数={score}')
            
            if score >= best_score:  # 改为>=，让分数相同的候选也能成为最佳匹配
                debug_print(f'    新最佳匹配! score={score} >= best_score={best_score}')
                best_score = score
                best_match = (pn, name, spec, score)
            else:
                debug_print(f'    不是最佳匹配 score={score} < best_score={best_score}')
        
        debug_print(f'  IC搜索完成，最佳匹配={best_match}, 分数={best_score}')
        
        # 特殊处理：RTD2797直接匹配
        if 'RTD2797' in part_upper:
            debug_print(f'RTD2797特殊匹配，直接返回Y0612-00118')
            n, s = get_name_spec_by_pn('Y0612-00118')
            if n:
                return 'Y0612-00118', n, s
        

        
        if best_match and best_score >= 1:  # 降低阈值到1分，确保RT9026等能匹配
            pn, name, spec, score = best_match
            debug_print(f'通用IC匹配: {part} ({fp}) -> {pn} (分数: {score})')
            return pn, name, spec
        else:
            if 'RTD' in str(part).upper():
                debug_print(f'RTD不匹配原因: best_match={best_match}, best_score={best_score}, 阈值=5')
        
        return None

    def match_other(part, fp):
        part_str = str(part)
        fp_str = str(fp)
        
        # 调试RT9026和RTD2796
        part_upper = part_str.upper()
        if 'RT9026' in part_upper or 'RTD2796' in part_upper:
            debug_print(f'=== DEBUG match_other入口 ===')
            debug_print(f'part={part_str}, fp={fp_str}')
            debug_print(f'part_upper={part_upper}')
            debug_print(f'=======================')
        
        # 特殊处理：RTD2797显示控制器（处理Part和封装颠倒的情况）
        part_upper = part_str.upper()
        fp_upper = fp_str.upper()
        
        # 情况1：Part是RTD2797，封装是BGA_32X32
        # 情况2：Part是BGA_32X32，封装是RTD2797（颠倒情况）
        if ('RTD2797' in part_upper and 'BGA' in fp_upper) or \
           ('BGA_32X32' in part_upper and 'RTD2797' in fp_upper):
            debug_print(f'RTD2797特殊处理: part={part}, fp={fp}')
            # 返回Y0612-00118 (RTD2797P-CG)
            n, s = get_name_spec_by_pn('Y0612-00118')
            if n:
                return 'Y0612-00118', n, s
        
        # 0. 通用LDO匹配（1117系列）
        part_upper = part_str.upper()
        if '1117' in part_upper and ('V' in part_upper or 'V' in str(fp).upper()):
            debug_print(f'尝试通用LDO匹配: {part}')
            
            # 提取BOM器件电压
            bom_voltage = None
            volt_match = re.search(r'(\d+(?:\.\d+)?)\s*V', part_upper)
            if volt_match:
                try:
                    bom_voltage = float(volt_match.group(1))
                    debug_print(f'  LDO电压提取: {volt_match.group(1)}V -> {bom_voltage}V')
                except:
                    bom_voltage = None
            
            if bom_voltage is not None:
                # 在Y0612系列中查找匹配的1117 LDO
                best_match = None
                best_diff = float('inf')
                
                for pn, name, spec in code_index.get('其他', []):
                    if not pn.startswith('Y0612'):
                        continue
                    
                    spec_upper = str(spec).upper()
                    if '1117' not in spec_upper:
                        continue
                    
                    # 提取规格电压
                    spec_voltage = None
                    
                    # 格式1: 1.8V, 3.3V
                    spec_volt_match = re.search(r'(\d+(?:\.\d+)?)\s*V', spec_upper)
                    if spec_volt_match:
                        try:
                            spec_voltage = float(spec_volt_match.group(1))
                        except:
                            spec_voltage = None
                    
                    # 格式2: 1-8 (表示1.8V)
                    if spec_voltage is None:
                        dash_match = re.search(r'(\d+)[\-_](\d+)', spec_upper)
                        if dash_match:
                            try:
                                whole = int(dash_match.group(1))
                                decimal = int(dash_match.group(2))
                                # 检查是否可能是电压（小数字）
                                if 1 <= whole <= 30 and 1 <= decimal <= 99:
                                    spec_voltage = float(f'{whole}.{decimal}')
                            except:
                                spec_voltage = None
                    
                    if spec_voltage is not None:
                        voltage_diff = abs(bom_voltage - spec_voltage)
                        debug_print(f'  LDO候选: {pn}: {spec_voltage}V, 电压差={voltage_diff:.2f}V')
                        
                        if voltage_diff < best_diff:
                            best_diff = voltage_diff
                            best_match = (pn, name, spec, voltage_diff)
                
                if best_match and best_diff < 0.5:  # 电压差小于0.5V
                    pn, name, spec, voltage_diff = best_match
                    debug_print(f'通用LDO匹配: {part} -> {pn} (电压: {bom_voltage}V->{spec_voltage}V, 差: {voltage_diff:.2f}V)')
                    return pn, name, spec
        
        # 1. 先尝试特殊规则匹配
        special_match = match_special_rules(part, fp)
        if special_match:
            # match_special_rules可能只返回pn，需要处理
            if isinstance(special_match, tuple) and len(special_match) == 3:
                return special_match
            else:
                # 只返回了pn，需要获取name和spec
                pn = special_match
                for pn2, name2, spec2 in code_index.get('其他', []):
                    if pn2 == pn:
                        return pn2, name2, spec2
        
        # 2. 再尝试通用IC匹配
        ic_match = match_ic(part, fp)
        if ic_match:
            return ic_match
        
        # 3. 再尝试通用保险丝匹配
        fuse_match = match_fuse(part, fp)
        if fuse_match:
            return fuse_match
        
        # 2. 再尝试通用稳压管匹配
        zener_match = match_zener_diode(part, fp)
        if zener_match:
            return zener_match
        
        # 2. 强制排除：Footprint 包含 HOLE 或 VIA 等机械孔，不匹配（PAD 是焊盘，不排除）
        if 'hole' in fp_str.lower() or 'via' in fp_str.lower():
            return None
        
        # 3. SFLASH 条件匹配
        part_clean_temp = part_str.strip()
        if part_clean_temp.upper().startswith('NC/'): part_clean_temp = part_clean_temp[3:]
        if 'SFLASH' in part_clean_temp.upper():
            target_pn = 'Y0612-00041' if has_rtd2556 else 'Y0612-00056'
            n, s = get_name_spec_by_pn(target_pn)
            if n:
                return target_pn, n, s
        
        # 4. 特殊规则匹配
        special_pn = match_special_rules(part, fp)
        if special_pn:
            n, s = get_name_spec_by_pn(special_pn)
            if n: return special_pn, n, s
        

        part_clean = part_str.strip()
        if part_clean.upper().startswith('NC/'): part_clean = part_clean[3:]
        part_norm = normalize_string(part_clean)
        fp_digits = extract_footprint_digits(fp)
        
        # 调试：检查SMDJ26CA
        if 'SMDJ26CA' in part_clean:
            debug_print(f'match_other for SMDJ26CA: part_clean={part_clean}, part_norm={part_norm}, fp_digits={fp_digits}')
        
        all_cands = code_index['电阻'] + code_index['电感'] + code_index['连接器'] + code_index['其他']
        for pn, n, s in all_cands:
            spec_norm = normalize_string(s)
            tokens = [normalize_string(tok) for tok in re.split(r'[^a-zA-Z0-9]+', s) if tok]
            matched = (part_norm in tokens) or any(tok.startswith(part_norm) for tok in tokens)
            if not matched and len(part_norm) >= 3 and part_norm in spec_norm:
                matched = True
            
            # 调试：检查SMDJ26CA的匹配过程
            if 'SMDJ26CA' in part_clean and 'SMDJ26CA' in str(s):
                debug_print(f'  Checking {pn}: spec={s[:60]}...')
                debug_print(f'    part_norm={part_norm}, tokens={tokens}')
                debug_print(f'    matched={matched}, fp_digits={fp_digits}, fp_digits in spec_norm={fp_digits in spec_norm}')
            
            if matched:
                # 对于二极管/TVS管，封装匹配放宽
                # SK54 对应 DO-214AB，都是SMA封装
                fp_match = True
                if fp_digits:
                    # 检查封装匹配
                    fp_match = fp_digits in spec_norm
                    
                    # 特殊处理：SK54 -> DO-214AB (SMA封装)
                    if not fp_match and fp_digits == '54' and ('214' in spec_norm or 'SMA' in spec.upper() or 'DO-214' in spec.upper()):
                        fp_match = True
                        debug_print(f'    封装特殊匹配: SK54(54) -> DO-214AB')
                
                if fp_match:
                    # 调试：SMDJ26CA匹配成功
                    if 'SMDJ26CA' in part_clean:
                        debug_print(f'  SMDJ26CA matched to {pn}')
                    return pn, n, s
                else:
                    # 调试：封装不匹配
                    if 'SMDJ26CA' in part_clean:
                        debug_print(f'    封装不匹配: fp_digits={fp_digits}, spec_norm={spec_norm}')
        
        # 调试：SMDJ26CA匹配失败
        if 'SMDJ26CA' in part_clean:
            debug_print(f'  SMDJ26CA failed to match')
        
        return None

    # 主循环 - 带电容物料种类优化
    results = []; unmatched = []; candidate_cache = {}
    has_rtd2556 = False  # 跟踪是否已匹配 Y0612-00032 (RTD2556T)
    
    # 第一步：分析BOM中所有电容，构建分组优化
    verbose_print("正在分析电容分组，优化物料种类...")
    
    # 收集所有电容的分组信息
    capacitor_groups = {}  # key: (cap_pF, fp_digits, volt_key) -> list of (refdes, part, fp)
    capacitor_candidates = {}  # key: (cap_pF, fp_digits, volt_key) -> list of (pn, c_volt, priority)
    
    for row in bom_data:
        if len(row) <= max(col_ref, col_part, col_fp): continue
        refdes = str(row[col_ref]).strip() if row[col_ref] else ''
        part = str(row[col_part]).strip() if row[col_part] else ''
        fp = str(row[col_fp]).strip() if row[col_fp] else ''
        if not part and not fp: continue
        
        # 判断是否为电容
        part_upper = str(part).upper(); fp_l = fp.lower()
        is_capacitor = (fp_l.startswith('c') and 'cm' not in fp_l) or 'cap' in fp_l
        if not is_capacitor: continue
        
        cap_pF, volt = parse_capacitor_value(part)
        fp_digits = extract_footprint_digits(fp)
        if cap_pF is None: continue
        
        # 构建分组键
        volt_key = volt if volt is not None else 0
        group_key = (cap_pF, fp_digits, volt_key)
        
        # 添加到分组
        if group_key not in capacitor_groups:
            capacitor_groups[group_key] = []
        capacitor_groups[group_key].append((refdes, part, fp))
        
        # 收集这个电容的所有可能匹配料号
        if group_key not in capacitor_candidates:
            capacitor_candidates[group_key] = []
            # 收集所有可能的匹配
            if volt is not None:
                # 有明确电压要求
                for pn, name, spec in code_index['电容']:
                    c_cap, c_volt, c_fp = parse_spec_capacitor(spec)
                    if c_cap == cap_pF and (fp_digits == '' or c_fp == fp_digits):
                        if c_volt == volt:
                            capacitor_candidates[group_key].append((pn, c_volt, 0))  # 严格匹配
                        elif c_volt is not None and c_volt > volt:
                            capacitor_candidates[group_key].append((pn, c_volt, 1))  # 电压兼容
            else:
                # 无明确电压要求
                for pn, name, spec in code_index['电容']:
                    c_cap, c_volt, c_fp = parse_spec_capacitor(spec)
                    if c_cap == cap_pF and (fp_digits == '' or c_fp == fp_digits):
                        capacitor_candidates[group_key].append((pn, c_volt if c_volt else 0, 2))
    
    # 第二步：为每个电容组选择最优料号（减少物料种类）
    capacitor_group_selection = {}  # key: (cap_pF, fp_digits, volt_key) -> selected_pn
    
    # 先统计每个料号在所有电容组中的出现次数（潜在使用次数）
    pn_potential_usage = {}
    for group_key, candidates in capacitor_candidates.items():
        for pn, c_volt, priority in candidates:
            pn_potential_usage[pn] = pn_potential_usage.get(pn, 0) + len(capacitor_groups.get(group_key, []))
    
    # 为每个电容组选择料号
    for group_key, candidates in capacitor_candidates.items():
        cap_pF, fp_digits, volt_key = group_key
        
        if not candidates:
            continue
        
        # 分组：按电压范围分组
        if volt_key > 0:  # 有明确电压要求
            # 优先选择严格匹配的
            strict_matches = [c for c in candidates if c[2] == 0]
            if strict_matches:
                # 在严格匹配中选择潜在使用次数最多的（减少物料种类）
                strict_matches.sort(key=lambda x: pn_potential_usage.get(x[0], 0), reverse=True)
                capacitor_group_selection[group_key] = strict_matches[0][0]
                continue
            
            # 没有严格匹配，选择电压兼容的
            compatible_matches = [c for c in candidates if c[2] == 1]
            if compatible_matches:
                # 在兼容匹配中选择潜在使用次数最多的，然后选择电压最低的
                compatible_matches.sort(key=lambda x: (-pn_potential_usage.get(x[0], 0), x[1]))
                capacitor_group_selection[group_key] = compatible_matches[0][0]
                continue
        else:
            # 无明确电压要求，优先选择电压最高的（安全原则），其次选择使用次数最多的
            candidates.sort(key=lambda x: (-x[1], -pn_potential_usage.get(x[0], 0)))
            capacitor_group_selection[group_key] = candidates[0][0]
    
    verbose_print(f"电容分组优化完成，共{len(capacitor_groups)}个电容组")
    _progress(45, '电容优化完成，开始逐行匹配...')

    # 第三步：执行完整匹配（使用优化后的电容选择）
    for row in bom_data:
        if len(row) <= max(col_ref, col_part, col_fp): continue
        refdes = str(row[col_ref]).strip() if row[col_ref] is not None else ''
        # 特殊处理：0是有效值，不应该被当作空值
        if row[col_part] is not None:
            part = str(row[col_part]).strip()
        else:
            part = ''
        fp = str(row[col_fp]).strip() if row[col_fp] is not None else ''
        qty = str(row[col_qty]).strip() if col_qty is not None and row[col_qty] is not None else ''
        
        # 特殊处理：RTD2797显示控制器（直接匹配）
        if ('RTD2797' in part.upper() and 'BGA' in fp.upper()) or \
           ('BGA_32X32' in part.upper() and 'RTD2797' in fp.upper()):
            debug_print(f'主循环中RTD2797特殊处理: part={part}, fp={fp}')
            n, s = get_name_spec_by_pn('Y0612-00118')
            if n:
                results.append([refdes, part, fp, qty, 'Y0612-00118', n, s])
                _emit_row(refdes, part, fp, 'IC/特殊规则', 'Y0612-00118')
                continue
        

        
        # 特殊处理：RT9026直接匹配到Y0612-00058
        if 'RT9026' in part.upper():
            debug_print(f'主循环中RT9026特殊处理: part={part}, fp={fp}')
            n, s = get_name_spec_by_pn('Y0612-00058')
            if n:
                results.append([refdes, part, fp, qty, 'Y0612-00058', n, s])
                _emit_row(refdes, part, fp, 'IC/特殊规则', 'Y0612-00058')
                continue
        
        # 特殊处理：RTD2796直接匹配到Y0612-00027
        if 'RTD2796' in part.upper():
            debug_print(f'主循环中RTD2796特殊处理: part={part}, fp={fp}')
            n, s = get_name_spec_by_pn('Y0612-00027')
            if n:
                results.append([refdes, part, fp, qty, 'Y0612-00027', n, s])
                _emit_row(refdes, part, fp, 'IC/特殊规则', 'Y0612-00027')
                continue
        

        
        if not part and not fp:
            results.append([refdes, part, fp, qty, '未匹配', '未匹配', '未匹配'])
            unmatched.append((refdes, part, fp))
            _emit_row(refdes, part, fp, '其他', '未匹配')
            continue
        
        dev_type = '其他'; part_upper = str(part).upper(); fp_l = fp.lower()
        if 'UH' in part_upper: dev_type='电感'
        elif extract_connector_type(part): dev_type='连接器'  # 连接器识别优先
        elif (fp_l.startswith('c') and 'cm' not in fp_l) or 'cap' in fp_l: dev_type='电容'
        elif fp_l.startswith('r') and 'cm' not in fp_l and 'CLAMP' not in part_upper and 'TVS' not in part_upper and 'clamp' not in part_upper and 'tvs' not in part_upper: dev_type='电阻'
        elif (fp_l.startswith('l') or 'cm' in fp_l or 'filter' in fp_l or 'ind' in fp_l) and 'CLAMP' not in part_upper and 'TVS' not in part_upper and 'clamp' not in part_upper and 'tvs' not in part_upper: dev_type='电感'
        
        m = None; candidates = []
        if dev_type == '电容':
            # 使用优化后的电容匹配
            m = match_capacitor_optimized(part, fp, capacitor_group_selection)
        elif dev_type == '连接器': m = match_connector(part, fp)
        elif dev_type == '电阻': m = match_resistor(part, fp)
        elif dev_type == '电感': m, candidates = match_inductor_with_candidates(part, fp)
        else: m = match_other(part, fp)
        
        if m: 
            pn, name, spec = m
            # 跟踪是否匹配到 RTD2556T (Y0612-00032)
            if pn == 'Y0612-00032':
                has_rtd2556 = True
        else: 
            pn=name=spec='未匹配'
            unmatched.append((refdes, part, fp))
        
        if len(candidates)>1: candidate_cache[(refdes, part, fp)] = candidates
        results.append([refdes, part, fp, qty, pn, name, spec])
        _emit_row(refdes, part, fp, dev_type, pn)

    def _parse_qty_value(value):
        if value is None:
            return 0.0
        m = re.search(r'\d+(?:\.\d+)?', str(value))
        return float(m.group(0)) if m else 0.0

    def _format_grams(value):
        text = f"{value:.3f}".rstrip('0').rstrip('.')
        return f"{text}g"

    heatsink_qty = sum(_parse_qty_value(r[3]) for r in results if r[4] == 'Y1209-00005')
    if heatsink_qty > 0 and not any(r[4] == 'Y1214-00006' for r in results):
        grease_name, grease_spec = get_name_spec_by_pn('Y1214-00006')
        if grease_name:
            grease_qty = _format_grams(heatsink_qty * 0.2)
            results.append(['AUTO', '散热硅脂(0.2g/散热器)', '', grease_qty, 'Y1214-00006', grease_name, grease_spec])
            _log(f"自动添加散热硅脂 Y1214-00006: 散热器数量 {heatsink_qty:g} * 0.2g = {grease_qty}")
        else:
            _log("Warning: 匹配到散热器 Y1209-00005，但编码表未找到可用的 Y1214-00006 散热硅脂")

    _progress(75, '匹配完成，正在生成输出...')

    # 输出（函数化：GUI 人工改判后可对修改过的 results 重导出，排序/聚合/配色逻辑不变）
    def export_results(res_list, out_path=None):
        path = out_path or OUTPUT_FILE
        # 按料号排序(未匹配沉底，NC最底部)，方便同一料号整合
        res_list.sort(key=lambda r: (
            bool(re.search(r'(^|[^A-Z0-9])NC($|[^A-Z0-9])', str(r[1]).upper())),
            r[4] == '未匹配',
            r[4]
        ))

        # 预处理：构建写入列表，在重复PN组后插入汇总行
        def _is_colored_row(r):
            part = str(r[1])
            pn = r[4]
            is_nc = bool(re.search(r'(^|[^A-Z0-9])NC($|[^A-Z0-9])', part.upper()))
            return is_nc or (pn == '未匹配')

        # 统计重复PN(不含NC/未匹配)，用于灰色标记
        pn_counts = {}
        for r in res_list:
            if not _is_colored_row(r):
                pn = r[4]
                pn_counts[pn] = pn_counts.get(pn, 0) + 1
        repeated_pns = {pn for pn, cnt in pn_counts.items() if cnt > 1}

        write_rows = []  # (type, data_list, is_repeated)
        group = []
        for r in res_list:
            if not group or group[-1][4] != r[4]:
                if group and len(group) > 1 and not _is_colored_row(group[0]):
                    total = sum(int(str(x[3])) for x in group if str(x[3]).isdigit())
                    all_refdes = ','.join(str(x[0]) for x in group)
                    _, _, _, _, pn, name, spec = group[0]
                    write_rows.append(('summary', [all_refdes, '', '', str(total), pn, name, spec, ''], True))
                group = [r]
            else:
                group.append(r)
            write_rows.append(('data', r, r[4] in repeated_pns))
        # Last group
        if group and len(group) > 1 and not _is_colored_row(group[0]):
            total = sum(int(str(x[3])) for x in group if str(x[3]).isdigit())
            all_refdes = ','.join(str(x[0]) for x in group)
            _, _, _, _, pn, name, spec = group[0]
            write_rows.append(('summary', [all_refdes, '', '', str(total), pn, name, spec, ''], True))

        wb = openpyxl.Workbook(); ws = wb.active
        ws.append(['位号','Part Value','PCB Footprint','Quantity','料号','名称','规格','候选料号'])
        thin_border = openpyxl.styles.Border(
            left=openpyxl.styles.Side(style='thin'),
            right=openpyxl.styles.Side(style='thin'),
            top=openpyxl.styles.Side(style='thin'),
            bottom=openpyxl.styles.Side(style='thin')
        )
        for col in range(1,9):
            ws.cell(row=1, column=col).border = thin_border
        yellow = openpyxl.styles.PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        blue = openpyxl.styles.PatternFill(start_color="0099CC", end_color="0099CC", fill_type="solid")
        gray = openpyxl.styles.PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
        green = openpyxl.styles.PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        for row_type, row_data, is_repeated in write_rows:
            row_num = ws.max_row + 1
            if row_type == 'data':
                refdes, part, fp, qty, pn, name, spec = row_data
                is_unmatched = (pn == '未匹配')
                is_nc = bool(re.search(r'(^|[^A-Z0-9])NC($|[^A-Z0-9])', str(part).upper()))
                cands = candidate_cache.get((refdes, part, fp), [])
                cand_str = '; '.join(f"{c[0]}: {c[1][:40]}" for c in cands[:5])
                ws.append([refdes, part, fp, qty, pn, name, spec, cand_str])
                if is_nc:
                    for col in range(1,9): ws.cell(row=row_num, column=col).fill = blue
                elif is_unmatched:
                    for col in range(1,9): ws.cell(row=row_num, column=col).fill = yellow
                elif is_repeated:
                    for col in range(1,9): ws.cell(row=row_num, column=col).fill = gray
            else:  # summary row
                ws.append(row_data)
                for col in range(1,9):
                    cell = ws.cell(row=row_num, column=col)
                    cell.fill = green
                    cell.font = openpyxl.styles.Font(bold=True)
            for col in range(1,9):
                ws.cell(row=row_num, column=col).border = thin_border
        # ===== 格式化 =====
        _progress(85, '正在格式化输出...')
        # 对齐设置: 表头居中，Qty/料号/名称居中，其他靠左
        align_center = openpyxl.styles.Alignment(horizontal='center', vertical='center', wrap_text=True)
        align_left = openpyxl.styles.Alignment(horizontal='left', vertical='center', wrap_text=True)
        center_cols = {4, 5, 6}  # Quantity, 料号, 名称
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
            for cell in row:
                col_idx = cell.column
                if cell.row == 1:
                    cell.alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center', wrap_text=True)
                elif col_idx in center_cols:
                    cell.alignment = align_center
                else:
                    cell.alignment = align_left

        # 列宽(cm换算, 1cm≈5.5字符): 位号/规格=8cm, PartValue/Footprint=3cm, 候选=2cm
        col_widths = {
            1: 44,   # 位号 8cm
            2: 16.5, # Part Value 3cm
            3: 16.5, # PCB Footprint 3cm
            4: 10,   # Quantity
            5: 16,   # 料号
            6: 14,   # 名称
            7: 44,   # 规格 8cm
            8: 11,   # 候选料号 2cm
        }
        for col, w in col_widths.items():
            ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = w

        # 行高: 逐列计算溢出行数取最大值，保守估算确保不截断
        # 每列实际可容纳半角字符数(列宽对应cm，Calibri 11pt下保守取值)
        import math
        ws.row_dimensions[1].height = 22
        col_eff = {
            1: 36,  # 位号 8cm -> 44名义宽，实际约36半角字/行
            2: 13,  # Part Value 3cm -> 16.5名义宽
            3: 13,  # PCB Footprint 3cm
            7: 36,  # 规格 8cm
            8: 9,   # 候选料号 2cm -> 11名义宽
        }
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            max_lines = 1
            for cell in row:
                if cell.column in col_eff and cell.value:
                    content = str(cell.value)
                    char_w = sum(2 if ord(c) > 127 else 1 for c in content)
                    lines = math.ceil(char_w / col_eff[cell.column])
                    max_lines = max(max_lines, lines)
            ws.row_dimensions[row[0].row].height = max(30, max_lines * 15)

        wb.save(path)
        return path

    export_results(results)
    _progress(100, '完成')
    _log(f"\n完成。输出: {OUTPUT_FILE}, 总{len(results)}行, 未匹配{len(unmatched)}条")
    if candidate_cache:
        _log(f"  注意：有 {len(candidate_cache)} 行存在多个候选料号，已列在'候选料号'列")
    return {
        'output_file': OUTPUT_FILE,
        'results': results,
        'unmatched': unmatched,
        'code_index': code_index,
        'candidate_cache': candidate_cache,
        'export': export_results,
    }

if __name__ == '__main__':
    main()
