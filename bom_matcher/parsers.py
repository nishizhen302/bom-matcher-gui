# -*- coding: utf-8 -*-
"""解析层：BOM写法与编码表规格 → 结构化字段。全部为纯函数，可独立单测。

所有值统一基准单位：电阻Ω、电容pF、电感uH。
"""
import re

CHIP_SIZES = {'0402', '0603', '0805', '1206', '1210', '1812', '2010', '2512'}
METRIC_TO_INCH = {'1005': '0402', '1608': '0603', '2012': '0805', '3216': '1206',
                  '3225': '1210', '4532': '1812', '5025': '2010', '6332': '2512'}

# 值与精度之间常见分隔符：空格 % / _ -（"100R-5%"、"3K-5%"写法）
_SEP = r'[\s%/_\-]'
_SEP_R = r'[\s%/_\-,，+]'   # 阻值另认逗号（'240R,1%'）和 '+-1%'

NC_RE = re.compile(r'(^|[^A-Z0-9])NC($|[^A-Z0-9])')


def norm(s):
    """归一化：小写、去掉字母数字汉字以外的一切"""
    return re.sub(r'[^a-z0-9一-鿿]', '', str(s).lower())


def is_nc(part):
    return bool(NC_RE.search(str(part).upper()))


def strip_nc(part):
    """去掉NC标记，返回剩余有效内容（可能为空串）"""
    s = re.sub(r'\bNC\b', '', str(part), flags=re.I)
    return s.strip(' ./_-,，')


def tokens_of(text):
    """按非字母数字切分并归一化，同时给出相邻token的连写形式。
    'B4B-XH-A' -> ['b4b','xh','a','b4bxha']（连写放最后，用于整型号匹配）"""
    raw = [t for t in re.split(r'[^0-9A-Za-z]+', str(text)) if t]
    toks = [t.lower() for t in raw]
    joined = ''.join(toks)
    if len(raw) > 1 and joined:
        toks.append(joined)
    return toks


# ---------------- 封装 ----------------

def parse_package(fp_or_spec, spec_mode=False):
    """返回 (归一化封装, 尺寸dims)。
    spec_mode=True 时用于编码表规格文本（只认标准封装token，避免误抓型号数字）"""
    s = str(fp_or_spec or '').upper()
    if not s:
        return '', None
    dims = parse_dims(s)
    # 铝电解：CAPAE660X770N / AE660 -> AE660(直径6.60mm)；CAPAE1040X1000N -> AE1040(10.40mm)
    m = re.search(r'AE\s*(\d{3,4})(?=X|\b|_|$)', s)
    if m:
        return 'AE' + m.group(1), dims
    # IPC-7351 命名（RESC1608X60N / CAPC3216X140N / LEDC1608X102N）里是公制尺寸，换成英制
    m = re.match(r'^(?:RES|CAP|IND|LED)C(\d{4})X', s)
    if m and m.group(1) in METRIC_TO_INCH:
        return METRIC_TO_INCH[m.group(1)], dims
    # 贴片尺寸：第一个命中已知尺寸表的4位数字
    for m4 in re.findall(r'\d{4}', s):
        if m4 in CHIP_SIZES:
            return m4, dims
    if spec_mode:
        # 规格侧只识别 SOD/SOT/SOIC 等标准封装词
        m = re.search(r'\b(SOD|SOT|SOIC|SOP|MSOP|QSOP|QFN|BGA|DO)[-_ ]?(\d+\w{0,2})', s)
        if m:
            return (m.group(1) + m.group(2)).lower(), dims
        return '', dims
    m = re.search(r'^(SOD|SOT|SOIC|SOP|MSOP|QSOP|QFN|BGA|DO|TO)[-_]?(\d+\w{0,4})', s)
    if m:
        return (m.group(1) + m.group(2)).lower(), dims
    return norm(s)[:24], dims


def parse_dims(s):
    """尺寸段：12.5X12.5 / 12_2x12_2 / 6.3X7.7mm -> (12.5,12.5)。>60mm视为非尺寸。"""
    m = re.search(r'(\d+(?:[._]\d+)?)\s*[xX*]\s*(\d+(?:[._]\d+)?)', str(s))
    if not m:
        return None
    try:
        a = float(m.group(1).replace('_', '.'))
        b = float(m.group(2).replace('_', '.'))
    except ValueError:
        return None
    if a <= 0 or b <= 0 or a > 60 or b > 60:
        return None
    return (a, b)


def is_chip_size(pkg):
    return pkg in CHIP_SIZES


# ---------------- 电阻 ----------------

def parse_r_value(text):
    """'100R-5%'/'4K7'/'10K'/'0'/'100R*4' -> (欧姆值, 阵列个数) 或 None"""
    # Ω 的几种写法一律去掉：Ω / Ω(U+2126) / OHM / 繁体系统画的板子读出来的乱码 U+E5A1(Big5 的 Ω)
    s = re.sub(r'OHMS?|[ΩΩ]', '', str(text).strip().upper())
    if not s:
        return None
    array_count = 1
    m = re.search(r'\*(\d+)', s)
    if m:
        array_count = int(m.group(1))
        s = re.sub(r'\*\d+', '', s)
    for tok in re.split(_SEP_R, s):
        tok = tok.strip()
        if not tok:
            continue
        if tok.isdigit() and len(tok) == 4:
            continue  # 0402 之类封装数字
        m = re.fullmatch(r'(\d+)([KMGR])(\d+)', tok)  # 4K7 / 2R2 / 8K2
        if m:
            val = float(m.group(1) + '.' + m.group(3))
            return val * _R_MULT[m.group(2)], array_count
        m = re.fullmatch(r'(\d+\.\d+)([KMGR])\d+', tok)  # '4.99K1%' 少了分隔符：末尾数字是精度
        if m:
            return float(m.group(1)) * _R_MULT[m.group(2)], array_count
        m = re.fullmatch(r'(\d+(?:\.\d+)?)([KMGR]?)', tok)  # 10K / 100R / 0
        if m:
            return float(m.group(1)) * _R_MULT.get(m.group(2), 1), array_count
    return None


_R_MULT = {'K': 1e3, 'M': 1e6, 'G': 1e9, 'R': 1, '': 1}


def parse_tolerance(text):
    m = re.search(r'(\d+(?:\.\d+)?)\s*%', str(text))
    return float(m.group(1)) if m else None


# ---------------- 电容 ----------------

def parse_c_value(text):
    """'100nF-16V'/'0.1uF/50V'/'22pF' -> (pF, 电压V或None) 或 (None,None)"""
    s = str(text).strip().upper()
    m = re.search(r'([\d\.]+)\s*(UF|U|NF|N|PF|P)(?![A-Z])', s)
    if not m:
        return None, None
    val = float(m.group(1))
    pf = int(round(val * {'UF': 1e6, 'U': 1e6, 'NF': 1e3, 'N': 1e3, 'PF': 1, 'P': 1}[m.group(2)]))
    mv = re.search(r'[\s,，/\-]\s*(\d+(?:\.\d+)?)\s*V\b', s[m.end():])
    volt = float(mv.group(1)) if mv else None
    return pf, volt


def parse_spec_c(spec):
    """编码表电容规格 -> (pF, V, 封装)"""
    s = str(spec or '')
    pkg, _ = parse_package(s, spec_mode=True)
    mv = re.search(r'(\d+(?:\.\d+)?)V', s, re.I)
    volt = float(mv.group(1)) if mv else None
    m = re.search(r'(?<![a-zA-Z])(\d+(?:\.\d+)?)\s*(UF|U|NF|N|PF|P)\b', s, re.I)
    if not m:
        return None, volt, pkg
    val = float(m.group(1))
    pf = int(round(val * {'UF': 1e6, 'U': 1e6, 'NF': 1e3, 'N': 1e3, 'PF': 1, 'P': 1}[m.group(2).upper()]))
    return pf, volt, pkg


# ---------------- 电感/磁珠 ----------------

def parse_l_value(text):
    """'CR32-10uH'/'15uH'/'4R7'/'1uH_10A'/'4.7u_2.8A' -> uH 或 None"""
    s = str(text).strip().upper()
    # 单位后面不能再接字母；'_' 不算（'1uH_10A' 曾因 \\b 认不出）
    m = re.search(r'(\d+(?:\.\d+)?)\s*(UH|MH|NH|U)(?![A-Z])', s)
    if m:
        val = float(m.group(1))
        return val * {'UH': 1, 'U': 1, 'MH': 1e3, 'NH': 1e-3}[m.group(2)]
    m = re.search(r'(\d+)R(\d+)', s)  # 4R7 -> 4.7uH
    if m:
        return float(m.group(1) + '.' + m.group(2))
    return None


def parse_series(text):
    """型号系列前缀：'CR32-10uH' -> 'CR32'；无则''"""
    tok = re.split(r'[\s/_\-,，]', str(text).strip().upper())[0]
    m = re.match(r'^([A-Z]{1,6}\d{1,4}[A-Z]{0,2})$', tok)
    return m.group(1) if m and not re.match(r'^\d', tok) else ''


def is_bead(part):
    u = str(part).upper()
    return ('BEAD' in u or 'FEB' in u or u.startswith(('MPZ', 'BLM')) or '100MHZ' in u)


def bead_ohms(part):
    """磁珠阻抗 Ω：'220R/100MHz'、'470Ohm/0603'、'BEAD/100R/0805'；Murata 'BLM21PG300SN1D' 300=30×10^0；
    TDK 'MPZ2012S101A' 101=10×10^1。认不出返回 None"""
    u = str(part).upper()
    m = re.match(r'(?:BLM\d{2}[A-Z]{2}|MPZ\d{4}[A-Z])(\d{2})(\d)', u)
    if m:
        return int(m.group(1)) * 10 ** int(m.group(2))
    m = re.search(r'(\d+)\s*(?:R|OHM|Ω|Ω|)', u)   # '46R0C-010' 也要认出 46
    return int(m.group(1)) if m else None


def parse_current_a(text):
    """'1uH_10A' / '4.7u_2.8A' / '4.7uH,4.5A' -> 安培；没有返回 None"""
    m = re.search(r'(?<![A-Z\d.])(\d+(?:\.\d+)?)\s*A(?![A-Z])', str(text).upper())
    return float(m.group(1)) if m else None


# ---------------- 连接器 ----------------

def connector_type(s):
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
    if re.match(r'^MPC[-\s]', s) or 'DC JACK' in s or 'DC_RA' in s: return 'DC_JACK'
    return ''


def header_pins(part, fp):
    """HEADER_15*2W / con-2_0-2x15p-v -> (rows, cols) 或 None"""
    m = re.search(r'HEADER[_-](\d+)[*Xx](\d+)', str(part).upper())
    if m:
        return m.group(1), m.group(2)
    m = re.search(r'HEADER[\s_-]*(\d+)(?![\d*Xx])', str(part).upper())   # 'HEADER 6' 单排
    if m:
        return m.group(1), '1'
    m = re.search(r'(\d+)[Xx](\d+)P', str(fp).upper())
    if m:
        return m.group(1), m.group(2)
    return None


def fp_pitch_mm(fp):
    """封装名里的间距：'HDRV6W50P200_1X6' -> 2.0；'CON-2_0-2X15P-V' -> 2.0；认不出 None"""
    s = str(fp).upper()
    m = re.search(r'P(\d{3})(?:_|$)', s)            # IPC 排针 P200/P254/P125
    if m:
        return int(m.group(1)) / 100
    m = re.search(r'(?:CON|J)[-_]?(\d)[._](\d{1,2})', s)   # CON-2_0 / J-1_25
    if m:
        return float(f'{m.group(1)}.{m.group(2)}')
    return None


def fp_orient(fp):
    """封装名里的方向：HDRV…/-V → 'V'(立式)；HDRRA…/-RA/-H → 'RA'(卧式)；看不出 None"""
    s = str(fp).strip().upper()
    if re.search(r'HDRRA|[-_]RA(?:[-_]|$)|[-_]H(?:[-_]|$)', s):
        return 'RA'
    if re.search(r'HDRV|[-_]V(?:[-_]|$)', s):
        return 'V'
    return None


def spec_orient(spec):
    s = str(spec)
    if '卧' in s:
        return 'RA'
    if '立' in s:
        return 'V'
    return None


def spec_pitch_mm(spec):
    m = re.search(r'(\d+(?:\.\d+)?)\s*MM', str(spec).upper())
    return float(m.group(1)) if m else None


# ---------------- 其他 ----------------

def parse_freq_mhz(text):
    """'14.318MHz-20ppm' / '24.576M' -> MHz 或 None"""
    s = str(text).strip().upper()
    m = re.search(r'(\d+(?:\.\d+)?)\s*M(?:HZ)?(?![A-Z0-9])', s)
    return float(m.group(1)) if m else None


def spec_freqs_mhz(spec):
    """编码表规格里所有频率值（MHz）：'SMD3225,14.318MHz,20ppM' -> [14.318]；KHz 换算成 MHz"""
    s = str(spec).upper()
    out = [float(x) for x in re.findall(r'(\d+(?:\.\d+)?)\s*MHZ', s)]
    out += [float(x) / 1000 for x in re.findall(r'(\d+(?:\.\d+)?)\s*KHZ', s)]
    return out


def parse_volt(text):
    """提取电压：3.3V / 1-8(→1.8) / 1V8(→1.8)"""
    s = str(text).upper()
    m = re.search(r'(\d+(?:\.\d+)?)\s*V\b', s)
    if m:
        return float(m.group(1))
    m = re.search(r'(\d+)V(\d+)', s)
    if m:
        return float(m.group(1) + '.' + m.group(2))
    m = re.search(r'(\d+)[\-_](\d+)', s)
    if m:
        w, d = int(m.group(1)), int(m.group(2))
        if 1 <= w <= 30 and 1 <= d <= 99:
            return float(f'{w}.{d}')
    return None


def parse_spec_r(spec):
    """编码表电阻规格 -> (欧姆, 封装, 精度%, 是否排阻)"""
    s = str(spec or '')
    fields = [f for f in re.split(r'[,;，\s]+', s) if f]
    ohms = None
    for f in fields:
        fu = f.upper()
        m = re.fullmatch(r'(\d+)([KMGR])(\d+)', fu)
        if m:
            ohms = float(m.group(1) + '.' + m.group(3)) * _R_MULT[m.group(2)]
            break
        m = re.fullmatch(r'(\d+(?:\.\d+)?)([KMG]?)R?', fu)
        if m and not (m.group(1).isdigit() and len(m.group(1)) == 4 and not m.group(2)):
            ohms = float(m.group(1)) * _R_MULT.get(m.group(2), 1)
            break
    pkg, _ = parse_package(s, spec_mode=True)
    tol = parse_tolerance(s)
    is_array = bool(re.search(r'\d{4}X\d\b', s.upper()))
    return ohms, pkg, tol, is_array


def parse_spec_l(spec):
    """编码表电感规格 -> uH 或 None"""
    s = str(spec or '').upper()
    m = re.search(r'(\d+(?:\.\d+)?)\s*(UH|MH|NH)\b', s)
    if m:
        return float(m.group(1)) * {'UH': 1, 'MH': 1e3, 'NH': 1e-3}[m.group(2)]
    m = re.search(r'(\d+)R(\d+)', s)
    if m:
        return float(m.group(1) + '.' + m.group(2))
    return None
