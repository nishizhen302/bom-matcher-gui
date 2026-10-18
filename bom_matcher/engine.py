# -*- coding: utf-8 -*-
"""匹配引擎：行分类 → 别名规则 → 分类别匹配（硬约束过滤 + 软打分排序）。

设计约束：
- 匹配只使用解析后的结构化字段，不做原始字符串子串碰撞；
- 结果与编码表行序无关（候选池按料号排序，平分时取料号最小者）；
- 永远返回 最佳+候选+理由，绝不静默取第一个命中。
"""
import re

from .model import (Item, Match, KIND_BLANK, KIND_MECH, KIND_NORMAL,
                    CAT_R, CAT_C, CAT_L, CAT_CONN, CAT_OTHER)
from . import parsers as P
from . import rules as R


IC_PREFIXES = [
    'SGM', 'RTD', 'RT', 'MP', 'LT', 'MAX', 'LM', 'TPS', 'AP', 'PI', 'AD',
    'STM', 'AT', 'XC', 'MT', 'TC', 'SP', 'BU', 'KA', 'DR', 'W25',
    'MX', 'CY', 'PW', 'GL', 'IT', 'DS', 'AW', 'JW', 'YT', 'TX', 'SY',
]

MECH_FP_KEYS = ('mtg', 'hole', 'via', 'mark', 'npth', 'pad_test')
MECH_REF_PREFIXES = {'TP', 'MARK'}
# 按位号前缀定类别（2026-10-18 用户确认）：封装名不规范时（PR0402、PC0603、RESC1608、CON…）
# 只看封装首字母会判错。RD/RDP/RN 可能是排阻，是否排阻再看阻值 *N 或封装 RP4
REF_CATEGORY = {
    'R': CAT_R, 'RD': CAT_R, 'RDP': CAT_R, 'RN': CAT_R, 'RP': CAT_R,
    'C': CAT_C, 'CD': CAT_C, 'CE': CAT_C, 'EC': CAT_C,
    'L': CAT_L, 'LW': CAT_L, 'FB': CAT_L, 'CM': CAT_L,
}
# 这些位号一定不是阻容感，判类别时不看封装首字母（U 开头的一律算芯片）
CHIP_REF_PREFIXES = {'IC', 'Q', 'D', 'Y', 'X', 'ZD', 'LED', 'VD', 'VT',
                     'CN', 'CNF', 'J', 'JP', 'P', 'HDMI', 'DP', 'USB', 'VGA', 'DVI'}


def _ref_prefix(refdes):
    m = re.match(r'[A-Za-z]+', str(refdes).strip())
    return m.group().upper() if m else ''
MECH_PARTS = {'ORIENTATION', 'TP', 'TEST POINT', 'TESTPOINT'}
NON_PART_PREFIXES = ('Y0616', 'Y12')
# 这些词只说明器件种类，不是型号：曾把 'DIODE' 配成 B340LB、'TVS' 配成 SMAJ28A、'LED' 配成某颗双色灯
GENERIC_WORDS = {
    'diode', 'led', 'tvs', 'esd', 'zener', 'conn', 'con', 'connector', 'header', 'cap',
    'res', 'ind', 'fuse', 'sw', 'switch', 'button', 'key', 'mark', 'ph', 'dnp', 'nc',
    'crystal', 'xtal', 'ic', 'mos', 'mosfet', 'npn', 'pnp', 'transistor', 'bead', 'jumper',
    'flash', 'eeprom', 'ldo', 'mcu', 'opamp', 'dram', 'sram',
}


def _model_like(pj):
    """归一化后的 BOM 写法能不能当型号去编码表里搜。"""
    if pj in GENERIC_WORDS:
        return False
    if re.fullmatch(r'\d+(?:mm|mhz|khz|v|a|p|pin|w|uf|nf|pf|uh|r|k|ohm)', pj):
        return False   # '2.00mm' -> '200mm' 曾配成 200mm*130mm 的光板
    if pj.isdigit() and len(pj) < 6:
        return False   # '100'；'53398-0871' 这类长料号保留
    if re.fullmatch(r'\d+(?:x\d+)?p(?:in)?|\d+x\d+', pj):
        return False   # 针数写法 '2x6P' 曾把 0.5mm 30P 的 FPC 座配成 2.0mm 12P 排针
    if re.fullmatch(r'(?:lqfp|tqfp|qfp|qfn|dfn|bga|tssop|ssop|msop|sop|soic|sot|sod|dip|do|to)\d+[a-z]{0,2}'
                    r'|sma|smb|smc|smd\d*', pj):
        return False   # 'VL817S/QFN76' 拆出的 'QFN76' 是封装，不能拿去搜型号（SMBJ15CA 这类型号不受影响）
    return True


def _model_keys(body):
    """先整串，再 '/' 分出的各段；带数字的段优先（'FLASH/BY25D80ASTIG' 先试 BY25D80ASTIG）。"""
    segs = [P.norm(s) for s in re.split(r'[/\\]', str(body)) if s.strip()]
    segs.sort(key=lambda s: not re.search(r'\d', s))
    keys = [P.norm(body)] + segs
    return list(dict.fromkeys(k for k in keys if k))


def _token_hits(pj, pools, pkg_ok=lambda e: True):
    """型号在编码表 token 里的命中：完全相等的优先，其次 BOM 型号是 token 前缀（MP2307 → MP2307DN）。"""
    exact, prefix = [], []
    for pool in pools:
        for e in pool:
            if not pkg_ok(e):
                continue
            if pj in e.tokens:
                exact.append(e)
            elif len(pj) >= 4 and any(t.startswith(pj) for t in e.tokens):
                prefix.append(e)
    return exact or prefix


def _model_match(it, hits):
    """一个命中直接用；多个命中标需确认+候选，不默默取第一个（曾把 DDR3 静默配成两颗内存中的一颗）。"""
    if not hits:
        return None
    e = hits[0]
    if len(hits) == 1:
        return Match(e.pn, e.name, e.spec, reason=f'型号{it.body}匹配')
    return Match(e.pn, e.name, e.spec, confidence='review',
                 reason=f'型号{it.body}命中{len(hits)}个料,请确认',
                 candidates=[(c.pn, c.name, c.spec) for c in hits[1:4]])


class Engine:
    def __init__(self, entries, extra_alias=None, banned=None, debug=lambda *a: None, learned=None):
        self.learned = learned or {}       # 人工选择记忆 {写法键: {pn, board, date...}}
        self.learned_used = 0
        self.learned_skipped = []          # 记过但料号已不在编码表里（停用/删除）
        banned = (banned or set()) | R.BUILTIN_BANNED
        self.entries = [e for e in entries if e.pn not in banned]
        self.by_pn = {e.pn: e for e in self.entries}
        # 匹配候选池不含 PCB 光板(Y0616) 和辅料(Y12)：BOM 元件按型号搜时曾把 '100'、'2.00mm'
        # 配成光板。辅料只经强制规则（如 HEATSINK）按料号直取，走 by_pn 不受影响
        self.by_cat = {}
        for e in self.entries:
            if not e.pn.startswith(NON_PART_PREFIXES):
                self.by_cat.setdefault(e.category, []).append(e)
        self.alias_rules = list(extra_alias or []) + R.ALIAS_RULES
        self.debug = debug
        self._cap_selection = {}
        self._has_2556 = False

    # ---------- 行分类 ----------

    def build_item(self, refdes, part, fp, qty):
        it = Item(refdes=refdes, part=part, fp=fp, qty=qty)
        if not refdes and not part and not fp:
            it.kind = KIND_BLANK
            return it
        fl = fp.strip().lower()
        if (any(k in fl for k in MECH_FP_KEYS) or fl.startswith(('tp_', 'tp-')) or fl == 'tp'
                or part.strip().upper() in MECH_PARTS or _ref_prefix(refdes) in MECH_REF_PREFIXES):
            it.kind = KIND_MECH
            return it
        it.is_nc = P.is_nc(part)
        it.body = P.strip_nc(part) if it.is_nc else part.strip()
        it.category = self._detect_category(it.body, fp, refdes)
        it.package, it.dims = P.parse_package(fp)
        return it

    def _detect_category(self, body, fp, refdes=''):
        u = body.upper()
        fl = fp.strip().lower()
        pre = _ref_prefix(refdes)
        if pre in REF_CATEGORY:
            return REF_CATEGORY[pre]
        # 位号是芯片/连接器等（U/Q/D/Y/CN…）时，封装首字母不作数：
        # 2026-10-18 U3 'LT87121' 封装 'LT87121_QFN64_EPAD' 以 l 开头被当成电感，型号明明能精确命中却落成未匹配
        if pre in CHIP_REF_PREFIXES or pre.startswith('U'):
            fl = ''
        esd_like = (any(k in u for k in ('CLAMP', 'TVS', 'ESD'))
                    or any(k in fl for k in ('clamp', 'tvs', 'esd')))
        if 'UH' in u:
            return CAT_L
        if P.is_bead(body):
            return CAT_L
        if P.connector_type(body):
            return CAT_CONN
        if (fl.startswith('c') and 'cm' not in fl) or 'cap' in fl:
            return CAT_C
        if fl.startswith('r') and 'cm' not in fl and not esd_like:
            return CAT_R
        if (fl.startswith('l') or 'cm' in fl or 'filter' in fl or 'ind' in fl) and not esd_like:
            return CAT_L
        return CAT_OTHER

    # ---------- 预扫描 ----------

    def prescan(self, items):
        self._has_2556 = any('RTD2556' in it.body.upper() for it in items
                             if it.kind == KIND_NORMAL)
        self._build_cap_selection(items)

    # ---------- 主匹配 ----------

    def match(self, it):
        """返回 Match 或 None"""
        if not it.body:
            return None
        m = self._match_learned(it)
        if m:
            return m
        m = self._match_alias(it)
        if m:
            return m
        if 'SFLASH' in it.body.upper():
            pn = R.SFLASH_WITH_2556 if self._has_2556 else R.SFLASH_DEFAULT
            e = self.by_pn.get(pn)
            if e:
                return Match(e.pn, e.name, e.spec, reason='SFLASH规则' +
                             ('(板载RTD2556)' if self._has_2556 else ''))
        matcher = {CAT_R: self._match_r, CAT_C: self._match_c,
                   CAT_L: self._match_l, CAT_CONN: self._match_conn}.get(it.category, self._match_other)
        return matcher(it)

    def nearest(self, it):
        """未匹配时给出候选提示"""
        pj = P.norm(it.body)
        out = []
        if len(pj) >= 5:
            head = pj[:5]
            for e in self.by_cat.get(CAT_OTHER, []) + self.by_cat.get(CAT_CONN, []):
                if any(len(t) >= 5 and t[:5] == head for t in e.tokens):
                    out.append(e)
        if not out:
            ctype = P.connector_type(it.body) or P.connector_type(it.fp)
            if ctype and ctype not in ('JST_PH', 'JST_XH', 'DC_JACK'):
                out = [e for e in self.by_cat.get(CAT_CONN, []) if ctype in e.spec.upper()]
        return out[:3]

    # ---------- 人工选择记忆（比内置规则优先：用户改过的就是引擎配错或配不上的） ----------

    def _match_learned(self, it):
        if not self.learned:
            return None
        from .learned import key_of
        rec = self.learned.get(key_of(it.part, it.fp))
        if not rec:
            return None
        e = self.by_pn.get(rec.get('pn'))
        if not e:
            self.learned_skipped.append((it.refdes, it.part, rec.get('pn')))
            return None
        self.learned_used += 1
        src = '、'.join(x for x in (rec.get('board'), rec.get('date')) if x)
        return Match(e.pn, e.name, e.spec, reason='沿用人工选择' + (f'({src})' if src else ''))

    # ---------- 别名规则 ----------

    def _match_alias(self, it):
        pu = it.body.upper()
        fu = it.fp.upper()
        for part_kw, fp_kw, pns in self.alias_rules:
            if part_kw and part_kw not in pu:
                continue
            if fp_kw and fp_kw not in fu:
                continue
            for pn in pns:
                e = self.by_pn.get(pn)
                if e:
                    return Match(e.pn, e.name, e.spec, reason=f'规则: {part_kw}→{pn}')
        return None

    # ---------- 电阻 ----------

    def _match_r(self, it):
        parsed = P.parse_r_value(it.body)
        if parsed is None:
            return None
        # 按实际阻值比（0.1%容差）：取整会把 0.15Ω/0.039Ω、2.2Ω/2.4Ω 当成同一个
        ohms, count = parsed[0], parsed[1]
        if count == 1:
            # 排阻封装 RP4_0402 而值里没写 *4（RDP 位号常见）
            m = re.match(r'(?i)^R[PNA](\d)[_\-]', it.fp.strip())
            if m and int(m.group(1)) > 1:
                count = int(m.group(1))
        pool = self.by_cat.get(CAT_R, [])
        want_array = count > 1
        # 封装看不出贴片尺寸（R_4）或没填：不按封装筛，但必须人确认
        pkg_known = P.is_chip_size(it.package)

        def hit(e, need_array):
            if e.r_ohms is None or abs(e.r_ohms - ohms) > ohms * 1e-3:
                return False
            if e.is_array != need_array:
                return False
            if pkg_known and e.package != it.package:
                return False
            return True

        cands = sorted([e for e in pool if hit(e, want_array)],
                       key=lambda e: (e.tol if e.tol is not None else 999, e.pn))
        note = ''
        if not cands and want_array:
            cands = sorted([e for e in pool if hit(e, False)],
                           key=lambda e: (e.tol if e.tol is not None else 999, e.pn))
            note = f'(需用{count}个)'
        if not cands:
            return None
        e = cands[0]
        reason = f'阻值{_fmt_r(ohms)}+封装{it.package if pkg_known else "未知"}' + \
                 (f'+阵列x{count}' if want_array and e.is_array else '') + note
        if not pkg_known:
            reason += f';封装{it.fp or "(空)"}看不出尺寸,请确认'
        return Match(e.pn, e.name, e.spec, reason=reason,
                     confidence='exact' if pkg_known else 'review',
                     candidates=[(c.pn, c.name, c.spec) for c in cands[1:4]])

    # ---------- 电容 ----------

    def _build_cap_selection(self, items):
        """全局电容分组优化：同(容值,封装,电压)组统一选料，减少物料种类。"""
        groups = {}      # key -> 行数
        cand_map = {}    # key -> [(entry, priority)]
        pool = self.by_cat.get(CAT_C, [])
        for it in items:
            if it.kind != KIND_NORMAL or it.category != CAT_C:
                continue
            pf, volt = P.parse_c_value(it.body)
            if pf is None:
                continue
            key = (pf, it.package, volt if volt is not None else 0)
            groups[key] = groups.get(key, 0) + 1
            if key in cand_map:
                continue
            cands = []
            for e in pool:
                if e.c_pf != pf:
                    continue
                if it.package and e.package and e.package != it.package:
                    continue
                if it.package and not e.package:
                    continue
                if volt is not None:
                    if e.volt == volt:
                        cands.append((e, 0))
                    elif e.volt is not None and e.volt > volt:
                        cands.append((e, 1))
                else:
                    cands.append((e, 2))
            cand_map[key] = cands
        usage = {}
        for key, cands in cand_map.items():
            for e, _ in cands:
                usage[e.pn] = usage.get(e.pn, 0) + groups.get(key, 0)
        for key, cands in cand_map.items():
            if not cands:
                continue
            _, _, volt_key = key
            if volt_key:
                strict = [e for e, p in cands if p == 0]
                if strict:
                    strict.sort(key=lambda e: (-usage.get(e.pn, 0), e.pn))
                    self._cap_selection[key] = strict[0].pn
                    continue
                compat = [e for e, p in cands if p == 1]
                if compat:
                    compat.sort(key=lambda e: (-usage.get(e.pn, 0), e.volt, e.pn))
                    self._cap_selection[key] = compat[0].pn
            else:
                allc = [e for e, _ in cands]
                allc.sort(key=lambda e: (-(e.volt or 0), -usage.get(e.pn, 0), e.pn))
                self._cap_selection[key] = allc[0].pn
        # 合并：同(容值,封装)选出了多个料号时，若最高电压者能覆盖某组的电压要求，
        # 该组统一到它（高压可替代低压，减少物料种类）
        clusters = {}
        for key in self._cap_selection:
            clusters.setdefault((key[0], key[1]), []).append(key)
        for keys in clusters.values():
            pns = {self._cap_selection[k] for k in keys}
            if len(pns) <= 1:
                continue
            best = max((self.by_pn[p] for p in pns), key=lambda e: (e.volt or 0, e.pn))
            if best.volt is None:
                continue
            for k in keys:
                if k[2] == 0 or best.volt >= k[2]:
                    self._cap_selection[k] = best.pn

    def _match_c(self, it):
        pf, volt = P.parse_c_value(it.body)
        if pf is None:
            return None
        # 铝电解（AE封装）：直径+电压兼容
        if it.package.startswith('AE'):
            return self._match_c_aluminum(it, pf, volt, int(it.package[2:]) / 100.0)
        # 其他电解写法：CAPACITOR_POL / ECAP-T6_5X6_5 / EC080-… / 位号 CE、EC。有尺寸按直径，没有只比容值耐压
        if not P.is_chip_size(it.package) and (
                re.search(r'(?i)POL|ECAP|^EC\d', it.fp.strip())
                or _ref_prefix(it.refdes) in ('CE', 'EC')):
            return self._match_c_aluminum(it, pf, volt, it.dims[0] if it.dims else None)
        key = (pf, it.package, volt if volt is not None else 0)
        pn = self._cap_selection.get(key)
        if pn:
            e = self.by_pn[pn]
            if volt is None:
                reason, conf = f'容值+封装匹配,无电压要求取{e.volt:g}V' if e.volt else '容值+封装匹配', 'exact'
            elif e.volt == volt:
                reason, conf = '容值+电压+封装严格匹配', 'exact'
            else:
                reason, conf = f'电压{e.volt:g}V≥{volt:g}V兼容代用', 'compat'
            return Match(e.pn, e.name, e.spec, reason=reason, confidence=conf)
        return None

    def _match_c_aluminum(self, it, pf, volt, diam):
        best, best_score = None, 1e9
        for e in self.by_cat.get(CAT_C, []):
            if e.c_pf is None or e.dims is None:
                continue
            if diam is not None and abs(e.dims[0] - diam) > 0.6:
                continue
            cap_diff = abs(e.c_pf - pf) / pf
            if cap_diff > 0.5:
                continue
            if volt is not None and (e.volt is None or e.volt < volt):
                continue
            score = cap_diff + (((e.volt - volt) / volt * 0.1) if volt else 0)
            if score < best_score:
                best, best_score = e, score
        if best:
            same = abs(best.c_pf - pf) <= pf * 1e-3
            if diam is None:
                return Match(best.pn, best.name, best.spec, confidence='review',
                             reason=f'电解:封装{it.fp}无直径信息,按容值耐压选,请确认尺寸')
            return Match(best.pn, best.name, best.spec,
                         reason=f'铝电解:直径{diam:g}mm+容值' + ('一致' if same else '近似,请确认'),
                         confidence='compat' if same else 'review')
        return None

    # ---------- 电感/磁珠 ----------

    def _match_l(self, it):
        if P.is_bead(it.body):
            return self._match_bead(it)
        uh = P.parse_l_value(it.body)
        if uh is None:
            # FB/LW 位号、值里没有电感量（'470Ohm/0603'）：按磁珠处理
            if _ref_prefix(it.refdes) in ('FB', 'LW') and P.bead_ohms(it.body) is not None:
                return self._match_bead(it)
            return None
        series = P.parse_series(it.body)
        pool = [e for e in self.by_cat.get(CAT_L, [])
                if e.l_uh is not None and abs(e.l_uh - uh) < 0.01]
        if it.package in P.CHIP_SIZES:
            pool = [e for e in pool if not e.package or e.package == it.package]
        if not pool:
            return None
        # BOM 写了电流（'1uH_10A'）：只在额定电流够的里面选；都不够就标需确认（曾把 10A 配成 8.5A）
        need_a = P.parse_current_a(it.body)
        short_a = False
        if need_a:
            enough = [e for e in pool if (P.parse_current_a(e.spec) or 0) >= need_a]
            if enough:
                pool = enough
            else:
                short_a = True

        def score(e):
            s = 0
            if series and any(t.startswith(series.lower()) for t in e.tokens):
                s += 100
            if it.dims and e.dims:
                s += max(0, 40 - 10 * abs(it.dims[0] - e.dims[0]))
            elif it.package in P.CHIP_SIZES and e.package == it.package:
                s += 50
            return s

        pool.sort(key=lambda e: (-score(e), e.pn))
        e = pool[0]
        conf, reason = 'exact', f'感值{uh:g}uH'
        if series:
            if score(e) >= 100:
                reason += f'+系列{series}'
            else:
                conf = 'review'
                reason += f';系列{series}在编码表无对应,按感值近似,请确认'
        elif len(pool) > 1 and not it.dims:
            conf = 'review'
            reason += ';无封装尺寸信息,多候选请确认'
        if short_a:
            conf = 'review'
            reason += f';编码表里同感值的额定电流都不到{need_a:g}A,请确认'
        elif need_a:
            reason += f'+电流≥{need_a:g}A'
        return Match(e.pn, e.name, e.spec, reason=reason, confidence=conf,
                     candidates=[(c.pn, c.name, c.spec) for c in pool[1:5]])

    def _match_bead(self, it):
        pn = R.BEAD_BY_PACKAGE.get(it.package)
        e = self.by_pn.get(pn) if pn else None
        if not e:
            for c in self.by_cat.get(CAT_L, []):
                if c.pn.startswith('Y0606') and c.package == it.package:
                    e = c
                    break
        if not e:
            return None
        conf, reason = 'exact', f'磁珠按封装{it.package}映射'
        ohm = P.bead_ohms(it.body)
        if ohm is not None and abs(ohm - 100) > 15:
            conf = 'review'
            reason += f';阻抗{ohm}Ω与库存100Ω不同,请确认'
        return Match(e.pn, e.name, e.spec, reason=reason, confidence=conf)

    # ---------- 连接器 ----------

    def _match_conn(self, it):
        pools = [self.by_cat.get(CAT_CONN, []), self.by_cat.get(CAT_OTHER, [])]
        pj = P.norm(it.body)
        # 1) 整型号token匹配（同样按 '/' 分段再试）
        for seg in _model_keys(it.body):
            if len(seg) >= 4 and _model_like(seg):
                m = _model_match(it, _token_hits(seg, pools))
                if m:
                    return m
        # 2) HEADER针数
        ctype = P.connector_type(it.body)
        if ctype == 'HEADER':
            pins = P.header_pins(it.body, it.fp)
            if pins:
                rows, cols = pins
                total = str(int(rows) * int(cols))
                # 数字两侧不能再接数字：'2P' 不能命中 '12P'、'2*15' 不能命中 '12*15'
                # 总针数前面不能是 '*'/'X'：'2*6p' 是双排 12 针，曾被当成 6 针
                needles = [rf'(?<![\d.]){rows}\*{cols}(?!\d)', rf'(?<![\d.]){cols}\*{rows}(?!\d)',
                           rf'(?<![\d.]){rows}X{cols}(?!\d)', rf'(?<![\d.]){cols}X{rows}(?!\d)',
                           rf'(?<![\d.*X]){total}P(?:IN)?(?![A-Z0-9])']
                # 只在连接器/插座料里找（'6P' 曾可能命中 SOT23-6P 之类芯片）；间距对得上才算完全匹配
                pitch, orient = P.fp_pitch_mm(it.fp), P.fp_orient(it.fp)
                hits = [e for pool in pools for e in pool
                        if e.pn.startswith(('Y0613', 'Y0614'))
                        and any(re.search(n, e.spec.upper()) for n in needles)]
                # 间距一致、方向一致（立式/卧式）的排前面；两样都核对上才算完全匹配
                pitch_ok = lambda e: pitch is not None and P.spec_pitch_mm(e.spec) == pitch
                orient_ok = lambda e: orient is not None and P.spec_orient(e.spec) == orient
                hits.sort(key=lambda e: (not pitch_ok(e), not orient_ok(e)))
                if hits:
                    e = hits[0]
                    ok = pitch_ok(e) and orient_ok(e)
                    miss = [w for w, f in (('间距', pitch_ok), ('立/卧式', orient_ok)) if not f(e)]
                    return Match(e.pn, e.name, e.spec,
                                 confidence='exact' if ok else 'review',
                                 reason=f'HEADER {rows}x{cols}={total}针' +
                                        (f'+间距{pitch:g}mm+方向一致' if ok else f',{"/".join(miss)}未核对上,请确认'),
                                 candidates=[(c.pn, c.name, c.spec) for c in hits[1:4]])
        # 3) 类型关键词
        if ctype in ('DVI', 'USB', 'HDMI', 'DP', 'RJ45', 'VGA', 'ETHERNET'):
            for pool in pools:
                for e in pool:
                    if ctype in e.spec.upper():
                        # 只按接口类型对上（贴片/卧式、A/B 型都没比），必须人看
                        return Match(e.pn, e.name, e.spec, reason=f'类型{ctype}关键词,型号未比对,请确认',
                                     confidence='review')
        return None

    # ---------- 其他（IC/二极管/保险丝/晶振/ESD…） ----------

    def _match_other(self, it):
        body_u = it.body.upper()
        pool = self.by_cat.get(CAT_OTHER, [])
        # 晶振
        m = self._match_crystal(it, body_u)
        if m:
            return m
        # 1117 LDO
        if '1117' in body_u:
            m = self._match_ldo(it, body_u)
            if m:
                return m
        # 保险丝 XA/YV
        m = self._match_fuse(it, body_u)
        if m:
            return m
        # 稳压管
        m = self._match_zener(it, body_u)
        if m:
            return m
        # 通用型号token匹配（整串不中再按 '/' 分段试：'VL817S/QFN76'、'1N4148/15AHEQ003B'）
        pj = P.norm(it.body)
        for seg in _model_keys(it.body):
            if len(seg) >= 3 and _model_like(seg):
                m = _model_match(it, _token_hits(seg, [pool, self.by_cat.get(CAT_CONN, [])],
                                                 lambda e: self._pkg_ok(it, e)))
                if m:
                    return m
        # IC前缀+数字段一致
        m = self._match_ic(it, body_u, pj)
        if m:
            return m
        # ESD兜底：类别词在Part、真实型号在Footprint
        fj = P.norm(it.fp)
        if (fj and fj != pj and len(fj) >= 8 and re.search(r'[a-z]', fj)
                and re.search(r'\d', fj)
                and not re.match(r'^(sot|sod|soic|sop|qsop|msop|qfn|bga|dip|to\d|do\d|smd|led|cap|con)', fj)):
            for e in pool:
                if any(t == fj or t.startswith(fj) for t in e.tokens):
                    return Match(e.pn, e.name, e.spec,
                                 reason=f'封装列型号{it.fp}匹配', confidence='compat')
        return None

    def _pkg_ok(self, it, e):
        if P.is_chip_size(it.package) and P.is_chip_size(e.package):
            return it.package == e.package
        return True

    def _match_crystal(self, it, body_u):
        freq = P.parse_freq_mhz(it.body)
        if freq is None:
            return None
        if not ('MHZ' in body_u or 'PPM' in body_u or it.fp.upper().startswith('SMD')):
            return None
        needle = f'{freq:g}'
        fpn = P.norm(it.fp)
        # 按规格里「xxMHz」的数值比（0.1%容差）：原来找子串，'25' 会命中 SMD3225、'12' 会命中 12pF
        hits = [e for e in self.by_cat.get(CAT_OTHER, [])
                if e.pn.startswith('Y0611')
                and any(abs(f - freq) <= freq * 1e-3 for f in P.spec_freqs_mhz(e.spec))]
        if not hits:
            return None
        # 封装命中的优先；一旦有封装命中，其他封装的就不再作为候选干扰
        pkg_hits = [e for e in hits if (fpn and fpn in e.norm) or _fits_fp(e, it.fp)]
        if pkg_hits:
            pkg_hits.sort(key=lambda e: e.pn)
            e = pkg_hits[0]
            return Match(e.pn, e.name, e.spec,
                         reason=f'晶振频率{needle}MHz+封装{it.fp}',
                         candidates=[(c.pn, c.name, c.spec) for c in pkg_hits[1:4]])
        hits.sort(key=lambda e: e.pn)
        e = hits[0]
        return Match(e.pn, e.name, e.spec, confidence='review',
                     reason=f'晶振频率{needle}MHz;封装无精确对应,请确认',
                     candidates=[(c.pn, c.name, c.spec) for c in hits[1:4]])

    def _match_ldo(self, it, body_u):
        # 先去掉型号里的 1117 再取电压：'AMS1117-1.2' 曾把 '1117-1' 当电压格式，结果认不出
        rest = body_u.replace('1117', ' ')
        volt = P.parse_volt(rest)
        if volt is None:
            m = re.search(r'(?<![\d.])(\d{1,2}(?:\.\d+)?)(?![\d.])', rest)
            volt = float(m.group(1)) if m and 0.5 <= float(m.group(1)) <= 15 else None
        if volt is None:
            return None
        best, best_diff = None, 0.5
        for e in self.by_cat.get(CAT_OTHER, []):
            if not e.pn.startswith('Y0612') or '1117' not in e.spec.upper():
                continue
            if e.volt is None:
                continue
            d = abs(e.volt - volt)
            if d < best_diff:
                best, best_diff = e, d
        if best:
            return Match(best.pn, best.name, best.spec,
                         reason=f'1117 LDO电压{volt:g}V')
        return None

    def _match_fuse(self, it, body_u):
        m = re.search(r'(\d+(?:\.\d+)?)A\s*[/\-_]\s*(\d+(?:\.\d+)?)V', body_u)
        if not m:
            return None
        cur, volt = float(m.group(1)), float(m.group(2))
        best, best_diff = None, 1e9
        for e in self.by_cat.get(CAT_OTHER, []):
            if not e.pn.startswith('Y0607'):
                continue
            su = e.spec.upper()
            mv = re.search(r'(\d+(?:\.\d+)?)\s*V', su)
            mc = re.search(r'(\d+(?:\.\d+)?)\s*A\b', su)
            if not mv or not mc:
                continue
            if abs(float(mv.group(1)) - volt) > 0.1:
                continue
            if it.package and P.is_chip_size(it.package):
                if it.package not in su:
                    continue
            d = abs(float(mc.group(1)) - cur)
            if d < best_diff:
                best, best_diff = e, d
        if best:
            return Match(best.pn, best.name, best.spec,
                         reason=f'保险丝{cur:g}A/{volt:g}V')
        return None

    def _match_zener(self, it, body_u):
        if not any(k in body_u for k in ('UDZ', 'BZX', 'ZENER', '稳压')):
            return None
        mv = re.search(r'(\d+(?:\.\d+)?)', body_u)
        if not mv:
            return None
        volt = float(mv.group(1))
        cands = []
        for e in self.by_cat.get(CAT_OTHER, []):
            if not e.pn.startswith('Y0608'):
                continue
            su = e.spec.upper()
            if not any(k in su for k in ('UDZ', 'BZX', 'ZENER', '稳压')):
                continue
            ev = re.search(r'(\d+(?:\.\d+)?)\s*V', su)
            if not ev:
                continue
            d = abs(float(ev.group(1)) - volt)
            if d <= 0.5:
                cands.append((d, e))
        if cands:
            cands.sort(key=lambda x: (x[0], x[1].pn))
            e = cands[0][1]
            return Match(e.pn, e.name, e.spec, reason=f'稳压管{volt:g}V')
        return None

    def _match_ic(self, it, body_u, pj):
        prefix = None
        for p in IC_PREFIXES:
            if body_u.startswith(p):
                if len(p) <= 2 and p.isalpha():
                    nxt = body_u[len(p):len(p) + 1]
                    if not nxt.isdigit():
                        continue
                prefix = p
                break
        if not prefix:
            return None
        # 数字段里的 X 是通配（RTD27X5 = RTD2785/2775…），逐位比对，X 对任何数字都算相同
        mb = re.search(r'\d[\dX]*', body_u[len(prefix):])
        bom_num = mb.group().rstrip('X') if mb else ''
        if len(bom_num.replace('X', '')) < 2:
            return None
        scored = []
        for e in self.by_cat.get(CAT_OTHER, []):
            if not e.pn.startswith('Y0612'):
                continue
            su = e.spec.upper()
            pos = su.find(prefix)
            if pos < 0:
                continue
            ms = re.search(r'\d+', su[pos + len(prefix):])
            spec_num = ms.group() if ms else ''
            common = 0
            for a, b in zip(bom_num, spec_num):
                if a != b and a != 'X':
                    break
                common += 1
            if common >= 2:
                scored.append((common, e))
        if not scored:
            return None
        top = max(c for c, _ in scored)
        best = [e for c, e in scored if c == top]
        if len(best) > 1:
            # 一样像的不止一个：封装对得上的排前面，标需确认、其余进候选，不默默挑一个
            best.sort(key=lambda e: (not _fits_fp(e, it.fp), e.pn))
            e = best[0]
            return Match(e.pn, e.name, e.spec, confidence='review',
                         reason=f'IC前缀{prefix}+数字段{bom_num[:top]}近似,有{len(best)}个同样接近,请确认',
                         candidates=[(c.pn, c.name, c.spec) for c in best[1:4]])
        e = best[0]
        # 型号只是近似（SGM61430→SGM61431），是另一颗料，必须标色让人看
        return Match(e.pn, e.name, e.spec, confidence='review',
                     reason=f'IC前缀{prefix}+数字段{bom_num[:top]}近似,请确认')


def _fits_fp(e, fp):
    """编码表规格里的封装字样能否在 BOM 封装名里找到：
    'RTD2785T-CG,LQFP216' ↔ 'LQFP216_24X24_DB'；'SMD3225,14.318MHz' ↔ 'CRY-D3225'（看4位尺寸）"""
    fpn = P.norm(fp)
    if not fpn:
        return False
    for t in e.tokens[:-1] if len(e.tokens) > 1 else e.tokens:   # 末尾是整串连写，不算
        if re.fullmatch(r'(lqfp|tqfp|qfp|qfn|dfn|bga|tssop|ssop|msop|sop|soic|sot|sod|dip|smd)\d+[a-z0-9]*', t) \
                and t in fpn:
            return True
    m = re.search(r'(?<!\d)(\d{4})(?!\d)', fp)
    return bool(m and re.search(rf'(?<!\d){m.group(1)}(?!\d)', e.spec))


def _fmt_r(ohms):
    if ohms >= 1e6 and ohms % 1e5 == 0:
        return f'{ohms/1e6:g}M'
    if ohms >= 1000:
        return f'{ohms/1000:g}K'
    return f'{ohms:g}R'
