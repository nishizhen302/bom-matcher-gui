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

MECH_FP_KEYS = ('mtg', 'hole', 'via')
MECH_PARTS = {'ORIENTATION', 'TP', 'TEST POINT', 'TESTPOINT'}


class Engine:
    def __init__(self, entries, extra_alias=None, banned=None, debug=lambda *a: None):
        banned = banned or set()
        self.entries = [e for e in entries if e.pn not in banned]
        self.by_pn = {e.pn: e for e in self.entries}
        self.by_cat = {}
        for e in self.entries:
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
        fl = fp.lower()
        if (any(k in fl for k in MECH_FP_KEYS) or fl.startswith('tp_') or fl == 'tp'
                or part.strip().upper() in MECH_PARTS):
            it.kind = KIND_MECH
            return it
        it.is_nc = P.is_nc(part)
        it.body = P.strip_nc(part) if it.is_nc else part.strip()
        it.category = self._detect_category(it.body, fp)
        it.package, it.dims = P.parse_package(fp)
        return it

    def _detect_category(self, body, fp):
        u = body.upper()
        fl = fp.lower()
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
        ohms, count = int(round(parsed[0])), parsed[1]
        pool = self.by_cat.get(CAT_R, [])
        want_array = count > 1

        def hit(e, need_array):
            if e.r_ohms is None or int(round(e.r_ohms)) != ohms:
                return False
            if e.is_array != need_array:
                return False
            if it.package and e.package and e.package != it.package:
                return False
            if it.package and not e.package:
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
        reason = f'阻值{_fmt_r(ohms)}+封装{it.package or "任意"}' + \
                 (f'+阵列x{count}' if want_array and e.is_array else '') + note
        return Match(e.pn, e.name, e.spec, reason=reason,
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
            return self._match_c_aluminum(it, pf, volt)
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

    def _match_c_aluminum(self, it, pf, volt):
        diam = int(it.package[2:]) / 100.0
        best, best_score = None, 1e9
        for e in self.by_cat.get(CAT_C, []):
            if e.c_pf is None or e.dims is None:
                continue
            if abs(e.dims[0] - diam) > 0.6:
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
            return Match(best.pn, best.name, best.spec,
                         reason=f'铝电解:直径{diam:g}mm+容值近似', confidence='compat')
        return None

    # ---------- 电感/磁珠 ----------

    def _match_l(self, it):
        if P.is_bead(it.body):
            return self._match_bead(it)
        uh = P.parse_l_value(it.body)
        if uh is None:
            return None
        series = P.parse_series(it.body)
        pool = [e for e in self.by_cat.get(CAT_L, [])
                if e.l_uh is not None and abs(e.l_uh - uh) < 0.01]
        if it.package in P.CHIP_SIZES:
            pool = [e for e in pool if not e.package or e.package == it.package]
        if not pool:
            return None

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
        m = re.search(r'(\d+)\s*R', it.body.upper())
        if m and abs(int(m.group(1)) - 100) > 15:
            conf = 'review'
            reason += f';阻抗{m.group(1)}Ω与库存100Ω不同,请确认'
        return Match(e.pn, e.name, e.spec, reason=reason, confidence=conf)

    # ---------- 连接器 ----------

    def _match_conn(self, it):
        pools = [self.by_cat.get(CAT_CONN, []), self.by_cat.get(CAT_OTHER, [])]
        pj = P.norm(it.body)
        # 1) 整型号token匹配
        if len(pj) >= 4:
            for pool in pools:
                for e in pool:
                    if any(t == pj or t.startswith(pj) for t in e.tokens):
                        return Match(e.pn, e.name, e.spec, reason=f'型号{it.body}匹配')
        # 2) HEADER针数
        ctype = P.connector_type(it.body)
        if ctype == 'HEADER':
            pins = P.header_pins(it.body, it.fp)
            if pins:
                rows, cols = pins
                total = str(int(rows) * int(cols))
                needles = [f'{rows}*{cols}', f'{cols}*{rows}', f'{rows}X{cols}',
                           f'{cols}X{rows}', f'{total}P', f'{total}PIN']
                for pool in pools:
                    for e in pool:
                        su = e.spec.upper()
                        if any(n in su for n in needles):
                            return Match(e.pn, e.name, e.spec,
                                         reason=f'HEADER {rows}x{cols}={total}针')
        # 3) 类型关键词
        if ctype in ('DVI', 'USB', 'HDMI', 'DP', 'RJ45', 'VGA', 'ETHERNET'):
            for pool in pools:
                for e in pool:
                    if ctype in e.spec.upper():
                        return Match(e.pn, e.name, e.spec, reason=f'类型{ctype}关键词',
                                     confidence='compat')
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
        # 通用型号token匹配
        pj = P.norm(it.body)
        if len(pj) >= 3:
            for e in pool + self.by_cat.get(CAT_CONN, []):
                if not self._pkg_ok(it, e):
                    continue
                if any(t == pj or (len(pj) >= 4 and t.startswith(pj)) for t in e.tokens):
                    return Match(e.pn, e.name, e.spec, reason=f'型号{it.body}匹配')
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
        hits = [e for e in self.by_cat.get(CAT_OTHER, [])
                if e.pn.startswith('Y0611') and needle in e.spec]
        if not hits:
            return None
        # 封装命中的优先；一旦有封装命中，其他封装的就不再作为候选干扰
        pkg_hits = [e for e in hits if fpn and fpn in e.norm]
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
        volt = P.parse_volt(body_u)
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
        mb = re.search(r'\d+', body_u[len(prefix):])
        bom_num = mb.group() if mb else ''
        if len(bom_num) < 2:
            return None
        best, best_common = None, 1
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
                if a != b:
                    break
                common += 1
            if common >= 2 and common > best_common:
                best, best_common = e, common
        if best:
            return Match(best.pn, best.name, best.spec, confidence='compat',
                         reason=f'IC前缀{prefix}+数字段{bom_num[:best_common]}近似,请确认')
        return None


def _fmt_r(ohms):
    if ohms >= 1e6 and ohms % 1e5 == 0:
        return f'{ohms/1e6:g}M'
    if ohms >= 1000:
        return f'{ohms/1000:g}K'
    return f'{ohms:g}R'
