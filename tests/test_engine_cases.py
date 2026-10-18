# -*- coding: utf-8 -*-
"""引擎误判回归：每个用例对应一个真实板子上踩到的坑（编码表条目按 F27.5 原文摘录）。"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from bom_matcher.engine import Engine
from bom_matcher.loader import _build_entry
from bom_matcher.model import CAT_L, CAT_OTHER

CODE = [
    ('Y0612-00114', '集成电路', 'LT87121，VBO转DP，QFN64，龙迅'),
    ('Y0612-00140', '集成电路', 'LT8712SX, DP克隆芯片, QFN88, 龙迅'),
    ('Y0612-00027', '集成电路', 'RTD2796-CG BGA_32X32 PIP/PBP，REALTEK'),
    ('Y0612-00106', '集成电路', 'RTD2785T-CG,LQFP216,REALTEK'),
    ('Y0612-00108', '集成电路', 'RTD2775YT,LQFP216,REALTEK'),
    ('Y0612-00133', '集成电路', 'RTD2735-CG,REALTEK'),
    ('Y0611-00003', '石英晶振', 'HC-49S(DIP)，14.3181MHz,20ppM,20PF,鸿星'),
    ('Y0611-00004', '石英晶振', 'SMD3225,14.318MHz,20ppM,20pF,30ohm,晶峰'),
    ('Y0604-00010', '贴片电感', 'SMPA0503-4R7MT，4.7uH,4.5A,5.4X5.2X2.8mm,天诚'),
    ('Y0611-00011', '贴片晶振', '25MHz,20ppm,20PF,30ohm,SMD3225,晶峰'),
    ('Y0611-00016', '贴片晶振', 'XL2EL89COI-111YLC,SMD3225,24.576MHz,10ppM,12pF,YXC'),
    ('Y0601-00001', '贴片电阻', 'RL1206FR-070R039L,0R039,±1%,1/4W,1206，国巨'),
    ('Y0601-00177', '贴片电阻', 'RL1206FR-070R15L,0R15,±1%,1/4W,1206,国巨'),
    ('Y0608-00009', '二极管', 'B340LB-13-F，40V,3A,150℃ ,SMB/DO-214AA,DIODES'),
    ('Y0616-00187', 'PCB', '32010客户转接板光板，100*100mm'),
    ('Y0616-00099', 'PCB', 'RTD2719AD板光板，200mm*130mm'),
    ('Y0612-00064', '集成电路', 'DRAM NT5CC64M16GP-DI 96-Ball BGA 64Mx16 DDR3'),
    ('Y0612-00075', '集成电路', 'MT41K128M16JT-125IT:K,DDR3,1 x16'),
    ('Y0614-00027', '连接器', 'B2B-PH-K-S，2P,2.0mm，插件立式，JST'),
    ('Y0614-00035', '连接器', 'B2B-PH-K-S，2P，2.0mm，插件立式，JST'),
    ('Y0612-00105', '集成电路', 'SGM61431,36V,3A,SOIC-8N,SGMICRO'),
    ('Y0601-00070', '贴片电阻', 'RC0603FR-070RL,0R,±1%,1/10W,0603,国巨'),
    ('Y0601-00093', '贴片电阻', 'RC0402FR-07100RL,100R,±1%,1/16W,0402,国巨'),
    ('Y0601-00165', '贴片排阻', 'YC124-JR-07100RL，100R，±5%，1/16W,0402X4，国巨'),
    ('Y0601-00006', '贴片电阻', 'RC0603FR-071KL,1K,±1%,1/10W,0603,国巨'),
    ('Y0602-00068', '贴片电容', 'CL05B104KA5NNNC，100nF,25V,±10%,X7R,0402,三星'),
    ('Y0614-00066', '连接器', 'A2006WV-2x6P,12P,2.0mm,插件立式,长江'),
    ('Y0601-00019', '贴片电阻', 'RC0402FR-07240RL,240R,±1%,1/16W,0402,国巨'),
    ('Y0601-00173', '贴片电阻', 'RC0603FR-074K99L,4K99,±1%,1/10W, 0603，国巨'),
    ('Y0601-00005', '贴片电阻', 'RC0603FR-074K7L,4K7,±1%,1/10W,0603,国巨'),
    ('Y0606-00006', '磁珠', 'BLM18EG101TN1D,100Ω@100MHz，2A，0603，Murata'),
    ('Y0604-00035', '贴片电感', 'FXL0530-1R0-M,1uH,8.5A,5.4X5.2X2.8mm,长江微电'),
    ('Y0612-00050', '集成电路', 'LD1117S12TR,1.2V,950mA,1.1V(Dropout),SOT-223,ST'),
    ('Y0608-00023', '二极管', '1N4148WS,75V,0.15A,150℃,SOD-323,长电'),
    ('Y0602-00026', '电解电容', 'UWT1H221MNL1GS-1010 ,220uF,220mA,20%,50V,10X10mm,NICHICON'),
    ('Y0614-00028', '连接器', 'B12B-PHDSS，2*6p,2.0mm,双排插件立式，JST'),
    ('Y0614-00029', '连接器', 'S6B-PH-K-S，6P，2.0mm，插件卧式，JST'),
    ('Y0614-00053', '连接器', 'B6B-PH-K-S,6P，2.0mm，插件立式，JST'),
]


def _match(ref, part, fp):
    eng = Engine([_build_entry(*c) for c in CODE])
    it = eng.build_item(ref, part, fp, '1')
    return it, eng.match(it)


def test_ic_footprint_starting_with_l_is_not_inductor():
    # VBO_EDP U3：封装 LT87121_QFN64_EPAD 以 l 开头，曾被当电感→未匹配，近似候选还混进 LT8712SX
    it, m = _match('U3', 'LT87121', 'LT87121_QFN64_EPAD')
    assert it.category == CAT_OTHER
    assert m.pn == 'Y0612-00114' and m.confidence == 'exact' and not m.candidates


def test_inductor_still_by_footprint():
    it, m = _match('L1', '4.7uH', 'L0603')
    assert it.category == CAT_L


def test_wildcard_ic_ties_flagged_not_silently_picked():
    # RTD2785 板 U5 'RTD27X5'：曾被配成 RTD2796-CG(BGA)，X 应当通配，并列时封装对得上的排前+需确认
    it, m = _match('U5', 'RTD27X5', 'LQFP216_24X24_DB')
    assert m.pn == 'Y0612-00106' and m.confidence == 'review'
    assert [c[0] for c in m.candidates] == ['Y0612-00108', 'Y0612-00133']


def test_crystal_size_from_footprint():
    # 2796 板 Y1：封装 CRY-D3225，曾首选插件 HC-49S，应认出 3225 尺寸选贴片
    it, m = _match('Y1', '14.318MHz-20ppm', 'CRY-D3225')
    assert m.pn == 'Y0611-00004' and m.confidence != 'review'


def test_crystal_frequency_compared_by_value():
    # 25MHz 曾因规格里的 'SMD3225' 配成 14.318MHz；F30.2 下 12MHz 曾因 '12pF' 配成 24.576MHz
    assert _match('X1', '25MHz', 'SMD3225')[1].pn == 'Y0611-00011'
    assert _match('Y1', '12MHz', 'SMD3225')[1] is None


def test_low_ohm_resistor_not_rounded():
    # 0.15Ω 曾因取整成 0 配成 0.039Ω
    assert _match('R14', '0.15 OHM +-1% 1/4W', 'R1206')[1].pn == 'Y0601-00177'
    assert _match('R9', '0.039 OHM +-1% 1/4W', 'R1206')[1].pn == 'Y0601-00001'


def test_generic_words_and_pcb_not_matched():
    assert _match('D1', 'DIODE', 'TO252-DPAK')[1] is None        # 曾配成 B340LB
    assert _match('CN4', '2.00mm', '2_0X4P')[1] is None          # 曾配成 RTD2719AD 光板
    assert _match('U9', '100', 'X')[1] is None                   # 曾配成 100*100mm 光板


def test_multiple_hits_flagged():
    _, m = _match('UD1', 'DDR3', 'DDR3')
    assert m.confidence == 'review' and [c[0] for c in m.candidates] == ['Y0612-00075']


def test_banned_duplicate_and_near_ic_flagged():
    assert _match('CN19', 'B2B-PH-K-S', 'B2B-PH-K-S')[1].pn == 'Y0614-00035'
    assert _match('U4', 'SGM61430', 'SOIC8_EPAD')[1].confidence == 'review'


def test_category_by_refdes_and_metric_package():
    # RESC1608X60N 是公制=0603，曾因认不出封装 191 行电阻全未配
    it, m = _match('R11', '0R', 'RESC1608X60N')
    assert it.package == '0603' and m.pn == 'Y0601-00070'
    # PR0402 封装 P 开头，曾被当成"其他"
    it, m = _match('R5', '1K', 'PR0603')
    assert it.category == 'R' and m.pn == 'Y0601-00006'
    # 位号 C、封装 PC0402
    assert _match('C3', '100nF', 'PC0402')[0].category == 'C'


def test_resistor_array_from_footprint():
    # RDP 位号值写 100R、封装 RP4_0402，曾被当成单颗电阻
    assert _match('RDP1', '100R', 'RP4_0402')[1].pn == 'Y0601-00165'


def test_unknown_resistor_package_needs_review():
    _, m = _match('R1', '1K', '')
    assert m.pn == 'Y0601-00006' and m.confidence == 'review'
    _, m = _match('R267', '1K', 'R_4')
    assert m.confidence == 'review'


def test_testpoint_and_mark_are_not_purchased():
    from bom_matcher.model import KIND_MECH
    assert _match('TP1', 'VCC_3_3', 'TP-R-0_66')[0].kind == KIND_MECH
    assert _match('P1', 'HEADER 2', 'MARK100D300_ALL')[0].kind == KIND_MECH


def test_pin_count_is_not_a_model():
    # CN '2x6P' 封装是 0.5mm 30P FPC，曾被配成 2.0mm 12P 排针
    assert _match('CN59', '2x6P', 'con-FPC30P_50_2125x395x100')[1] is None


def test_resistor_value_notations():
    assert _match('RD13', '240R,1%', 'R0402')[1].pn == 'Y0601-00019'       # 逗号
    assert _match('R82', '4.99K1%', 'R0603')[1].pn == 'Y0601-00173'        # 少分隔符
    assert _match('R8', '4.7K', 'RESC1608X60N')[1].pn == 'Y0601-00005'   # 繁体系统的 Ω 乱码
    assert _match('R9', '4.7KΩ', 'R0603')[1].pn == 'Y0601-00005'


def test_bead_impedance_and_refdes():
    _, m = _match('FB1', '470Ohm/0603', 'L0603')        # FB 位号、值只写阻抗
    assert m.pn == 'Y0606-00006' and m.confidence == 'review'
    _, m = _match('L4', 'BLM21PG300SN1D', 'L0603')     # BLM 型号解出 30Ω
    assert m.confidence == 'review' and '30Ω' in m.reason


def test_inductor_current_and_underscore():
    _, m = _match('L1', '1uH_10A', 'ind_smd2_540x520x280')   # '_10A' 曾让 uH 认不出；8.5A 不够 10A
    assert m.pn == 'Y0604-00035' and m.confidence == 'review'


def test_ldo_voltage_and_model_suffix():
    assert _match('U19', 'AMS1117-1.2', 'SOT223')[1].pn == 'Y0612-00050'
    assert _match('D10', '1N4148/15AHEQ003B', 'SOD2513X90N')[1].pn == 'Y0608-00023'
    assert _match('U1', 'XYZ123/QFN76', 'QFN76')[1] is None    # 拆出的 QFN76 是封装，不拿去搜


def test_large_aluminum_diameter():
    # CAPAE1040X1000N 直径 10.4mm，曾被读成 1.04mm
    assert _match('C56', '220uF/50V', 'CAPAE1040X1000N')[1].pn == 'Y0602-00026'


def test_header_single_row_pitch_and_orientation():
    # 'HEADER 6' 曾配成双排 2*6=12 针，修完又配到卧式；封装 HDRV=立式
    assert _match('CN18', 'HEADER 6', 'HDRV6W50P200_1X6_1400X450X600')[1].pn == 'Y0614-00053'
