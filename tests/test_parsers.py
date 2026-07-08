# -*- coding: utf-8 -*-
"""解析层单测——每个用例对应一个历史真实踩坑点。"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from bom_matcher import parsers as P


def test_resistor_dash_tolerance():
    # 2026-07-08: "100R-5%" 曾错配到排阻
    assert P.parse_r_value('100R-5%') == (100.0, 1)
    assert P.parse_r_value('3K-5%') == (3000.0, 1)
    assert P.parse_r_value('4K7') == (4700.0, 1)
    assert P.parse_r_value('0R') == (0.0, 1)
    assert P.parse_r_value('100R*4/0402') == (100.0, 4)


def test_capacitor_dash_voltage():
    assert P.parse_c_value('100nF-16V') == (100000, 16.0)
    assert P.parse_c_value('0.1uF/50V') == (100000, 50.0)
    assert P.parse_c_value('10uF/0805') == (10000000, None)  # 0805是封装不是电压
    assert P.parse_c_value('1u/10V/Y') == (1000000, 10.0)


def test_inductor_value_not_fooled_by_model_number():
    # 旧版把 SCRH127-100M 的 "100M" 当成 100mH，漏掉真正的 10uH
    assert P.parse_spec_l('SCRH127-100M-PF-CB,10uH,5.4A,12.5X12.5X8mm,TDK') == 10.0
    assert P.parse_l_value('CR32-10uH') == 10.0
    assert P.parse_series('CR32-10uH') == 'CR32'


def test_bead_dash_100mhz():
    # "90R-100MHz" 旧版只认 "/100MHz" 写法
    assert P.is_bead('90R-100MHz')
    assert P.is_bead('BEAD_0805')
    assert P.is_bead('MPZ2012S101A')
    assert not P.is_bead('100R')


def test_dims_not_confused_with_ratings():
    # 旧版拿封装数字撞规格任意数字（'27'撞'Isat 27A'）
    assert P.parse_dims('SMPA2313-100MT,10uH, Isat 27A, 23.5X22X13mm') == (23.5, 22.0)
    assert P.parse_dims('COIL-SPII27QR-330') is None
    assert P.parse_dims('l-coil-12_2x12_2') == (12.2, 12.2)


def test_spec_resistor_fields():
    ohms, pkg, tol, arr = P.parse_spec_r('RC0402FR-07100RL,100R,±1%,1/16W,0402,国巨')
    assert (ohms, pkg, tol, arr) == (100.0, '0402', 1.0, False)
    ohms, pkg, tol, arr = P.parse_spec_r('YC124-JR-07100RL，100R，±5%，1/16W,0402X4，国巨')
    assert (ohms, pkg, arr) == (100.0, '0402', True)


def test_nc_handling():
    assert P.is_nc('NC/10K') and P.is_nc('10K/NC') and P.is_nc('nc.') and P.is_nc('NC')
    assert not P.is_nc('RE-FUSE/49FKA-002')  # NC不能误伤子串
    assert P.strip_nc('NC/10K') == '10K'
    assert P.strip_nc('nc.') == ''


def test_freq():
    assert P.parse_freq_mhz('14.318MHz-20ppm') == 14.318
    assert P.parse_freq_mhz('24.576M') == 24.576
