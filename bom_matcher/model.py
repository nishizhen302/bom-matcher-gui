# -*- coding: utf-8 -*-
"""数据结构：BOM行(Item)、编码表条目(Entry)、匹配结果(Match)。"""
from dataclasses import dataclass, field


# 行类别
KIND_BLANK = 'blank'      # 空行，直接丢弃
KIND_MECH = 'mech'        # 非采购项：安装孔/测试点/过孔
KIND_NORMAL = 'normal'    # 正常物料行

# 器件类别
CAT_R, CAT_C, CAT_L, CAT_CONN, CAT_OTHER = 'R', 'C', 'L', 'CONN', 'OTHER'


@dataclass
class Item:
    """一条BOM行（结构化后）"""
    refdes: str
    part: str            # 原始 Part Value
    fp: str              # 原始 Footprint
    qty: str
    kind: str = KIND_NORMAL
    is_nc: bool = False
    body: str = ''       # 去掉NC标记后的有效内容
    category: str = CAT_OTHER
    package: str = ''    # 归一化封装 '0402'/'sot23'/'AE660'/...
    dims: tuple = None   # footprint 中的尺寸 (mm)


@dataclass
class Entry:
    """一条编码表条目（结构化后）"""
    pn: str
    name: str
    spec: str
    category: str = CAT_OTHER
    discontinued: bool = False
    # 解析出的字段（按类别取用）
    r_ohms: float = None
    c_pf: int = None
    l_uh: float = None
    volt: float = None
    tol: float = None
    package: str = ''
    dims: tuple = None
    is_array: bool = False
    tokens: tuple = ()   # 规格里的归一化token（型号匹配用）
    norm: str = ''       # 规格整体归一化文本


@dataclass
class Match:
    pn: str
    name: str
    spec: str
    reason: str = ''             # 人读的匹配依据
    confidence: str = 'exact'    # exact / compat / review
    candidates: list = field(default_factory=list)  # [(pn, name, spec), ...]
