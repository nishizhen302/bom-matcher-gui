# -*- coding: utf-8 -*-
"""人工选择记忆：「Part 写法 + 封装」→ 料号。

人工复核里点「采用」、或在 Excel 里改过料号后生成生产版，都记一笔；
下次任何板子出现同样写法，引擎优先沿用（比内置规则优先）。
料号后来在编码表里停用/删除的不沿用，由引擎记进 skipped 供日志提示。
"""
import json
import os
from datetime import date

from . import parsers as P


def key_of(part, fp):
    """NC 标记不算写法的一部分；大小写、分隔符差异视为同一种写法。"""
    body = P.strip_nc(part) if P.is_nc(part) else str(part or '')
    return P.norm(body) + '|' + P.norm(fp or '')


def load(path):
    try:
        with open(path, encoding='utf-8') as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except (OSError, ValueError):
        return {}


def save(path, store):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp = path + '.tmp'
    with open(tmp, 'w', encoding='utf-8') as f:
        json.dump(store, f, ensure_ascii=False, indent=1)
    os.replace(tmp, path)   # 写一半断电也不会留下坏文件


def remember(store, part, fp, pn, name='', spec='', board=''):
    """记一条；同一写法再次人工指定时以最新为准。返回是否有变化。"""
    k = key_of(part, fp)
    if not k.split('|')[0]:
        return False
    old = store.get(k)
    if old and old.get('pn') == pn:
        return False
    store[k] = {'part': str(part).strip(), 'fp': str(fp or '').strip(), 'pn': pn,
                'name': name, 'spec': spec, 'board': board, 'date': date.today().isoformat()}
    return True
