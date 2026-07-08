# -*- coding: utf-8 -*-
"""BOM 匹配助手 — pywebview 壳
界面: gui/index.html (方案A暗色/方案B浅色双皮肤)
引擎: engine.py (bom-matcher skill, Y06/Y12 料号体系, 配物料编码表F27类表格)
"""
import json
import os
import re
import subprocess
import sys
import threading
import time
from datetime import datetime

import webview

import engine


def resource_path(rel):
    base = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, rel)


CONFIG_DIR = os.path.join(os.environ.get('APPDATA', os.path.expanduser('~')), 'BomMatcher')
CONFIG_FILE = os.path.join(CONFIG_DIR, 'config.json')


def _load_config():
    try:
        with open(CONFIG_FILE, encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return {}


def _save_config(cfg):
    try:
        os.makedirs(CONFIG_DIR, exist_ok=True)
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


class Api:
    def __init__(self):
        self._window = None
        self.bom_file = None
        self.code_file = None
        self.output_file = None
        self.session = None
        self.busy = False

    # ---------- 窗口控制 ----------
    def win_minimize(self):
        self._window.minimize()

    def win_close(self):
        self._window.destroy()

    # ---------- 偏好设置（皮肤等，存 %APPDATA%\BomMatcher，EXE 每次解压路径不同导致
    # 页面 localStorage 不跨启动，必须落盘到固定位置） ----------
    def get_pref(self, key):
        return _load_config().get(key)

    def set_pref(self, key, value):
        cfg = _load_config()
        cfg[key] = value
        _save_config(cfg)
        return True

    # ---------- 文件选择 ----------
    def _file_info(self, path):
        try:
            size = os.path.getsize(path)
        except OSError:
            size = 0
        if size >= 1024 * 1024:
            size_s = f"{size / 1024 / 1024:.1f} MB"
        else:
            size_s = f"{size / 1024:.1f} KB"
        return {'path': path, 'name': os.path.basename(path), 'size': size_s}

    def _pick_excel(self):
        r = self._window.create_file_dialog(
            webview.OPEN_DIALOG, allow_multiple=False,
            file_types=('Excel 文件 (*.xlsx;*.xls)', '所有文件 (*.*)'))
        if r:
            return r[0] if isinstance(r, (list, tuple)) else r
        return None

    def _default_output(self):
        stamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        base = os.path.dirname(self.bom_file) if self.bom_file else os.path.expanduser('~')
        return os.path.join(base, f'BOM_匹配结果_{stamp}.xlsx')

    def pick_bom(self):
        p = self._pick_excel()
        if not p:
            return None
        return self.set_file('bom', p)

    def pick_code(self):
        p = self._pick_excel()
        if not p:
            return None
        return self.set_file('code', p)

    def set_file(self, kind, path):
        """kind: bom|code。拖放与对话框共用入口。"""
        if not path or not os.path.isfile(path):
            return {'error': '文件不存在'}
        if not path.lower().endswith(('.xlsx', '.xls')):
            return {'error': '请选择 Excel 文件 (.xlsx/.xls)'}
        if kind == 'bom':
            self.bom_file = path
            if not self.output_file:
                self.output_file = self._default_output()
        else:
            self.code_file = path
        info = self._file_info(path)
        info['output'] = self.output_file or ''
        return info

    def pick_output(self):
        stamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        r = self._window.create_file_dialog(
            webview.SAVE_DIALOG, save_filename=f'BOM_匹配结果_{stamp}.xlsx',
            file_types=('Excel 文件 (*.xlsx)',))
        if r:
            p = r[0] if isinstance(r, (list, tuple)) else r
            if not p.lower().endswith('.xlsx'):
                p += '.xlsx'
            self.output_file = p
            return p
        return None

    # ---------- 匹配 ----------
    def start_match(self):
        if self.busy:
            return {'error': '正在匹配中'}
        if not self.bom_file or not self.code_file:
            return {'error': '请先选择 BOM 清单和物料编码表'}
        if not self.output_file:
            self.output_file = self._default_output()
        self.busy = True
        self.session = None
        threading.Thread(target=self._worker, daemon=True).start()
        return {'ok': True, 'output': self.output_file}

    def _push(self, ev):
        try:
            payload = json.dumps(ev, ensure_ascii=False)
            self._window.evaluate_js(f'window.engineEvent && window.engineEvent({payload})')
        except Exception:
            pass

    def _worker(self):
        t0 = time.time()
        delay = [0.0]
        bom_total = [0]
        row_n = [0]
        cat_stats = {}
        dtype_map = {}

        def row_cb(ev):
            row_n[0] += 1
            d = cat_stats.setdefault(ev['dtype'], {'total': 0, 'matched': 0})
            d['total'] += 1
            if ev['matched']:
                d['matched'] += 1
            dtype_map[(ev['ref'], ev['part'], ev['fp'])] = ev['dtype']
            self._push({'type': 'row', 'index': row_n[0] - 1,
                        'total': max(bom_total[0], row_n[0]),
                        'ref': ev['ref'], 'part': ev['part'], 'fp': ev['fp'],
                        'dtype': ev['dtype'], 'matched': ev['matched'],
                        'code': ev['pn'] if ev['matched'] else ''})
            if delay[0] > 0:
                time.sleep(delay[0])

        def log_cb(msg):
            m = re.search(r'BOM:\s*(\d+)\s*行', str(msg))
            if m:
                bom_total[0] = max(int(m.group(1)), 1)
                # 流式日志可视化节奏: 总耗时目标 ≤3s
                delay[0] = min(0.02, 3.0 / bom_total[0])
            self._push({'type': 'logmsg', 'message': str(msg).strip()})

        def prog_cb(pct, msg=''):
            self._push({'type': 'progress', 'pct': pct, 'msg': msg})

        engine._set_row_callback(row_cb)
        try:
            res = engine.process_bom(
                self.bom_file, self.code_file, self.output_file,
                log_callback=log_cb, progress_callback=prog_cb)
            # 引擎导出时会按料号原地重排序 results —— 复核行必须持有行对象引用
            review_rows = [r for r in res['results'] if r[4] == '未匹配']
            res['review_rows'] = review_rows
            res['dtype_map'] = dtype_map
            res['confirmed'] = set()
            self.session = res
            matched = sum(1 for r in res['results']
                          if r[4] != '未匹配' and r[0] != 'AUTO')
            total = row_n[0]
            self._push({'type': 'done', 'output_file': str(res['output_file']),
                        'total': total, 'matched': matched,
                        'unmatched': len(review_rows), 'cat_stats': cat_stats})
            self._push({'type': 'elapsed', 'seconds': round(time.time() - t0, 2)})
        except Exception as e:
            self._push({'type': 'error', 'message': f'{type(e).__name__}: {e}'})
        finally:
            engine._set_row_callback(None)
            engine._set_log_callback(None)
            engine._set_progress_callback(None)
            self.busy = False

    # ---------- 人工复核 ----------
    def _row_dtype(self, r):
        return self.session['dtype_map'].get((r[0], r[1], r[2]), '其他')

    def get_unmatched(self):
        """复核行按 review_rows 序号引用（导出会重排 results，行对象引用不受影响）"""
        if not self.session:
            return []
        out = []
        for i, r in enumerate(self.session['review_rows']):
            note = ''
            if r[4] != '未匹配':
                note = '人工指定'
            elif i in self.session['confirmed']:
                note = '人工确认无匹配'
            out.append({'index': i, 'ref': r[0], 'part': r[1], 'fp': r[2],
                        'dtype': self._row_dtype(r), 'note': note})
        # 已人工指定的行不再出现在复核列表
        return [o for o in out if o['note'] != '人工指定']

    def get_candidates(self, row_index):
        """给未匹配行在编码表里找近似候选料号"""
        if not self.session:
            return []
        try:
            r = self.session['review_rows'][int(row_index)]
        except (IndexError, ValueError):
            return []
        part, fp = str(r[1]), str(r[2])
        dtype = self._row_dtype(r)
        code_index = self.session['code_index']
        cats = {'电容': ['电容'], '电阻': ['电阻'], '电感': ['电感'],
                '连接器': ['连接器']}.get(dtype, list(code_index.keys()))

        part_clean = re.sub(r'^NC[/_]?', '', part.strip(), flags=re.I)
        part_norm = re.sub(r'[^a-z0-9]', '', part_clean.lower())
        part_upper = part_clean.upper()
        fp_digits_m = re.search(r'(\d{4})', fp)
        fp_digits = fp_digits_m.group(1) if fp_digits_m else ''
        tokens = [t for t in re.findall(r'[A-Z0-9\.]{2,}', part_upper)
                  if not t.isdigit() or len(t) >= 3]

        scored = []
        for cat in cats:
            for pn, name, spec, in code_index.get(cat, []):
                if '停用' in str(spec):
                    continue
                spec_upper = str(spec).upper() + ' ' + str(name).upper()
                spec_norm = re.sub(r'[^a-z0-9]', '', str(spec).lower())
                score = 0
                hits = []
                if part_norm and len(part_norm) >= 4 and part_norm[:6] in spec_norm:
                    score += 5; hits.append('型号相符')
                tok_hit = sum(1 for t in tokens if t in spec_upper)
                if tok_hit:
                    score += tok_hit
                    hits.append(f'{tok_hit}处关键词')
                if fp_digits and fp_digits in spec_upper:
                    score += 2; hits.append('封装出现')
                if score > 0:
                    scored.append((score, pn, name, spec, hits))
        scored.sort(key=lambda x: -x[0])
        return [{'code': p, 'name': n, 'spec': s, 'hits': h, 'score': sc}
                for sc, p, n, s, h in scored[:8]]

    def apply_override(self, row_index, code, name, spec):
        """人工指定料号: 改行对象 + 重导出 Excel（引擎会重新排序聚合）"""
        if not self.session:
            return {'error': '无匹配会话'}
        try:
            r = self.session['review_rows'][int(row_index)]
        except (IndexError, ValueError):
            return {'error': '行不存在'}
        r[4], r[5], r[6] = code, name, spec
        return self._reexport()

    def mark_confirmed_nomatch(self, row_index):
        """确认无匹配: 仅界面标记（该引擎输出无备注列，Excel 保持黄色未匹配）"""
        if not self.session:
            return {'error': '无匹配会话'}
        self.session['confirmed'].add(int(row_index))
        return {'ok': True}

    def _reexport(self):
        try:
            self.session['export'](self.session['results'])
        except PermissionError:
            return {'error': '写入失败：结果文件正被 Excel 占用，请先关闭它'}
        except Exception as e:
            return {'error': f'导出失败: {e}'}
        res = self.session['results']
        bom_rows = [r for r in res if r[0] != 'AUTO']
        unmatched = sum(1 for r in bom_rows if r[4] == '未匹配')
        cat = {}
        for r in bom_rows:
            d = cat.setdefault(self._row_dtype(r), {'total': 0, 'matched': 0})
            d['total'] += 1
            if r[4] != '未匹配':
                d['matched'] += 1
        return {'ok': True, 'total': len(bom_rows),
                'matched': len(bom_rows) - unmatched,
                'unmatched': unmatched, 'cat_stats': cat}

    # ---------- 拖放（必须 Python 侧绑定 DOM drop，pywebview 才会捕获原生文件路径） ----------
    def _on_drop(self, e):
        try:
            files = e.get('dataTransfer', {}).get('files', [])
            if not files:
                return
            path = files[0].get('pywebviewFullPath')
            # 拖放落点由 JS 在 dragover 时记录（DOM 事件序列化不含可靠的 target 信息）
            zone = self._window.evaluate_js('window.__lastDropZone')
            if not path or zone not in ('bom', 'code'):
                return
            info = self.set_file(zone, path)
            payload = json.dumps(info, ensure_ascii=False)
            if info.get('error'):
                self._window.evaluate_js(f'toast({json.dumps(info["error"], ensure_ascii=False)}, true)')
            else:
                self._window.evaluate_js(f'fileLoaded({json.dumps(zone)}, {payload})')
        except Exception:
            pass

    def bind_dnd(self):
        from webview.dom import DOMEventHandler
        self._window.dom.document.events.drop += DOMEventHandler(
            self._on_drop, prevent_default=True, stop_propagation=False)

    # ---------- 结果文件 ----------
    def open_result(self):
        if self.output_file and os.path.isfile(self.output_file):
            os.startfile(self.output_file)
            return True
        return False

    def open_folder(self):
        if self.output_file and os.path.isfile(self.output_file):
            subprocess.Popen(['explorer', '/select,', self.output_file])
            return True
        return False


def main():
    api = Api()
    window = webview.create_window(
        'BOM 匹配助手',
        resource_path(os.path.join('gui', 'index.html')),
        js_api=api,
        width=980, height=700,
        frameless=True, easy_drag=False,
        resizable=False,
        background_color='#0d1117',
    )
    api._window = window
    window.events.loaded += api.bind_dnd
    webview.start(debug='--debug' in sys.argv)


if __name__ == '__main__':
    main()
