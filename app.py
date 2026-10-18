# -*- coding: utf-8 -*-
"""BOM 匹配助手 — pywebview 壳
界面: gui/index.html (方案A暗色/方案B浅色双皮肤)
引擎: engine.py (bom-matcher skill, Y06/Y12 料号体系, 配物料编码表F27类表格)
"""
import json
import os
import re
import shutil
import subprocess
import sys
import threading
import time
import urllib.error
import urllib.request
import webbrowser
from datetime import datetime

import webview

import dsn_reader
import engine
import excel_bridge
import version
from bom_matcher import learned as LR


def resource_path(rel):
    base = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, rel)


CONFIG_DIR = os.path.join(os.environ.get('APPDATA', os.path.expanduser('~')), 'BomMatcher')
CONFIG_FILE = os.path.join(CONFIG_DIR, 'config.json')
# 编码表存一份副本：桌面上的原文件被挪走/删掉也照样能用
CODE_DIR = os.path.join(CONFIG_DIR, 'code_table')
# 人工选择记忆：「Part 写法 + 封装」→ 料号，下次同样写法优先沿用
LEARNED_FILE = os.path.join(CONFIG_DIR, 'learned.json')


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
        self.final_file = None
        self.final_output = None
        self.output_file = None
        self.session = None
        self.busy = False
        self._lookup_cache = (None, None)    # ((路径, 修改时间), (料号表, 停用集合))

    # ---------- 窗口控制 ----------
    def win_minimize(self):
        self._window.minimize()

    def win_close(self):
        self._window.destroy()

    # ---------- 检查更新（仓库公开，匿名对比 GitHub main 分支） ----------
    def _gh_json(self, url):
        req = urllib.request.Request(url, headers={
            'User-Agent': 'BomMatcher-UpdateCheck',
            'Accept': 'application/vnd.github+json'})
        with urllib.request.urlopen(req, timeout=6) as r:
            return json.loads(r.read().decode('utf-8'))

    def check_update(self):
        """对比本机 exe 打包时的提交与远端 main 最新提交。
        返回 has_update / behind_by（落后几个提交）/ 当前版本号等。"""
        info = version.get_local()
        result = {
            'current': info['version'], 'sha': info['sha'], 'date': info['date'],
            'has_update': False, 'behind_by': 0, 'latest_sha': '',
            'url': version.REPO_URL, 'error': None,
        }
        sha_full = info.get('sha_full')
        if not sha_full:
            result['error'] = '无法确定当前版本'
            return result
        try:
            latest = self._gh_json(
                f'https://api.github.com/repos/{version.REPO}/commits/main')
            latest_sha = latest.get('sha', '')
            result['latest_sha'] = latest_sha[:7]
            if latest_sha and latest_sha != sha_full:
                try:
                    cmp = self._gh_json(
                        f'https://api.github.com/repos/{version.REPO}'
                        f'/compare/{sha_full}...{latest_sha}')
                    if cmp.get('status') in ('ahead', 'diverged'):
                        result['has_update'] = True
                        result['behind_by'] = cmp.get('ahead_by', 0)
                except Exception:
                    # 无法比对（多半是本机提交未推送），只要 sha 不同就提示有更新
                    result['has_update'] = True
        except urllib.error.URLError:
            result['error'] = '网络不可用，暂时无法检查更新'
        except Exception as e:
            result['error'] = f'检查失败: {type(e).__name__}'
        return result

    def open_repo(self):
        webbrowser.open(version.REPO_URL)
        return True

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

    def _pick_excel(self, allow_dsn=False):
        types = ('Excel 文件 (*.xlsx;*.xls)', '所有文件 (*.*)')
        if allow_dsn:
            types = ('BOM Excel 或 OrCAD 原理图 (*.xlsx;*.xls;*.dsn)',) + types
        r = self._window.create_file_dialog(
            webview.OPEN_DIALOG, allow_multiple=False, file_types=types)
        if r:
            return r[0] if isinstance(r, (list, tuple)) else r
        return None

    def _default_output(self):
        stamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        base = os.path.dirname(self.bom_file) if self.bom_file else os.path.expanduser('~')
        return os.path.join(base, f'BOM_匹配结果_{stamp}.xlsx')

    def pick_bom(self):
        p = self._pick_excel(allow_dsn=True)
        if not p:
            return None
        return self.set_file('bom', p)

    def pick_code(self):
        p = self._pick_excel()
        if not p:
            return None
        return self.set_file('code', p)

    def pick_final(self):
        p = self._pick_excel()
        if not p:
            return None
        return self.set_file('final', p)

    def set_file(self, kind, path):
        """kind: bom|code|final。拖放与对话框共用入口。"""
        if not path or not os.path.isfile(path):
            return {'error': '文件不存在'}
        if kind == 'bom' and path.lower().endswith('.dsn'):
            # 拖入时先读一遍：OrCAD 缺失/打不开/没编号当场报，不等到点开始
            try:
                parts = dsn_reader.read_parts(path)
            except dsn_reader.DsnError as e:
                return {'error': str(e)}
            n_rows = len(dsn_reader.group_rows(parts))
            self.bom_file = path
            self.output_file = self._default_output()
            info = self._file_info(path)
            info['meta'] = f'OrCAD 原理图 · {len(parts)} 个元件 → {n_rows} 行'
            return info
        if not path.lower().endswith(('.xlsx', '.xls')):
            return {'error': '请选择 Excel 文件 (.xlsx/.xls)' +
                    ('，或 OrCAD 原理图 (.DSN)' if kind == 'bom' else '')}
        if kind == 'bom':
            self.bom_file = path
            # 输出固定跟 BOM 源文件同路径（UI 已去掉输出文件框）
            self.output_file = self._default_output()
        elif kind == 'final':
            self.final_file = path
        else:
            return self._remember_code(path)
        return self._file_info(path)

    # ---------- 编码表记忆：拖一次存副本，下次启动自动载入；拖新的就替换 ----------
    def _code_lookup(self, path):
        key = (path, os.path.getmtime(path))
        if self._lookup_cache[0] != key:
            self._lookup_cache = (key, engine.load_code_lookup(path))
        return self._lookup_cache[1]

    def _remember_code(self, path):
        try:
            lookup, _stopped = self._code_lookup(path)
        except Exception as e:
            return {'error': f'这个文件读不成编码表：{e}'}
        try:
            os.makedirs(CODE_DIR, exist_ok=True)
            copy = os.path.join(CODE_DIR, os.path.basename(path))
            if os.path.abspath(copy) != os.path.abspath(path):
                for old in os.listdir(CODE_DIR):
                    os.remove(os.path.join(CODE_DIR, old))
                shutil.copy2(path, copy)
            cfg = _load_config()
            cfg['code_table'] = {'source': path, 'copy': copy,
                                 'source_mtime': os.path.getmtime(path),
                                 'saved': datetime.now().strftime('%Y-%m-%d %H:%M')}
            _save_config(cfg)
            self.code_file = copy
        except OSError:
            self.code_file = path    # 存副本失败也不耽误这次用
        info = self._file_info(path)
        info['meta'] = f'{len(lookup)} 条料号 · 已记住，下次自动载入'
        return info

    def restore_code(self):
        """启动时载入上次的编码表。原文件还在且改过 → 用新的重新记一次。"""
        ct = _load_config().get('code_table') or {}
        src, copy = ct.get('source'), ct.get('copy')
        if src and os.path.isfile(src) and os.path.getmtime(src) != ct.get('source_mtime'):
            info = self._remember_code(src)
            if not info.get('error'):
                info['meta'] = '原文件有更新，已重新载入 · ' + info['meta'].split(' · ')[0]
                return info
        if copy and os.path.isfile(copy):
            self.code_file = copy
            info = self._file_info(copy)
            info['meta'] = f'上次的编码表 · {ct.get("saved", "")} 载入 · 变了就拖新的'
            return info
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
                log_callback=log_cb, progress_callback=prog_cb,
                learned=LR.load(LEARNED_FILE))
            # 引擎导出时会按料号原地重排序 results —— 复核行必须持有行对象引用
            review_rows = [r for r in res['results'] if r[4] == '未匹配']
            res['review_rows'] = review_rows
            res['dtype_map'] = dtype_map
            res['confirmed'] = set()
            self.session = res
            # 结果表直接接进生产版区：Excel 里审完保存，回来点「生成生产版」即可，不用再拖
            self.final_file = str(res['output_file'])
            final_info = self._file_info(self.final_file)
            final_info['meta'] = '本次匹配结果 · 在 Excel 里审完保存后点下面按钮'
            matched = sum(1 for r in res['results']
                          if r[4] != '未匹配' and r[0] != 'AUTO')
            total = row_n[0]
            self._push({'type': 'done', 'output_file': str(res['output_file']),
                        'final': final_info,
                        'total': total, 'matched': matched,
                        'unmatched': len(review_rows), 'cat_stats': cat_stats})
            self._push({'type': 'elapsed', 'seconds': round(time.time() - t0, 2)})
        except dsn_reader.DsnError as e:
            self._push({'type': 'error', 'message': str(e)})
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
        store = LR.load(LEARNED_FILE)
        board = os.path.splitext(os.path.basename(self.bom_file or ''))[0]
        if LR.remember(store, r[1], r[2], code, name, spec, board):
            try:
                LR.save(LEARNED_FILE, store)
            except OSError:
                pass
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
            if not path or zone not in ('bom', 'code', 'final'):
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

    # ---------- 生产版（审核后清理导出） ----------
    def make_final(self, mode='check'):
        """把审核后结果 Excel 清理聚合成 <原名>_生产版.xlsx（同路径）。
        不依赖匹配会话，历史结果文件也可用；不自动打开，由界面提示完成状态。
        结果表还在 Excel 里开着时直接问 Excel 存没存（不让用户自己"确认"）：
          mode='check' 已保存/没开着→生成；有改动→回 need=dirty；问不到→回 need=unknown
          mode='save'  替用户在 Excel 里保存后生成
          mode='force' 仅用于问不到的情况（WPS/残留锁文件），用户声明没开着"""
        from bom_matcher.finalize import finalize
        if not self.final_file:
            return {'error': '请先拖入审核后的匹配结果文件'}
        name = os.path.basename(self.final_file)
        st = excel_bridge.status(self.final_file)
        if st == excel_bridge.DIRTY:
            if mode != 'save':
                return {'need': 'dirty', 'file': name}
            err = excel_bridge.save(self.final_file)
            if err:
                return {'error': err}
        elif st == excel_bridge.UNKNOWN and mode != 'force':
            return {'need': 'unknown', 'file': name}
        logs = []
        lookup, stopped = {}, set()
        if self.code_file and os.path.isfile(self.code_file):
            try:
                lookup, stopped = self._code_lookup(self.code_file)
            except Exception:
                logs.append('警告: 编码表读取失败，手填的料号没法自动补名称/规格')
        else:
            logs.append('警告: 没载入编码表，手填的料号没法自动补名称/规格')
        try:
            out = finalize(self.final_file, log=logs.append,
                           code_lookup=lookup, stopped=stopped)
        except PermissionError:
            return {'error': '写入失败：生产版文件正被 Excel 占用，请先关闭它'}
        except ValueError as e:
            return {'error': str(e)}
        except Exception as e:
            return {'error': f'生成失败: {type(e).__name__}: {e}'}
        n_learned = self._learn_from_result(self.final_file, lookup)
        if n_learned:
            logs.append(f'记住了 {n_learned} 条人工选择，下次同样的写法直接配上')
        warn = '\n'.join(m for m in logs if m.startswith('警告'))
        info = '；'.join(m for m in logs if '补齐' in m or '记住了' in m)
        self.final_output = out
        return {'ok': True, 'output': out, 'name': os.path.basename(out),
                'warn': warn, 'info': info}

    def _learn_from_result(self, result_file, lookup):
        """Excel 里改过料号的行记进人工选择；料号必须在编码表里（敲错的不记）。"""
        from bom_matcher.finalize import collect_manual_picks
        try:
            picks = collect_manual_picks(result_file)
        except Exception:
            return 0
        store, n = LR.load(LEARNED_FILE), 0
        board = os.path.basename(os.path.dirname(os.path.abspath(result_file)))
        for part, fp, pn, _ref in picks:
            rec = lookup.get(pn.upper())
            if rec and LR.remember(store, part, fp, pn.upper(), rec[0], rec[1], board):
                n += 1
        if n:
            try:
                LR.save(LEARNED_FILE, store)
            except OSError:
                return 0
        return n

    # ---------- 人工选择记忆管理 ----------
    def list_learned(self):
        store = LR.load(LEARNED_FILE)
        items = [dict(v, key=k) for k, v in store.items()]
        items.sort(key=lambda x: x.get('date', ''), reverse=True)
        return items

    def forget_learned(self, key):
        store = LR.load(LEARNED_FILE)
        if store.pop(key, None) is None:
            return {'error': '这条已经不在了'}
        LR.save(LEARNED_FILE, store)
        return {'ok': True, 'count': len(store)}

    def learned_count(self):
        return len(LR.load(LEARNED_FILE))

    def open_final(self):
        if self.final_output and os.path.isfile(self.final_output):
            os.startfile(self.final_output)
            return True
        return False

    def open_final_folder(self):
        if self.final_output and os.path.isfile(self.final_output):
            subprocess.Popen(['explorer', '/select,', self.final_output])
            return True
        return False

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
