# -*- coding: utf-8 -*-
"""版本信息 —— 供"检查更新"使用。

版本号 = git 提交数（rev-list --count HEAD），单调递增，无需手动维护。
- 打包成 exe 时：由 BomMatcher.spec 在构建前生成 _build_info.py，焊进 exe。
- 源码直接跑（python app.py）时：现场读 git 仓库。
"""
import os
import subprocess

REPO = 'nishizhen302/bom-matcher-gui'
REPO_URL = 'https://github.com/nishizhen302/bom-matcher-gui'

_HERE = os.path.dirname(os.path.abspath(__file__))


def _from_build_info():
    try:
        import _build_info as b
        return {'version': b.VERSION, 'sha': b.SHA, 'sha_full': b.SHA_FULL,
                'date': b.DATE, 'source': 'build'}
    except Exception:
        return None


def _from_git():
    def run(args):
        return subprocess.run(
            ['git'] + args, cwd=_HERE, capture_output=True,
            text=True, timeout=5).stdout.strip()
    try:
        count = run(['rev-list', '--count', 'HEAD'])
        if not count:
            return None
        return {
            'version': count,
            'sha': run(['rev-parse', '--short', 'HEAD']),
            'sha_full': run(['rev-parse', 'HEAD']),
            'date': run(['log', '-1', '--format=%cd', '--date=format:%Y-%m-%d']),
            'source': 'git',
        }
    except Exception:
        return None


def get_local():
    return _from_build_info() or _from_git() or {
        'version': '?', 'sha': '', 'sha_full': '', 'date': '', 'source': 'unknown'}
