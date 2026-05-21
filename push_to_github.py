# -*- coding: utf-8 -*-
"""把本地文件推送到GitHub"""
import subprocess, os
from datetime import datetime

GIT_PATH = r"C:\Program Files\Git\cmd\git.exe"
REPO_DIR = r"C:\Users\Administrator\Desktop\df-sol-dashboard"

def git(args, silent=False):
    result = subprocess.run(
        [GIT_PATH] + args,
        cwd=REPO_DIR, capture_output=True, text=True, encoding='utf-8', errors='replace'
    )
    if not silent:
        if result.stdout.strip(): print('  >', result.stdout.strip()[:300])
        if result.stderr.strip(): print('  >', result.stderr.strip()[:300])
    return result.returncode

def ts():
    return datetime.now().strftime('%H:%M:%S')

print(f"[{ts()}] 对齐分支...")
git(['branch', '-M', 'main'])

print(f"[{ts()}] 拉取远端...")
git(['pull', 'origin', 'main', '--allow-unrelated-histories', '--no-rebase'])

print(f"[{ts()}] 添加文件...")
git(['add', 'index.html', 'data.json', 'sync_data.py', 'extract_data.py', 'push_to_github.py'])

# 读取data.json里的build_time作为commit信息
try:
    import json
    with open(os.path.join(REPO_DIR, 'data.json'), 'r', encoding='utf-8') as f:
        d = json.load(f)
    msg = f"数据更新 update={d.get('update_time','')} build={d.get('build_time','')}"
except:
    msg = f"数据更新 {datetime.now().strftime('%Y-%m-%d %H:%M')}"

print(f"[{ts()}] 提交: {msg}")
git(['commit', '-m', msg])

print(f"[{ts()}] 推送到GitHub...")
ret = git(['push', 'origin', 'main'])

if ret == 0:
    print(f"[{ts()}] [OK] 推送成功！")
    print(f"  访问地址: https://Ziqian0512.github.io/df-sol-dashboard/")
else:
    print(f"[{ts()}] [FAIL] 推送失败")

input("按回车键退出...")
