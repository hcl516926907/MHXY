"""
One-time migration script: move flat YYYYMMDD folders to YYYY/MM/DD hierarchy.

Before: data/20260402/  analysis/20260402/
After:  data/2026/04/02/  analysis/2026/04/02/
"""
import os
import re
import shutil

ROOT = os.path.dirname(os.path.abspath(__file__))

for folder in ("data", "analysis"):
    src_root = os.path.join(ROOT, folder)
    if not os.path.isdir(src_root):
        print(f"  跳过（不存在）：{folder}/")
        continue

    moved = 0
    for name in sorted(os.listdir(src_root)):
        if not re.fullmatch(r"\d{8}", name):
            continue
        src = os.path.join(src_root, name)
        dst = os.path.join(src_root, name[:4], name[4:6], name[6:8])
        if os.path.exists(dst):
            print(f"  [跳过，目标已存在] {folder}/{name} → {name[:4]}/{name[4:6]}/{name[6:8]}/")
            continue
        os.makedirs(os.path.dirname(dst), exist_ok=True)
        shutil.move(src, dst)
        print(f"  {folder}/{name} → {folder}/{name[:4]}/{name[4:6]}/{name[6:8]}/")
        moved += 1
    print(f"  {folder}/: 已迁移 {moved} 个目录")

os.makedirs(os.path.join(ROOT, "data", "archive"), exist_ok=True)
print("\n迁移完成。data/archive/ 已创建。")
