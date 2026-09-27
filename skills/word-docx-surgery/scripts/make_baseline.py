#!/usr/bin/env python3
"""开工前给一个目录下的 Word 文件拍一张「只读基线」快照。

    python make_baseline.py --base <项目目录> --out 基线.tsv
    python make_baseline.py --base . --out 基线.tsv --ext .doc,.docx,.docm

输出每行：<mtime 的 epoch 秒> \t <字节数> \t <相对 base 的路径>
之后用 verify_readonly.py --base <base> --tsv <基线.tsv> 复验。

安全设计：**基线文件已存在时拒绝覆盖**（必须 --force）。
这条基线的作用是证明"我没动过原件"，一旦可以随手覆盖它就没有证明力了。
要登记"确实发生的外部改动"，请走 verify_readonly.py --ack，不要去改基线。
"""
import argparse
import os
import sys

DEFAULT_EXT = (".doc", ".docx", ".docm", ".dot", ".dotx", ".dotm")
SKIP_DIRS = {".git", "__pycache__", "node_modules", ".venv", "venv"}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--base", default=".", help="要拍快照的目录（递归）")
    ap.add_argument("--out", required=True, help="写出的基线 TSV")
    ap.add_argument("--ext", default=",".join(DEFAULT_EXT), help="参与统计的扩展名，逗号分隔")
    ap.add_argument("--force", action="store_true", help="允许覆盖已存在的基线文件")
    args = ap.parse_args()
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

    base = os.path.abspath(args.base)
    out = os.path.abspath(args.out)
    if not os.path.isdir(base):
        print("不是目录: %s" % base, file=sys.stderr)
        return 2
    if os.path.exists(out) and not args.force:
        print("基线文件已存在: %s" % out, file=sys.stderr)
        print("拒绝覆盖：基线可以随手覆盖就没有证明力。重新采集请加 --force；"
              "登记外部改动请用 verify_readonly.py --ack。", file=sys.stderr)
        return 2

    exts = tuple(e.strip().lower() for e in args.ext.split(",") if e.strip())
    rows = []
    for root, dirs, files in os.walk(base):
        dirs[:] = [d for d in dirs if d not in SKIP_DIRS]
        for name in files:
            if not name.lower().endswith(exts):
                continue
            path = os.path.join(root, name)
            try:
                st = os.stat(path)
            except OSError:
                continue
            rel = os.path.relpath(path, base).replace("\\", "/")
            rows.append((st.st_mtime, st.st_size, rel))
    rows.sort(key=lambda r: r[2])

    with open(out, "w", encoding="utf-8", newline="\n") as fh:
        for mtime, size, rel in rows:
            fh.write("%.0f\t%d\t%s\n" % (mtime, size, rel))
    print("已写入 %s：%d 个文件（%s）" % (out, len(rows), " ".join(exts)))
    print("复验: python verify_readonly.py --base %s --tsv %s" % (base, out))
    return 0


if __name__ == "__main__":
    sys.exit(main())
