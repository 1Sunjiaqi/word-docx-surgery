#!/usr/bin/env python3
"""字节级最小替换：只改动 <w:t> 里的那一段文字，其余字节原样保留。

    python docx_min_edit.py <docx> --find "旧文字" --replace "新文字" [--occurrence N | --all | --dry-run]

和"整份 document.xml 重写"的做法相比：这里只在原始字节流上做一次拼接，
XML 声明、命名空间前缀、属性顺序、其它部件全部不动 —— 所以扰动最小、最容易核对。

已知局限（会如实报错，不会瞎猜）：
  Word 会把一句话拆成多个 run（例如"合同"单独一个 run），此时整句在 XML 里不连续，
  本工具会报"未找到连续文本"，请改用 docx_ledger.py 先看清结构再定位。
"""
import argparse
import os
import shutil
import sys
import zipfile

DOC_PART = "word/document.xml"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("path")
    ap.add_argument("--find", required=True)
    ap.add_argument("--replace", default="")
    ap.add_argument("--occurrence", type=int)
    ap.add_argument("--all", action="store_true")
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--out")
    args = ap.parse_args()
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

    with zipfile.ZipFile(args.path) as z:
        items = [(i, z.read(i.filename)) for i in z.infolist()]
    xml = dict((i.filename, d) for i, d in items)[DOC_PART]

    needle = (">" + args.find + "<").encode("utf-8")
    hits = []
    start = 0
    while True:
        k = xml.find(needle, start)
        if k < 0:
            break
        hits.append(k)
        start = k + 1
    print("命中 %d 处（%s 中连续出现的 %r）" % (len(hits), DOC_PART, args.find))
    if not hits:
        print("未找到连续文本：可能被 Word 拆成多个 run。请先用 docx_ledger.py 看结构。")
        return 1
    if args.dry_run:
        return 0

    if args.all:
        targets = hits
    elif args.occurrence:
        if not 1 <= args.occurrence <= len(hits):
            print("--occurrence 超出范围", file=sys.stderr)
            return 2
        targets = [hits[args.occurrence - 1]]
    else:
        print("必须给 --all / --occurrence N / --dry-run 之一", file=sys.stderr)
        return 2

    new_needle = (">" + args.replace + "<").encode("utf-8")
    for k in reversed(targets):
        xml = xml[:k] + new_needle + xml[k + len(needle):]

    dst = args.out or args.path
    if args.out:
        shutil.copy2(args.path, dst)
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as z:
        for info, data in items:
            if info.filename == DOC_PART:
                data = xml
            z.writestr(info, data)
    print("已替换 %d 处 -> %r（写出 %s）" % (len(targets), args.replace, dst))
    return 0


if __name__ == "__main__":
    sys.exit(main())
