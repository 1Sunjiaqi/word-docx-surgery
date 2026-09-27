#!/usr/bin/env python3
"""只读约束自检：把目录下的 Word 文件与你开工前的基线逐条比对。

    python scripts/verify_readonly.py --base <要守住的目录> --tsv <基线.tsv>
    python scripts/verify_readonly.py --base <目录> --tsv <基线.tsv> --ack <已登记外部改动.tsv>

基线由 make_baseline.py 在开工前生成。默认值按 WORDKIT_BASE / WORDKIT_BASELINE 环境变量取，
再退回"当前目录 + word文件基线_mtime_size.tsv"；换机器用时请显式给参，别依赖默认值。

基线文件每行：<mtime 的 epoch 秒> \t <字节数> \t <相对 base 的路径>
判定：大小必须完全一致；mtime 允许 2 秒内的容差（文件系统精度）。

--ack 登记表每行：<相对路径> \t <期望的新字节数> \t <说明（谁在何时改的、改了什么、证据在哪）>
「#」开头的行与空行忽略。登记过的文件**不篡改基线**，只在比对时按"已登记的外部改动"放行；
如果它的大小和登记里写的新大小也对不上（说明后来又被人动过），照样算失败。

退出码：0 全部一致（或只剩已登记的差异） / 1 有差异。
"""
import argparse
import os
import sys

TOL = 2.0


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--base", default=os.environ.get("WORDKIT_BASE", "."))
    ap.add_argument("--tsv", default=os.environ.get("WORDKIT_BASELINE", "word文件基线_mtime_size.tsv"))
    ap.add_argument("--ack", help="已登记的外部改动 TSV（见文件头说明）")
    args = ap.parse_args()
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

    if not os.path.exists(args.tsv):
        print("读不到基线文件: %s" % args.tsv, file=sys.stderr)
        print("先用 make_baseline.py 在开工前生成基线，或用 --tsv 指定路径。", file=sys.stderr)
        return 1
    rows = []
    with open(args.tsv, encoding="utf-8") as fh:
        for line in fh:
            line = line.rstrip("\n")
            if not line.strip():
                continue
            parts = line.split("\t")
            if len(parts) < 3:
                continue
            rows.append((float(parts[0]), int(parts[1]), "\t".join(parts[2:])))

    ack = {}
    if args.ack:
        if not os.path.exists(args.ack):
            print("读不到登记表: %s" % args.ack, file=sys.stderr)
            return 1
        with open(args.ack, encoding="utf-8") as fh:
            for line in fh:
                line = line.rstrip("\n")
                if not line.strip() or line.lstrip().startswith("#"):
                    continue
                parts = line.split("\t")
                if len(parts) < 3:
                    continue
                ack[parts[0]] = (int(parts[1]), "\t".join(parts[2:]))

    missing, size_bad, mtime_bad, acked = [], [], [], []
    for mtime, size, rel in rows:
        path = os.path.join(args.base, rel)
        if not os.path.exists(path):
            missing.append(rel)
            continue
        st = os.stat(path)
        if rel in ack:
            want_new, why = ack[rel]
            if st.st_size == want_new:
                acked.append((rel, size, st.st_size, why))
            else:
                size_bad.append((rel, want_new, st.st_size))
            continue
        if st.st_size != size:
            size_bad.append((rel, size, st.st_size))
        elif abs(st.st_mtime - mtime) > TOL:
            mtime_bad.append((rel, mtime, st.st_mtime))

    print("基线文件: %s" % args.tsv)
    print("比对范围: %s" % args.base)
    print("共 %d 个 Word 文件" % len(rows))
    for rel, old, new, why in acked:
        print("  已登记的外部改动: %s  %d -> %d 字节" % (rel, old, new))
        print("     %s" % why)
    for rel in missing:
        print("  缺失: %s" % rel)
    for rel, want, got in size_bad:
        print("  大小变了: %s  %d -> %d" % (rel, want, got))
    for rel, want, got in mtime_bad:
        print("  mtime 变了: %s  %.0f -> %.0f（差 %.0f 秒）"
              % (rel, want, got, got - want))
    bad = len(missing) + len(size_bad) + len(mtime_bad)
    if bad:
        print("结论: 有 %d 处不一致 —— 只读约束可能被破坏。" % bad)
        return 1
    if acked:
        print("结论: 除上述 %d 处已登记的外部改动外，其余 %d 个文件与基线一致。"
              % (len(acked), len(rows) - len(acked)))
        return 0
    print("结论: 全部一致，未改动任何既有项目的 Word 文件。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
