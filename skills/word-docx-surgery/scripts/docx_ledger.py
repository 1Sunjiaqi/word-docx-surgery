#!/usr/bin/env python3
"""改动闸门：把两份 .docx 的差异摊成一份可核对的清单，并断言"只许改这些地方"。

    python docx_ledger.py <基线.docx> <候选.docx>
    python docx_ledger.py <基线.docx> <候选.docx> --allow "TBL#16" --allow "TBL#26"
    python docx_ledger.py <基线.docx> <候选.docx> --max-changes 30 --json 清单.json

两种模式：
  * 报告模式（不给 --allow）：把所有改动列出来，始终返回 0；
  * 闸门模式（给了 --allow）：位置不匹配白名单的改动算违例，有违例返回 1。

退出码：0 通过 / 1 有违例 / 2 用法或读取失败。

白名单用正则匹配"改动位置"，例如：
  --allow "TBL#16"                  只第 16 张表
  --allow "TBL#1[6-9]|TBL#2[0-6]"   第 16~19 与 20~26 张表
  --allow "word/media/image1[3-9]"  只允许新增这些图片
  --allow "COMMENT"                 允许批注变化
"""
import argparse
import difflib
import hashlib
import json
import os
import re
import sys
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import docx_textconv as tc  # noqa: E402

BLOCK_RE = re.compile(r"^(#\d{4})(?: (TBL#\d+))?")
CELL_RE = re.compile(r"^(r\d{2}c\d{2})")
# 正文 XML 的体积变化是"正文改动"的副产品，默认不计入闸门（用 --strict-parts 打开）
GATE_IGNORED_PARTS = {"word/document.xml"}


def sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def _digest(data):
    return hashlib.sha256(data).hexdigest()


def _fmt(fingerprint):
    """(字节数, sha256) -> 可读串；等长改写会额外打标记。"""
    if fingerprint is None:
        return None
    size, h = fingerprint
    return "%d 字节 sha=%s" % (size, h[:12])


def part_map(path):
    """部件名 -> (字节数, 内容 sha256)。**必须带内容哈希**：

    只记 file_size 会对等长改写失明；页眉、页脚、属性和其他 XML 部件都可能保持
    相同长度而内容已变。
    """
    out = {}
    if zipfile.is_zipfile(path):
        with zipfile.ZipFile(path) as z:
            for info in z.infolist():
                out[info.filename] = (info.file_size, _digest(z.read(info.filename)))
    return out


def part_changes(base, cand):
    changes = []
    for name in sorted(set(base) | set(cand)):
        if name not in base:
            kind, old, new = "add", None, _fmt(cand[name])
        elif name not in cand:
            kind, old, new = "del", _fmt(base[name]), None
        elif base[name] != cand[name]:
            kind = "mod"
            old, new = _fmt(base[name]), _fmt(cand[name])
            if base[name][0] == cand[name][0]:
                # 长度相同但内容不同 —— 旧版闸门看不见这一类
                old += "  <== 长度相同，仅内容不同"
                new += "  <=="
        else:
            continue
        changes.append({"kind": kind, "location": "PART " + name, "old": old, "new": new})
    return changes


def _anchor(lines, idx):
    """往上找最近的 '#NNNN [TBL#k]' 锚点，返回精简形式。"""
    for i in range(min(idx, len(lines) - 1), -1, -1):
        m = BLOCK_RE.match(lines[i])
        if m:
            return m.group(1) + (" " + m.group(2) if m.group(2) else "")
    return ""


def line_changes(old_lines, new_lines, anchor_from_blocks):
    """通用行级差异；anchor_from_blocks=True 时用 '#NNNN TBL#k' 定位。"""
    changes = []
    sm = difflib.SequenceMatcher(a=old_lines, b=new_lines, autojunk=False)
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            continue
        old = [l.strip() for l in old_lines[i1:i2]]
        new = [l.strip() for l in new_lines[j1:j2]]
        for k in range(max(len(old), len(new))):
            o = old[k] if k < len(old) else None
            n = new[k] if k < len(new) else None
            first = o or n or ""
            if anchor_from_blocks:
                anchor = _anchor(old_lines if i1 else new_lines, i1 - 1 if i1 else j1 - 1)
                cm = CELL_RE.match(first)
                loc = (anchor + " > " + cm.group(1)) if (anchor and cm) else (anchor or cm and cm.group(1) or "#?")
                if not loc:
                    loc = "#?"
            else:
                loc = first.split("|")[0].strip() or "#?"
            changes.append({"kind": tag, "location": loc, "old": o, "new": n})
    return changes


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("baseline")
    ap.add_argument("candidate")
    ap.add_argument("--allow", action="append", default=[])
    ap.add_argument("--max-changes", type=int)
    ap.add_argument("--strict-parts", action="store_true",
                    help="把 word/document.xml 的体积变化也算作改动（默认忽略，因为它是正文改动的副产品）")
    ap.add_argument("--json")
    ap.add_argument("--quiet", action="store_true")
    args = ap.parse_args()
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

    for p in (args.baseline, args.candidate):
        if not os.path.exists(p):
            print("读不到文件: %s" % p, file=sys.stderr)
            return 2

    pc = part_changes(part_map(args.baseline), part_map(args.candidate))
    if not args.strict_parts:
        pc = [c for c in pc if c["location"] not in
              set("PART " + n for n in GATE_IGNORED_PARTS)]
    sec_a = tc.canonical_sections(args.baseline)
    sec_b = tc.canonical_sections(args.candidate)
    cc = line_changes(sec_a["comments"], sec_b["comments"], anchor_from_blocks=False)
    bc = line_changes(sec_a["body"], sec_b["body"], anchor_from_blocks=True)
    all_changes = pc + cc + bc

    def allowed(loc):
        return any(re.search(pat, loc) for pat in args.allow)

    violations = [c for c in all_changes if args.allow and not allowed(c["location"])]

    if not args.quiet:
        print("基线: %s" % args.baseline)
        print("候选: %s  sha256=%s" % (args.candidate, sha256(args.candidate)[:12]))
        print("部件级 %d 处 / 批注 %d 处 / 正文 %d 处" % (len(pc), len(cc), len(bc)))
        if pc:
            print("\n-- 部件（图片与内嵌对象都在这里暴露）--")
            for c in pc:
                print("   [%s] %s  %s -> %s" % (
                    {"add": "+", "del": "-", "mod": "~"}[c["kind"]],
                    c["location"], c["old"] or "-", c["new"] or "-"))
        if cc:
            print("\n-- 批注 --")
            for c in cc:
                print("   %s" % c["location"])
                print("      - %s" % (c["old"] or "(无)"))
                print("      + %s" % (c["new"] or "(无)"))
        if bc:
            print("\n-- 正文 --")
            last = None
            for c in bc:
                if c["location"] != last:
                    print("   %s" % c["location"])
                    last = c["location"]
                print("      - %s" % (c["old"] or "(无)"))
                print("      + %s" % (c["new"] or "(无)"))
        if not all_changes:
            print("\n无改动。")

    verdict = "pass"
    if args.max_changes is not None and len(all_changes) > args.max_changes:
        verdict = "fail"
        print("\n[闸门] 改动数 %d 超过上限 %d" % (len(all_changes), args.max_changes))
    if violations:
        verdict = "fail"
        print("\n[闸门] %d 处改动不在白名单内：" % len(violations))
        for c in violations[:50]:
            print("   ! %s" % c["location"])
        if len(violations) > 50:
            print("   ...（其余 %d 处略）" % (len(violations) - 50))
    elif args.allow:
        print("\n[闸门] 全部改动都在白名单内（%d 处）。" % len(all_changes))

    if args.json:
        with open(args.json, "w", encoding="utf-8") as fh:
            json.dump({
                "baseline": args.baseline,
                "candidate": args.candidate,
                "candidate_sha256": sha256(args.candidate),
                "part_changes": pc,
                "comment_changes": cc,
                "body_changes": bc,
                "allow": args.allow,
                "violations": violations,
                "verdict": verdict,
            }, fh, ensure_ascii=False, indent=2)
        print("\n清单已写入: %s" % args.json)

    print("结论: %s" % ("通过" if verdict == "pass" else "未通过"))
    return 0 if verdict == "pass" else 1


if __name__ == "__main__":
    sys.exit(main())
