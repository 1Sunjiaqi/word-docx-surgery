#!/usr/bin/env python3
"""语义层比对：把「Word 重新保存」的噪声与真实改动分开。

闸门一（`docx_ledger.py`）**故意**按原始字节比较部件，这是最大的灵敏度，
不该改；代价是候选件只要被 Word 打开并保存过一次，就会报出一大片改动：

    一次真实交付里，闸门一报 48 处（22 部件 + 6 批注 + 20 正文），
    其中真实改动只有 6 处，其余 42 处都是 Word 自己的规范化。

本脚本是闸门一的**语义层补充，不替代它**：把下面这些"Word 自己的写法"归一掉，
再看还剩什么。归一规则每一条都对应一次实际观测（见 `references/Word重存噪声.md`）：

    - rsid* / w14:paraId / w14:textId   每次保存都重新生成
    - xml:space                         Word 按需增删
    - w:lastRenderedPageBreak           重新分页后写进 run
    - 相邻且 rPr 相同的 run               Word 会合并、也会在分页处再拆开
    - 同一个 run 内相邻的 w:t             Word 会把一段文字重新切分到不同 w:t
    - 书签 w:bookmarkStart/End           只比较书签名多重集（id 会整体重编号）
    - XML 命名空间前缀                    按 Clark 名比较，与前缀无关

**不归一**（必须报出来交给人判断）：文字、段落属性、run 级字符格式、
批注增删、书签增删、域的结构改写、样式表改动。

用法：
    python docx_semantic_diff.py 基线.docx 候选.docx
    python docx_semantic_diff.py 基线.docx 候选.docx --quiet --json 语义.json
    python docx_semantic_diff.py 基线.docx 候选.docx --max-real 0   # 断言"只是重存"

退出码：0 通过 / 1 超过 --max-real / 2 用法或读取失败
"""

import argparse
import copy
import hashlib
import json
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
DOC_PART = "word/document.xml"

# 归一后仍要单独比较的元素（不能只是"看不见"）
COMMENT_MARKS = ("commentRangeStart", "commentRangeEnd", "commentReference")
RUN_MARKERS = ("fldChar", "instrText", "commentReference", "br", "tab", "drawing",
               "object", "pict", "footnoteReference", "endnoteReference")
# Word 每次保存都会改写的元数据部件：单独列出，不计入"内容改动"
META_PARTS = {"docProps/app.xml", "docProps/core.xml", "docProps/custom.xml",
              "word/settings.xml"}


def local(tag):
    return tag.rsplit("}", 1)[-1]


def sha(data):
    return hashlib.sha256(data).hexdigest()


def parts(path):
    with zipfile.ZipFile(path) as z:
        return {i.filename: z.read(i.filename) for i in z.infolist()}


# --------------------------------------------------------------------------
# 归一
# --------------------------------------------------------------------------

def strip_noise(el, bookmarks):
    """就地去掉 Word 每次保存都会重写的属性/元素。"""
    for key in list(el.attrib):
        name = local(key)
        if key == XML_SPACE or name.startswith("rsid") or name in ("paraId", "textId"):
            del el.attrib[key]
    for child in list(el):
        name = local(child.tag)
        if name == "lastRenderedPageBreak":
            el.remove(child)
            continue
        if name == "bookmarkStart":
            bookmarks.append(child.get(W + "name") or "")
            el.remove(child)
            continue
        if name == "bookmarkEnd":
            el.remove(child)
            continue
        strip_noise(child, bookmarks)


def _rpr_sig(run):
    rpr = run.find(W + "rPr")
    return canon(rpr) if rpr is not None else ""


def _mergeable(run):
    return all(local(ch.tag) not in RUN_MARKERS for ch in run)


def coalesce(el):
    """相邻同 rPr 的 run 合并；同一 run 内相邻的 w:t 合并。"""
    for parent in list(el.iter()):
        i = 0
        while i < len(parent) - 1:
            a, b = parent[i], parent[i + 1]
            if (a.tag == W + "r" and b.tag == W + "r" and _mergeable(a) and _mergeable(b)
                    and _rpr_sig(a) == _rpr_sig(b)):
                for child in list(b):
                    if child.tag != W + "rPr":
                        a.append(child)
                parent.remove(b)
                continue
            i += 1
        for run in [c for c in parent if c.tag == W + "r"]:
            j = 0
            kids = list(run)
            while j < len(kids) - 1:
                a, b = kids[j], kids[j + 1]
                if a.tag == W + "t" and b.tag == W + "t":
                    a.text = (a.text or "") + (b.text or "")
                    run.remove(b)
                    kids = list(run)
                    continue
                j += 1
    return el


def normalized_el(el):
    """归一后的副本：剥噪声 + 合并 run/w:t。"""
    clone = copy.deepcopy(el)
    strip_noise(clone, [])
    coalesce(clone)
    return clone


def canon(el):
    """与命名空间前缀无关的确定性文本表示。"""
    out = []

    def rec(node):
        out.append("<" + node.tag)
        for key in sorted(node.attrib):
            out.append(" %s=%s" % (key, node.attrib[key]))
        out.append(">")
        if node.text:
            out.append(node.text)
        for child in node:
            rec(child)
        out.append("</" + node.tag + ">")

    rec(el)
    return "".join(out)


def normalized(data):
    """返回 (归一后的 digest, 书签名多重集)；非 XML 部件返回 (原始 digest, None)。"""
    try:
        root = ET.fromstring(data)
    except ET.ParseError:
        return sha(data), None
    bookmarks = []
    strip_noise(root, bookmarks)
    coalesce(root)
    return sha(canon(root).encode("utf-8")), sorted(bookmarks)


# --------------------------------------------------------------------------
# 正文逐块比对
# --------------------------------------------------------------------------

def text_of(el):
    return "".join(t.text or "" for t in el.iter(W + "t"))


def para_props(p):
    ppr = p.find(W + "pPr")
    if ppr is None:
        return {}
    return {local(ch.tag): dict((local(k), v) for k, v in ch.attrib.items()) for ch in ppr}


def run_sigs(p):
    return [_rpr_sig(r) for r in p.findall(W + "r")]


def comment_ids(el):
    return sorted(ch.get(W + "id") for ch in el.iter()
                  if local(ch.tag) in COMMENT_MARKS and ch.get(W + "id"))


def first_cell_diff(tbl_a, tbl_b):
    """表格：定位到第一个不同的单元格，返回 rXXcYY 与两侧文本。"""
    for ri, (ra, rb) in enumerate(zip(tbl_a.findall(W + "tr"), tbl_b.findall(W + "tr"))):
        for ci, (ca, cb) in enumerate(zip(ra.findall(W + "tc"), rb.findall(W + "tc"))):
            ta, tb = text_of(ca), text_of(cb)
            if ta != tb:
                return "r%02dc%02d" % (ri, ci), ta, tb
    return "", "", ""


def diff_body(base_part, cand_part):
    ra, rb = ET.fromstring(base_part), ET.fromstring(cand_part)
    ba, bb = ra.find(W + "body"), rb.find(W + "body")
    ca, cb = list(ba), list(bb)
    same_shape = (len(ca) == len(cb)
                  and [local(x.tag) for x in ca] == [local(x.tag) for x in cb])

    changes = []
    if same_shape:
        pairs = [(a, b, i, i) for i, (a, b) in enumerate(zip(ca, cb))]
        anchor_mode = "按位置配对（两侧正文字块形状一致）"
    else:
        # 段落顺序/数量变过：按下标配对会静默错配（真实踩过），
        # 改用「标签 + 归一文本 + 出现序号」的内容锚点。
        anchor_mode = "按内容锚点配对（两侧正文字块形状不同，禁止按下标配对）"
        index = {}
        for j, el in enumerate(cb):
            key = (local(el.tag), re.sub(r"\s+", "", text_of(el)))
            index.setdefault(key, []).append(j)
        pairs, used = [], set()
        for i, el in enumerate(ca):
            key = (local(el.tag), re.sub(r"\s+", "", text_of(el)))
            j = next((x for x in index.get(key, []) if x not in used), None)
            if j is None:
                changes.append({"kind": "del", "tag": local(el.tag), "index_base": i,
                                "text_base": text_of(el)[:120], "real": True})
                continue
            used.add(j)
            pairs.append((el, cb[j], i, j))
        for j, el in enumerate(cb):
            if j not in used:
                changes.append({"kind": "add", "tag": local(el.tag), "index_cand": j,
                                "text_cand": text_of(el)[:120], "real": True})

    for a, b, i, j in pairs:
        na, nb = normalized_el(a), normalized_el(b)
        if canon(na) == canon(nb):
            continue
        entry = {"kind": "mod", "tag": local(a.tag), "index_base": i, "index_cand": j,
                 "text_base": text_of(na)[:120], "text_cand": text_of(nb)[:120]}
        if local(a.tag) == "p":
            pa, pb = para_props(na), para_props(nb)
            prop_diff = {}
            for key in sorted(set(pa) | set(pb)):
                if pa.get(key) != pb.get(key):
                    prop_diff[key] = [pa.get(key), pb.get(key)]
            if prop_diff:
                entry["prop_diff"] = prop_diff
            sa, sb = run_sigs(na), run_sigs(nb)
            if sa != sb:
                entry["runs"] = [len(sa), len(sb)]
            ka, kb = comment_ids(na), comment_ids(nb)
            if ka != kb:
                entry["comment_ids"] = [ka, kb]
        else:
            loc, ta, tb = first_cell_diff(na, nb)
            if loc:
                entry["cell"] = loc
                entry["text_base"], entry["text_cand"] = ta[:120], tb[:120]
        # 一条"真实"的判定：文字 / 段落属性 / run 级格式 / 批注标记 任一变了才算
        entry["real"] = bool(entry.get("text_base") != entry.get("text_cand")
                             or entry.get("prop_diff") or entry.get("cell")
                             or entry.get("comment_ids") or entry.get("runs"))
        changes.append(entry)
    return changes, anchor_mode


# --------------------------------------------------------------------------

def main():
    ap = argparse.ArgumentParser(description="Word 重存噪声与真实改动的语义层区分")
    ap.add_argument("baseline")
    ap.add_argument("candidate")
    ap.add_argument("--json")
    ap.add_argument("--quiet", action="store_true")
    ap.add_argument("--max-real", type=int,
                    help="真实改动数超过这个值就返回 1（0 = 断言只是重存）")
    args = ap.parse_args()
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

    try:
        pa, pc = parts(args.baseline), parts(args.candidate)
    except (OSError, zipfile.BadZipFile) as exc:
        print("读取失败: %s" % exc)
        return 2

    noise_parts, real_parts, meta_parts, added, removed = [], [], [], [], []
    for name in sorted(set(pa) | set(pc)):
        if name not in pa:
            added.append(name)
            continue
        if name not in pc:
            removed.append(name)
            continue
        if sha(pa[name]) == sha(pc[name]):
            continue
        na, _ = normalized(pa[name])
        nb, _ = normalized(pc[name])
        if na == nb:
            noise_parts.append(name)
        elif name in META_PARTS:
            meta_parts.append(name)
        else:
            real_parts.append(name)

    # 书签名多重集
    _, bm_a = normalized(pa.get(DOC_PART, b"")) if DOC_PART in pa else (None, [])
    _, bm_b = normalized(pc.get(DOC_PART, b"")) if DOC_PART in pc else (None, [])
    bm_added = sorted(set(bm_b or []) - set(bm_a or []))
    bm_removed = sorted(set(bm_a or []) - set(bm_b or []))

    body_changes, anchor_mode = ([], "")
    if DOC_PART in pa and DOC_PART in pc:
        body_changes, anchor_mode = diff_body(pa[DOC_PART], pc[DOC_PART])

    real_body = [c for c in body_changes if c.get("kind") != "mod" or c.get("real", True)]
    n_real = (len(real_parts) + len(added) + len(removed) + len(real_body)
              + len(bm_added) + len(bm_removed))

    if not args.quiet:
        print("基线: %s" % args.baseline)
        print("候选: %s" % args.candidate)
        print("部件：字节有变化 %d 个 → 归一后仍不同（真实）%d 个，元数据 %d 个，纯重存噪声 %d 个"
              % (len(noise_parts) + len(real_parts) + len(meta_parts),
                 len(real_parts), len(meta_parts), len(noise_parts)))
        if real_parts:
            print("   真实部件: %s" % ", ".join(real_parts))
        if meta_parts:
            print("   元数据部件（Word 每次保存都会写，不计入内容改动）: %s"
                  % ", ".join(meta_parts))
        if noise_parts:
            print("   噪声部件: %s" % ", ".join(noise_parts[:12])
                  + (" …" if len(noise_parts) > 12 else ""))
        if added or removed:
            print("   新增 %s / 删除 %s" % (added or "-", removed or "-"))
        if bm_added or bm_removed:
            print("书签：+%s -%s" % (bm_added or "-", bm_removed or "-"))
        print("正文：归一后仍有 %d 处真实改动（%s）" % (len(real_body), anchor_mode))
        for c in real_body[:80]:
            where = "#%s" % c.get("index_cand", c.get("index_cand", "?"))
            print("   [%s] %s %s" % (c["kind"], where, c.get("tag", "")))
            if c.get("cell"):
                print("      格 %s" % c["cell"])
            if c.get("text_base") != c.get("text_cand"):
                print("      - %s" % (c.get("text_base") or "(空)"))
                print("      + %s" % (c.get("text_cand") or "(空)"))
            if c.get("prop_diff"):
                for k, v in c["prop_diff"].items():
                    print("      段属性 %s: %s -> %s" % (k, v[0], v[1]))
            if c.get("comment_ids"):
                print("      批注 %s -> %s" % tuple(c["comment_ids"]))
            if c.get("runs"):
                print("      run 数 %d -> %d" % tuple(c["runs"]))
        if real_body and len(real_body) > 80:
            print("   …其余 %d 处略" % (len(real_body) - 80))
        print("\n真实改动合计 %d 处" % n_real)

    verdict = "pass"
    if args.max_real is not None and n_real > args.max_real:
        verdict = "fail"
        print("\n[闸门] 真实改动 %d 处，超过上限 %d" % (n_real, args.max_real))

    if args.json:
        with open(args.json, "w", encoding="utf-8") as fh:
            json.dump({"baseline": args.baseline, "candidate": args.candidate,
                       "raw_parts_changed": (len(noise_parts) + len(real_parts)
                                             + len(meta_parts) + len(added) + len(removed)),
                       "noise_parts": noise_parts, "real_parts": real_parts,
                       "meta_parts": meta_parts,
                       "parts_added": added, "parts_removed": removed,
                       "bookmarks_added": bm_added, "bookmarks_removed": bm_removed,
                       "body_real_changes": real_body, "real_total": n_real,
                       "anchor_mode": anchor_mode, "verdict": verdict},
                      fh, ensure_ascii=False, indent=2)
        print("清单已写入: %s" % args.json)
    if not args.quiet:
        print("结论: %s" % ("通过" if verdict == "pass" else "未通过"))
    return 0 if verdict == "pass" else 1


if __name__ == "__main__":
    sys.exit(main())
