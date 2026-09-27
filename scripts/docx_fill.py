# -*- coding: utf-8 -*-
"""docx_fill.py —— 按"填充规格"往 .docx 的指定表格单元/正文段落写文本。

设计原则：
  1. 只动规格里点名的位置，别处一个字节都不碰；
  2. 保留原段落格式：复用既有 <w:p> 的 <w:pPr>（含 pStyle / numPr 编号），
     并按原 run 切分做最小文本替换（每个 run 的 <w:rPr> 原样保留）；
  3. 段落数不一致时按需克隆/删除段落，克隆出的段落保留同一套 pPr，
     并重新生成 paraId（避免重复 ID）；
  4. --restart-lists：给每个被填的格新建独立编号实例（startOverride=1），
     让每张表的 "1) 2) 3)" 都从 1 开始，不再沿用文前面的计数；
  5. 全程不拉起 Word，纯 OOXML 层操作，因此在沙箱内也能跑。

用法:
  python docx_fill.py 输入.docx 规格.json --out 输出.docx [--dry-run] [--restart-lists] [--report 报告.json]

规格.json:
{
  "cells": [{"table": 21, "row": 2, "col": 1, "paras": ["第一段", "第二段"]}],
  "paragraphs": [{"anchor": "锚点子串", "text": "整段替换后的完整文本"}]
}
退出码: 0 成功 / 1 定位失败或规格与文档不符 / 2 参数错误
"""
import sys, os, io, json, re, copy, random, difflib, zipfile, argparse
import xml.etree.ElementTree as ET

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
w = "{%s}" % W
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
ROOT_RE = re.compile(rb"<(?![/?!])([A-Za-z0-9_.\-]+:)?([A-Za-z0-9_.\-]+)\b[^>]*>")
NS_DECL_RE = re.compile(rb'xmlns:([A-Za-z0-9_.\-]+)="([^"]*)"')
NS_NAME_RE = re.compile(rb'xmlns:([A-Za-z0-9_.\-]+)="')


def register_namespaces(data):
    """按原样注册前缀，避免序列化时改写前缀导致 mc:Ignorable 失效。"""
    for _, (prefix, uri) in ET.iterparse(io.BytesIO(data), events=("start-ns",)):
        if not prefix:
            continue
        try:
            ET.register_namespace(prefix, uri)
        except (ValueError, KeyError):
            pass


def preserve_root_namespaces(new_xml, orig_xml):
    """把原根标签上的 xmlns 声明补回新根标签。

    ElementTree 只序列化"实际用到"的命名空间，会丢掉 w15/w16 这类声明；
    但根上的 mc:Ignorable="w14 w15 ..." 仍引用它们，缺声明会让 Word 判定文件损坏。
    """
    mo, mn = ROOT_RE.search(orig_xml), ROOT_RE.search(new_xml)
    if not mo or not mn:
        return new_xml
    decls = dict(NS_DECL_RE.findall(mo.group(0)))
    have = set(NS_NAME_RE.findall(mn.group(0)))
    missing = [(k, v) for k, v in decls.items() if k not in have]
    if not missing:
        return new_xml
    add = b"".join(b' xmlns:%s="%s"' % (k, v) for k, v in missing)
    fixed = mn.group(0)[:-1] + add + b">"
    return new_xml.replace(mn.group(0), fixed, 1)


def serialize(root, orig_xml):
    body = ET.tostring(root, encoding="UTF-8", xml_declaration=False)
    out = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + body
    return preserve_root_namespaces(out, orig_xml)


def para_text(p):
    return "".join(t.text or "" for t in p.iter(w + "t"))


def numpr_of(p):
    npr = p.find(w + "pPr/" + w + "numPr")
    if npr is None:
        return None
    nid, ilvl = npr.find(w + "numId"), npr.find(w + "ilvl")
    return (nid.get(w + "val") if nid is not None else None,
            ilvl.get(w + "val") if ilvl is not None else "0")


def _set_run_text(run, text):
    """把 run 的文本设成 text：只换掉它的 <w:t>。

    rPr 与 lastRenderedPageBreak 之类的非文本子元素一律保留，绝不整段清空 run。
    """
    ts = run.findall(w + "t")
    if ts:
        at = list(run).index(ts[0])
        for t in ts:
            run.remove(t)
    else:
        at = 1 if run.find(w + "rPr") is not None else 0
    if not text:
        return
    t = ET.Element(w + "t")
    if text != text.strip() or "  " in text:
        t.set(XML_SPACE, "preserve")
    t.text = text
    run.insert(at, t)


def _run_sig(r):
    """一个 run 的字符格式指纹（rPr 的字节形式；没有 rPr 记为 b""）。"""
    rpr = r.find(w + "rPr")
    return b"" if rpr is None else ET.tostring(rpr, encoding="UTF-8")


def set_para_text(p, text):
    """把段落文本设成 text：保留 pPr，并在需要时保留原有 run 切分与每个 run 的 rPr。

    规则（由渲染结果与逐页像素比对总结出来的三条）：

    1. 文本本来就相同 —— 一个字节都不动（noop）。「无论如何都重写一遍」会把没改
       文字的段落也拍平成单 run，制造并不存在的格式变化。
    2. 段内编辑，且各 run 的格式不一致 —— 保留 run 切分，把新文本按字符对齐切回各
       run。多 run 段落里，数字所在 run 可能带有独立字体设置；拍平后数字会继承首
       run 字体，渲染时出现真实字形错位。
    3. 其余（单 run／各 run 格式相同／整段重写）—— 收敛到首个 run，删掉其余 run。
       这是 Word 里「选中整段重打」的语义。

    返回 (status, n_old, n_new)：noop=未动；keep=保留 run 结构；flat=收敛到首 run；
    new=段落原本没有 run，新建一个。
    """
    runs = p.findall(w + "r")
    old = "".join(para_text(r) for r in runs)
    if text == old:
        return ("noop", len(runs), len(runs))
    if not runs:
        _set_run_text(ET.SubElement(p, w + "r"), text)
        return ("new", 0, 1)

    ops = difflib.SequenceMatcher(None, old, text, autojunk=False).get_opcodes()
    matched = sum(i2 - i1 for tag, i1, i2, _j1, _j2 in ops if tag == "equal")
    mixed = len({_run_sig(r) for r in runs}) > 1
    if len(runs) == 1 or not mixed or (old and matched * 2 < len(old)):
        _set_run_text(runs[0], text)
        for r in runs[1:]:
            p.remove(r)
        return ("flat", len(runs), 1)

    owner = []                       # 原文本第 k 个字符属于第几个 run
    for i, r in enumerate(runs):
        owner.extend([i] * len(para_text(r)))
    pieces = [""] * len(runs)
    for tag, i1, i2, j1, j2 in ops:
        if tag == "delete":
            continue
        if tag == "equal":
            for k in range(j2 - j1):
                pieces[owner[i1 + k]] += text[j1 + k]
        else:                        # replace / insert：新字挂到该位置所属的 run
            pieces[owner[i1] if i1 < len(owner) else len(runs) - 1] += text[j1:j2]
    for r, piece in zip(runs, pieces):
        _set_run_text(r, piece)
    return ("keep", len(runs), len(runs))


def clone_para(p):
    new = copy.deepcopy(p)
    for k in list(new.attrib):
        if k.endswith("}paraId") or k.endswith("}textId"):
            new.attrib.pop(k)
    return new


def fill_cell(tc, paras):
    olds = tc.findall(w + "p")
    if not olds:
        raise ValueError("单元格没有任何 <w:p>")
    n_old, n_new = len(olds), len(paras)
    tmpl = olds[0]
    if n_new > n_old:
        for _ in range(n_new - n_old):
            new = clone_para(tmpl)
            kids = list(tc)
            tc.insert(kids.index(olds[-1]) + 1, new)
            olds = tc.findall(w + "p")
    elif n_new < n_old:
        for p in olds[n_new:]:
            tc.remove(p)
    statuses = [set_para_text(p, text)[0]
                for p, text in zip(tc.findall(w + "p"), paras)]
    return n_old, n_new, statuses


def find_tables(body):
    out, n = {}, 0
    for el in body:
        if el.tag == w + "tbl":
            n += 1
            out[n] = el
    return out


def cell_of(tbl, row, col):
    trs = tbl.findall(w + "tr")
    if row >= len(trs):
        raise ValueError("行号 %d 超界（共 %d 行）" % (row, len(trs)))
    tcs = trs[row].findall(w + "tc")
    if col >= len(tcs):
        raise ValueError("列号 %d 超界（第 %d 行共 %d 格）" % (col, row, len(tcs)))
    return tcs[col]


def find_paragraph(body, anchor):
    hits = [p for p in body.iter(w + "p") if anchor in para_text(p)]
    if not hits:
        raise ValueError("正文里找不到包含 %r 的段落" % anchor)
    if len(hits) > 1:
        raise ValueError("%r 命中 %d 个段落，锚点不唯一" % (anchor, len(hits)))
    return hits[0]


def restart_cell_list(nroot, tc, next_id):
    """给该格编号段新建独立编号实例（startOverride=1），返回新 numId 或 None。"""
    src = None
    for p in tc.findall(w + "p"):
        np = numpr_of(p)
        if np and np[0]:
            src = np
            break
    if src is None:
        return None, next_id
    src_id, ilvl = src
    abs_id = None
    for n in nroot.findall(w + "num"):
        if n.get(w + "numId") == src_id:
            a = n.find(w + "abstractNumId")
            abs_id = a.get(w + "val") if a is not None else None
    if abs_id is None:
        return None, next_id
    new_id = str(next_id)
    next_id += 1
    newnum = ET.Element(w + "num", {w + "numId": new_id})
    ET.SubElement(newnum, w + "abstractNumId", {w + "val": abs_id})
    ov = ET.SubElement(newnum, w + "lvlOverride", {w + "ilvl": ilvl})
    ET.SubElement(ov, w + "startOverride", {w + "val": "1"})
    nums = nroot.findall(w + "num")
    kids = list(nroot)
    nroot.insert(kids.index(nums[-1]) + 1, newnum)
    for p in tc.findall(w + "p"):
        np = p.find(w + "pPr/" + w + "numPr")
        if np is None:
            continue
        nid = np.find(w + "numId")
        if nid is not None and nid.get(w + "val") == src_id:
            nid.set(w + "val", new_id)
    return (src_id, new_id), next_id


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("src")
    ap.add_argument("spec")
    ap.add_argument("--out", required=True)
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--restart-lists", action="store_true",
                    help="给每个被填的格新建独立编号实例，列表从 1) 重新开始")
    ap.add_argument("--report")
    args = ap.parse_args()

    spec = json.load(open(args.spec, encoding="utf-8"))
    with zipfile.ZipFile(args.src) as z:
        names = z.namelist()
        parts = {n: z.read(n) for n in names}

    doc_xml = parts["word/document.xml"]
    register_namespaces(doc_xml)
    root = ET.fromstring(doc_xml)
    body = root.find(w + "body")
    tables = find_tables(body)

    report = {"src": os.path.abspath(args.src), "cells": [], "paragraphs": [],
              "restart_lists": []}
    filled = []

    for c in spec.get("cells", []):
        tn, row, col = int(c["table"]), int(c["row"]), int(c["col"])
        if tn not in tables:
            print("!! 找不到 TBL#%d" % tn)
            return 1
        tc = cell_of(tables[tn], row, col)
        before = [para_text(p) for p in tc.findall(w + "p")]
        n_old, n_new, st = fill_cell(tc, c["paras"])
        filled.append(tc)
        report["cells"].append({"loc": "TBL#%d > r%02dc%02d" % (tn, row, col),
                                "paras_old": n_old, "paras_new": n_new,
                                "paras_noop": st.count("noop"),
                                "paras_flat": st.count("flat"),
                                "paras_keep": st.count("keep"),
                                "before": before, "after": c["paras"]})

    for pp in spec.get("paragraphs", []):
        p = find_paragraph(body, pp["anchor"])
        before = para_text(p)
        status, n_runs, _ = set_para_text(p, pp["text"])
        report["paragraphs"].append({"anchor": pp["anchor"], "status": status,
                                     "runs": n_runs, "before": before,
                                     "after": pp["text"]})

    parts["word/document.xml"] = serialize(root, doc_xml)

    if args.restart_lists:
        num_xml = parts.get("word/numbering.xml")
        if num_xml is None:
            print("!! 没有 word/numbering.xml，无法重置列表编号")
            return 1
        register_namespaces(num_xml)
        nroot = ET.fromstring(num_xml)
        existing = [int(n.get(w + "numId")) for n in nroot.findall(w + "num")]
        next_id = (max(existing) + 1) if existing else 1
        for tn, row, col in [(int(c["table"]), int(c["row"]), int(c["col"]))
                             for c in spec.get("cells", [])]:
            tc = cell_of(tables[tn], row, col)
            res, next_id = restart_cell_list(nroot, tc, next_id)
            if res:
                report["restart_lists"].append(
                    {"loc": "TBL#%d > r%02dc%02d" % (tn, row, col),
                     "from_numId": res[0], "to_numId": res[1]})
        parts["word/document.xml"] = serialize(root, doc_xml)
        parts["word/numbering.xml"] = serialize(nroot, num_xml)

    print("规格: %d 个单元 + %d 个正文段落" % (len(report["cells"]), len(report["paragraphs"])))
    for c in report["cells"]:
        flag = ("段数 %d->%d" % (c["paras_old"], c["paras_new"])
                if c["paras_old"] != c["paras_new"] else "段数不变")
        print("  %-22s %s  %s" % (c["loc"], flag, " | ".join(c["after"])[:70]))
    for pp in report["paragraphs"]:
        print("  正文锚点 %r [%s 保留 %d 个 run] -> %s"
              % (pp["anchor"][:20], pp["status"], pp["runs"], pp["after"][:60]))
    print("run 处理: 原文未动 %d 段；保留 run 结构 %d 段；收敛到首个 run %d 段"
          % (sum(c["paras_noop"] for c in report["cells"]),
             sum(c["paras_keep"] for c in report["cells"]),
             sum(c["paras_flat"] for c in report["cells"])))
    if report["restart_lists"]:
        print("列表重置: %d 格" % len(report["restart_lists"]))
        for r in report["restart_lists"]:
            print("  %-22s numId %s -> %s" % (r["loc"], r["from_numId"], r["to_numId"]))

    if args.report:
        json.dump(report, open(args.report, "w", encoding="utf-8"),
                  ensure_ascii=False, indent=1)

    if args.dry_run:
        print("[dry-run] 未写文件")
        return 0

    with zipfile.ZipFile(args.out, "w", zipfile.ZIP_DEFLATED) as z:
        for n in names:
            z.writestr(n, parts[n])
    print("已写出:", os.path.abspath(args.out))
    return 0


if __name__ == "__main__":
    sys.exit(main())
