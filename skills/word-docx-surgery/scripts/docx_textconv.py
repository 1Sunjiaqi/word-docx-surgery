#!/usr/bin/env python3
"""把 .docx 解包成「可逐行比较」的规范化文本。

三种用法：
  * 作为 git textconv 驱动（由 git 自动调用）：
        <python> docx_textconv.py <docx 路径>
  * 作为库，取分段结果（被 docx_ledger.py 复用）：
        import docx_textconv as tc
        sec = tc.canonical_sections(path)   # {"parts": [...], "comments": [...], "body": [...]}
        lines = tc.canonical_lines(path)    # 上面三段拼成的完整文本

设计要点：
  1. 只打印"有意义"的东西（文字、样式、编号、域、批注、部件与图片清单），
     不打印 rsid / paraId 这类每次 Word 保存都会变的属性 ——
     所以差异清单里出现的每一行都对应一次真实改动。
  2. 表格按 行/列 逐个单元格打印，表按出现顺序编号（TBL#k），
     这样能直接说出"第几张表的哪一格变了"。
  3. 只读，绝不改动任何文件。
"""
import hashlib
import os
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
WS = re.compile(r"\s+")
OLE2_MAGIC = b"\xd0\xcf\x11\xe0"


FULLWIDTH_SPACE = "\u3000"
NBSP = "\u00a0"


def norm(s):
    """折叠空白，但**先**把全角空格/不换行空格换成可见标记。

    否则"把半角空格改成全角空格"这种改动会被 \\s+ 折叠抹平而看不见。
    """
    s = (s or "").replace(FULLWIDTH_SPACE, "<U+3000>").replace(NBSP, "<U+00A0>")
    return WS.sub(" ", s).strip()


def para_text(p):
    parts = []
    for node in p.iter():
        if node.tag == W + "t":
            parts.append(node.text or "")
        elif node.tag == W + "tab":
            parts.append("\t")
        elif node.tag == W + "br":
            parts.append(" ")
        elif node.tag == W + "instrText":
            parts.append("{FIELD:" + norm(node.text) + "}")
    return norm("".join(parts))


def para_props(p):
    out = []
    ppr = p.find(W + "pPr")
    if ppr is not None:
        st = ppr.find(W + "pStyle")
        if st is not None:
            out.append("style=" + (st.get(W + "val") or "?"))
        npr = ppr.find(W + "numPr")
        if npr is not None:
            ilvl = npr.find(W + "ilvl")
            nid = npr.find(W + "numId")
            out.append("num=%s/%s" % (
                nid.get(W + "val") if nid is not None else "-",
                ilvl.get(W + "val") if ilvl is not None else "-"))
        if ppr.find(W + "pageBreakBefore") is not None:
            out.append("pageBreakBefore")
        if ppr.find(W + "keepNext") is not None:
            out.append("keepNext")
        if ppr.find(W + "cantSplit") is not None:
            out.append("cantSplit")
    return " ".join(out)


def canon_rpr(rpr):
    """把一个 <w:rPr> 规范化成可比较的短串；None 表示这个 run 不写 rPr（继承段落/默认）。"""
    if rpr is None:
        return "-"
    parts = []
    for el in rpr:
        tag = el.tag.rsplit("}", 1)[-1]
        attrs = ",".join("%s=%s" % (k.rsplit("}", 1)[-1], v)
                         for k, v in sorted(el.attrib.items()))
        parts.append(tag + ("(" + attrs + ")" if attrs else ""))
    return "&".join(sorted(parts))


def para_run_sig(p):
    """(run 个数, 各 run 字符格式指纹)。

    run 级格式原先完全不在规范文本里 —— 偷偷把一个 run 加粗/改红，闸门一个字都看不见。
    指纹只覆盖 rPr 的内容，rsid 是元素属性、不在其中。
    """
    runs = p.findall(W + "r")
    if not runs:
        return None
    fps = "\x1f".join(canon_rpr(r.find(W + "rPr")) for r in runs)
    return (len(runs), hashlib.md5(fps.encode("utf-8")).hexdigest()[:8])


def sig_text(sigs):
    """把若干段落的 run 签名拼成 runs=3,1 rpr=ab,cd 形式。"""
    sigs = [x for x in sigs if x]
    if not sigs:
        return ""
    return "runs=%s rpr=%s" % (",".join(str(x[0]) for x in sigs),
                               ",".join(x[1] for x in sigs))


def cell_props(cell):
    out = []
    tcpr = cell.find(W + "tcPr")
    if tcpr is not None:
        gs = tcpr.find(W + "gridSpan")
        if gs is not None:
            out.append("span=" + (gs.get(W + "val") or "?"))
        if tcpr.find(W + "vMerge") is not None:
            out.append("vMerge")
    return " ".join(out)


def _emit(el, out, idx, indent, tbl_counter):
    if el.tag == W + "p":
        bits = [para_props(el), sig_text([para_run_sig(el)])]
        props = " ".join(b for b in bits if b)
        out.append("%s#%04d P%s | %s" % (
            indent, idx, (" " + props) if props else "", para_text(el)))
        return
    if el.tag == W + "tbl":
        tbl_counter[0] += 1
        rows = el.findall(W + "tr")
        widths = [len(tr.findall(W + "tc")) for tr in rows]
        out.append("%s#%04d TBL#%d %d行 每行格数=%s" % (
            indent, idx, tbl_counter[0], len(rows), widths))
        for ri, tr in enumerate(rows):
            for ci, cell in enumerate(tr.findall(W + "tc")):
                props = cell_props(cell)
                paras = cell.findall(W + "p")
                txts = [para_text(p) for p in paras]
                txts = [t for t in txts if t] or [""]
                nums = []
                for p in paras:
                    npr = p.find(W + "pPr/" + W + "numPr")
                    if npr is None:
                        nums.append("-")
                        continue
                    nid = npr.find(W + "numId")
                    ilvl = npr.find(W + "ilvl")
                    nums.append("%s/%s" % (
                        nid.get(W + "val") if nid is not None else "-",
                        ilvl.get(W + "val") if ilvl is not None else "-"))
                bits = [props, sig_text([para_run_sig(p) for p in paras])]
                if any(x != "-" for x in nums):
                    bits.append("nums=" + ",".join(nums))
                props = " ".join(b for b in bits if b)
                out.append("%s    r%02dc%02d%s | %s" % (
                    indent, ri, ci, (" " + props) if props else "",
                    " / ".join(txts)))
                for sub in cell.findall(W + "tbl"):
                    _emit(sub, out, idx, indent + "        ", tbl_counter)
        return
    if el.tag == W + "sectPr":
        out.append("%s#%04d SECTPR" % (indent, idx))


def canonical_sections(path):
    """返回 {'parts': [...], 'comments': [...], 'body': [...]}。"""
    if not os.path.exists(path):
        return {"parts": [], "comments": ["(文件不存在: %s)" % path], "body": []}
    with open(path, "rb") as fh:
        magic = fh.read(4)
    if magic == OLE2_MAGIC:
        return {"parts": [], "comments": [], "body": [
            "(二进制 .doc（OLE2），本工具不解析；如需比较请先转 .docx)"]}
    if not zipfile.is_zipfile(path):
        return {"parts": [], "comments": [], "body": ["(不是 zip/OOXML 容器)"]}

    with zipfile.ZipFile(path) as z:
        names = sorted(z.namelist())
        if "word/document.xml" not in names:
            return {"parts": [], "comments": [], "body": ["(缺 word/document.xml)"]}
        parts = ["PART  %s %d" % (n, z.getinfo(n).file_size) for n in names]
        comments = []
        try:
            croot = ET.fromstring(z.read("word/comments.xml"))
            clist = croot.findall(W + "comment")
            if not clist:
                comments.append("(无)")
            for c in clist:
                comments.append("COMMENT #%s | 作者=%s | 日期=%s | %s" % (
                    c.get(W + "id"), c.get(W + "author"), c.get(W + "date"),
                    " / ".join(para_text(p) for p in c.findall(W + "p"))))
        except KeyError:
            comments.append("(无 comments.xml)")
        root = ET.fromstring(z.read("word/document.xml"))

    body = []
    b = root.find(W + "body")
    if b is not None:
        idx = 0
        tbl_counter = [0]
        for el in b:
            if el.tag in (W + "p", W + "tbl", W + "sectPr"):
                _emit(el, body, idx, "", tbl_counter)
                idx += 1
    return {"parts": parts, "comments": comments, "body": body}


def canonical_lines(path):
    sec = canonical_sections(path)
    out = ["== 部件 =="]
    out += sec["parts"]
    out += ["", "== 批注 =="]
    out += sec["comments"]
    out += ["", "== 正文块 =="]
    out += sec["body"]
    return out


def main():
    if len(sys.argv) < 2:
        print("usage: docx_textconv.py <file.docx>", file=sys.stderr)
        return 2
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stdout.write("\n".join(canonical_lines(sys.argv[1])) + "\n")
    return 0


if __name__ == "__main__":
    sys.exit(main())
