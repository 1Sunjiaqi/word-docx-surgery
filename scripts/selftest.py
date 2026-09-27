#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""自检：证明这套脚本搬到哪里都咬得住改动。只在临时目录里造样本，不碰你的文档。

    python scripts/selftest.py [--keep]

七项断言，结论取自 --json 的机器可读输出，不靠人眼看屏：
  A 正文改字     白名单命中=通过(0)；白名单不命中=拦下(1)
  B run 级格式   只改字符格式、一个字没动 -> 仍被看见
  C 空格宽窄     半角空格换成全角 U+3000 -> 仍被看见
  D 等长部件     字节数不变、内容变了 -> 仍被看见，并标注"长度相同"
  E 格内编号     单元格内 numPr 的 numId 变了 -> 仍被看见
  F 最小替换     --dry-run 报命中数；--occurrence N 只改那一处；跨 run 的句子如实报"未找到"
  G 只读自检     未登记的外部改动=失败(1)；登记后=通过(0)；登记后又被改=失败(1)；基线拒绝被覆盖

退出码：0 全部通过 / 1 有未通过项。
"""
import argparse
import json
import os
import shutil
import subprocess
import sys
import tempfile
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
PY = sys.executable

CT = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
</Types>
"""

RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"""

NUMBERING = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum>
<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
<w:num w:numId="2"><w:abstractNumId w:val="0"/></w:num>
</w:numbering>
"""

DOC = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>标题</w:t></w:r></w:p>
<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>第一条</w:t></w:r></w:p>
<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>加粗的行</w:t></w:r></w:p>
<w:tbl>
<w:tr><w:tc><w:p><w:r><w:t>甲</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>乙</w:t></w:r></w:p></w:tc></w:tr>
<w:tr><w:tc><w:p><w:r><w:t>A 半角空格</w:t></w:r></w:p></w:tc><w:tc><w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>丁</w:t></w:r></w:p></w:tc></w:tr>
</w:tbl>
<w:sectPr/>
</w:body>
</w:document>
"""

RESULTS = []


def check(name, ok, detail=""):
    RESULTS.append((name, bool(ok), detail))
    print("  [%s] %s%s" % ("OK  " if ok else "FAIL", name, ("  <- " + detail) if detail else ""))


def run(script, *args):
    cmd = [PY, os.path.join(HERE, script)] + [str(a) for a in args]
    env = dict(os.environ, PYTHONIOENCODING="utf-8")
    proc = subprocess.run(cmd, capture_output=True, env=env)
    out = proc.stdout.decode("utf-8", "replace") + proc.stderr.decode("utf-8", "replace")
    return proc.returncode, out


def build(path, document=DOC, numbering=NUMBERING):
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/document.xml", document)
        z.writestr("word/numbering.xml", numbering)


def ledger(tmp, baseline, candidate, tag, allow=None):
    jpath = os.path.join(tmp, "ledger_%s.json" % tag)
    args = [baseline, candidate, "--quiet", "--json", jpath]
    if allow:
        args += ["--allow", allow]
    rc, out = run("docx_ledger.py", *args)
    data = {}
    if os.path.exists(jpath):
        with open(jpath, encoding="utf-8") as fh:
            data = json.load(fh)
    return rc, data, out


def case_a(tmp, base):
    cand = os.path.join(tmp, "A.docx")
    build(cand, DOC.replace("<w:t>甲</w:t>", "<w:t>甲甲</w:t>"))
    rc, data, _ = ledger(tmp, base, cand, "a_ok", allow="TBL#1 > r00")
    hit = any("TBL#1 > r00c00" in c["location"] for c in data.get("body_changes", []))
    check("A1 命中白名单的改动 -> 退出码 0", rc == 0 and hit,
          "rc=%s 命中=%s" % (rc, hit))
    rc, data, _ = ledger(tmp, base, cand, "a_bad", allow="TBL#9")
    n = len(data.get("violations", []))
    check("A2 不在白名单的改动 -> 退出码 1", rc == 1 and n >= 1,
          "rc=%s 越界=%d 处" % (rc, n))


def case_b(tmp, base):
    cand = os.path.join(tmp, "B.docx")
    build(cand, DOC.replace("<w:rPr><w:b/></w:rPr>",
                            "<w:rPr><w:b/><w:color w:val=\"FF0000\"/></w:rPr>"))
    rc, data, _ = ledger(tmp, base, cand, "b")
    hit = any(c["old"] != c["new"] and "rpr=" in (c["old"] or "")
              and "rpr=" in (c["new"] or "") and "加粗的行" in (c["new"] or "")
              for c in data.get("body_changes", []))
    check("B 只改 run 字符格式（文字未动）-> 被看见", hit,
          "正文改动 %d 处" % len(data.get("body_changes", [])))


def case_c(tmp, base):
    cand = os.path.join(tmp, "C.docx")
    build(cand, DOC.replace("<w:t>A 半角空格</w:t>", "<w:t>A\u3000半角空格</w:t>"))
    rc, data, _ = ledger(tmp, base, cand, "c")
    hit = any("<U+3000>" in (c["new"] or "") for c in data.get("body_changes", []))
    check("C 半角空格 -> 全角 U+3000 -> 被看见", hit,
          "正文改动 %d 处" % len(data.get("body_changes", [])))


def case_d(tmp, base):
    same_len = NUMBERING.replace('<w:num w:numId="1"><w:abstractNumId w:val="0"/>',
                                 '<w:num w:numId="1"><w:abstractNumId w:val="1"/>')
    if len(same_len) != len(NUMBERING):
        check("D 等长改写（等长部件内容变了）-> 被看见", False, "样本没造出等长改写")
        return
    cand = os.path.join(tmp, "D.docx")
    build(cand, DOC, same_len)
    rc, data, _ = ledger(tmp, base, cand, "d")
    # 等长改写的标记打在 old 一侧（"长度相同，仅内容不同"），new 一侧只回指一个 <=="
    hit = any(c["location"] == "PART word/numbering.xml"
              and "长度相同" in (c["old"] or "") and (c["new"] or "").endswith("<==")
              for c in data.get("part_changes", []))
    check("D 等长改写（长度一样、内容变了）-> 被看见", hit,
          "部件改动 %d 处" % len(data.get("part_changes", [])))


def case_e(tmp, base):
    cand = os.path.join(tmp, "E.docx")
    build(cand, DOC.replace('<w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>丁</w:t>',
                            '<w:numId w:val="2"/></w:numPr></w:pPr><w:r><w:t>丁</w:t>'))
    rc, data, _ = ledger(tmp, base, cand, "e")
    hit = any("nums=1/0" in (c["old"] or "") and "nums=2/0" in (c["new"] or "")
              for c in data.get("body_changes", []))
    check("E 单元格内列表编号实例被换 -> 被看见", hit,
          "正文改动 %d 处" % len(data.get("body_changes", [])))


def case_f(tmp, base):
    work = os.path.join(tmp, "F.docx")
    shutil.copy2(base, work)
    rc, out = run("docx_min_edit.py", work, "--find", "甲", "--dry-run")
    check("F1 --dry-run 报出命中数", rc == 0 and "命中 1 处" in out,
          out.strip().splitlines()[0] if out.strip() else "")
    out2 = os.path.join(tmp, "F2.docx")
    rc, out = run("docx_min_edit.py", work, "--find", "甲", "--replace", "甲甲",
                  "--occurrence", "1", "--out", out2)
    rc2, data, _ = ledger(tmp, base, out2, "f")
    check("F2 --occurrence N 只改那一处", rc == 0 and rc2 == 0
          and len(data.get("body_changes", [])) == 1,
          "正文改动 %d 处（期望 1）" % len(data.get("body_changes", [])))
    rc, out = run("docx_min_edit.py", work, "--find", "标题第一条", "--dry-run")
    check("F3 被拆成多个 run 的句子 -> 如实报未找到（不瞎猜）",
          rc == 1 and "未找到" in out, "rc=%s" % rc)


def case_g(tmp):
    gdir = os.path.join(tmp, "G")
    os.makedirs(gdir)
    a = os.path.join(gdir, "a.doc")
    b = os.path.join(gdir, "b.docx")
    with open(a, "wb") as fh:
        fh.write(b"doc-a")
    with open(b, "wb") as fh:
        fh.write(b"docx-b")
    tsv = os.path.join(tmp, "基线.tsv")
    rc, _ = run("make_baseline.py", "--base", gdir, "--out", tsv)
    rc2, out = run("verify_readonly.py", "--base", gdir, "--tsv", tsv)
    check("G1 基线生成 -> 复验通过", rc == 0 and rc2 == 0, "rc=%s/%s" % (rc, rc2))
    with open(a, "ab") as fh:
        fh.write(b"X")
    rc, out = run("verify_readonly.py", "--base", gdir, "--tsv", tsv)
    check("G2 未登记的外部改动 -> 退出码 1", rc == 1 and "大小变了" in out,
          "rc=%s" % rc)
    ack = os.path.join(tmp, "登记.tsv")
    with open(ack, "w", encoding="utf-8", newline="\n") as fh:
        fh.write("a.doc\t%d\t自检脚本模拟的外部改动\n" % os.path.getsize(a))
    rc, out = run("verify_readonly.py", "--base", gdir, "--tsv", tsv, "--ack", ack)
    check("G3 显式登记后 -> 退出码 0（基线未被篡改）",
          rc == 0 and "已登记的外部改动" in out, "rc=%s" % rc)
    with open(a, "ab") as fh:
        fh.write(b"Y")
    rc, out = run("verify_readonly.py", "--base", gdir, "--tsv", tsv, "--ack", ack)
    check("G4 登记之后又被改 -> 仍然退出码 1", rc == 1, "rc=%s" % rc)
    rc, out = run("make_baseline.py", "--base", gdir, "--out", tsv)
    check("G5 基线已存在时拒绝覆盖", rc != 0 and "拒绝覆盖" in out, "rc=%s" % rc)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--keep", action="store_true", help="保留临时目录以便查看样本")
    args = ap.parse_args()
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

    tmp = tempfile.mkdtemp(prefix="wordkit_selftest_")
    print("== 自检（临时目录 %s）==" % tmp)
    print("解释器: %s" % PY)
    base = os.path.join(tmp, "base.docx")
    build(base)
    for fn in (lambda: case_a(tmp, base), lambda: case_b(tmp, base),
               lambda: case_c(tmp, base), lambda: case_d(tmp, base),
               lambda: case_e(tmp, base), lambda: case_f(tmp, base),
               lambda: case_g(tmp)):
        fn()

    bad = [r for r in RESULTS if not r[1]]
    print("\n== 结论 ==")
    print("通过 %d / %d" % (len(RESULTS) - len(bad), len(RESULTS)))
    for name, _, detail in bad:
        print("  未通过: %s  %s" % (name, detail))
    if not bad:
        print("这套脚本在当前解释器与本目录下可用：闸门看得见四类真实改动，拦得住越界。")
    if args.keep:
        print("临时目录保留在: %s" % tmp)
    else:
        shutil.rmtree(tmp, ignore_errors=True)
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main())
