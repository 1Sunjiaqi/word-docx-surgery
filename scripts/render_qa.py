#!/usr/bin/env python3
"""渲染核对：副本 -> Word 更新域 -> PDF -> 逐页 PNG -> 联系表 + 清单。

    python render_qa.py <要看的.docx> [--out 目录] [--dpi 110] [--no-fields]
    python render_qa.py <要看的.docx> --compare 上一次的_manifest.json

做的事：
  1. 把输入复制成 <out>/_source.docx —— 永远不碰你给的那份原件（原件 sha256 会记录下来做对照）；
  2. 用本机 Word（COM）打开副本：更新域与目录、重新分页（--no-fields 可跳过）；
  3. 另存 PDF，然后关闭且不保存；
  4. 用 pdftoppm（Codex 运行时自带 poppler）或 pypdfium2 把 PDF 拆成逐页 PNG；
  5. 拼一张联系表，方便一眼扫版面；
  6. 写 manifest.json：页数、每页 PNG 的 sha256、输入文件的 sha256。
     下次渲染同一文档时加 --compare，就能说出"只有第 3、22 页不一样"。

依赖：pywin32 + 本机 Word（必需）；pdftoppm 或 pypdfium2（二选一）；Pillow（可选，仅联系表）。
"""
import argparse
import glob
import hashlib
import json
import os
import shutil
import subprocess
import sys
import time

PAGE_RE = "^\\s*#\\d{4}"
DOC_PART = "word/document.xml"


def sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def find_pdftoppm():
    cands = []
    env = os.environ.get("WORDKIT_PDFTOPPM")
    if env:
        cands.append(env)
    home = os.path.expanduser("~")
    cands += sorted(glob.glob(os.path.join(
        home, ".cache", "codex-runtimes", "*", "dependencies", "native",
        "poppler", "**", "pdftoppm.exe"), recursive=True))
    which = shutil.which("pdftoppm")
    if which:
        cands.append(which)
    for c in cands:
        if c and os.path.exists(c):
            return c
    return None


def word_export(src, pdf_path, update_fields=True):
    """用本机 Word 打开副本、更新域、导出 PDF；返回 (页数, 是否更新了域)。"""
    import win32com.client as wc

    app = wc.DispatchEx("Word.Application")
    app.Visible = False
    app.DisplayAlerts = 0
    doc = None
    updated = False
    try:
        doc = app.Documents.Open(os.path.abspath(src), ReadOnly=True,
                                 AddToRecentFiles=False)
        if update_fields:
            try:
                doc.Fields.Update()
                for toc in doc.TablesOfContents:
                    toc.Update()
                doc.Repaginate()
                updated = True
            except Exception as exc:  # 更新失败不该让整件事失败，但要如实上报
                print("  ! 更新域时出错（PDF 仍会导出）：%s" % exc)
        pages = int(doc.ComputeStatistics(2))
        doc.SaveAs2(os.path.abspath(pdf_path), FileFormat=17)
        doc.Close(False)
        doc = None
    finally:
        if doc is not None:
            try:
                doc.Close(False)
            except Exception:
                pass
        try:
            app.Quit()
        except Exception:
            pass
    return pages, updated


def rasterize(pdf_path, out_dir, dpi):
    os.makedirs(out_dir, exist_ok=True)
    for old in glob.glob(os.path.join(out_dir, "*.png")):
        os.remove(old)
    pop = find_pdftoppm()
    if pop:
        prefix = os.path.join(out_dir, "page")
        subprocess.run([pop, "-r", str(dpi), "-png", os.path.abspath(pdf_path), prefix],
                       check=True, capture_output=True)
        pngs = sorted(glob.glob(os.path.join(out_dir, "page*.png")))
        return pngs, "pdftoppm: %s" % pop
    try:
        import pypdfium2 as pdfium
    except ImportError:
        raise SystemExit(
            "既没有 pdftoppm，也没有 pypdfium2，无法把 PDF 拆成图片。\n"
            "装一个：python -m pip install pypdfium2")
    doc = pdfium.PdfDocument(pdf_path)
    pngs = []
    for i in range(len(doc)):
        img = doc[i].render(scale=dpi / 72).to_pil()
        p = os.path.join(out_dir, "page-%02d.png" % (i + 1))
        img.save(p)
        pngs.append(p)
    return pngs, "pypdfium2"


def contact_sheet(pngs, out_path, cols=4, width=380):
    try:
        from PIL import Image, ImageDraw
    except ImportError:
        return None
    thumbs = []
    for p in pngs:
        im = Image.open(p).convert("RGB")
        h = int(im.height * width / im.width)
        thumbs.append(im.resize((width, h)))
    if not thumbs:
        return None
    rows = (len(thumbs) + cols - 1) // cols
    cell_h = max(t.height for t in thumbs) + 22
    sheet = Image.new("RGB", (cols * width, rows * cell_h), "white")
    draw = ImageDraw.Draw(sheet)
    for i, t in enumerate(thumbs):
        x = (i % cols) * width
        y = (i // cols) * cell_h
        sheet.paste(t, (x, y + 20))
        draw.rectangle([x, y, x + width - 1, y + cell_h - 1], outline="#bbbbbb")
        draw.text((x + 6, y + 4), "第 %d 页" % (i + 1), fill="black")
    sheet.save(out_path)
    return out_path


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("path")
    ap.add_argument("--out")
    ap.add_argument("--dpi", type=int, default=110)
    ap.add_argument("--no-fields", action="store_true")
    ap.add_argument("--no-contact-sheet", action="store_true")
    ap.add_argument("--compare")
    args = ap.parse_args()
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

    src = os.path.abspath(args.path)
    if not os.path.exists(src):
        print("读不到文件: %s" % src, file=sys.stderr)
        return 2
    stem = os.path.splitext(os.path.basename(src))[0]
    out = os.path.abspath(args.out or os.path.join(os.path.dirname(src), "_qa_" + stem))
    os.makedirs(out, exist_ok=True)

    copy_path = os.path.join(out, "_source.docx")
    shutil.copy2(src, copy_path)

    pdf_path = os.path.join(out, "rendered.pdf")
    t0 = time.time()
    print("== 渲染核对 ==")
    print("输入: %s" % src)
    print("      sha256=%s（核对前后应当一致）" % sha256(src)[:16])
    print("副本: %s" % copy_path)
    pages, updated = word_export(copy_path, pdf_path, not args.no_fields)
    print("页数: %d（Word 统计）  域已更新=%s  耗时 %.1fs"
          % (pages, "是" if updated else "否", time.time() - t0))

    pngs, engine = rasterize(pdf_path, os.path.join(out, "pages"), args.dpi)
    print("逐页 PNG: %d 张（%s, %d dpi）-> %s"
          % (len(pngs), engine, args.dpi, os.path.join(out, "pages")))
    if len(pngs) != pages:
        print("  ! 注意：PNG 张数(%d) 与 Word 页数(%d) 不一致，请人工看一眼"
              % (len(pngs), pages))

    sheet = None
    if not args.no_contact_sheet:
        sheet = contact_sheet(pngs, os.path.join(out, "contact_sheet.png"))
        print("联系表: %s" % (sheet or "（未生成：没装 Pillow）"))

    manifest = {
        "input": src,
        "input_sha256": sha256(src),
        "pages": pages,
        "fields_updated": updated,
        "dpi": args.dpi,
        "rasterizer": engine,
        "pdf": pdf_path,
        "contact_sheet": sheet,
        "png_sha256": dict((os.path.basename(p), sha256(p)) for p in pngs),
        "generated_at": time.strftime("%Y-%m-%d %H:%M:%S"),
    }
    man_path = os.path.join(out, "manifest.json")
    with open(man_path, "w", encoding="utf-8") as fh:
        json.dump(manifest, fh, ensure_ascii=False, indent=2)
    print("清单: %s" % man_path)

    after = sha256(src)
    if after == manifest["input_sha256"]:
        print("原件校验: sha256 与开始时一致，未被改动。")
    else:
        print("原件校验: !! sha256 变了（%s -> %s），请立刻检查 %s"
              % (manifest["input_sha256"][:16], after[:16], src))

    if args.compare:
        with open(args.compare, encoding="utf-8") as fh:
            old = json.load(fh)
        diff = sorted(
            set(k for k in set(old["png_sha256"]) | set(manifest["png_sha256"])
                if old["png_sha256"].get(k) != manifest["png_sha256"].get(k)))
        print("\n== 与 %s 对比 ==" % os.path.basename(args.compare))
        print("页数: %s -> %s" % (old["pages"], pages))
        if not diff:
            print("逐页 PNG 全部逐字节一致：版面没有变化。")
        else:
            print("有变化的是 %d 页: %s" % (len(diff), "、".join(diff)))
    return 0


if __name__ == "__main__":
    sys.exit(main())
