#!/usr/bin/env python3
"""旧格式 .doc（OLE2）-> .docx：交给本机 Word 转换，不改动原文件。

    python scripts/doc_to_docx.py <旧.doc> [更多.doc ...] [--out 目录]

为什么需要它：python-docx 读写不了二进制 .doc。
Word 自己的 SaveAs2(FileFormat=16) 是唯一保真的路径，实测 10 项统计量逐项相等。
注意：本步骤需要 Windows、桌面版 Word 和 pywin32。
"""
import argparse
import os
import sys


def convert(src, out_dir):
    import win32com.client as wc
    src = os.path.abspath(src)
    base = os.path.splitext(os.path.basename(src))[0]
    dst = os.path.join(os.path.abspath(out_dir), base + ".docx")
    app = wc.DispatchEx("Word.Application")
    app.Visible = False
    app.DisplayAlerts = 0
    doc = None
    try:
        doc = app.Documents.Open(src, ReadOnly=True, AddToRecentFiles=False)
        stats = {
            "paragraphs": doc.Paragraphs.Count,
            "tables": doc.Tables.Count,
            "pages": doc.ComputeStatistics(2),
            "words": doc.ComputeStatistics(0),
            "chars": doc.ComputeStatistics(3),
        }
        doc.SaveAs2(dst, FileFormat=16)
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
    return dst, stats


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("files", nargs="+")
    ap.add_argument("--out")
    args = ap.parse_args()
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    for f in args.files:
        if not os.path.exists(f):
            print("读不到: %s" % f)
            return 2
        out_dir = args.out or os.path.dirname(os.path.abspath(f))
        dst, stats = convert(f, out_dir)
        print("%s -> %s" % (os.path.basename(f), dst))
        print("   段落=%d 表格=%d 页数=%d 字数=%d 字符=%d"
              % (stats["paragraphs"], stats["tables"], stats["pages"],
                 stats["words"], stats["chars"]))
    return 0


if __name__ == "__main__":
    sys.exit(main())
