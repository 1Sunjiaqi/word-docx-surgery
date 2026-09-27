# word-docx-surgery

在既有 `.doc` / `.docx` 上做局部修改，并用机器可核对的证据证明“只改了申报过的那几处”。

[English](README.md)

## 它解决什么

`word-docx-surgery` 面向真实 Word 交付件：原件必须保持不动，只允许修改明确列出的位置。
它不是只检查“文字看起来对不对”，而是用两道闸门约束整次改动：

1. `docx_ledger.py` 证明内容与包内部件的每处变化都在白名单内。
2. `render_qa.py` 让桌面版 Word 更新域、重新分页、导出 PDF，并比较不同版本之间的变化页。

## 环境

核心编辑与验证：

- Python 3.11 或更高版本。
- 不需要第三方 Python 包。

真实渲染与 `.doc -> .docx` 转换：

- Windows 和桌面版 Microsoft Word。
- `pywin32`。
- Poppler `pdftoppm` 或 `pypdfium2`，用于把 PDF 拆成逐页 PNG。
- `Pillow` 可选，仅用于生成页面联系表。

安装可选的渲染依赖：

```bash
python -m pip install -r requirements-render.txt
```

## 安装为 Codex Skill

把仓库直接克隆到 Codex skills 目录：

```powershell
git clone <repository-url> "$env:USERPROFILE\.codex\skills\word-docx-surgery"
```

使用 `CODEX_HOME` 时：

```bash
git clone <repository-url> "${CODEX_HOME}/skills/word-docx-surgery"
```

之后可以用 `$word-docx-surgery` 调用，或直接描述“只改这些位置并证明没有改别处”的任务。

## 首次自检

在处理真实文档前先运行：

```bash
python scripts/selftest.py
```

预期结果为 `14 / 14` 通过。它会在临时目录生成合成 OOXML 样本，不读取任何真实文档。

## 基本流程

先看结构：

```bash
python scripts/docx_textconv.py document.docx > structure.txt
python scripts/docx_min_edit.py document.docx --find "旧文字" --dry-run
```

先做一处样本：

```bash
python scripts/docx_min_edit.py copy.docx --find "旧文字" --replace "新文字" --occurrence 1 --out sample.docx
python scripts/docx_ledger.py baseline.docx sample.docx --allow "TBL#1 > r03c01"
```

再渲染核对：

```bash
python scripts/render_qa.py sample.docx --out qa_sample
python scripts/render_qa.py sample.docx --compare qa_baseline/manifest.json
```

白名单应尽量窄，并且能反查到原始要求。不要使用无限制的全文替换。

## 仓库结构

```text
.
├─ SKILL.md
├─ agents/openai.yaml
├─ references/
├─ scripts/
├─ requirements-render.txt
└─ .github/workflows/validate.yml
```

核心脚本只使用 Python 标准库，由跨平台 CI 自检覆盖。真实 Word 渲染需要在装有桌面版
Word 的 Windows 机器上单独验证。

## 隐私边界

本仓库不包含用户文档、文档正文摘录、会话日志或私人验证输出。自检使用临时目录中的
合成 OOXML 样本。

## 许可证

MIT
