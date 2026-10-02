# Word DOCX Surgery

[English](README.en.md) | 简体中文

在既有 `.doc` / `.docx` 上做局部修改，并用机器可核对的证据证明“只改了申报过的那几处”。

本仓库同时提供 Codex / ChatGPT skill 与 plugin。Agent 可直接读取
[`skills/word-docx-surgery/SKILL.md`](skills/word-docx-surgery/SKILL.md)，或先读取机器友好的
[`llms.txt`](llms.txt)。

## 安装

### Codex / ChatGPT Plugin

先把本仓库注册为 plugin marketplace，再安装 plugin：

```bash
codex plugin marketplace add 1Sunjiaqi/word-docx-surgery --sparse .agents/plugins
codex plugin add word-docx-surgery@word-docx-surgery
```

安装后可在新任务中用 `$word-docx-surgery` 显式调用，也可以直接描述一个需要核对改动范围的 Word 任务。

### 使用 Skill Installer

在 Codex 中让 `$skill-installer` 从本仓库安装：

```text
$skill-installer https://github.com/1Sunjiaqi/word-docx-surgery/tree/main/skills/word-docx-surgery
```

### 手动安装

把 [`skills/word-docx-surgery`](skills/word-docx-surgery) 整个目录复制到用户的 skills 目录：

```powershell
git clone --depth 1 https://github.com/1Sunjiaqi/word-docx-surgery.git "$env:TEMP\word-docx-surgery"
Copy-Item -Recurse "$env:TEMP\word-docx-surgery\skills\word-docx-surgery" "$env:USERPROFILE\.codex\skills\word-docx-surgery"
```

## 它解决什么

`word-docx-surgery` 面向真实 Word 交付件：原件必须保持不动，只允许修改明确列出的位置。
它不是只检查“文字看起来对不对”，而是用两道闸门约束整次改动：

1. `docx_ledger.py` 证明内容与包内部件的每处变化都在白名单内。
2. `render_qa.py` 让桌面版 Word 更新域、重新分页、导出 PDF，并比较不同版本之间的变化页。

候选件被人用 Word 打开保存过时，中间加一层 `docx_semantic_diff.py`：它按语义归一
（rsid、run 切分、页眉页脚、书签 id 等），把 Word 自己写的噪声与真实改动分开——
实测一对文件上闸门一报 48 处，真实改动只有 8 处正文 + 4 个批注部件。它只是解释层，
**"有没有越界"仍以闸门一的字节判等为准**。

适合：

- 在既有 Word 文档中填表、改措辞、重置列表编号或插入指定段落。
- 需要交付“改动清单 + 渲染结论”，而不是只交一份看似正确的文件。
- 需要独立复核“原件没变、改动没有越界、版面变化符合预期”。

不适合：

- 从零创建 Word 文档。
- 整篇格式转换。
- 没有明确白名单，却要求全文替换或重排。

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
python -m pip install -r skills/word-docx-surgery/requirements-render.txt
```

## 首次自检

在处理真实文档前先运行：

```bash
python skills/word-docx-surgery/scripts/selftest.py
```

预期结果为 `18 / 18` 通过。它会在临时目录生成合成 OOXML 样本，不读取任何真实文档。

## 基本流程

先看结构：

```bash
python skills/word-docx-surgery/scripts/docx_textconv.py document.docx > structure.txt
python skills/word-docx-surgery/scripts/docx_min_edit.py document.docx --find "旧文字" --dry-run
```

先做一处样本：

```bash
python skills/word-docx-surgery/scripts/docx_min_edit.py copy.docx --find "旧文字" --replace "新文字" --occurrence 1 --out sample.docx
python skills/word-docx-surgery/scripts/docx_ledger.py baseline.docx sample.docx --allow "TBL#1 > r03c01"
```

再渲染核对：

```bash
python skills/word-docx-surgery/scripts/render_qa.py sample.docx --out qa_sample
python skills/word-docx-surgery/scripts/render_qa.py sample.docx --compare qa_baseline/manifest.json
```

白名单应尽量窄，并且能反查到原始要求。不要使用无限制的全文替换。

## 仓库结构

```text
.
├─ plugin.json                    # 可移植 Agent Plugins 清单
├─ .codex-plugin/plugin.json      # Codex 兼容清单
├─ .agents/plugins/marketplace.json
├─ skills/word-docx-surgery/
│  ├─ SKILL.md
│  ├─ agents/openai.yaml
│  ├─ references/
│  ├─ scripts/
│  └─ requirements-render.txt
├─ llms.txt
└─ .github/workflows/validate.yml
```

`plugin.json`、`.codex-plugin/plugin.json` 和 `agents/openai.yaml` 中的名称、描述和关键词用于
plugin 展示与隐式匹配；`SKILL.md` 的 `description` 决定 agent 何时加载完整工作流。

## 隐私边界

本仓库不包含用户文档、文档正文摘录、会话日志或私人验证输出。自检使用临时目录中的
合成 OOXML 样本。Plugin 本身在本机执行，不建立远程上传或数据收集服务。详见
[`PRIVACY.md`](PRIVACY.md)。

## 许可证

MIT
