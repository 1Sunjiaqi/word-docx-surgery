---
name: word-docx-surgery
description: "Use when making surgical edits to an existing .doc or .docx deliverable (fill selected table cells, reword text, reset list numbering, insert declared paragraphs) and proving with machine-checkable evidence that only the declared locations changed. Trigger words: Word, DOCX, OOXML, allowlist, change ledger, render QA, source integrity. 当真实 Word 交付件必须局部改动、且改动范围需要被证明时使用；不用于从零写新文档、整篇格式转换或无白名单的全文替换。"
metadata:
  short-description: "Word/DOCX surgical edits with change-ledger and render-QA proof | 在既有 Word 上做可核对的局部改动"
---

# Word 局部改动：两道闸门

场景：**在一份既有的 .doc/.docx 上做局部增删改，其余部分一个字都不许动。**

`python-docx` 在批注、域、编号、书签上没有 API，所以真实交付件只能在 OOXML 层动手术 —— 代价是
"我改对了"这句话必须由工具证明，不能由执行者自述。真实交付中最不能接受的往往不是
"没改对"，而是**改动来历不明、没有解释就留在文件里**。

## 三条铁律

1. **原件只读。** 一切改动在副本上做；渲染核对也开副本（`render_qa.py` 自动复制并回验原件 sha256）。
2. **只改白名单。** 改动范围在开工前写死；收工必须由 `docx_ledger.py` 证明没有越界。
3. **没渲染过不许叫"最终版"。** 内容对 ≠ 版面对；分页、续表位置、图片缩放只有渲染能看出来。

## 两道闸门（缺一道就等于没有核对）

```powershell
# 闸门一：只改了该改的（越界改动 -> 退出码 1）
python scripts/docx_ledger.py 基线.docx 候选.docx --allow "TBL#1 > r03c01" --json 清单.json

# 闸门二：版面符合预期（页数 + "有变化的页" 正是预期那几页）
python scripts/render_qa.py 候选.docx --out qa_候选
python scripts/render_qa.py 候选.docx --compare qa_基线/manifest.json
```

两道闸门都通过 = 这次改动**可核对**；只跑一道、或靠人眼看 = 不算。

## 主流程

1. `git init` + `git tag baseline`：先冻结节点（一个项目只做一次）。
2. `git worktree add <目录> -b task/<任务名> baseline`：一个任务一条支线，互不污染。
   **版本用分支和 tag 表达，不要用文件名**（"版本1/版本2/- 副本"很容易让来源失焦）。
3. **先看结构再动手**：`scripts/docx_textconv.py 目标.docx` 数表格、认 `TBL#k` 与 `rXXcYY`；
   `scripts/docx_min_edit.py 目标.docx --find "某句话" --dry-run` 看它在全篇出现几次。
4. **先做 1 处样本**，交出两道闸门的结论，等人确认后再批量。禁止 `--all` / 全文替换。
5. 填表类任务先写**规格文件**（JSON：第几张表、第几行第几列、每段写什么），再一次填完，
   不要手工逐格敲：`scripts/docx_fill.py`。
6. 收工跑两道闸门 + `scripts/verify_readonly.py`（证明既有的原件一个都没被碰）。每一步都 commit。
7. 交付的是"改动清单 + 渲染结论"，不是"已完成"三个字。

**候选件被人用 Word 打开保存过时**，加一步：闸门一必然噪声很大（实测 48 处里真实改动只有十几处），
跑 `scripts/docx_semantic_diff.py` 把 Word 自己的写法摘出来，再照着它写白名单。
顺序不能反：**"有没有越界"以闸门一的字节判等为准，语义层只用来解释和写白名单**。
噪声清单与实测数字：`references/Word重存噪声.md`。

逐步命令原文、禁止清单（每条对应一次真实事故）、失败样例：`references/工作流.md`。

## 闸门看得见什么（本 skill 的核心知识）

规范文本必须能表示下面这几类改动，缺一类就等于放行：

| 类别 | 表示方式 | 不表示会怎样 |
|---|---|---|
| 文字内容 | 正文/格文本 | 改字看不见（最基础） |
| 段落属性 | `pStyle / numPr / pageBreakBefore / keepNext / cantSplit` | "把某格列表换个编号实例"看不见 |
| run 级字符格式 | 每段的 `runs=N rpr=<指纹>` | 把一个 run 加粗变红看不见 |
| 空格的宽窄 | 先 `U+3000`/`U+00A0` 换可见标记，再折叠空白 | 半角空格换全角看不见 |
| 部件内容 | `(size, sha256)` 按内容判等 | **只比长度 = 对等长改写完全失明** |
| 段落的先后 | 两侧形状不同时改用内容锚点配对 | 按下标配对会**静默错配**：真的换了 6 处只报 3 处 |

出现新的改动类别（域、书签、批注、内容控件、图片、样式）时，先造一个**只有这一类改动**的副本，
跑一遍闸门，看它报不报。边界与失明点、只读登记机制：`references/闸门边界.md`。

## 交给另一个会话的任务卡

> 在 `<worktree 路径>` 里、从 tag `<起点>` 出发完成 `<任务>`；只允许改 `<白名单>`；
> 先做 1 处样本 + 两道闸门结论给我确认，再批量；结束前给 `docx_ledger.py` 和
> `render_qa.py --compare` 的结论；每一步 commit。

任务卡模板、"先补验收标准再执行"的写法、独立复核怎么安排：`references/任务卡与验收.md`。

## 环境前提（换机器先读这一节）

核心脚本只需要 Python 3.11+ 标准库。闸门二与 `.doc -> .docx` 转换还需要
**Windows + 桌面版 Microsoft Word + pywin32**；PDF 光栅化可用 `pdftoppm` 或
`pypdfium2`，Pillow 仅用于可选的页面联系表。

域更新时间、目录重排和真实分页只有桌面版 Word 算得出来。拉起 Word 的步骤在受限环境中
可能挂起：先正常尝试，长时间没有输出就停止；清理时只结束自己启动的进程，不要按进程名
批量终止 Python。**拿不到桌面 Word 而退回 LibreOffice 之类的替代渲染器时，有两个已知假象**
（不认 `STYLEREF`，题注会显示成 `Error: Reference source not found`；放不下的 `cantSplit` 行会被整行丢掉），
见 `references/环境事实.md` §2.1。更多兼容性说明见同一文件。

## 换到新机器/新解释器后先自检

```powershell
python scripts/selftest.py
```

它在临时目录里造样本，验证 A~H 八类改动是否都能被闸门咬住、噪声有没有被误判成改动（18 项断言）。
全部 OK 才说明这套脚本在当前环境里可用；有 FAIL 就先修环境，不要开始改真文档。
