# Word DOCX Surgery

English | [简体中文](README.md)

Make surgical edits to existing `.doc` and `.docx` files, then prove that only the
declared locations changed.

This repository is both a Codex / ChatGPT skill and an installable plugin. Agents can
read [`skills/word-docx-surgery/SKILL.md`](skills/word-docx-surgery/SKILL.md) directly
or start with the machine-readable [`llms.txt`](llms.txt).

## Install

### Codex / ChatGPT Plugin

Add the repository as a plugin marketplace, then install the plugin:

```bash
codex plugin marketplace add 1Sunjiaqi/word-docx-surgery --sparse .agents/plugins
codex plugin add word-docx-surgery@word-docx-surgery
```

Invoke it explicitly as `$word-docx-surgery`, or describe a Word task whose change
scope must be verified.

### Skill Installer

Ask `$skill-installer` to install from this repository:

```text
$skill-installer https://github.com/1Sunjiaqi/word-docx-surgery/tree/main/skills/word-docx-surgery
```

### Manual Install

Copy [`skills/word-docx-surgery`](skills/word-docx-surgery) into your Codex skills
directory:

```powershell
git clone --depth 1 https://github.com/1Sunjiaqi/word-docx-surgery.git "$env:TEMP\word-docx-surgery"
Copy-Item -Recurse "$env:TEMP\word-docx-surgery\skills\word-docx-surgery" "$env:USERPROFILE\.codex\skills\word-docx-surgery"
```

## What It Solves

`word-docx-surgery` is for real Word deliverables where the original document must
otherwise remain untouched. It combines direct OOXML edits with two
machine-checkable gates:

1. `docx_ledger.py` proves that every content and package-part change is inside the
   declared allowlist.
2. `render_qa.py` asks desktop Word to update fields, repaginate, export a PDF, and
   compare the changed pages across renders.

When the candidate has been opened and saved by Word, run `docx_semantic_diff.py` in
between: it normalizes the ways Word rewrites the same document (rsids, run splitting,
headers and footers, bookmark ids) and separates that noise from real changes. On one
real pair, gate 1 reported 48 changes while only 8 body changes and 4 comment parts were
real. It is an explanation layer only — **out-of-scope edits are still decided by the
byte-level comparison of gate 1**.

Use it for:

- Filling cells, rewording text, resetting list numbering, or inserting selected
  paragraphs in an existing Word document.
- Delivering a change ledger and render conclusion, not merely a file that appears
  correct.
- Independently checking that the source stayed unchanged, the edit stayed inside
  scope, and the page changes match expectations.

Do not use it for:

- Creating a document from scratch.
- Whole-document format conversion.
- Unrestricted find-and-replace without a declared allowlist.

## Requirements

Core editing and verification:

- Python 3.11 or newer.
- No third-party Python packages.

Real rendering and `.doc -> .docx` conversion:

- Windows with desktop Microsoft Word.
- `pywin32`.
- Poppler `pdftoppm` or `pypdfium2` for PDF rasterization.
- `Pillow` is optional and only used for the contact sheet.

Install the optional rendering dependencies:

```bash
python -m pip install -r skills/word-docx-surgery/requirements-render.txt
```

## First Check

Run the synthetic self-test before editing a real document:

```bash
python skills/word-docx-surgery/scripts/selftest.py
```

Expected result: `18 / 18` checks pass. The test builds synthetic OOXML fixtures in
a temporary directory and does not read real documents.

## Basic Workflow

Inspect the document structure:

```bash
python skills/word-docx-surgery/scripts/docx_textconv.py document.docx > structure.txt
python skills/word-docx-surgery/scripts/docx_min_edit.py document.docx --find "old text" --dry-run
```

Make a single-location sample edit:

```bash
python skills/word-docx-surgery/scripts/docx_min_edit.py copy.docx --find "old text" --replace "new text" --occurrence 1 --out sample.docx
python skills/word-docx-surgery/scripts/docx_ledger.py baseline.docx sample.docx --allow "TBL#1 > r03c01"
```

Render and compare the result:

```bash
python skills/word-docx-surgery/scripts/render_qa.py sample.docx --out qa_sample
python skills/word-docx-surgery/scripts/render_qa.py sample.docx --compare qa_baseline/manifest.json
```

The allowlist should always be narrow and traceable to the original request. Do not
use unrestricted whole-document replacement.

## Repository Layout

```text
.
├─ plugin.json                    # Portable Agent Plugins manifest
├─ .codex-plugin/plugin.json      # Codex compatibility manifest
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

The names, descriptions, and keywords in `plugin.json`,
`.codex-plugin/plugin.json`, and `agents/openai.yaml` drive plugin presentation and
implicit matching. The `SKILL.md` description decides when an agent loads the full
workflow.

## Privacy

This repository contains no user documents, extracted document text, session logs,
or private verification output. The self-test builds synthetic OOXML fixtures in a
temporary directory. The plugin runs locally and does not create a remote upload or
telemetry service. See [`PRIVACY.md`](PRIVACY.md).

## License

MIT
