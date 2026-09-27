# word-docx-surgery

Make surgical edits to existing `.doc` and `.docx` files, then prove that only the
declared locations changed.

[简体中文](README.zh-CN.md)

## Overview

`word-docx-surgery` is a Codex skill for real Word deliverables where the original
document must otherwise remain untouched. It combines direct OOXML edits with two
machine-checkable gates:

1. `docx_ledger.py` proves that every content and package-part change is inside the
   declared allowlist.
2. `render_qa.py` asks desktop Word to update fields, repaginate, export a PDF, and
   compare the changed pages across renders.

The skill is intentionally evidence-oriented. A change is not considered complete
merely because the edited text looks correct.

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
python -m pip install -r requirements-render.txt
```

## Install As A Codex Skill

Clone this repository directly into the Codex skills directory:

```powershell
git clone <repository-url> "$env:USERPROFILE\.codex\skills\word-docx-surgery"
```

On systems where `CODEX_HOME` is set:

```bash
git clone <repository-url> "${CODEX_HOME}/skills/word-docx-surgery"
```

Then invoke it as `$word-docx-surgery` or describe the required proof-oriented edit.

## Quick Check

Run the synthetic self-test before editing a real document:

```bash
python scripts/selftest.py
```

Expected result: `14 / 14` checks pass.

## Basic Workflow

Inspect the document structure:

```bash
python scripts/docx_textconv.py document.docx > structure.txt
python scripts/docx_min_edit.py document.docx --find "old text" --dry-run
```

Make a single-location sample edit:

```bash
python scripts/docx_min_edit.py copy.docx --find "old text" --replace "new text" --occurrence 1 --out sample.docx
python scripts/docx_ledger.py baseline.docx sample.docx --allow "TBL#1 > r03c01"
```

Render and compare the result:

```bash
python scripts/render_qa.py sample.docx --out qa_sample
python scripts/render_qa.py sample.docx --compare qa_baseline/manifest.json
```

The allowlist should always be narrow and traceable to the original request. Do not
use unrestricted whole-document replacement.

## Repository Layout

```text
.
├─ SKILL.md
├─ agents/openai.yaml
├─ references/
├─ scripts/
├─ requirements-render.txt
└─ .github/workflows/validate.yml
```

The core scripts use only the Python standard library and are covered by the
cross-platform CI self-test. Word rendering must be validated separately on a
Windows machine with desktop Word installed.

## Privacy

This repository contains no user documents, extracted document text, session logs,
or private verification output. The self-test builds synthetic OOXML fixtures in a
temporary directory.

## License

MIT
