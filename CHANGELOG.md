# Changelog

## 0.3.0

- Added `scripts/docx_semantic_diff.py`: separates the ways Word rewrites a document on
  save (rsids, `w14:paraId`, `xml:space`, `w:lastRenderedPageBreak`, run splitting and
  merging, bookmark id renumbering, namespace declarations) from real changes. Measured
  on a real delivery: gate 1 reported 48 changes, of which only 8 body changes and 4
  comment parts were real. `--max-real 0` asserts "this was only a re-save".
- Added `references/Word重存噪声.md` with the measured noise catalogue and the rules for
  what must *not* be normalized away.
- `references/闸门边界.md`: banned index-based paragraph pairing once paragraph order
  has changed (a real swap of 6 paragraph marks was silently missed by the index-paired
  comparison and only appeared under content anchors); documented that comment `paraId`
  drift is a hint, not a failure.
- `references/环境事实.md`: documented the two known falsehoods of LibreOffice-style
  fallback rendering (no `STYLEREF` support, so captions render as an error; `cantSplit`
  rows dropped at page boundaries, so row counting lies).
- Self-test grew from 14 to 18 assertions: new case H checks that 7 classes of re-save
  rewriting are classified as noise, that the check is not vacuous, and that a single real
  edit is still localized to the exact cell.

## 0.2.0

- Reorganized the repository as an installable Agent Plugins package.
- Added portable and Codex compatibility manifests.
- Added a repo marketplace for `codex plugin marketplace add`.
- Made Chinese the default README and moved English to `README.en.md`.
- Added `llms.txt`, plugin keywords, richer discovery metadata, and public
  privacy/terms documents.

## 0.1.0

- Initial public release of the standalone skill.
