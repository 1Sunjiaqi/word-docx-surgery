# Official Plugin Directory Submission

The repository is packaged for the universal ChatGPT / Codex Plugins Directory, but
listing there requires an OpenAI Platform submission:

https://platform.openai.com/plugins

## Prerequisites

- An OpenAI Platform organization with `Apps Management: Write`.
- A verified individual or business identity for the publisher.
- A production-ready listing name, descriptions, logo, website, support URL, privacy
  policy, and terms.
- Five positive and three negative test cases.

## Proposed Listing

- Name: `Word DOCX Surgery`
- Short description: `Surgical .doc/.docx edits with change and render gates.`
- Category: `Productivity`
- Repository: https://github.com/1Sunjiaqi/word-docx-surgery
- Privacy: https://github.com/1Sunjiaqi/word-docx-surgery/blob/main/PRIVACY.md
- Terms: https://github.com/1Sunjiaqi/word-docx-surgery/blob/main/TERMS.md

## Positive Test Cases

1. Fill two declared table cells in a synthetic `.docx`; expect one scoped edit and a
   passing `docx_ledger.py` result.
2. Replace one exact wording occurrence; expect a dry run to identify the occurrence,
   followed by a single-location edit and source-integrity check.
3. Reset list numbering for one declared paragraph; expect the paragraph-property
   change to appear in the ledger and render comparison.
4. Insert one declared paragraph before a known bookmark or structural anchor; expect
   the edit to stay inside the allowlist.
5. Render a `.doc` or `.docx` in desktop Word and compare before/after page manifests;
   expect page count and changed-page evidence to match the declared result.

## Negative Test Cases

1. Ask the plugin to create a Word document from scratch; expect it to redirect to a
   general document-authoring workflow rather than this surgery workflow.
2. Ask for unrestricted whole-document find-and-replace; expect a refusal to proceed
   without a narrow, declared allowlist.
3. Ask it to claim success without running the ledger or render gate; expect the
   result to remain explicitly unverified.
