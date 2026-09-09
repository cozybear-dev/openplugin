---
name: word-memo-from-sheet
description: Draft a one-page financial or operational memo in Word from Excel selection data. Use when the user wants a narrative, board note, or memo from spreadsheet numbers.
license: Apache-2.0
compatibility: Requires OpenPlugin Word or Excel host
metadata:
  openplugin/hosts: "excel,word"
  openplugin/tools: "excel.getSummary excel.readRange word.insertParagraphs word.getOutline"
  openplugin/version: "1.0"
---

# Memo from sheet

When the host is Excel:

1. Read the selection with `excel.readRange`.
2. Draft a memo (headline, 3–6 bullets, one closing sentence).
3. Tell the user to paste it into Word or switch to Word and ask again. Do not claim you wrote into another file.

When the host is Word:

1. If the user pasted numbers in the prompt, use those. Otherwise ask for the figures.
2. Queue `word.insertParagraphs` at the end with a Heading-style title paragraph plus body paragraphs.
3. Keep it to one page. Cite the numbers you used.
