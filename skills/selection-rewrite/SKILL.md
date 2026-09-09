---
name: selection-rewrite
description: Rewrite, shorten, or change the tone of the current selection in Excel, Word, or PowerPoint. Use when the user asks to rewrite, edit, summarize, or restyle selected text.
license: Apache-2.0
metadata:
  openplugin/hosts: "excel,word,powerpoint"
  openplugin/tools: "word.getSelection word.replaceSelection excel.readRange excel.writeRange ppt.getSelection ppt.setShapeText"
  openplugin/version: "1.0"
---

# Selection rewrite

1. Read only the current selection (do not load the whole document).
2. Rewrite to match the user's instruction (tone, length, language).
3. Queue a replacement of the selection with the rewritten text.
4. Keep facts, numbers, and names unchanged unless asked.
5. Tell the user to click Apply.
