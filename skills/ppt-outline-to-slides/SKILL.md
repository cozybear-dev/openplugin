---
name: ppt-outline-to-slides
description: Turn an outline or selected text into structured PowerPoint slides. Use when the user wants slides, a deck, or an outline converted to PowerPoint.
license: Apache-2.0
compatibility: Requires OpenPlugin PowerPoint host
metadata:
  openplugin/hosts: "powerpoint,word"
  openplugin/tools: "ppt.getSlideTree ppt.addSlide ppt.setShapeText ppt.setNotes word.getSelection"
  openplugin/version: "1.0"
---

# Outline to slides

1. Get the outline from the user message or the current selection.
2. Split into 3–8 slides. One idea per slide. Title ≤ 8 words. 3–5 bullets.
3. Queue `ppt.addSlide` for each slide. Do not restyle the whole deck.
4. Add speaker notes only if the user asked.
5. Prefer new slides over editing existing ones unless asked to replace.
