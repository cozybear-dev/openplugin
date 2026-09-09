import type { HostKind } from "@openplugin/core";

export type PromptChip = { label: string; prompt: string };

export const CHIPS: Record<HostKind, PromptChip[]> = {
  excel: [
    { label: "Summarize this sheet", prompt: "Summarize the active sheet in three bullets. Cite cells." },
    { label: "Clean the selection", prompt: "Clean the selected range: fix headers, types, and blanks. Queue a table." },
    { label: "Explain the formula", prompt: "Explain the formula in the selection and cite precedents." }
  ],
  word: [
    { label: "Rewrite the selection", prompt: "Rewrite the selection so it is clearer and tighter. Keep the facts." },
    { label: "One-page memo", prompt: "Turn this document into a one-page memo with a headline and 4 bullets." },
    { label: "Tighten for execs", prompt: "Tighten the selection for an executive reader. Shorter sentences." }
  ],
  powerpoint: [
    { label: "Three slides from outline", prompt: "Add three slides from this outline. One idea per slide, short bullets." },
    { label: "Tighten this slide", prompt: "Tighten the bullets on the current slide. Max 8 words each." },
    { label: "Draft speaker notes", prompt: "Draft speaker notes for the current slide. Two short sentences." }
  ]
};
