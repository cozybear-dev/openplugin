import { parseSkillMarkdown, SkillRegistry, type Skill } from "@openplugin/core";
import excelRangeCleanup from "../../../skills/excel-range-cleanup/SKILL.md?raw";
import pptOutlineToSlides from "../../../skills/ppt-outline-to-slides/SKILL.md?raw";
import selectionRewrite from "../../../skills/selection-rewrite/SKILL.md?raw";
import wordMemoFromSheet from "../../../skills/word-memo-from-sheet/SKILL.md?raw";

function skill(md: string, folder: string): Skill {
  return parseSkillMarkdown(md, folder);
}

export function bundledSkills(): SkillRegistry {
  return SkillRegistry.fromSkills([
    skill(selectionRewrite, "selection-rewrite"),
    skill(excelRangeCleanup, "excel-range-cleanup"),
    skill(wordMemoFromSheet, "word-memo-from-sheet"),
    skill(pptOutlineToSlides, "ppt-outline-to-slides")
  ]);
}
