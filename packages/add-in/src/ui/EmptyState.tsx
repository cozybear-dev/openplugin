import { ArrowRight24Regular } from "@fluentui/react-icons";
import type { HostKind } from "@openplugin/core";
import { CHIPS } from "./prompts";

const copy = {
  excel: { title: "Make something", emphasis: "of your data.", description: "Find the insight. Untangle the formula. Leave the busywork here.", details: ["The story behind the numbers", "A little order, cell by cell", "The logic, made clear"] },
  word: { title: "Find the words.", emphasis: "Keep your voice.", description: "A sharper draft, a clearer argument, a little more room to think.", details: ["Clarity without losing your voice", "The essentials, on one page", "A sharper point of view"] },
  powerpoint: { title: "Shape the story.", emphasis: "Own the room.", description: "Give your ideas a structure, your slides a focus, and your delivery a rhythm.", details: ["A beginning, a middle, a point", "One clear idea at a time", "The words between the slides"] }
};

export function EmptyState(props: { hostKind: HostKind; onPick: (prompt: string) => void }) {
  const content = copy[props.hostKind];
  return (
    <main className="op-empty">
      <div className="op-welcome-heading"><span className="op-eyebrow">A WORKING COMPANION</span><svg className="op-page-motif" viewBox="0 0 60 52" fill="none" aria-hidden><path d="M8 13h30v35H8zM15 7h30v35M22 1h30v35" /><path d="M15 24h16M15 30h12M15 36h16" /></svg></div>
      <h1>{content.title}<em>{content.emphasis}</em></h1>
      <p className="op-welcome-description">{content.description}</p>
      <div className="op-starters-label"><span>START WITH A SMALL STEP</span><span>01 — 03</span></div>
      <div className="op-chips">
        {CHIPS[props.hostKind].map((chip, index) => (
          <button className="op-starter" key={chip.label} onClick={() => props.onPick(chip.prompt)}>
            <span className="op-starter-number">0{index + 1}</span>
            <span className="op-starter-copy"><strong>{chip.label}</strong><span>{content.details[index]}</span></span>
            <ArrowRight24Regular aria-hidden />
          </button>
        ))}
      </div>
      <p className="op-empty-note">Your document. Your direction.</p>
    </main>
  );
}
