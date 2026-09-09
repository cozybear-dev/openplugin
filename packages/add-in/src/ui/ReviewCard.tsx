import { Body1, Button, Caption1, Checkbox } from "@fluentui/react-components";
import type { DiffHunk } from "@openplugin/core";
import { jumpTo } from "../citations";
import type { HostKind } from "@openplugin/core";

export function ReviewPanel(props: {
  hostKind: HostKind;
  hunks: DiffHunk[];
  selected: Set<string>;
  onToggle: (id: string, on: boolean) => void;
  onApply: () => void;
  onReject: () => void;
}) {
  const n = props.hunks.length;
  const selectedCount = [...props.selected].filter((id) => props.hunks.some((h) => h.id === id)).length;
  return (
    <section className="op-review">
      <Caption1 className="op-review-kicker">
        {n} {n === 1 ? "change" : "changes"} · review before they hit the document
      </Caption1>
      <div className="op-hunks">
        {props.hunks.map((hunk) => (
          <article key={hunk.id} className="op-hunk">
            <div className="op-hunk-head">
              <Checkbox
                checked={props.selected.has(hunk.id)}
                onChange={(_, d) => props.onToggle(hunk.id, Boolean(d.checked))}
                label={hunk.title}
              />
              {!hunk.reversible && <Caption1>Not auto-revertable</Caption1>}
            </div>
            {hunk.kind === "grid" && (
              <GridDiff hostKind={props.hostKind} hunk={hunk} />
            )}
            {hunk.kind === "text" && <TextDiff hunk={hunk} />}
            {hunk.kind === "note" && <Caption1>{hunk.detail}</Caption1>}
          </article>
        ))}
      </div>
      <div className="op-review-actions">
        <Button appearance="primary" disabled={selectedCount === 0} onClick={props.onApply}>
          Apply selected ({selectedCount})
        </Button>
        <Button appearance="subtle" onClick={props.onReject}>
          Reject all
        </Button>
      </div>
    </section>
  );
}

function GridDiff(props: {
  hostKind: HostKind;
  hunk: Extract<DiffHunk, { kind: "grid" }>;
}) {
  const changed = props.hunk.cells.filter((c) => c.changed);
  const rows = (changed.length ? changed : props.hunk.cells).slice(0, 40);
  return (
    <div className="op-grid">
      <div className="op-grid-row op-grid-head">
        <span>Cell</span>
        <span>Before</span>
        <span>After</span>
      </div>
      {rows.map((cell) => (
        <button
          key={cell.address}
          type="button"
          className={`op-grid-row ${cell.changed ? "is-changed" : ""}`}
          onClick={() =>
            void jumpTo(props.hostKind, {
              text: cell.address,
              kind: "cell",
              address: cell.address.split("!").pop(),
              sheet: cell.address.includes("!") ? cell.address.split("!")[0] : undefined
            })
          }
        >
          <span className="op-cite">{cell.address}</span>
          <span className="op-before">{fmt(cell.before)}</span>
          <span className="op-after">{fmt(cell.after)}</span>
        </button>
      ))}
      {props.hunk.truncated && <Caption1>Showing the first cells only.</Caption1>}
    </div>
  );
}

function TextDiff(props: { hunk: Extract<DiffHunk, { kind: "text" }> }) {
  return (
    <p className="op-textdiff">
      {props.hunk.parts.map((part, i) => (
        <span key={i} className={`op-t op-t-${part.type}`}>
          {part.text}
        </span>
      ))}
    </p>
  );
}

function fmt(value: unknown): string {
  if (value == null || value === "") return "·";
  return String(value);
}

export function AppliedBar(props: {
  summary: string;
  canRevert: boolean;
  onRevert: () => void;
}) {
  return (
    <section className="op-applied">
      <Body1>Applied · {props.summary}</Body1>
      <Button appearance="subtle" disabled={!props.canRevert} onClick={props.onRevert}>
        Revert
      </Button>
    </section>
  );
}
