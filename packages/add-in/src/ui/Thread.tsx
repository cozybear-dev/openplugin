import { Caption1, Spinner } from "@fluentui/react-components";
import type { HostKind } from "@openplugin/core";
import { findCitations, jumpTo, type Citation } from "../citations";
import type { StoredLine } from "../history";

export function Thread(props: {
  hostKind: HostKind;
  lines: StoredLine[];
  busy: boolean;
}) {
  return (
    <div className="op-thread">
      {props.lines.map((line, i) => (
        <article key={i} className={`op-msg op-msg-${line.kind}`}>
          {line.kind === "tool" ? (
            <Caption1>{line.text}</Caption1>
          ) : line.kind === "assistant" ? (
            <CitedText hostKind={props.hostKind} text={line.text} />
          ) : (
            <div>{line.text}</div>
          )}
        </article>
      ))}
      {props.busy && (
        <div className="op-msg op-msg-tool">
          <Spinner size="tiny" label="Working" />
        </div>
      )}
    </div>
  );
}

function CitedText(props: { hostKind: HostKind; text: string }) {
  const parts = splitCitations(props.text);
  return (
    <div>
      {parts.map((part, i) =>
        part.citation ? (
          <button
            key={i}
            type="button"
            className="op-cite"
            onClick={() => void jumpTo(props.hostKind, part.citation!)}
          >
            {part.text}
          </button>
        ) : (
          <span key={i}>{part.text}</span>
        )
      )}
    </div>
  );
}

function splitCitations(text: string): Array<{ text: string; citation?: Citation }> {
  const cites = findCitations(text);
  if (!cites.length) return [{ text }];
  const out: Array<{ text: string; citation?: Citation }> = [];
  let cursor = 0;
  for (const c of cites) {
    const idx = text.indexOf(c.text, cursor);
    if (idx < 0) continue;
    if (idx > cursor) out.push({ text: text.slice(cursor, idx) });
    out.push({ text: c.text, citation: c });
    cursor = idx + c.text.length;
  }
  if (cursor < text.length) out.push({ text: text.slice(cursor) });
  return out;
}
