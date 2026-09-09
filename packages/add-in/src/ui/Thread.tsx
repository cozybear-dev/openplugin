import { Button, Caption1, Spinner, Tooltip } from "@fluentui/react-components";
import { ArrowCounterclockwise24Regular } from "@fluentui/react-icons";
import type { HostKind } from "@openplugin/core";
import { useLayoutEffect, useRef } from "react";
import type { StoredLine } from "../history";
import { MarkdownMessage } from "./MarkdownMessage";
import { isPinnedToBottom, scrollToBottom } from "./scroll";

export function Thread(props: {
  hostKind: HostKind;
  lines: StoredLine[];
  busy: boolean;
  onRestore?: (userLineIndex: number) => void;
}) {
  const scrollerRef = useRef<HTMLDivElement>(null);
  const pinRef = useRef(true);
  const prevCountRef = useRef(props.lines.length);

  useLayoutEffect(() => {
    const el = scrollerRef.current;
    if (!el) return;
    const last = props.lines[props.lines.length - 1];
    if (last?.kind === "user" && props.lines.length > prevCountRef.current) pinRef.current = true;
    prevCountRef.current = props.lines.length;
    if (pinRef.current) scrollToBottom(el);
  }, [props.lines, props.busy]);

  return (
    <div
      className="op-thread"
      ref={scrollerRef}
      onScroll={() => {
        const el = scrollerRef.current;
        if (el) pinRef.current = isPinnedToBottom(el);
      }}
    >
      {props.lines.map((line, i) => (
        <article key={i} className={`op-msg op-msg-${line.kind}`}>
          {line.kind === "activity" || line.kind === "tool" ? (
            <Caption1 className={line.kind === "activity" && line.phase === "error" ? "op-activity-error" : undefined}>
              {line.kind === "activity" ? activityPrefix(line.activity) : "• "}
              {line.kind === "activity" ? line.label : line.text}
              {line.kind === "activity" && line.detail && line.detail !== line.label ? (
                <span className="op-activity-detail"> — {line.detail}</span>
              ) : null}
            </Caption1>
          ) : line.kind === "assistant" ? (
            <MarkdownMessage hostKind={props.hostKind} text={line.text} />
          ) : line.kind === "user" ? (
            <div className="op-msg-user-row">
              <div>{line.text}</div>
              {props.onRestore && (
                <Tooltip content="Restore chat and undo OpenPlugin edits after this message" relationship="label">
                  <Button
                    appearance="subtle"
                    size="small"
                    aria-label="Restore to this message"
                    icon={<ArrowCounterclockwise24Regular />}
                    onClick={() => props.onRestore?.(i)}
                  />
                </Tooltip>
              )}
            </div>
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

function activityPrefix(kind: string): string {
  if (kind === "read") return "Reading · ";
  if (kind === "write") return "Drafting · ";
  if (kind === "skill") return "Skill · ";
  if (kind === "search") return "Web · ";
  if (kind === "fetch") return "Fetch · ";
  return "• ";
}

