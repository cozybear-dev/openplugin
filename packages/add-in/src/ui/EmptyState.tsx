import { Body1, Button, Caption1 } from "@fluentui/react-components";
import type { HostKind } from "@openplugin/core";
import { CHIPS } from "./prompts";

export function EmptyState(props: { hostKind: HostKind; onPick: (prompt: string) => void }) {
  return (
    <div className="op-empty">
      <Body1>
        Describe an edit. OpenPlugin reads the selection, drafts a change, and waits for you to apply
        it.
      </Body1>
      <Caption1>Try</Caption1>
      <div className="op-chips">
        {CHIPS[props.hostKind].map((chip) => (
          <Button key={chip.label} size="small" onClick={() => props.onPick(chip.prompt)}>
            {chip.label}
          </Button>
        ))}
      </div>
    </div>
  );
}
