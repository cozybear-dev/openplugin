import { Button, Caption1, Tooltip } from "@fluentui/react-components";
import { Add24Regular, Settings24Regular } from "@fluentui/react-icons";
import type { HostKind } from "@openplugin/core";

export function Header(props: {
  hostKind: HostKind;
  contextLabel: string;
  onNewChat: () => void;
  onOpenSettings: () => void;
}) {
  return (
    <header className="op-header">
      <div className="op-brand">
        <svg className="op-mark" viewBox="0 0 36 40" fill="none" aria-hidden><path d="M15 5H6v30h9M21 5h9v30h-9" /><path d="M13 13h10v14H13z" /><circle cx="18" cy="20" r="2" /></svg>
        <div>
          <div className="op-title">OpenPlugin</div>
          <Caption1 className="op-context">{props.hostKind === "powerpoint" ? "PowerPoint" : props.hostKind === "word" ? "Word" : "Excel"} · WORKSPACE</Caption1>
        </div>
      </div>
      <div className="op-header-actions">
        <Tooltip content="New chat" relationship="label">
          <Button aria-label="New chat" appearance="subtle" icon={<Add24Regular />} onClick={props.onNewChat} />
        </Tooltip>
        <Tooltip content="Settings" relationship="label">
          <Button aria-label="Settings" appearance="subtle" icon={<Settings24Regular />} onClick={props.onOpenSettings} />
        </Tooltip>
      </div>
    </header>
  );
}
