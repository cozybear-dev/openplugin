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
        <span className="op-mark" aria-hidden />
        <div>
          <div className="op-title">OpenPlugin</div>
          <Caption1 className="op-context">{props.contextLabel || props.hostKind}</Caption1>
        </div>
      </div>
      <div className="op-header-actions">
        <Tooltip content="New chat" relationship="label">
          <Button appearance="subtle" icon={<Add24Regular />} onClick={props.onNewChat} />
        </Tooltip>
        <Tooltip content="Settings" relationship="label">
          <Button appearance="subtle" icon={<Settings24Regular />} onClick={props.onOpenSettings} />
        </Tooltip>
      </div>
    </header>
  );
}
