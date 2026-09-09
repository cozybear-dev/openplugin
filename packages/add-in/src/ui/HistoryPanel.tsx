import {
  Button,
  Caption1,
  Dialog,
  DialogActions,
  DialogBody,
  DialogContent,
  DialogSurface,
  DialogTitle,
  Field,
  Input,
  Menu,
  MenuItem,
  MenuList,
  MenuPopover,
  MenuTrigger
} from "@fluentui/react-components";
import { MoreHorizontal24Regular } from "@fluentui/react-icons";
import { useState } from "react";
import type { ConversationRecord } from "../conversation-store";

export function HistoryPanel(props: {
  conversations: ConversationRecord[];
  activeId: string;
  error: string | null;
  onClose: () => void;
  onSelect: (id: string) => void;
  onRename: (id: string, title: string) => void;
  onDelete: (id: string) => void;
}) {
  const [pendingDelete, setPendingDelete] = useState<ConversationRecord | null>(null);
  const [pendingRename, setPendingRename] = useState<ConversationRecord | null>(null);
  const [renameValue, setRenameValue] = useState("");

  return (
    <div className="op-settings">
      <div className="op-settings-head">
        <div>
          <div className="op-eyebrow">YOUR THREADS</div>
          <h1>History</h1>
        </div>
        <Button appearance="subtle" size="small" onClick={props.onClose}>
          Done
        </Button>
      </div>
      {props.error ? <Caption1>{props.error}</Caption1> : null}
      {props.conversations.length === 0 ? (
        <Caption1>No saved chats yet</Caption1>
      ) : (
        <ul className="op-history-list">
          {props.conversations.map((chat) => (
            <li key={chat.id} className={`op-history-row${chat.id === props.activeId ? " is-active" : ""}`}>
              <button
                type="button"
                className="op-history-open"
                aria-label={chat.title}
                onClick={() => props.onSelect(chat.id)}
              >
                <strong>{chat.title}</strong>
                <span className="op-history-meta">
                  {formatRelativeTime(chat.updatedAt)}
                  {chat.documentTitle ? ` · ${chat.documentTitle}` : ""}
                </span>
              </button>
              <Menu>
                <MenuTrigger disableButtonEnhancement>
                  <Button
                    appearance="subtle"
                    size="small"
                    icon={<MoreHorizontal24Regular />}
                    aria-label={`Actions for ${chat.title}`}
                  />
                </MenuTrigger>
                <MenuPopover>
                  <MenuList>
                    <MenuItem
                      onClick={() => {
                        setPendingRename(chat);
                        setRenameValue(chat.title);
                      }}
                    >
                      Rename
                    </MenuItem>
                    <MenuItem onClick={() => setPendingDelete(chat)}>Delete</MenuItem>
                  </MenuList>
                </MenuPopover>
              </Menu>
            </li>
          ))}
        </ul>
      )}

      <Dialog
        modalType="alert"
        open={pendingDelete !== null}
        onOpenChange={(_, data) => {
          if (!data.open) setPendingDelete(null);
        }}
      >
        <DialogSurface className="op-restore-dialog">
          <DialogBody>
            <DialogTitle>Delete “{pendingDelete?.title}”?</DialogTitle>
            <DialogContent>This cannot be undone.</DialogContent>
            <DialogActions>
              <Button appearance="secondary" onClick={() => setPendingDelete(null)}>
                Cancel
              </Button>
              <Button
                appearance="primary"
                onClick={() => {
                  const id = pendingDelete?.id;
                  setPendingDelete(null);
                  if (id) props.onDelete(id);
                }}
              >
                Delete
              </Button>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>

      <Dialog
        modalType="alert"
        open={pendingRename !== null}
        onOpenChange={(_, data) => {
          if (!data.open) setPendingRename(null);
        }}
      >
        <DialogSurface className="op-restore-dialog">
          <DialogBody>
            <DialogTitle>Rename chat</DialogTitle>
            <DialogContent>
              <Field label="Title">
                <Input value={renameValue} onChange={(_, d) => setRenameValue(d.value)} />
              </Field>
            </DialogContent>
            <DialogActions>
              <Button appearance="secondary" onClick={() => setPendingRename(null)}>
                Cancel
              </Button>
              <Button
                appearance="primary"
                onClick={() => {
                  const id = pendingRename?.id;
                  const title = renameValue.trim();
                  setPendingRename(null);
                  if (id && title) props.onRename(id, title);
                }}
              >
                Save
              </Button>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>
    </div>
  );
}

export function formatRelativeTime(iso: string, now = Date.now()): string {
  const then = Date.parse(iso);
  if (!Number.isFinite(then)) return "";
  const delta = Math.max(0, now - then);
  if (delta < 60_000) return "just now";
  if (delta < 3_600_000) return `${Math.floor(delta / 60_000)}m ago`;
  if (delta < 86_400_000) return `${Math.floor(delta / 3_600_000)}h ago`;
  return new Date(then).toISOString().slice(0, 10);
}
