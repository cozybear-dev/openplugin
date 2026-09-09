import {
  Button,
  Menu,
  MenuItem,
  MenuList,
  MenuPopover,
  MenuTrigger,
  Textarea,
  Tooltip
} from "@fluentui/react-components";
import { Add24Regular, Send24Regular, Square24Filled } from "@fluentui/react-icons";
import type { SkillCatalogEntry } from "@openplugin/core";

export function Composer(props: {
  value: string;
  busy: boolean;
  skills: SkillCatalogEntry[];
  model: string;
  disabled: boolean;
  onChange: (value: string) => void;
  onSend: () => void;
  onStop: () => void;
  onInsertSkill: (name: string) => void;
}) {
  return (
    <div className="op-composer">
      <div className="op-composer-box">
        <Textarea
          value={props.value}
          disabled={props.disabled}
          textarea={{ className: "op-composer-input" }}
          placeholder="Ask OpenPlugin to edit the document…"
          onChange={(_, d) => props.onChange(d.value)}
          onKeyDown={(e) => {
            if (e.key === "Enter" && !e.shiftKey) {
              e.preventDefault();
              if (!props.busy) props.onSend();
            }
          }}
        />
        <div className="op-composer-bar">
          <Menu>
            <MenuTrigger disableButtonEnhancement>
              <Button appearance="subtle" size="small" icon={<Add24Regular />}>
                Skills
              </Button>
            </MenuTrigger>
            <MenuPopover>
              <MenuList>
                {props.skills.length === 0 ? (
                  <MenuItem disabled>No skills for this host</MenuItem>
                ) : (
                  props.skills.map((s) => (
                    <MenuItem key={s.name} onClick={() => props.onInsertSkill(s.name)}>
                      {s.name}
                    </MenuItem>
                  ))
                )}
              </MenuList>
            </MenuPopover>
          </Menu>
          <span className="op-model">{props.model || "No model"}</span>
          {props.busy ? (
            <Tooltip content="Stop" relationship="label">
              <Button appearance="subtle" icon={<Square24Filled />} onClick={props.onStop} />
            </Tooltip>
          ) : (
            <Button
              appearance="primary"
              icon={<Send24Regular />}
              disabled={props.disabled || !props.value.trim()}
              onClick={props.onSend}
            >
              Send
            </Button>
          )}
        </div>
      </div>
    </div>
  );
}
