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
import { slashSuggestions, type SkillCatalogEntry } from "@openplugin/core";

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
  const names = props.skills.map((s) => s.name);
  const suggestions = slashSuggestions(props.value, names);
  const showSlash = props.value.startsWith("/") && !props.value.includes("\n");

  return (
    <div className="op-composer">
      {showSlash && suggestions.length > 0 && (
        <div className="op-slash">
          {suggestions.map((name) => (
            <button
              key={name}
              type="button"
              className="op-slash-item"
              onMouseDown={(e) => {
                e.preventDefault();
                props.onInsertSkill(name);
              }}
            >
              /{name}
              <span>{props.skills.find((s) => s.name === name)?.description}</span>
            </button>
          ))}
        </div>
      )}
      <div className="op-composer-box">
        <Textarea
          value={props.value}
          disabled={props.disabled}
          textarea={{ className: "op-composer-input" }}
          placeholder="Ask OpenPlugin, or type / for a skill…"
          onChange={(_, d) => props.onChange(d.value)}
          onKeyDown={(e) => {
            if (e.key === "Tab" && showSlash && suggestions[0]) {
              e.preventDefault();
              props.onInsertSkill(suggestions[0]);
              return;
            }
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
                      /{s.name}
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
