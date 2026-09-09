import {
  Button,
  Menu,
  MenuItem,
  MenuList,
  MenuPopover,
  MenuTrigger,
  Textarea,
  ToggleButton,
  Tooltip
} from "@fluentui/react-components";
import { Add24Regular, Globe24Regular, Send24Regular, Square24Filled } from "@fluentui/react-icons";
import { slashSuggestions, type ModelInfo, type SkillCatalogEntry } from "@openplugin/core";
import { ModelPicker } from "./ModelPicker";

export function Composer(props: {
  value: string;
  busy: boolean;
  skills: SkillCatalogEntry[];
  model: string;
  models: ModelInfo[];
  disabled: boolean;
  webSearch: boolean;
  onChange: (value: string) => void;
  onSend: () => void;
  onStop: () => void;
  onInsertSkill: (name: string) => void;
  onToggleWebSearch: (on: boolean) => void;
  onModelChange: (model: string) => void;
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
<div className="op-model-row"><span className="op-eyebrow">MODEL</span><ModelPicker label="Chat model" model={props.model} models={props.models} disabled={props.disabled || props.busy} placement="above" onChange={props.onModelChange} /></div>
      <div className="op-composer-box">
        <Textarea
          aria-label="Message OpenPlugin"
          value={props.value}
          disabled={props.disabled}
          textarea={{ className: "op-composer-input" }}
          placeholder="What would you like to work on?"
          onChange={(_, d) => props.onChange(d.value)}
          onKeyDown={(e) => {
            if (e.key === "Tab" && showSlash && suggestions[0]) {
              e.preventDefault();
              props.onInsertSkill(suggestions[0]);
              return;
            }
            if (e.key === "Enter" && !e.shiftKey && !e.nativeEvent.isComposing) {
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
                  <MenuItem disabled>No slash skills for this host</MenuItem>
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
          <ToggleButton
            className="op-web-toggle"
            size="small"
            appearance="subtle"
            checked={props.webSearch}
            icon={<Globe24Regular />}
            onClick={() => props.onToggleWebSearch(!props.webSearch)}
          >
            Web
          </ToggleButton>
          {props.busy ? (
            <Tooltip content="Stop" relationship="label">
              <Button appearance="subtle" icon={<Square24Filled />} onClick={props.onStop} />
            </Tooltip>
          ) : (
            <Button
              className="op-send"
              aria-label="Send message"
              appearance="primary"
              icon={<Send24Regular />}
              disabled={props.disabled || !props.value.trim()}
              onClick={props.onSend}
            >
            </Button>
          )}
        </div>
      </div>
      <div className="op-composer-hint"><span>Type <kbd>/</kbd> for skills</span><span>Shift + Enter for a new line</span></div>
    </div>
  );
}
