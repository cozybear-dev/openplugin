import {
  Button,
  Combobox,
  Menu,
  MenuItem,
  MenuList,
  MenuPopover,
  MenuTrigger,
  Option,
  Textarea,
  ToggleButton,
  Tooltip
} from "@fluentui/react-components";
import { Add24Regular, Globe24Regular, Send24Regular, Square24Filled } from "@fluentui/react-icons";
import { slashSuggestions, type ModelInfo, type SkillCatalogEntry } from "@openplugin/core";
import { useEffect, useMemo, useState } from "react";

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
  const [modelText, setModelText] = useState(props.model);
  useEffect(() => {
    setModelText(props.model);
  }, [props.model]);
  const filteredModels = useMemo(
    () => filterModels(props.models, modelText, props.model),
    [props.models, modelText, props.model]
  );

  function commitTypedModel() {
    const typed = modelText.trim();
    const id = typed ? modelIdFromInput(props.models, typed) : "";
    setModelText(id);
    if (id !== props.model) props.onModelChange(id);
  }

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
                Add
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
          <div className="op-model">
            <Combobox
              aria-label="Model"
              size="small"
              freeform
              placeholder="No model"
              value={modelText}
              selectedOptions={props.model ? [props.model] : []}
              disabled={props.disabled}
              onOptionSelect={(_, data) => {
                if (data.optionValue == null) return;
                setModelText(data.optionValue);
                if (data.optionValue !== props.model) props.onModelChange(data.optionValue);
              }}
              onChange={(e) => setModelText(e.target.value)}
              onBlur={commitTypedModel}
              onKeyDown={(e) => {
                if (e.key === "Enter") commitTypedModel();
              }}
            >
              {filteredModels.map((m) => (
                <Option key={m.id} value={m.id} text={m.name ?? m.id}>
                  {m.name ?? m.id}
                </Option>
              ))}
            </Combobox>
          </div>
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

function modelIdFromInput(models: ModelInfo[], raw: string): string {
  const exact = models.find((m) => m.id === raw);
  if (exact) return exact.id;
  const byName = models.filter((m) => (m.name ?? m.id) === raw);
  return byName.length === 1 ? byName[0]!.id : raw;
}

function filterModels(models: ModelInfo[], query: string, selected: string): ModelInfo[] {
  const q = query.trim().toLowerCase();
  if (!q || q === selected.trim().toLowerCase()) return models;
  return models.filter(
    (m) => m.id.toLowerCase().includes(q) || (m.name?.toLowerCase().includes(q) ?? false)
  );
}
