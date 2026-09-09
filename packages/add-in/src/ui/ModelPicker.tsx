import { Combobox, Option } from "@fluentui/react-components";
import { useMemo, useState } from "react";
import type { ModelInfo } from "@openplugin/core";

/** Search is a draft. Only choosing an option changes the active model. */
export function ModelPicker(props: {
  label: string;
  model: string;
  models: ModelInfo[];
  disabled?: boolean;
  loading?: boolean;
  placement?: "above" | "below";
  onChange: (id: string) => void;
}) {
  const [open, setOpen] = useState(false);
  const [query, setQuery] = useState("");
  const options = useMemo(() => {
    const byId = new Map(props.models.map((model) => [model.id, model]));
    if (props.model && !byId.has(props.model)) byId.set(props.model, { id: props.model });
    return [...byId.values()];
  }, [props.models, props.model]);
  const selected = options.find((model) => model.id === props.model);
  const search = query.trim().toLowerCase();
  const filtered = options.filter((model) => `${model.id} ${model.name ?? ""}`.toLowerCase().includes(search));
  const customId = query.trim();
  const showCustom = customId && !options.some((model) => model.id === customId || model.name === customId);

  return (
    <Combobox
      className="op-model-picker"
      aria-label={props.label}
      title={props.model || "Choose a model"}
      freeform
      autoComplete="off"
      open={open}
      value={open ? query : selected?.name ?? props.model}
      selectedOptions={props.model ? [props.model] : []}
      disabled={props.disabled}
      placeholder={open ? "Search models or enter an ID…" : props.loading ? "Loading models…" : "Choose a model…"}
      positioning={{ position: props.placement ?? "below", align: "start", matchTargetSize: "width" }}
      listbox={{ className: "op-model-options" }}
      onOpenChange={(_, data) => {
        setOpen(data.open);
        if (data.open) setQuery("");
      }}
      onChange={(event) => { setQuery(event.target.value); setOpen(true); }}
      onOptionSelect={(_, data) => {
        if (!data.optionValue) return;
        props.onChange(data.optionValue);
        setOpen(false);
        setQuery("");
      }}
    >
      {filtered.map((model) => (
        <Option key={model.id} value={model.id} text={model.name ?? model.id}>
          <span className="op-model-option"><strong>{model.name ?? model.id}</strong>{model.name && model.name !== model.id && <small>{model.id}</small>}</span>
        </Option>
      ))}
      {showCustom && <Option value={customId} text={customId}>Use model ID: {customId}</Option>}
      {!filtered.length && !showCustom && <Option disabled>{props.loading ? "Loading models…" : "Enter a model ID to get started"}</Option>}
    </Combobox>
  );
}
