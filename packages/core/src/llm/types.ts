export type HostKind = "excel" | "word" | "powerpoint";

export type ProviderConfig = {
  baseUrl: string;
  apiKey?: string;
  model: string;
  headers?: Record<string, string>;
  timeoutMs?: number;
  maxOutputTokens?: number;
};

export type ModelInfo = { id: string; name?: string };

export type FunctionToolDefinition = {
  type: "function";
  function: {
    name: string;
    description: string;
    parameters: Record<string, unknown>;
  };
};

export type NativeSearchToolDefinition =
  | { type: "openrouter:web_search"; parameters?: Record<string, unknown> }
  | { type: "web_search" };

export type ToolDefinition = FunctionToolDefinition | NativeSearchToolDefinition;

export type ToolCall = {
  id: string;
  type: "function";
  function: { name: string; arguments: string };
};

export type ChatMessage =
  | { role: "system"; content: string }
  | { role: "user"; content: string }
  | { role: "assistant"; content: string | null; tool_calls?: ToolCall[] }
  | { role: "tool"; tool_call_id: string; content: string };

export type StreamEvent =
  | { type: "text"; delta: string }
  | { type: "tool_call"; index: number; id?: string; name?: string; argumentsDelta: string }
  | { type: "done"; message: ChatMessage }
  | { type: "error"; error: LlmError };

export type ChatResult = {
  message: Extract<ChatMessage, { role: "assistant" }>;
  finishReason?: string;
};

export type LlmErrorCode =
  | "cors"
  | "auth"
  | "context_length"
  | "no_tools"
  | "network"
  | "http"
  | "abort"
  | "parse";

export class LlmError extends Error {
  readonly name = "LlmError";
  constructor(
    message: string,
    readonly code: LlmErrorCode,
    readonly status?: number
  ) {
    super(message);
  }
}
