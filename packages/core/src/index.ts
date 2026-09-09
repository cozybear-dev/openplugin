export { chatCompletions, chatUrl, classifyHttpError, LlmError } from "./llm/client.js";
export type {
  ChatMessage,
  ChatResult,
  HostKind,
  LlmErrorCode,
  ProviderConfig,
  StreamEvent,
  ToolCall,
  ToolDefinition
} from "./llm/types.js";

export { parseSkillMarkdown, SkillParseError } from "./skills/parse.js";
export type { Skill, SkillMeta } from "./skills/parse.js";
export { SkillRegistry } from "./skills/registry.js";
export type { SkillCatalogEntry } from "./skills/registry.js";

export { Changeset } from "./tools/changeset.js";
export type { Change, ExcelChange, PptChange, WordChange } from "./tools/changeset.js";
export { executeHostTool, listToolDefinitions } from "./tools/registry.js";

export { compileSnapshot, estimateTokens, snapshotToPrompt } from "./context/compiler.js";
export type { DocumentSnapshot } from "./context/compiler.js";

export { assertPolicy, OPEN_POLICY } from "./policy.js";
export type { Policy } from "./policy.js";

export { runAgent, SYSTEM_PROMPT } from "./agent/loop.js";
export type { AgentEvent, AgentResult } from "./agent/loop.js";

export type { HostAdapter, RawFacts } from "./hosts/types.js";
export { sliceGrid, truncateGrid, writeIntoGrid } from "./hosts/types.js";
export { FakeExcelHost } from "./hosts/fake-excel.js";
export { FakeWordHost } from "./hosts/fake-word.js";
export { FakePowerPointHost } from "./hosts/fake-powerpoint.js";
