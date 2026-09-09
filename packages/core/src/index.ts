export { chatCompletions, chatUrl, classifyHttpError, listModels, LlmError, modelsUrl } from "./llm/client.js";
export type {
  ChatMessage,
  ChatResult,
  FunctionToolDefinition,
  HostKind,
  LlmErrorCode,
  ModelInfo,
  NativeSearchToolDefinition,
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
export { applyChangeset } from "./tools/apply.js";
export type { ApplyOpResult } from "./tools/apply.js";
export { executeHostTool, listToolDefinitions } from "./tools/registry.js";
export { cellDiff, invertChange, textDiff, type CellDelta, type DiffHunk, type TextPart } from "./tools/diff.js";
export { parseSlash, slashSuggestions } from "./skills/slash.js";
export type { SlashParse } from "./skills/slash.js";
export { describeActivity, describeTool } from "./tools/describe.js";
export type { ActivityDescription, ActivityKind } from "./tools/describe.js";
export { applyRestoreUndos, planRestore, resolveRestorePoint, truncateThread } from "./restore.js";
export type { RestorePoint, RestorableRevision, LineCheckpoint, RestoreUndoOutcome } from "./restore.js";
export {
  nativeSearchTool,
  resolveSearchBackend,
  runWebFetch,
  runWebSearch,
  supportsNativeSearch,
  WEB_SEARCH_TOOLS
} from "./search/index.js";
export type { ResolvedSearchBackend, SearchBackend, SearchHit, SearchProxy, SearchResult } from "./search/index.js";
export {
  MAP_CHUNK,
  buildExtractCall,
  buildMapCall,
  buildPromptCall,
  cacheKey,
  chunkRows,
  parseExtract,
  parseMappedArray,
  serializeRange
} from "./functions/excel-fn.js";

export { compileSnapshot, estimateTokens, snapshotToPrompt } from "./context/compiler.js";
export type { DocumentSnapshot } from "./context/compiler.js";

export { assertPolicy, mergePolicy, OPEN_POLICY } from "./policy.js";
export type { Policy } from "./policy.js";

export { runAgent, SYSTEM_PROMPT } from "./agent/loop.js";
export type { AgentEvent, AgentResult, WebSearchOptions } from "./agent/loop.js";

export type { HostAdapter, RawFacts } from "./hosts/types.js";
export {
  addressForGrid,
  addressToBounds,
  colLetter,
  excelMatrixAssign,
  normalizeGrid,
  sliceGrid,
  truncateGrid,
  writeIntoGrid
} from "./hosts/types.js";
export { FakeExcelHost } from "./hosts/fake-excel.js";
export { FakeWordHost } from "./hosts/fake-word.js";
export { FakePowerPointHost } from "./hosts/fake-powerpoint.js";
