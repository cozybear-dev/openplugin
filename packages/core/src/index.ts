export {
  chatCompletions,
  chatUrl,
  classifyHttpError,
  isTruncatedFinish,
  listModels,
  LlmError,
  modelsUrl
} from "./llm/client.js";
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

export { parseSkillMarkdown, serializeSkill, SkillParseError, SKILL_HOSTS } from "./skills/parse.js";
export type { Skill, SkillMeta } from "./skills/parse.js";
export { SkillRegistry } from "./skills/registry.js";
export type { SkillCatalogEntry } from "./skills/registry.js";
export { skillsForAgent, uniqueSkillName } from "./skills/manage.js";
export {
  compactTranscript,
  extractSkillMarkdown,
  skillifyMessages,
  stubSkillMarkdown,
  SKILLIFY_MAX_CHARS,
  SKILLIFY_MAX_TURNS
} from "./skills/skillify.js";
export type { SkillifyTurn } from "./skills/skillify.js";

export { Changeset } from "./tools/changeset.js";
export type { Change, ExcelChange, PptChange, WordChange } from "./tools/changeset.js";
export { applyChangeset, applyFailureMessage } from "./tools/apply.js";
export type { ApplyOpResult } from "./tools/apply.js";
export { rebaseExcelChanges } from "./tools/rebase.js";
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
  buildTranslateCall,
  cacheKey,
  chunkRows,
  flattenCells,
  parseExtract,
  parseMappedArray,
  reshapeCells,
  serializeRange
} from "./functions/excel-fn.js";

export { compileSnapshot, estimateTokens, snapshotToPrompt } from "./context/compiler.js";
export type { DocumentSnapshot } from "./context/compiler.js";

export { assertPolicy, mergePolicy, OPEN_POLICY } from "./policy.js";
export type { Policy } from "./policy.js";

export { CONTINUE_NUDGE, runAgent, SYSTEM_PROMPT } from "./agent/loop.js";
export type { AgentEvent, AgentResult, WebSearchOptions } from "./agent/loop.js";

export type { HostAdapter, RawFacts } from "./hosts/types.js";
export {
  addressForGrid,
  addressToBounds,
  boundsToAddress,
  canonicalizeExcelAddress,
  cellAddressInUsedRange,
  colLetter,
  excelMatrixAssign,
  formatApplySpec,
  normalizeGrid,
  numberFormatMatrix,
  sheetQualifiedAddress,
  sliceGrid,
  truncateGrid,
  writeIntoGrid,
  EXCEL_MAX_COLS,
  EXCEL_MAX_ROWS,
  FORMAT_CELL_CAP
} from "./hosts/types.js";
export type { ExcelBounds } from "./hosts/types.js";
export { FakeExcelHost } from "./hosts/fake-excel.js";
export { FakeWordHost } from "./hosts/fake-word.js";
export { FakePowerPointHost } from "./hosts/fake-powerpoint.js";
