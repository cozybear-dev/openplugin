import { MOCK_LLM_PORT } from "./mock-llm.js";
import { startMockLlm } from "./server.js";

await startMockLlm(MOCK_LLM_PORT);
console.log(`OpenPlugin mock LLM on http://127.0.0.1:${MOCK_LLM_PORT}/v1`);
