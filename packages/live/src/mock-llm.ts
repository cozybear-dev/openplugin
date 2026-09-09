export type ChatTurn = {
  role: string;
  content?: unknown;
  tool_calls?: unknown;
};

export type Completion = {
  content: string | null;
  tool_calls?: Array<{
    id: string;
    type: "function";
    function: { name: string; arguments: string };
  }>;
};

function textOf(content: unknown): string {
  if (typeof content === "string") return content;
  if (Array.isArray(content)) {
    return content
      .map((part) => (part && typeof part === "object" && "text" in part ? String(part.text) : ""))
      .join(" ");
  }
  return content == null ? "" : String(content);
}

export function completionForMessages(messages: ChatTurn[]): Completion {
  let lastUserIndex = -1;
  for (let i = messages.length - 1; i >= 0; i--) {
    if (messages[i]?.role === "user") {
      lastUserIndex = i;
      break;
    }
  }
  const lastUser = lastUserIndex >= 0 ? messages[lastUserIndex] : undefined;
  const afterUser = lastUserIndex >= 0 ? messages.slice(lastUserIndex + 1) : [];
  const hasToolAfter = afterUser.some((m) => m.role === "tool");
  const text = textOf(lastUser?.content).toLowerCase();

  if (hasToolAfter) {
    return { content: "Queued the change. Review and apply it." };
  }

  if (/heading 1|weekly update/.test(text)) {
    return {
      content: null,
      tool_calls: [
        {
          id: "call_apply_style",
          type: "function",
          function: {
            name: "word.applyStyle",
            arguments: JSON.stringify({ style: "Heading 1", paragraphIndex: 2 })
          }
        }
      ]
    };
  }

  if (/replace every.*q3|q3 with q4/.test(text)) {
    return {
      content: null,
      tool_calls: [
        {
          id: "call_search_replace",
          type: "function",
          function: {
            name: "word.searchReplace",
            arguments: JSON.stringify({ search: "Q3", replace: "Q4", all: true })
          }
        }
      ]
    };
  }

  if (/summary/.test(text) && /sheet|worksheet/.test(text)) {
    return {
      content: null,
      tool_calls: [
        {
          id: "call_create_summary",
          type: "function",
          function: {
            name: "excel.modifyWorkbook",
            arguments: JSON.stringify({ operation: "create", newName: "Summary" })
          }
        },
        {
          id: "call_write_summary",
          type: "function",
          function: {
            name: "excel.writeRange",
            arguments: JSON.stringify({ sheet: "Summary", address: "A1", values: [["Sales summary"]] })
          }
        }
      ]
    };
  }

  if (/speaker notes|set notes/.test(text)) {
    return {
      content: null,
      tool_calls: [
        {
          id: "call_set_notes",
          type: "function",
          function: {
            name: "ppt.setNotes",
            arguments: JSON.stringify({ slideIndex: 0, notes: "Welcome the team" })
          }
        }
      ]
    };
  }

  if (/move the last slide|reorder/.test(text)) {
    return {
      content: null,
      tool_calls: [
        {
          id: "call_reorder",
          type: "function",
          function: {
            name: "ppt.reorderSlides",
            arguments: JSON.stringify({ from: 1, to: 0 })
          }
        }
      ]
    };
  }

  if (/total/.test(text)) {
    return {
      content: null,
      tool_calls: [
        {
          id: "call_write_total",
          type: "function",
          function: {
            name: "excel.writeRange",
            arguments: JSON.stringify({
              sheet: "Sheet1",
              address: "A4",
              values: [["Total", "", 42]]
            })
          }
        }
      ]
    };
  }

  if (/paragraph/.test(text)) {
    return {
      content: null,
      tool_calls: [
        {
          id: "call_insert_para",
          type: "function",
          function: {
            name: "word.insertParagraphs",
            arguments: JSON.stringify({ paragraphs: ["Hello from OpenPlugin live."], location: "end" })
          }
        }
      ]
    };
  }

  if (/\bslide\b/.test(text)) {
    return {
      content: null,
      tool_calls: [
        {
          id: "call_add_slide",
          type: "function",
          function: {
            name: "ppt.addSlide",
            arguments: JSON.stringify({ title: "Outline", bullets: ["Live test"] })
          }
        }
      ]
    };
  }

  if (/hello|\bok\b|reply with ok/.test(text)) {
    return { content: "ok" };
  }

  return { content: "Ready." };
}

export function sseBody(completion: Completion): string {
  const chunks: string[] = [];
  if (completion.tool_calls?.length) {
    completion.tool_calls.forEach((tc, index) => {
      chunks.push(
        `data: ${JSON.stringify({
          choices: [
            {
              delta: {
                tool_calls: [
                  {
                    index,
                    id: tc.id,
                    type: "function",
                    function: { name: tc.function.name, arguments: "" }
                  }
                ]
              }
            }
          ]
        })}\n\n`
      );
      chunks.push(
        `data: ${JSON.stringify({
          choices: [
            {
              delta: {
                tool_calls: [{ index, function: { arguments: tc.function.arguments } }]
              }
            }
          ]
        })}\n\n`
      );
    });
    chunks.push(`data: ${JSON.stringify({ choices: [{ finish_reason: "tool_calls", delta: {} }] })}\n\n`);
  } else {
    const content = completion.content ?? "";
    chunks.push(`data: ${JSON.stringify({ choices: [{ delta: { content } }] })}\n\n`);
    chunks.push(`data: ${JSON.stringify({ choices: [{ finish_reason: "stop", delta: {} }] })}\n\n`);
  }
  chunks.push("data: [DONE]\n\n");
  return chunks.join("");
}

export function jsonBody(completion: Completion, model = "openplugin-live"): string {
  return JSON.stringify({
    id: "chatcmpl-live",
    object: "chat.completion",
    model,
    choices: [
      {
        index: 0,
        message: {
          role: "assistant",
          content: completion.content ?? "",
          ...(completion.tool_calls ? { tool_calls: completion.tool_calls } : {})
        },
        finish_reason: completion.tool_calls ? "tool_calls" : "stop"
      }
    ]
  });
}

export const MOCK_LLM_PORT = 8790;
export const MOCK_LLM_ORIGIN = `http://127.0.0.1:${MOCK_LLM_PORT}/v1`;
