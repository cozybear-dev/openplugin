import { describe, expect, it } from "vitest";
import { completionForMessages, sseBody } from "../src/mock-llm.js";

const user = (content: string) => ({ role: "user" as const, content });
const tool = (content: string) => ({ role: "tool" as const, content, tool_call_id: "call_write_total" });

describe("completionForMessages", () => {
  it("queues excel.writeRange when the user asks for a total", () => {
    const result = completionForMessages([user("Please add a total row")]);
    expect(result.tool_calls?.[0]?.function.name).toBe("excel.writeRange");
    const args = JSON.parse(result.tool_calls![0].function.arguments) as {
      sheet: string;
      address: string;
      values: unknown[][];
    };
    expect(args.sheet).toBe("Sheet1");
    expect(args.address).toBe("A4");
    expect(args.values[0]?.[0]).toBe("Total");
    expect(result.content).toBeNull();
  });

  it("returns a final reply after the writeRange tool result", () => {
    const result = completionForMessages([
      user("add a total"),
      { role: "assistant", content: null, tool_calls: [{ id: "call_write_total" }] },
      tool("queued")
    ]);
    expect(result.tool_calls).toBeUndefined();
    expect(result.content).toMatch(/apply/i);
  });

  it("answers connection tests with ok", () => {
    expect(completionForMessages([user("Reply with ok")]).content).toBe("ok");
  });

  it("queues a Word insert when asked for a paragraph", () => {
    const result = completionForMessages([user("insert a hello paragraph")]);
    expect(result.tool_calls?.[0]?.function.name).toBe("word.insertParagraphs");
  });

  it("queues a PowerPoint slide when asked to add a slide", () => {
    const result = completionForMessages([user("add a slide titled Outline")]);
    expect(result.tool_calls?.[0]?.function.name).toBe("ppt.addSlide");
  });

  it("queues a Summary sheet create when asked", () => {
    const result = completionForMessages([user("Create a new worksheet named Summary")]);
    expect(result.tool_calls?.map((c) => c.function.name)).toEqual([
      "excel.modifyWorkbook",
      "excel.writeRange"
    ]);
  });

  it("queues applyStyle by paragraph index for a heading request", () => {
    const result = completionForMessages([user("Turn the Weekly Update paragraph into Heading 1")]);
    expect(result.tool_calls?.[0]?.function.name).toBe("word.applyStyle");
    expect(JSON.parse(result.tool_calls![0].function.arguments)).toMatchObject({
      style: "Heading 1",
      paragraphIndex: 2
    });
  });

  it("queues searchReplace for Q3 to Q4", () => {
    const result = completionForMessages([user("Replace every occurrence of Q3 with Q4")]);
    expect(result.tool_calls?.[0]?.function.name).toBe("word.searchReplace");
  });

  it("queues speaker notes and slide reorder before generic addSlide", () => {
    expect(completionForMessages([user("Set speaker notes on slide 1")]).tool_calls?.[0]?.function.name).toBe(
      "ppt.setNotes"
    );
    expect(completionForMessages([user("Move the last slide to the first position")]).tool_calls?.[0]?.function.name).toBe(
      "ppt.reorderSlides"
    );
  });
});

describe("sseBody", () => {
  it("emits OpenAI-style tool_call deltas then DONE", () => {
    const body = sseBody({
      content: null,
      tool_calls: [
        {
          id: "call_1",
          type: "function",
          function: { name: "excel.writeRange", arguments: "{\"sheet\":\"Sheet1\"}" }
        }
      ]
    });
    expect(body).toContain("tool_calls");
    expect(body).toContain("excel.writeRange");
    expect(body.trim().endsWith("data: [DONE]")).toBe(true);
  });
});
