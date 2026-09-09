import type { HostKind, ToolDefinition } from "../llm/types.js";
import type { HostAdapter } from "../hosts/types.js";
import { truncateGrid } from "../hosts/types.js";
import type { Changeset } from "./changeset.js";

const META: ToolDefinition[] = [
  {
    type: "function",
    function: {
      name: "skills.list",
      description: "List installed skills (name + description only).",
      parameters: { type: "object", properties: {} }
    }
  },
  {
    type: "function",
    function: {
      name: "skills.load",
      description: "Load a skill's full SKILL.md instructions into context.",
      parameters: {
        type: "object",
        properties: { name: { type: "string" } },
        required: ["name"]
      }
    }
  },
  {
    type: "function",
    function: {
      name: "skills.readRef",
      description: "Read a file from an activated skill (references/ or assets/).",
      parameters: {
        type: "object",
        properties: {
          name: { type: "string" },
          path: { type: "string" }
        },
        required: ["name", "path"]
      }
    }
  },
  {
    type: "function",
    function: {
      name: "changeset.preview",
      description: "Show the pending document changeset the user will apply.",
      parameters: { type: "object", properties: {} }
    }
  }
];

const EXECUTE_JS: ToolDefinition = {
  type: "function",
  function: {
    name: "host.executeOfficeJs",
    description: "Gated escape hatch. Disabled unless the user enables advanced mode.",
    parameters: {
      type: "object",
      properties: { code: { type: "string" } },
      required: ["code"]
    }
  }
};

const BY_HOST: Record<HostKind, ToolDefinition[]> = {
  excel: [
    tool("excel.getSummary", "Workbook sheets, tables, and selection metadata — not cell values."),
    tool("excel.listSheets", "List worksheet names."),
    tool("excel.readRange", "Read a range. Large ranges are truncated.", {
      sheet: { type: "string" },
      address: { type: "string" }
    }, ["address"]),
    tool("excel.writeRange", "Queue a value write. Does not apply until the user confirms.", {
      sheet: { type: "string" },
      address: { type: "string" },
      values: { type: "array", items: { type: "array" } }
    }, ["sheet", "address", "values"]),
    tool("excel.setFormulas", "Queue formulas for a range.", {
      sheet: { type: "string" },
      address: { type: "string" },
      formulas: { type: "array", items: { type: "array", items: { type: "string" } } }
    }, ["sheet", "address", "formulas"]),
    tool("excel.listTables", "List tables in the workbook."),
    tool("excel.createTable", "Queue creating a table over a range.", {
      sheet: { type: "string" },
      address: { type: "string" },
      name: { type: "string" }
    }, ["sheet", "address"]),
    tool("excel.listCharts", "List charts."),
    tool("excel.createChart", "Queue a chart.", {
      sheet: { type: "string" },
      source: { type: "string" },
      chartType: { type: "string" }
    }, ["sheet", "source", "chartType"])
  ],
  word: [
    tool("word.getOutline", "Heading tree and paragraph count."),
    tool("word.getSelection", "Current selection text."),
    tool("word.replaceSelection", "Queue replacing the selection.", { text: { type: "string" } }, ["text"]),
    tool("word.insertParagraphs", "Queue inserting paragraphs.", {
      paragraphs: { type: "array", items: { type: "string" } },
      location: { type: "string", enum: ["start", "end", "afterSelection"] }
    }, ["paragraphs"]),
    tool("word.searchReplace", "Queue search and replace.", {
      search: { type: "string" },
      replace: { type: "string" },
      all: { type: "boolean" }
    }, ["search", "replace"]),
    tool("word.applyStyle", "Queue a style on the selection.", {
      style: { type: "string" },
      target: { type: "string", enum: ["selection", "heading"] }
    }, ["style"]),
    tool("word.insertTable", "Queue inserting a table.", {
      rows: { type: "number" },
      cols: { type: "number" }
    }, ["rows", "cols"]),
    tool("word.insertComment", "Queue a comment on the selection.", { text: { type: "string" } }, ["text"])
  ],
  powerpoint: [
    tool("ppt.getSlideTree", "Slide titles and shape counts."),
    tool("ppt.getSelection", "Selected shape text."),
    tool("ppt.setShapeText", "Queue setting shape text.", {
      slideIndex: { type: "number" },
      shapeName: { type: "string" },
      text: { type: "string" }
    }, ["slideIndex", "text"]),
    tool("ppt.addSlide", "Queue a new slide.", {
      title: { type: "string" },
      bullets: { type: "array", items: { type: "string" } }
    }, ["title"]),
    tool("ppt.setNotes", "Queue speaker notes.", {
      slideIndex: { type: "number" },
      notes: { type: "string" }
    }, ["slideIndex", "notes"]),
    tool("ppt.deleteSlide", "Queue deleting a slide.", { slideIndex: { type: "number" } }, ["slideIndex"])
  ]
};

function tool(
  name: string,
  description: string,
  properties: Record<string, unknown> = {},
  required: string[] = []
): ToolDefinition {
  return {
    type: "function",
    function: {
      name,
      description,
      parameters: { type: "object", properties, required }
    }
  };
}

export function listToolDefinitions(
  host: HostKind,
  opts: { executeJsEnabled?: boolean } = {}
): ToolDefinition[] {
  const list = [...META, ...BY_HOST[host]];
  if (opts.executeJsEnabled) list.push(EXECUTE_JS);
  return list;
}

export async function executeHostTool(
  host: HostAdapter,
  name: string,
  args: Record<string, unknown>,
  changeset: Changeset
): Promise<unknown> {
  switch (name) {
    case "excel.getSummary":
    case "excel.listSheets":
    case "word.getOutline":
    case "word.getSelection":
    case "ppt.getSlideTree":
    case "ppt.getSelection":
      return host.getRawFacts();
    case "excel.readRange": {
      if (!host.readRange) throw new Error("Host cannot read ranges.");
      return host.readRange({
        sheet: args.sheet as string | undefined,
        address: String(args.address)
      });
    }
    case "excel.listTables":
    case "excel.listCharts":
      return host.getRawFacts();
    case "excel.writeRange": {
      const sheet = String(args.sheet);
      const address = String(args.address);
      changeset.add({
        host: "excel",
        op: "writeRange",
        sheet,
        address,
        values: args.values as unknown[][],
        before: await captureGrid(host, sheet, address)
      });
      return { queued: true, preview: changeset.preview() };
    }
    case "excel.setFormulas": {
      const sheet = String(args.sheet);
      const address = String(args.address);
      changeset.add({
        host: "excel",
        op: "setFormulas",
        sheet,
        address,
        formulas: args.formulas as string[][],
        before: await captureGrid(host, sheet, address)
      });
      return { queued: true };
    }
    case "excel.createTable":
      changeset.add({
        host: "excel",
        op: "createTable",
        sheet: String(args.sheet),
        address: String(args.address),
        name: args.name as string | undefined
      });
      return { queued: true };
    case "excel.createChart":
      changeset.add({
        host: "excel",
        op: "createChart",
        sheet: String(args.sheet),
        source: String(args.source),
        chartType: String(args.chartType)
      });
      return { queued: true };
    case "word.replaceSelection":
      changeset.add({
        host: "word",
        op: "replaceSelection",
        text: String(args.text),
        beforeText: await captureText(host)
      });
      return { queued: true };
    case "word.insertParagraphs":
      changeset.add({
        host: "word",
        op: "insertParagraphs",
        paragraphs: args.paragraphs as string[],
        location: (args.location as "start" | "end" | "afterSelection") ?? "end"
      });
      return { queued: true };
    case "word.searchReplace":
      changeset.add({
        host: "word",
        op: "searchReplace",
        search: String(args.search),
        replace: String(args.replace),
        all: Boolean(args.all ?? true)
      });
      return { queued: true };
    case "word.applyStyle":
      changeset.add({
        host: "word",
        op: "applyStyle",
        style: String(args.style),
        target: (args.target as "selection" | "heading") ?? "selection"
      });
      return { queued: true };
    case "word.insertTable":
      changeset.add({
        host: "word",
        op: "insertTable",
        rows: Number(args.rows),
        cols: Number(args.cols)
      });
      return { queued: true };
    case "word.insertComment":
      changeset.add({ host: "word", op: "insertComment", text: String(args.text) });
      return { queued: true };
    case "ppt.setShapeText": {
      const slideIndex = Number(args.slideIndex);
      const shapeName = args.shapeName as string | undefined;
      changeset.add({
        host: "powerpoint",
        op: "setShapeText",
        slideIndex,
        shapeName,
        text: String(args.text),
        beforeText: host.readShapeText
          ? await host.readShapeText({ slideIndex, shapeName })
          : await captureText(host)
      });
      return { queued: true };
    }
    case "ppt.addSlide":
      changeset.add({
        host: "powerpoint",
        op: "addSlide",
        title: String(args.title),
        bullets: args.bullets as string[] | undefined
      });
      return { queued: true };
    case "ppt.setNotes": {
      const slideIndex = Number(args.slideIndex);
      changeset.add({
        host: "powerpoint",
        op: "setNotes",
        slideIndex,
        notes: String(args.notes),
        beforeText: host.readNotes ? await host.readNotes(slideIndex) : ""
      });
      return { queued: true };
    }
    case "ppt.deleteSlide":
      changeset.add({ host: "powerpoint", op: "deleteSlide", slideIndex: Number(args.slideIndex) });
      return { queued: true };
    case "host.executeOfficeJs":
      throw new Error("host.executeOfficeJs is disabled by policy.");
    default:
      return { truncated: false, values: truncateGrid([]).values };
  }
}

async function captureGrid(
  host: HostAdapter,
  sheet: string,
  address: string
): Promise<unknown[][] | undefined> {
  if (!host.readRange) return undefined;
  try {
    return (await host.readRange({ sheet, address })).values;
  } catch {
    return undefined;
  }
}

async function captureText(host: HostAdapter): Promise<string | undefined> {
  if (!host.readSelectionText) return undefined;
  try {
    return await host.readSelectionText();
  } catch {
    return undefined;
  }
}
