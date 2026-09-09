import type { HostKind, NativeSearchToolDefinition, ToolDefinition } from "../llm/types.js";
import type { HostAdapter } from "../hosts/types.js";
import { addressForGrid, normalizeGrid, truncateGrid } from "../hosts/types.js";
import { WEB_SEARCH_TOOLS } from "../search/index.js";
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
    tool(
      "excel.writeRange",
      "Queue a value write. Does not apply until the user confirms. `address` is the top-left (or any range); the actual write size is `values.length` × `values[0].length`. A 3×2 table at `A1` writes `A1:B3`.",
      {
        sheet: { type: "string" },
        address: { type: "string" },
        values: { type: "array", items: { type: "array" } }
      },
      ["sheet", "address", "values"]
    ),
    tool(
      "excel.setFormulas",
      "Queue formulas for a range. `address` is the top-left (or any range); the actual write size is `formulas.length` × `formulas[0].length`. A 3×2 table at `A1` writes `A1:B3`.",
      {
        sheet: { type: "string" },
        address: { type: "string" },
        formulas: { type: "array", items: { type: "array", items: { type: "string" } } }
      },
      ["sheet", "address", "formulas"]
    ),
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
    }, ["sheet", "source", "chartType"]),
    tool("excel.readCsv", "Read a range as CSV. Preferred for analysis.", {
      sheet: { type: "string" },
      address: { type: "string" }
    }, ["address"]),
    tool("excel.search", "Find text in the workbook and return matching addresses.", {
      query: { type: "string" },
      sheet: { type: "string" }
    }, ["query"]),
    tool("excel.listObjects", "List tables, charts, and pivot tables."),
    tool("excel.formatRange", "Queue number format or bold on a range.", {
      sheet: { type: "string" },
      address: { type: "string" },
      bold: { type: "boolean" },
      numberFormat: { type: "string" }
    }, ["sheet", "address"]),
    tool("excel.clearRange", "Queue clearing a range.", {
      sheet: { type: "string" },
      address: { type: "string" },
      clearType: { type: "string", enum: ["contents", "formats", "all"] }
    }, ["sheet", "address"]),
    tool("excel.copyRange", "Queue copying a range (formulas translated).", {
      sheet: { type: "string" },
      source: { type: "string" },
      dest: { type: "string" }
    }, ["sheet", "source", "dest"]),
    tool("excel.modifySheet", "Queue insert/delete/hide/freeze of rows or columns.", {
      sheet: { type: "string" },
      operation: { type: "string", enum: ["insert", "delete", "hide", "unhide", "freeze", "unfreeze"] },
      dimension: { type: "string", enum: ["rows", "columns"] },
      reference: { type: "string" },
      count: { type: "number" }
    }, ["sheet", "operation"]),
    tool("excel.modifyWorkbook", "Queue create/delete/rename/duplicate sheet.", {
      operation: { type: "string", enum: ["create", "delete", "rename", "duplicate"] },
      sheet: { type: "string" },
      newName: { type: "string" }
    }, ["operation"]),
    tool("excel.resizeRange", "Queue column width or row height.", {
      sheet: { type: "string" },
      address: { type: "string" },
      columnWidth: { type: "number" },
      rowHeight: { type: "number" }
    }, ["sheet", "address"]),
    tool("excel.createPivot", "Queue a pivot table.", {
      sheet: { type: "string" },
      source: { type: "string" },
      dest: { type: "string" },
      rows: { type: "array", items: { type: "string" } },
      columns: { type: "array", items: { type: "string" } },
      values: { type: "array" }
    }, ["sheet", "source", "dest", "rows", "values"])
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
    tool("word.insertComment", "Queue a comment on the selection.", { text: { type: "string" } }, ["text"]),
    tool("word.readParagraphs", "Read paragraphs by index range.", {
      start: { type: "number" },
      count: { type: "number" }
    }),
    tool("word.replaceParagraph", "Queue replacing a paragraph by index.", {
      index: { type: "number" },
      text: { type: "string" }
    }, ["index", "text"]),
    tool("word.find", "Find text in the document.", {
      query: { type: "string" },
      max: { type: "number" }
    }, ["query"]),
    tool("word.listComments", "List comments in the document."),
    tool("word.replyComment", "Queue a reply on a comment thread.", {
      commentIndex: { type: "number" },
      text: { type: "string" }
    }, ["commentIndex", "text"]),
    tool("word.resolveComment", "Queue resolving a comment.", {
      commentIndex: { type: "number" }
    }, ["commentIndex"]),
    tool("word.getRevisions", "Summarize tracked changes (read-only).")
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
    tool("ppt.deleteSlide", "Queue deleting a slide.", { slideIndex: { type: "number" } }, ["slideIndex"]),
    tool("ppt.readSlide", "Read all shape text on a slide.", { slideIndex: { type: "number" } }, ["slideIndex"]),
    tool("ppt.listLayouts", "List slide layouts from the template."),
    tool("ppt.duplicateSlide", "Queue duplicating a slide.", { slideIndex: { type: "number" } }, ["slideIndex"]),
    tool("ppt.reorderSlides", "Queue moving a slide.", {
      from: { type: "number" },
      to: { type: "number" }
    }, ["from", "to"]),
    tool("ppt.addChart", "Queue a native chart on a slide.", {
      slideIndex: { type: "number" },
      chartType: { type: "string" },
      categories: { type: "array", items: { type: "string" } },
      series: { type: "array" }
    }, ["slideIndex", "chartType", "categories", "series"])
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
  opts: {
    executeJsEnabled?: boolean;
    webSearch?: false | "function" | NativeSearchToolDefinition;
  } = {}
): ToolDefinition[] {
  const list: ToolDefinition[] = [...META, ...BY_HOST[host]];
  if (opts.executeJsEnabled) list.push(EXECUTE_JS);
  if (opts.webSearch === "function") list.push(...WEB_SEARCH_TOOLS);
  else if (opts.webSearch) list.push(opts.webSearch);
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
      const values = normalizeGrid(args.values);
      const address = addressForGrid(String(args.address), values);
      changeset.add({
        host: "excel",
        op: "writeRange",
        sheet,
        address,
        values,
        before: await captureGrid(host, sheet, address)
      });
      return { queued: true, preview: changeset.preview() };
    }
    case "excel.setFormulas": {
      const sheet = String(args.sheet);
      const formulas = normalizeGrid(args.formulas).map((row) => row.map((cell) => String(cell ?? "")));
      const address = addressForGrid(String(args.address), formulas);
      changeset.add({
        host: "excel",
        op: "setFormulas",
        sheet,
        address,
        formulas,
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
    case "excel.readCsv": {
      if (!host.readRange) throw new Error("Host cannot read ranges.");
      const grid = await host.readRange({
        sheet: args.sheet as string | undefined,
        address: String(args.address)
      });
      return { csv: toCsv(grid.values), truncated: grid.truncated };
    }
    case "excel.search": {
      if (host.search) {
        return host.search({ query: String(args.query), sheet: args.sheet as string | undefined });
      }
      return [];
    }
    case "excel.listObjects":
      return host.getRawFacts();
    case "excel.formatRange":
      changeset.add({
        host: "excel",
        op: "formatRange",
        sheet: String(args.sheet),
        address: String(args.address),
        bold: args.bold as boolean | undefined,
        numberFormat: args.numberFormat as string | undefined
      });
      return { queued: true };
    case "excel.clearRange": {
      const sheet = String(args.sheet);
      const address = String(args.address);
      changeset.add({
        host: "excel",
        op: "clearRange",
        sheet,
        address,
        clearType: (args.clearType as "contents" | "formats" | "all") ?? "contents",
        before: await captureGrid(host, sheet, address)
      });
      return { queued: true };
    }
    case "excel.copyRange": {
      const sheet = String(args.sheet);
      changeset.add({
        host: "excel",
        op: "copyRange",
        sheet,
        source: String(args.source),
        dest: String(args.dest),
        before: await captureGrid(host, sheet, String(args.dest))
      });
      return { queued: true };
    }
    case "excel.modifySheet":
      changeset.add({
        host: "excel",
        op: "modifySheet",
        sheet: String(args.sheet),
        operation: args.operation as "insert" | "delete" | "hide" | "unhide" | "freeze" | "unfreeze",
        dimension: args.dimension as "rows" | "columns" | undefined,
        reference: args.reference as string | undefined,
        count: args.count as number | undefined
      });
      return { queued: true };
    case "excel.modifyWorkbook":
      changeset.add({
        host: "excel",
        op: "modifyWorkbook",
        operation: args.operation as "create" | "delete" | "rename" | "duplicate",
        sheet: args.sheet as string | undefined,
        newName: args.newName as string | undefined
      });
      return { queued: true };
    case "excel.resizeRange":
      changeset.add({
        host: "excel",
        op: "resizeRange",
        sheet: String(args.sheet),
        address: String(args.address),
        columnWidth: args.columnWidth as number | undefined,
        rowHeight: args.rowHeight as number | undefined
      });
      return { queued: true };
    case "excel.createPivot":
      changeset.add({
        host: "excel",
        op: "createPivot",
        sheet: String(args.sheet),
        source: String(args.source),
        dest: String(args.dest),
        rows: (args.rows as string[]) ?? [],
        columns: args.columns as string[] | undefined,
        values: (args.values as Array<{ field: string; summarizeBy?: string }>) ?? []
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
    case "word.readParagraphs":
      if (host.readParagraphs) {
        return host.readParagraphs({ start: args.start as number | undefined, count: args.count as number | undefined });
      }
      return host.getRawFacts();
    case "word.replaceParagraph": {
      const index = Number(args.index);
      const paras = host.readParagraphs ? await host.readParagraphs({ start: index, count: 1 }) : [];
      changeset.add({
        host: "word",
        op: "replaceParagraph",
        index,
        text: String(args.text),
        beforeText: paras[0]?.text
      });
      return { queued: true };
    }
    case "word.find":
      if (host.findText) return host.findText({ query: String(args.query), max: args.max as number | undefined });
      return [];
    case "word.listComments":
      return host.listComments ? host.listComments() : [];
    case "word.replyComment":
      changeset.add({
        host: "word",
        op: "replyComment",
        commentIndex: Number(args.commentIndex),
        text: String(args.text)
      });
      return { queued: true };
    case "word.resolveComment":
      changeset.add({ host: "word", op: "resolveComment", commentIndex: Number(args.commentIndex) });
      return { queued: true };
    case "word.getRevisions":
      return host.getRevisions ? host.getRevisions() : [];
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
    case "ppt.readSlide":
      if (host.readSlide) return host.readSlide(Number(args.slideIndex));
      return host.getRawFacts();
    case "ppt.listLayouts":
      return host.listLayouts ? host.listLayouts() : [];
    case "ppt.duplicateSlide":
      changeset.add({ host: "powerpoint", op: "duplicateSlide", slideIndex: Number(args.slideIndex) });
      return { queued: true };
    case "ppt.reorderSlides":
      changeset.add({
        host: "powerpoint",
        op: "reorderSlides",
        from: Number(args.from),
        to: Number(args.to)
      });
      return { queued: true };
    case "ppt.addChart":
      changeset.add({
        host: "powerpoint",
        op: "addChart",
        slideIndex: Number(args.slideIndex),
        chartType: String(args.chartType),
        categories: (args.categories as string[]) ?? [],
        series: (args.series as Array<{ name: string; values: number[] }>) ?? []
      });
      return { queued: true };
    case "host.executeOfficeJs":
      throw new Error("host.executeOfficeJs is disabled by policy.");
    default:
      throw new Error(`Unknown host tool: ${name}`);
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

function toCsv(values: unknown[][]): string {
  return values
    .map((row) =>
      row
        .map((cell) => {
          const s = cell == null ? "" : String(cell);
          return /[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
        })
        .join(",")
    )
    .join("\n");
}

async function captureText(host: HostAdapter): Promise<string | undefined> {
  if (!host.readSelectionText) return undefined;
  try {
    return await host.readSelectionText();
  } catch {
    return undefined;
  }
}
