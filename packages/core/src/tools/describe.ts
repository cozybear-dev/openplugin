export type ActivityKind = "read" | "write" | "skill" | "search" | "fetch" | "tool";

export type ActivityDescription = {
  activity: ActivityKind;
  label: string;
  detail?: string;
};

function rangeLabel(args: Record<string, unknown>): string {
  const address = args.address ? String(args.address) : "";
  const sheet = args.sheet ? String(args.sheet) : "";
  return sheet && address ? `${sheet}!${address}` : address || sheet;
}

function hostName(name: string): string {
  return name.split(".")[0] ?? name;
}

export function describeActivity(name: string, args: Record<string, unknown> = {}): ActivityDescription {
  const range = rangeLabel(args);
  if (name === "web.search") {
    const query = String(args.query ?? "").trim();
    return { activity: "search", label: query ? `Searching the web: ${query}` : "Searching the web", detail: query };
  }
  if (name === "web.fetch") {
    const url = String(args.url ?? "");
    let host = url;
    try {
      host = url ? new URL(url).host + new URL(url).pathname : "page";
    } catch {
      host = url || "page";
    }
    return { activity: "fetch", label: `Opening ${host}`, detail: url };
  }
  if (name === "skills.load" || name === "skills.readRef") {
    const skill = String(args.name ?? "");
    return {
      activity: "skill",
      label: skill ? `Reading skill ${skill}` : "Reading skill",
      detail: args.path ? String(args.path) : skill
    };
  }
  if (name === "skills.list") return { activity: "skill", label: "Listing skills" };

  const writes = /^(excel|word|ppt)\.(write|set|create|insert|replace|apply|delete|clear|copy|modify|format|resize|add)/.test(
    name
  );
  const reads = /^(excel|word|ppt)\.(get|read|list|search|find)/.test(name);

  const label = describeTool(name, args);
  if (writes) return { activity: "write", label, detail: range || undefined };
  if (reads) return { activity: "read", label, detail: range || undefined };
  return { activity: "tool", label };
}

export function describeTool(name: string, args: Record<string, unknown> = {}): string {
  const address = args.address ? String(args.address) : "";
  const sheet = args.sheet ? String(args.sheet) : "";
  const range = sheet && address ? `${sheet}!${address}` : address || sheet;
  switch (name) {
    case "excel.readRange":
    case "excel.readCsv":
      return range ? `Reading ${range}` : "Reading range";
    case "excel.writeRange":
      return range ? `Drafting write to ${range}` : "Drafting a write";
    case "excel.getSummary":
    case "excel.listSheets":
    case "excel.listObjects":
      return "Scanning workbook";
    case "excel.search":
      return args.query ? `Searching workbook for “${args.query}”` : "Searching workbook";
    case "excel.setFormulas":
      return range ? `Drafting formulas in ${range}` : "Drafting formulas";
    case "excel.formatRange":
      return range ? `Drafting format for ${range}` : "Drafting formatting";
    case "excel.clearRange":
      return range ? `Drafting clear of ${range}` : "Drafting a clear";
    case "excel.copyRange":
      return "Drafting a copy";
    case "excel.modifySheet":
      return "Drafting sheet structure change";
    case "excel.modifyWorkbook":
      return "Drafting workbook structure change";
    case "excel.resizeRange":
      return range ? `Resizing ${range}` : "Drafting a resize";
    case "excel.createTable":
      return "Drafting a table";
    case "excel.createChart":
      return "Drafting a chart";
    case "excel.createPivot":
      return "Drafting a pivot table";
    case "word.getSelection":
      return "Reading selection";
    case "word.getOutline":
      return "Reading outline";
    case "word.readParagraphs":
      return "Reading paragraphs";
    case "word.find":
      return args.query ? `Finding “${args.query}”` : "Finding in document";
    case "word.listComments":
      return "Reading comments";
    case "word.getRevisions":
      return "Reading tracked changes";
    case "word.replaceSelection":
      return "Drafting a replacement";
    case "word.replaceParagraph":
      return "Drafting a paragraph replacement";
    case "word.insertParagraphs":
      return "Drafting paragraphs";
    case "word.replyComment":
      return "Drafting a comment reply";
    case "word.resolveComment":
      return "Resolving a comment";
    case "ppt.getSlideTree":
      return "Reading slides";
    case "ppt.getSelection":
      return "Reading selected shape";
    case "ppt.readSlide":
      return typeof args.slideIndex === "number" ? `Reading slide ${Number(args.slideIndex) + 1}` : "Reading slide";
    case "ppt.listLayouts":
      return "Reading layouts";
    case "ppt.addSlide":
      return args.title ? `Drafting slide “${args.title}”` : "Drafting a slide";
    case "ppt.setShapeText":
      return "Drafting slide text";
    case "ppt.duplicateSlide":
      return "Drafting a duplicated slide";
    case "ppt.reorderSlides":
      return "Drafting slide reorder";
    case "ppt.addChart":
      return "Drafting a slide chart";
    case "skills.load":
      return args.name ? `Using skill ${args.name}` : "Loading skill";
    case "skills.list":
      return "Listing skills";
    case "web.search":
      return args.query ? `Searching the web: ${args.query}` : "Searching the web";
    case "web.fetch":
      return args.url ? `Opening ${args.url}` : "Opening a page";
    case "changeset.preview":
      return "Preparing review";
    default:
      return name.replace(/^[a-z]+\./, "").replace(/([A-Z])/g, " $1").trim() || hostName(name);
  }
}
