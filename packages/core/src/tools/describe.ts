export function describeTool(name: string, args: Record<string, unknown> = {}): string {
  const address = args.address ? String(args.address) : "";
  const sheet = args.sheet ? String(args.sheet) : "";
  const range = sheet && address ? `${sheet}!${address}` : address || sheet;
  switch (name) {
    case "excel.readRange":
      return range ? `Reading ${range}` : "Reading range";
    case "excel.writeRange":
      return range ? `Drafting write to ${range}` : "Drafting a write";
    case "excel.getSummary":
    case "excel.listSheets":
      return "Scanning workbook";
    case "excel.setFormulas":
      return range ? `Drafting formulas in ${range}` : "Drafting formulas";
    case "excel.createTable":
      return "Drafting a table";
    case "excel.createChart":
      return "Drafting a chart";
    case "word.getSelection":
      return "Reading selection";
    case "word.getOutline":
      return "Reading outline";
    case "word.replaceSelection":
      return "Drafting a replacement";
    case "word.insertParagraphs":
      return "Drafting paragraphs";
    case "ppt.getSlideTree":
      return "Reading slides";
    case "ppt.getSelection":
      return "Reading selected shape";
    case "ppt.addSlide":
      return args.title ? `Drafting slide “${args.title}”` : "Drafting a slide";
    case "ppt.setShapeText":
      return "Drafting slide text";
    case "skills.load":
      return args.name ? `Using skill ${args.name}` : "Loading skill";
    case "skills.list":
      return "Listing skills";
    case "changeset.preview":
      return "Preparing review";
    default:
      return name.replace(/^[a-z]+\./, "").replace(/([A-Z])/g, " $1").trim();
  }
}
