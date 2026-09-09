export const CDP_PORT = 9222;
export const CDP_URL = `http://127.0.0.1:${CDP_PORT}`;
export const VITE_PORT = 3000;
export const WEBVIEW_ARGS = `--remote-debugging-port=${CDP_PORT} --remote-allow-origins=*`;

export const OFFICE_EXES = {
  excel: "EXCEL.EXE",
  word: "WINWORD.EXE",
  powerpoint: "POWERPNT.EXE"
} as const;

export const OFFICE_PATHS = {
  excel: String.raw`C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE`,
  word: String.raw`C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE`,
  powerpoint: String.raw`C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE`
} as const;

export const ADDIN_ID = "7e8a4c2f-9b31-4d6e-a1f0-2c5d8e9b4a17";

export const RECOVERY_HINT = [
  "Office did not expose the add-in WebView.",
  "Close Excel/Word/PowerPoint completely (including background processes).",
  "If this is the first sideload: trust office-addin-dev-certs and allow WebView localhost loopback.",
  "Then: npm run live -- <host> start --force"
].join(" ");
