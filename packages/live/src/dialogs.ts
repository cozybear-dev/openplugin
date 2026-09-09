import { spawn } from "node:child_process";
import { disableDebugging } from "office-addin-dev-settings";
import { ADDIN_ID } from "./constants.js";

export function isBlockingOfficeDialog(title: string): boolean {
  return /webview stop on load/i.test(title.trim());
}

export function buttonForBlockingDialog(title: string): "Cancel" | "OK" | null {
  if (isBlockingOfficeDialog(title)) return "Cancel";
  return null;
}

export async function disableOfficeJsDebugger(addinId = ADDIN_ID): Promise<void> {
  try {
    await disableDebugging(addinId);
  } catch (err) {
    console.warn("Could not clear Office add-in debugger flags.", err);
  }
  await setDebuggerDwords(addinId, {
    UseDirectDebugger: 0,
    UseWebDebugger: 0,
    OpenDevTools: 0
  });
}

function setDebuggerDwords(addinId: string, values: Record<string, number>): Promise<void> {
  const key = `HKCU\\SOFTWARE\\Microsoft\\Office\\16.0\\WEF\\Developer\\${addinId}`;
  return new Promise((resolve) => {
    const args = ["add", key, "/f"];
    const child = spawn("reg.exe", args, { windowsHide: true, stdio: "ignore" });
    child.on("close", () => {
      const writes = Object.entries(values).map(
        ([name, value]) =>
          new Promise<void>((done) => {
            const proc = spawn(
              "reg.exe",
              ["add", key, "/v", name, "/t", "REG_DWORD", "/d", String(value), "/f"],
              { windowsHide: true, stdio: "ignore" }
            );
            proc.on("close", () => done());
            proc.on("error", () => done());
          })
      );
      Promise.all(writes).then(() => resolve());
    });
    child.on("error", () => resolve());
  });
}

export function dismissBlockingOfficeDialogs(): Promise<string[]> {
  if (process.platform !== "win32") return Promise.resolve([]);
  const script = dismissScript();
  return new Promise((resolve) => {
    const child = spawn(
      "powershell.exe",
      ["-NoProfile", "-STA", "-Command", script],
      { windowsHide: true }
    );
    let out = "";
    child.stdout?.on("data", (chunk) => {
      out += String(chunk);
    });
    child.on("close", () => {
      const titles = out
        .split(/\r?\n/)
        .map((line) => line.trim())
        .filter(Boolean);
      resolve(titles);
    });
    child.on("error", () => resolve([]));
  });
}

export function watchBlockingDialogs(opts?: { intervalMs?: number; signal?: AbortSignal }): () => void {
  const intervalMs = opts?.intervalMs ?? 400;
  let stopped = false;
  const tick = async () => {
    if (stopped || opts?.signal?.aborted) return;
    const dismissed = await dismissBlockingOfficeDialogs();
    for (const title of dismissed) console.error(`[live] dismissed Office dialog: ${title}`);
    if (!stopped && !opts?.signal?.aborted) setTimeout(() => void tick(), intervalMs);
  };
  void tick();
  return () => {
    stopped = true;
  };
}

function dismissScript(): string {
  return `
Add-Type @"
using System;
using System.Text;
using System.Runtime.InteropServices;
public class LiveDialogs {
  public delegate bool EnumProc(IntPtr hWnd, IntPtr lParam);
  [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc cb, IntPtr l);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hWnd);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowText(IntPtr hWnd, StringBuilder s, int n);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern IntPtr FindWindowEx(IntPtr parent, IntPtr child, string cls, string win);
  [DllImport("user32.dll")] public static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);
  public const uint BM_CLICK = 0x00F5;
}
"@
function Click-NamedButton([IntPtr]$parent, [string]$name) {
  $btn = [LiveDialogs]::FindWindowEx($parent, [IntPtr]::Zero, "Button", $name)
  if ($btn -eq [IntPtr]::Zero) { $btn = [LiveDialogs]::FindWindowEx($parent, [IntPtr]::Zero, "Button", "&$name") }
  if ($btn -eq [IntPtr]::Zero) { return $false }
  [LiveDialogs]::SendMessage($btn, [LiveDialogs]::BM_CLICK, [IntPtr]::Zero, [IntPtr]::Zero) | Out-Null
  return $true
}
[LiveDialogs]::EnumWindows({
  param($h, $l)
  if (-not [LiveDialogs]::IsWindowVisible($h)) { return $true }
  $sb = New-Object System.Text.StringBuilder 512
  [LiveDialogs]::GetWindowText($h, $sb, $sb.Capacity) | Out-Null
  $title = $sb.ToString()
  if ($title -notmatch 'WebView Stop On Load') { return $true }
  if (Click-NamedButton $h 'Cancel') { Write-Output $title }
  return $true
}, [IntPtr]::Zero) | Out-Null
`;
}
