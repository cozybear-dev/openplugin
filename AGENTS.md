# Agent notes

## Live Office testing

After changing `packages/add-in`, `packages/host-*`, or Office.js apply/read paths, drive the real add-in. Do not ask the user to sideload.

This machine has Excel, Word, and PowerPoint desktop. The live driver attaches to the task pane WebView2 over CDP.

```
npm run live -- excel start --force
npm run live -- excel ui screenshot
npm run live -- excel host facts
npm run live -- excel app configure --mock
npm run live -- excel app send add a total
npm run live -- excel app apply
npm run live -- excel host range Sheet1 A4
npm run live -- excel stop --force
```

Or one shot: `npm run test:live` (Excel smoke: seed → mock LLM writeRange → apply → assert A4).

- Start Excel unless the bug is Word/PowerPoint-specific (`word` / `powerpoint` in place of `excel`).
- Assert document truth with `host range` / `host facts`, not pixels. Screenshot the task pane so layout and errors are visible.
- `--force` closes the Office host so it inherits `WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS`. Do not use it on a document the user is editing.
- If CDP attach fails, print the cert / loopback / "WebView Stop On Load" hint. Fake-host unit tests are not a substitute.
- Do not run two `live` commands at once. Each command connects over CDP; overlapping connects can drop the WebView.
- `--force` closes only that host (Excel / Word / PowerPoint). Switch hosts with `stop --force` first.
- Live start must not wait on "WebView Stop On Load". That dialog is the Office direct debugger; `live start` clears `UseDirectDebugger` and clicks Cancel if it still appears.

`packages/live/output/` holds task-pane screenshots from the driver.
