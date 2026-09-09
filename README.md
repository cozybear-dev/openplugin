# OpenPlugin

Open-source LLM agent for **Excel**, **Word**, and **PowerPoint**. One Office.js add-in. Any OpenAI-compatible endpoint. [Agent Skills](https://agentskills.io/specification) as the extension model.

The add-in does **not** proxy your documents through a vendor cloud. Your API key stays in this Office profile. The model sees a token-budgeted snapshot plus the tools you approve.

## Why it exists

Chat sidebars are a crowded market. OpenPlugin is a **document runtime for agents**:

- Typed Office tools and a previewable changeset (not generated JavaScript by default)
- Portable `SKILL.md` packs, not a proprietary prompt dump
- Any OpenAI-compatible base URL. Presets: **OpenRouter**, **Ollama (local)**, **Custom**.
- Enterprise pays for **governance** (allowlists, audit, Entra, air-gap) — not for the right to use a model

## Requirements

- Node.js 20+
- Microsoft 365 desktop (Windows or Mac) for sideload
- An OpenAI-compatible endpoint that supports `POST /chat/completions` with tool calling

## Quick start

```bash
npm install
npm test
npm run companion          # optional; required for most Ollama setups
npm run start:excel
```

`start:word` and `start:powerpoint` sideload the same add-in into the other hosts.

On first run, Office will trust a local HTTPS certificate from `office-addin-dev-certs`. Then:

1. Open the **OpenPlugin** ribbon tab → **Open**
2. Gear → pick **OpenRouter**, **Ollama (local)**, or **Custom OpenAI-compatible**
3. **Test connection**
4. Use a prompt chip or type. Review the change list. **Apply**

Excel custom functions (after sideload): `=OP.PROMPT("hello")`, `=OP.MAP(A2:A10, "uppercase")`, `=OP.EXTRACT(A1:D20, "name, amount")`.

### Manual sideload

If debugging tools do not launch Office, start the dev server and upload the manifest:

```bash
npm run dev
```

In Excel/Word/PowerPoint: **Insert → Add-ins → Upload My Add-in** → choose `manifest.xml`.

## Endpoint notes

The add-in calls your endpoint **from the Office WebView**. The server must send CORS headers for `https://localhost:3000` (dev) or your hosted origin.

| Setup | Usually works direct? |
|---|---|
| OpenRouter | Often yes |
| Custom (Azure, vLLM, …) | Paste the **full** `/v1` or `/chat/completions?api-version=` URL. Server must send CORS, or use the companion |
| Ollama on localhost | Usually **no** without CORS. Run `npm run companion` and keep the Ollama preset |

OSS builds have **no telemetry**.

## Skills

Skills live under `skills/` and follow the Agent Skills spec (`SKILL.md` + YAML frontmatter). Bundled in v0.1:

- `selection-rewrite`
- `excel-range-cleanup`
- `word-memo-from-sheet`
- `ppt-outline-to-slides`

Author a new folder, add `name` / `description`, and set `metadata.openplugin/hosts` to the hosts it supports.

## Repo

```
packages/core          Agent, LLM client, skills, tools, context compiler, Excel function helpers
packages/add-in        Task pane (React + Fluent UI)
packages/companion     Loopback CORS/Ollama proxy + audit + policy
packages/host-excel    Office.js Excel adapter
packages/host-word     Office.js Word adapter
packages/host-powerpoint
packages/enterprise    Policy format, AppSource notes (commercial)
skills/                Bundled Agent Skills
```

## License and enterprise

Core, add-in, hosts, companion (when it lands), and bundled skills are **Apache-2.0**.

`packages/enterprise` is the commercial governance layer (endpoint/model/skill allowlists, audit export, Entra ID, air-gapped installer). It is not required to use the agent.

See `CLA.md` if you contribute.

## Status

Sideloadable agent with a review-first task pane, OpenRouter / Ollama / custom presets, Excel `OP.*` functions, a loopback companion, and tenant `policy.json` via the companion. Tests run without Office (`npm test`).
