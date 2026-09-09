# OpenPlugin Enterprise

Commercial governance. The agent itself stays Apache-2.0.

Drop `policy.json` next to the companion data directory:

- Windows: `%LOCALAPPDATA%\OpenPlugin\policy.json`
- macOS/Linux: `~/.openplugin/policy.json`

Or set `OPENPLUGIN_HOME`. See `policy.example.json`.

The OSS add-in **loads** this file through the companion (`GET /policy`) and enforces it with `assertPolicy`. Users cannot widen allowlists.

## What you get

| Control | How |
|---|---|
| Endpoint / model / skill allowlists | `policy.json` |
| Disable generated JS | `denyExecuteJs: true` (default) |
| Require companion | `requireCompanion: true` |
| Audit log | `%LOCALAPPDATA%\OpenPlugin\audit.jsonl` |
| Internal skill catalog | `catalogUrl` pointing at `{ "skills": [{ "name", "url" }] }` |

## Optional Entra ID

The add-in does not require sign-in. To stamp audit rows with the signed-in Microsoft account, register an app in Entra ID, add `webApplicationInfo` to the manifest, and grant admin consent. If `Office.auth.getAccessToken` fails, the agent still runs and audit `user` is `"local"`.

## AppSource

See `appsource-notes.md`. This package does not submit the listing for you.
