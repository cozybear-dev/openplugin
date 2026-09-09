# AppSource listing notes (community add-in)

Not a submission. Checklist when you are ready:

- Host `dist/` on HTTPS (GitHub Pages or a static web app). Update every URL in `manifest.xml`.
- Privacy: keys stay on-device; no telemetry in OSS; companion is optional and loopback-only.
- Permissions: `ReadWriteDocument` only. No Graph until Outlook exists.
- AppDomains: production origin plus `http://127.0.0.1` for the companion.
- Support URL, icon 32/64/80, screenshots of the review card and Excel functions.
- Do not require the companion for AppSource validation — OpenRouter must work direct.
