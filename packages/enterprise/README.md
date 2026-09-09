# OpenPlugin Enterprise

Commercial governance package. Not required for the open-source agent.

Planned surface:

- Tenant policy (`allowedEndpoints`, `allowedModels`, `allowedSkills`, `denyExecuteJs`)
- Entra ID sign-in
- Audit export (JSONL / SIEM)
- Signed internal skill catalog
- Air-gapped installer and LTS

The OSS add-in exposes a `PolicyProvider` seam; this package is the paid implementation. Source-available under a commercial license — not Apache-2.0.
