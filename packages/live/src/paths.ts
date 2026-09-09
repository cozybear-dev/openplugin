import { existsSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

export function repoRoot(): string {
  let dir = dirname(fileURLToPath(import.meta.url));
  for (let i = 0; i < 8; i++) {
    if (existsSync(join(dir, "manifest.xml")) && existsSync(join(dir, "packages", "add-in"))) return dir;
    dir = dirname(dir);
  }
  throw new Error("Could not find the OpenPlugin repo root (manifest.xml).");
}

export function addInDir(): string {
  return join(repoRoot(), "packages", "add-in");
}

export function liveDir(): string {
  return join(repoRoot(), "packages", "live");
}

export function liveManifestPath(): string {
  return join(liveDir(), "manifest.live.xml");
}

export function outputDir(): string {
  return join(liveDir(), "output");
}
