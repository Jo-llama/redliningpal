// Shared wiring: config + store, used by the server and the CLI scripts.

import { existsSync, readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { createStore } from "./store.js";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

export function loadConfig(file = process.env.REDLINE_CONFIG) {
  const candidates = [file, path.join(root, "config/intake.json"), path.join(root, "config/intake.example.json")].filter(Boolean);
  const found = candidates.find((f) => existsSync(f));
  return { ...JSON.parse(readFileSync(found, "utf8")), _file: found };
}

export function createContext({ configFile, dataDir } = {}) {
  const config = loadConfig(configFile);
  const store = createStore(dataDir || process.env.REDLINE_DATA || path.join(root, "data"));
  store.upsertMatters(config.matters || []);
  return { config, store };
}
