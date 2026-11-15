// File-based registry: one JSON file plus content-addressed blobs.
// Good enough for a single-user prototype; the interface is what a
// DMS / SharePoint / Postgres backend would implement later.

import { mkdirSync, readFileSync, writeFileSync, renameSync, existsSync } from "node:fs";
import path from "node:path";
import crypto from "node:crypto";

export function sha256(buf) {
  return crypto.createHash("sha256").update(buf).digest("hex");
}

export function createStore(dataDir) {
  mkdirSync(path.join(dataDir, "files"), { recursive: true });
  mkdirSync(path.join(dataDir, "mail"), { recursive: true });
  const file = path.join(dataDir, "registry.json");
  let db = existsSync(file)
    ? JSON.parse(readFileSync(file, "utf8"))
    : { matters: [], versions: [], messages: [], triage: [], findingsCache: {} };

  const save = () => {
    const tmp = file + ".tmp";
    writeFileSync(tmp, JSON.stringify(db, null, 2));
    renameSync(tmp, file); // atomic replace
  };

  return {
    db: () => db,
    save,

    // ---- matters
    upsertMatters(matters) {
      for (const m of matters) {
        const i = db.matters.findIndex((x) => x.id === m.id);
        if (i >= 0) db.matters[i] = { ...db.matters[i], ...m }; else db.matters.push(m);
      }
      save();
    },
    matter: (id) => db.matters.find((m) => m.id === id),

    // ---- blobs
    putBlob(kind, buf, ext) {
      const h = sha256(buf);
      const p = path.join(dataDir, kind, `${h}${ext}`);
      if (!existsSync(p)) writeFileSync(p, buf);
      return h;
    },
    blobPath: (kind, hash, ext) => path.join(dataDir, kind, `${hash}${ext}`),
    readBlob: (kind, hash, ext) => readFileSync(path.join(dataDir, kind, `${hash}${ext}`)),

    // ---- versions
    versions: (matterId) => db.versions.filter((v) => !matterId || v.matterId === matterId),
    version: (id) => db.versions.find((v) => v.id === id),
    versionByHash: (hash) => db.versions.find((v) => v.fileHash === hash),
    addVersion(v) { db.versions.push(v); save(); return v; },
    updateVersion(id, patch) {
      const v = db.versions.find((x) => x.id === id);
      Object.assign(v, patch);
      save();
      return v;
    },

    // ---- messages & triage
    message: (messageId) => db.messages.find((m) => m.messageId === messageId),
    messageIds: (ids) => db.messages.filter((m) => ids.includes(m.messageId)),
    addMessage(m) { db.messages.push(m); save(); return m; },
    addTriage(t) { db.triage.push(t); save(); return t; },
    triage: (id) => db.triage.find((t) => t.id === id),
    openTriage: () => db.triage.filter((t) => t.status === "open"),
    closeTriage(id, patch) { Object.assign(db.triage.find((t) => t.id === id), { status: "closed", ...patch }); save(); },

    // ---- findings cache (key = hash of before/after/playbook version)
    cached: (key) => db.findingsCache[key],
    cache(key, finding) { db.findingsCache[key] = finding; },
  };
}
