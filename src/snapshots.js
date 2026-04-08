import crypto from "node:crypto";
import fs from "node:fs/promises";
import path from "node:path";

import { getSnapshotsDir } from "./config.js";
import { ensureDir, readJson, writeJson } from "./fs-utils.js";

export async function prepareSnapshotPaths({
  snapshotId = crypto.randomUUID(),
  fileName = "capture.png",
  annotatedFileName = "annotated.png",
} = {}) {
  const snapshotsDir = getSnapshotsDir();
  await ensureDir(snapshotsDir);

  const targetDir = path.join(snapshotsDir, snapshotId);
  await ensureDir(targetDir);

  return {
    snapshotId,
    targetDir,
    filePath: path.join(targetDir, fileName),
    annotatedPath: path.join(targetDir, annotatedFileName),
    metadataPath: path.join(targetDir, "snapshot.json"),
  };
}

export async function writeSnapshotMetadata({
  snapshotId,
  action,
  request,
  result,
  metadataPath,
}) {
  await writeJson(metadataPath, {
    snapshotId,
    createdAt: new Date().toISOString(),
    action,
    request,
    result,
  });
}

export async function readSnapshot(snapshotId) {
  const metadataPath = path.join(
    getSnapshotsDir(),
    snapshotId,
    "snapshot.json",
  );
  return await readJson(metadataPath);
}

export async function listSnapshots({ limit } = {}) {
  const snapshotsDir = getSnapshotsDir();
  await ensureDir(snapshotsDir);

  const entries = await fs.readdir(snapshotsDir, { withFileTypes: true });
  const snapshots = [];

  for (const entry of entries) {
    if (!entry.isDirectory()) {
      continue;
    }

    const snapshotId = entry.name;

    try {
      const metadata = await readSnapshot(snapshotId);
      snapshots.push(makeSnapshotSummary(metadata));
    } catch {
      snapshots.push({
        snapshotId,
        createdAt: null,
        action: null,
        path: path.join(snapshotsDir, snapshotId),
        metadataPath: path.join(snapshotsDir, snapshotId, "snapshot.json"),
        valid: false,
      });
    }
  }

  snapshots.sort((left, right) => {
    const leftTime = left.createdAt ? Date.parse(left.createdAt) : 0;
    const rightTime = right.createdAt ? Date.parse(right.createdAt) : 0;
    return rightTime - leftTime;
  });

  return typeof limit === "number" && limit > 0
    ? snapshots.slice(0, limit)
    : snapshots;
}

export async function resolveSnapshotReference(
  snapshotId,
  { requireElements = false } = {},
) {
  if (snapshotId && snapshotId !== "latest") {
    return snapshotId;
  }

  const snapshots = await listSnapshots();
  const latest = snapshots.find(
    (snapshot) =>
      snapshot.valid && (!requireElements || Number(snapshot.elementCount) > 0),
  );

  return latest?.snapshotId ?? null;
}

export async function cleanSnapshots({
  snapshotId,
  all = false,
  olderThanHours,
} = {}) {
  const snapshotsDir = getSnapshotsDir();
  await ensureDir(snapshotsDir);

  const targets = [];

  if (snapshotId) {
    targets.push(snapshotId);
  } else {
    const snapshots = await listSnapshots();

    for (const snapshot of snapshots) {
      if (all) {
        targets.push(snapshot.snapshotId);
        continue;
      }

      if (
        typeof olderThanHours === "number" &&
        olderThanHours > 0 &&
        snapshot.createdAt
      ) {
        const ageMs = Date.now() - Date.parse(snapshot.createdAt);
        if (ageMs >= olderThanHours * 60 * 60 * 1000) {
          targets.push(snapshot.snapshotId);
        }
      }
    }
  }

  const deleted = [];

  for (const id of [...new Set(targets)]) {
    const targetDir = path.join(snapshotsDir, id);
    await fs.rm(targetDir, { recursive: true, force: true });
    deleted.push(id);
  }

  return {
    deleted,
    count: deleted.length,
  };
}

export function resolveSnapshotElement(
  snapshot,
  { elementId, name, matchMode = "exact" } = {},
) {
  const elements = snapshot?.result?.elements;

  if (!Array.isArray(elements) || elements.length === 0) {
    throw new Error("Snapshot does not contain any indexed UI elements");
  }

  if (elementId) {
    const directMatch = elements.find((element) => element.id === elementId);

    if (!directMatch) {
      throw new Error(`Snapshot element '${elementId}' was not found`);
    }

    return directMatch;
  }

  if (!name) {
    throw new Error("Missing snapshot element selector");
  }

  const normalizedName = String(name);
  const matches = elements.filter((element) => {
    const candidate = element.name ?? "";

    if (matchMode === "contains") {
      return candidate.toLowerCase().includes(normalizedName.toLowerCase());
    }

    return candidate === normalizedName;
  });

  if (matches.length === 0) {
    throw new Error(`No snapshot element matched '${normalizedName}'`);
  }

  if (matches.length > 1) {
    throw new Error(
      `Snapshot selector '${normalizedName}' matched ${matches.length} elements; use --element-id`,
    );
  }

  return matches[0];
}

export function resolveSnapshotLabelTarget(
  snapshot,
  { name, matchMode = "contains" } = {},
) {
  if (!name) {
    throw new Error("Missing snapshot label selector");
  }

  const normalizedName = String(name);
  const elements = Array.isArray(snapshot?.result?.elements)
    ? snapshot.result.elements
    : [];
  const elementMatches = findSnapshotMatches(
    elements,
    normalizedName,
    matchMode,
    (element) => element?.name ?? "",
  );

  if (elementMatches.length === 1) {
    return {
      kind: "element",
      source: "ui-element",
      ...elementMatches[0],
    };
  }

  if (elementMatches.length > 1) {
    throw new Error(
      `Snapshot selector '${normalizedName}' matched ${elementMatches.length} UI elements; use a stricter label`,
    );
  }

  const ocrLines = Array.isArray(snapshot?.result?.ocr?.lines)
    ? snapshot.result.ocr.lines
    : [];
  const lineMatches = findSnapshotMatches(
    ocrLines,
    normalizedName,
    matchMode,
    (line) => line?.text ?? "",
  );

  if (lineMatches.length === 1) {
    return {
      kind: "ocr-line",
      source: "ocr-line",
      ...lineMatches[0],
    };
  }

  if (lineMatches.length > 1) {
    throw new Error(
      `Snapshot selector '${normalizedName}' matched ${lineMatches.length} OCR lines; use a stricter label`,
    );
  }

  const ocrWords = ocrLines.flatMap((line) =>
    (Array.isArray(line?.words) ? line.words : []).map((word) => ({
      ...word,
      lineText: line?.text ?? "",
    })),
  );
  const wordMatches = findSnapshotMatches(
    ocrWords,
    normalizedName,
    matchMode,
    (word) => word?.text ?? "",
  );

  if (wordMatches.length === 1) {
    return {
      kind: "ocr-word",
      source: "ocr-word",
      ...wordMatches[0],
    };
  }

  if (wordMatches.length > 1) {
    throw new Error(
      `Snapshot selector '${normalizedName}' matched ${wordMatches.length} OCR words; use a stricter label`,
    );
  }

  throw new Error(`No snapshot label target matched '${normalizedName}'`);
}

function makeSnapshotSummary(snapshot) {
  return {
    snapshotId: snapshot.snapshotId,
    createdAt: snapshot.createdAt ?? null,
    action: snapshot.action ?? null,
    path: snapshot.result?.path ?? snapshot.request?.path ?? null,
    annotatedPath:
      snapshot.result?.annotatedPath ?? snapshot.request?.annotatedPath ?? null,
    metadataPath: path.join(
      getSnapshotsDir(),
      snapshot.snapshotId,
      "snapshot.json",
    ),
    elementCount: Array.isArray(snapshot.result?.elements)
      ? snapshot.result.elements.length
      : 0,
    ocrLineCount: Array.isArray(snapshot.result?.ocr?.lines)
      ? snapshot.result.ocr.lines.length
      : 0,
    target: snapshot.result?.target ?? null,
    valid: true,
  };
}

function findSnapshotMatches(items, wanted, matchMode, readCandidate) {
  const normalizedWanted = String(wanted);

  return items.filter((item) => {
    const candidate = String(readCandidate(item) ?? "");

    if (matchMode === "contains") {
      return candidate.toLowerCase().includes(normalizedWanted.toLowerCase());
    }

    return candidate === normalizedWanted;
  });
}
