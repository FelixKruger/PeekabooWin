import fs from "node:fs/promises";
import path from "node:path";
import { setTimeout as delay } from "node:timers/promises";

import {
  focusWindow,
  getStoredSnapshot,
  scroll,
  scrollElementByLabel,
  seeUi,
  switchApp,
} from "./backend/tools.js";
import { getHarvestsDir } from "./config.js";
import { ensureDir, writeJson } from "./fs-utils.js";
import { resolveSnapshotReference } from "./snapshots.js";

export async function harvestScrollText(options = {}, dependencies = {}) {
  const settings = normalizeHarvestOptions(options);
  const capture = dependencies.capture ?? captureStoredSnapshot;
  const readSnapshot = dependencies.readSnapshot ?? readStoredSnapshot;
  const scrollAt = dependencies.scrollAt ?? scroll;
  const scrollByLabel = dependencies.scrollByLabel ?? scrollElementByLabel;
  const focusTarget = dependencies.focusTarget ?? bringTargetForward;
  const writeArtifacts = dependencies.writeArtifacts ?? writeHarvestArtifacts;

  let currentSnapshot = await resolveInitialSnapshot({
    settings,
    capture,
    readSnapshot,
  });

  if (!settings.snapshotId && extractOcrLines(currentSnapshot).length === 0) {
    try {
      const fallbackContext = deriveCaptureContext(settings, currentSnapshot);
      fallbackContext.mode = "screen";
      const fallbackSnapshot = await capture(fallbackContext);

      if (extractOcrLines(fallbackSnapshot).length > 0) {
        currentSnapshot = fallbackSnapshot;
      }
    } catch {
      // keep original snapshot if screen fallback fails
    }
  }

  if (settings.activateTarget) {
    await focusTarget({
      settings,
      snapshot: currentSnapshot,
    });
  }

  const steps = [];
  let collectedLines = [];
  let stalledSteps = 0;
  let stopReason = "max-steps";

  for (let stepIndex = 0; stepIndex < settings.maxSteps; stepIndex += 1) {
    const visibleLines = extractOcrLines(currentSnapshot);
    const mergeResult = mergeHarvestLines(collectedLines, visibleLines, {
      direction: settings.direction,
      overlapWindow: settings.overlapWindow,
      fuzzyThreshold: settings.fuzzyThreshold,
    });

    collectedLines = mergeResult.lines;

    const step = {
      index: stepIndex + 1,
      snapshotId: currentSnapshot.snapshotId,
      visibleLineCount: visibleLines.length,
      appendedLineCount: mergeResult.addedCount,
      overlapLineCount: mergeResult.overlapCount,
      firstVisibleLine: visibleLines[0]?.text ?? null,
      lastVisibleLine: visibleLines[visibleLines.length - 1]?.text ?? null,
    };
    steps.push(step);

    stalledSteps = mergeResult.addedCount > 0 ? 0 : stalledSteps + 1;

    if (stalledSteps >= settings.stopAfterStalledSteps) {
      stopReason = "stalled";
      break;
    }

    if (stepIndex === settings.maxSteps - 1) {
      stopReason = "max-steps";
      break;
    }

    await scrollHarvestTarget({
      settings,
      snapshot: currentSnapshot,
      scrollAt,
      scrollByLabel,
    });

    if (settings.pauseAfterScrollMs > 0) {
      await delay(settings.pauseAfterScrollMs);
    }

    currentSnapshot = await capture(
      deriveCaptureContext(settings, currentSnapshot),
    );
  }

  const transcriptLines = collectedLines.map((line) => line.text);
  const transcriptText = transcriptLines.join("\n");
  const artifacts = await writeArtifacts({
    settings,
    stopReason,
    steps,
    transcriptLines,
    transcriptText,
  });
  const preview = buildTextPreview(transcriptLines, {
    maxPreviewLines: settings.maxPreviewLines,
    maxPreviewChars: settings.maxPreviewChars,
  });

  return {
    ok: true,
    stopReason,
    direction: settings.direction,
    snapshotCount: steps.length,
    totalLines: transcriptLines.length,
    scrollLabel: settings.scrollLabel ?? null,
    outputPath: artifacts.outputPath,
    jsonPath: artifacts.jsonPath,
    previewTruncated: preview.truncated,
    textPreview: preview.text,
    steps,
  };
}

export function extractOcrLines(snapshot) {
  const rawLines = snapshot?.result?.ocr?.lines;

  if (!Array.isArray(rawLines)) {
    return [];
  }

  return rawLines
    .map((line, index) => {
      const text = String(line?.text ?? "")
        .replace(/\r?\n/g, " ")
        .trim();

      if (!text) {
        return null;
      }

      return {
        index,
        text,
        normalizedText: normalizeHarvestText(text),
        bounds: line?.bounds ?? null,
        center: line?.center ?? null,
      };
    })
    .filter(Boolean);
}

export function mergeHarvestLines(
  existingLines,
  incomingLines,
  { direction = "down", overlapWindow = 24, fuzzyThreshold } = {},
) {
  const existing = Array.isArray(existingLines) ? existingLines : [];
  const incoming = Array.isArray(incomingLines) ? incomingLines : [];

  if (existing.length === 0) {
    return {
      lines: [...incoming],
      overlapCount: 0,
      addedCount: incoming.length,
    };
  }

  if (incoming.length === 0) {
    return {
      lines: [...existing],
      overlapCount: 0,
      addedCount: 0,
    };
  }

  const maxOverlap = Math.min(existing.length, incoming.length, overlapWindow);
  let overlapCount = 0;

  for (let size = maxOverlap; size > 0; size -= 1) {
    const matches =
      direction === "up"
        ? sequenceMatches(existing.slice(0, size), incoming.slice(-size), {
            fuzzyThreshold,
          })
        : sequenceMatches(existing.slice(-size), incoming.slice(0, size), {
            fuzzyThreshold,
          });

    if (matches) {
      overlapCount = size;
      break;
    }
  }

  if (direction === "up") {
    return {
      lines: [
        ...incoming.slice(0, Math.max(0, incoming.length - overlapCount)),
        ...existing,
      ],
      overlapCount,
      addedCount: Math.max(0, incoming.length - overlapCount),
    };
  }

  return {
    lines: [...existing, ...incoming.slice(overlapCount)],
    overlapCount,
    addedCount: Math.max(0, incoming.length - overlapCount),
  };
}

export function normalizeHarvestText(text) {
  return String(text ?? "")
    .toLowerCase()
    .replace(/\s+/g, " ")
    .trim();
}

export function levenshteinRatio(a, b) {
  const strA = String(a ?? "");
  const strB = String(b ?? "");

  if (strA === strB) {
    return 1.0;
  }

  const lenA = strA.length;
  const lenB = strB.length;

  if (lenA === 0 || lenB === 0) {
    return 0.0;
  }

  let prevRow = new Array(lenB + 1);
  let currRow = new Array(lenB + 1);

  for (let j = 0; j <= lenB; j += 1) {
    prevRow[j] = j;
  }

  for (let i = 1; i <= lenA; i += 1) {
    currRow[0] = i;

    for (let j = 1; j <= lenB; j += 1) {
      const cost = strA[i - 1] === strB[j - 1] ? 0 : 1;
      currRow[j] = Math.min(
        prevRow[j] + 1,
        currRow[j - 1] + 1,
        prevRow[j - 1] + cost,
      );
    }

    [prevRow, currRow] = [currRow, prevRow];
  }

  const distance = prevRow[lenB];
  return 1.0 - distance / Math.max(lenA, lenB);
}

async function resolveInitialSnapshot({ settings, capture, readSnapshot }) {
  if (settings.snapshotId) {
    const resolvedSnapshotId = await resolveSnapshotReference(
      settings.snapshotId,
    );

    if (!resolvedSnapshotId) {
      throw new Error("No matching snapshot is available");
    }

    const snapshot = await readSnapshot(resolvedSnapshotId);
    return applySnapshotId(snapshot, resolvedSnapshotId);
  }

  return await capture(selectExplicitCaptureContext(settings));
}

async function captureStoredSnapshot(args) {
  const captureResult = await seeUi(args);
  return await readStoredSnapshot(captureResult.snapshotId);
}

async function readStoredSnapshot(snapshotId) {
  return await getStoredSnapshot({ snapshotId });
}

async function bringTargetForward({ settings, snapshot }) {
  const hwnd =
    settings.hwnd ?? snapshot?.result?.target?.hwnd ?? snapshot?.request?.hwnd;
  const title =
    settings.title ??
    snapshot?.result?.target?.title ??
    snapshot?.request?.title;

  if (hwnd || title) {
    await focusWindow({ hwnd, title });
    return;
  }

  const processId = settings.processId ?? snapshot?.request?.processId;
  const name = settings.processName ?? snapshot?.request?.processName;

  if (processId !== undefined || name) {
    await switchApp({
      processId,
      name,
      matchMode: settings.matchMode,
    });
  }
}

async function scrollHarvestTarget({
  settings,
  snapshot,
  scrollAt,
  scrollByLabel,
}) {
  if (settings.scrollLabel) {
    await scrollByLabel({
      label: settings.scrollLabel,
      snapshotId: snapshot.snapshotId,
      direction: settings.direction,
      ticks: settings.ticks,
      matchMode: settings.matchMode,
      profile: settings.profile,
    });
    return;
  }

  const point = resolveScrollPoint(settings, snapshot);

  await scrollAt({
    x: point.x,
    y: point.y,
    direction: settings.direction,
    ticks: settings.ticks,
    profile: settings.profile,
  });
}

function resolveScrollPoint(settings, snapshot) {
  if (Number.isFinite(settings.x) && Number.isFinite(settings.y)) {
    return {
      x: settings.x,
      y: settings.y,
    };
  }

  const bounds = snapshot?.result?.bounds;

  if (
    !bounds ||
    !Number.isFinite(bounds.left) ||
    !Number.isFinite(bounds.top)
  ) {
    throw new Error("Harvest scroll needs bounds or explicit x/y coordinates");
  }

  return {
    x: Math.round(bounds.left + bounds.width / 2),
    y: Math.round(bounds.top + bounds.height / 2),
  };
}

function deriveCaptureContext(settings, snapshot) {
  const request = snapshot?.request ?? {};
  const target = snapshot?.result?.target ?? {};
  const mode =
    settings.mode ?? request.mode ?? (target.hwnd ? "window" : "screen");

  return compactObject({
    mode,
    hwnd: settings.hwnd ?? target.hwnd ?? request.hwnd,
    title: settings.title ?? target.title ?? request.title,
    screenIndex: settings.screenIndex ?? request.screenIndex,
  });
}

function selectExplicitCaptureContext(settings) {
  return compactObject({
    mode: settings.mode,
    hwnd: settings.hwnd,
    title: settings.title,
    screenIndex: settings.screenIndex,
  });
}

function normalizeHarvestOptions(options) {
  const settings = {
    snapshotId: options.snapshotId,
    mode: options.mode,
    hwnd: options.hwnd,
    title: options.title,
    processId: options.processId,
    processName: options.processName,
    screenIndex: options.screenIndex,
    scrollLabel: options.scrollLabel ?? options.label,
    x: toOptionalNumber(options.x),
    y: toOptionalNumber(options.y),
    direction: options.direction === "up" ? "up" : "down",
    ticks: clampPositiveNumber(options.ticks, 3),
    maxSteps: clampPositiveNumber(options.maxSteps, 20),
    stopAfterStalledSteps: clampPositiveNumber(
      options.stopAfterStalledSteps,
      2,
    ),
    overlapWindow: clampPositiveNumber(options.overlapWindow, 24),
    fuzzyThreshold: toOptionalNumber(options.fuzzyThreshold),
    pauseAfterScrollMs: clampPositiveNumber(options.pauseAfterScrollMs, 350),
    maxPreviewLines: clampPositiveNumber(options.maxPreviewLines, 80),
    maxPreviewChars: clampPositiveNumber(options.maxPreviewChars, 12000),
    outputPath: options.outputPath,
    overwrite: options.overwrite === true,
    matchMode: options.matchMode === "exact" ? "exact" : "contains",
    profile: options.profile,
    activateTarget: options.activateTarget !== false,
  };

  const hasTarget =
    settings.snapshotId ||
    settings.hwnd ||
    settings.title ||
    settings.processId !== undefined ||
    settings.processName ||
    settings.mode ||
    settings.screenIndex !== undefined;

  if (!hasTarget) {
    throw new Error(
      "Harvest text requires a snapshot, window selector, or capture mode",
    );
  }

  return settings;
}

async function writeHarvestArtifacts({
  settings,
  stopReason,
  steps,
  transcriptLines,
  transcriptText,
}) {
  const { outputPath, jsonPath } = resolveHarvestOutputPaths(
    settings.outputPath,
  );
  await ensureWritableHarvestTargets({
    outputPath,
    jsonPath,
    overwrite: settings.overwrite,
  });
  await ensureDir(path.dirname(outputPath));
  await ensureDir(path.dirname(jsonPath));
  await fs.writeFile(
    outputPath,
    transcriptText ? `${transcriptText}\n` : "",
    "utf8",
  );
  await writeJson(jsonPath, {
    createdAt: new Date().toISOString(),
    stopReason,
    direction: settings.direction,
    outputPath,
    totalLines: transcriptLines.length,
    steps,
    lines: transcriptLines,
    text: transcriptText,
  });

  return {
    outputPath,
    jsonPath,
  };
}

async function ensureWritableHarvestTargets({
  outputPath,
  jsonPath,
  overwrite,
}) {
  if (overwrite) {
    return;
  }

  const existingTargets = [];

  if (await pathExists(outputPath)) {
    existingTargets.push(outputPath);
  }

  if (await pathExists(jsonPath)) {
    existingTargets.push(jsonPath);
  }

  if (existingTargets.length > 0) {
    throw new Error(
      `Harvest output already exists. Choose a new --output path or enable overwrite. Existing: ${existingTargets.join(", ")}`,
    );
  }
}

function resolveHarvestOutputPaths(outputPath) {
  if (outputPath) {
    const resolved = path.resolve(String(outputPath));
    const extension = path.extname(resolved).toLowerCase();

    if (extension === ".json") {
      return {
        outputPath: `${resolved.slice(0, -5)}.txt`,
        jsonPath: resolved,
      };
    }

    if (extension === ".txt") {
      return {
        outputPath: resolved,
        jsonPath: `${resolved.slice(0, -4)}.json`,
      };
    }

    return {
      outputPath: `${resolved}.txt`,
      jsonPath: `${resolved}.json`,
    };
  }

  const stamp = new Date().toISOString().replace(/[:.]/g, "-");
  const baseDir = path.join(getHarvestsDir(), stamp);

  return {
    outputPath: path.join(baseDir, "harvest.txt"),
    jsonPath: path.join(baseDir, "harvest.json"),
  };
}

function buildTextPreview(lines, { maxPreviewLines, maxPreviewChars }) {
  const selected = [];
  let totalChars = 0;
  let truncated = false;

  for (const line of lines) {
    const nextChars = totalChars + line.length + 1;

    if (selected.length >= maxPreviewLines || nextChars > maxPreviewChars) {
      truncated = true;
      break;
    }

    selected.push(line);
    totalChars = nextChars;
  }

  return {
    truncated,
    text: selected.join("\n"),
  };
}

function sequenceMatches(left, right, { fuzzyThreshold } = {}) {
  if (left.length !== right.length) {
    return false;
  }

  for (let index = 0; index < left.length; index += 1) {
    const leftText = left[index]?.normalizedText ?? "";
    const rightText = right[index]?.normalizedText ?? "";

    if (leftText === rightText) {
      continue;
    }

    if (
      Number.isFinite(fuzzyThreshold) &&
      levenshteinRatio(leftText, rightText) >= fuzzyThreshold
    ) {
      continue;
    }

    return false;
  }

  return true;
}

function applySnapshotId(snapshot, snapshotId) {
  if (snapshot?.snapshotId) {
    return snapshot;
  }

  return {
    ...snapshot,
    snapshotId,
  };
}

function toOptionalNumber(value) {
  if (value === undefined || value === null || value === "") {
    return undefined;
  }

  const numericValue = Number(value);
  return Number.isFinite(numericValue) ? numericValue : undefined;
}

async function pathExists(targetPath) {
  try {
    await fs.access(targetPath);
    return true;
  } catch {
    return false;
  }
}

function clampPositiveNumber(value, fallback) {
  const numericValue = Number(value);
  return Number.isFinite(numericValue) && numericValue > 0
    ? Math.round(numericValue)
    : fallback;
}

function compactObject(value) {
  return Object.fromEntries(
    Object.entries(value).filter(
      ([, entry]) => entry !== undefined && entry !== null && entry !== "",
    ),
  );
}
