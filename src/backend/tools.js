import { setTimeout as delay } from "node:timers/promises";

import {
  cleanSnapshots,
  listSnapshots,
  prepareSnapshotPaths,
  readSnapshot,
  resolveSnapshotReference,
  resolveSnapshotElement,
  resolveSnapshotLabelTarget,
  writeSnapshotMetadata,
} from "../snapshots.js";
import { pressHotkey } from "../hotkeys.js";
import {
  createMousePath,
  createTypingChunks,
  dragOptionsForProfile,
  isHumanPacedProfile,
  resolveInteractionProfile,
  typingDelayForChunk,
  waitForProfile,
} from "../interaction-profile.js";
import { invokeWindowsAutomation } from "./windows-automation.js";

export async function listWindows() {
  return await invokeWindowsAutomation("list-windows");
}

export async function listScreens() {
  return await invokeWindowsAutomation("list-screens");
}

export async function listStoredSnapshots({ limit } = {}) {
  return {
    snapshots: await listSnapshots({ limit }),
  };
}

export async function getStoredSnapshot({ snapshotId }) {
  if (!snapshotId) {
    throw new Error("Missing snapshotId");
  }

  return await readSnapshot(snapshotId);
}

export async function cleanStoredSnapshots(options = {}) {
  return await cleanSnapshots(options);
}

export async function focusWindow({ hwnd, title }) {
  return await invokeWindowsAutomation("focus-window", { hwnd, title });
}

export async function moveWindow({ hwnd, title, x, y }) {
  return await invokeWindowsAutomation("move-window", { hwnd, title, x, y });
}

export async function resizeWindow({ hwnd, title, width, height }) {
  return await invokeWindowsAutomation("resize-window", {
    hwnd,
    title,
    width,
    height,
  });
}

export async function setWindowBounds({ hwnd, title, x, y, width, height }) {
  return await invokeWindowsAutomation("set-window-bounds", {
    hwnd,
    title,
    x,
    y,
    width,
    height,
  });
}

export async function setWindowState({ hwnd, title, state }) {
  return await invokeWindowsAutomation("set-window-state", {
    hwnd,
    title,
    state,
  });
}

export async function waitForWindow({
  hwnd,
  title,
  processId,
  processName,
  matchMode = "contains",
  timeoutMs = 5000,
  pollMs = 200,
} = {}) {
  if (!hwnd && !title && processId === undefined && !processName) {
    throw new Error(
      "Window wait requires hwnd, title, processId, or processName",
    );
  }

  let attempts = 0;

  return await pollUntil({
    timeoutMs,
    pollMs,
    description: "Window",
    operation: async () => {
      attempts += 1;
      const result = await listWindows();
      const match = result.windows.find((window) =>
        matchesWindowSelector(window, {
          hwnd,
          title,
          processId,
          processName,
          matchMode,
        }),
      );

      if (!match) {
        return null;
      }

      return {
        window: match,
        attempts,
        elapsedMs: undefined,
      };
    },
  });
}

export async function listApps() {
  return await invokeWindowsAutomation("list-apps");
}

export async function listDialogs({
  hwnd,
  title,
  processId,
  processName,
  matchMode = "contains",
} = {}) {
  const result = await listWindows();
  const hasSelector = hwnd || title || processId !== undefined || processName;
  const dialogs = result.windows.filter((window) => {
    if (!isDialogWindow(window)) {
      return false;
    }

    if (!hasSelector) {
      return true;
    }

    return matchesWindowSelector(window, {
      hwnd,
      title,
      processId,
      processName,
      matchMode,
    });
  });

  return {
    dialogs,
  };
}

export async function switchApp({ processId, name, title, matchMode }) {
  return await invokeWindowsAutomation("switch-app", {
    processId,
    name,
    title,
    matchMode,
  });
}

export async function quitApp({ processId, name, title, matchMode }) {
  return await invokeWindowsAutomation("quit-app", {
    processId,
    name,
    title,
    matchMode,
  });
}

export async function launchApp({ command, args = [] }) {
  return await invokeWindowsAutomation("launch-app", { command, args });
}

export async function moveMouse({ x, y, profile }) {
  const resolvedProfile = resolveInteractionProfile(profile);

  if (!isHumanPacedProfile(resolvedProfile)) {
    return await moveMouseRaw({ x, y });
  }

  const current = await getCursorPositionRaw();
  const path = createMousePath(current, { x, y }, resolvedProfile);

  for (const point of path) {
    await moveMouseRaw(point);
    await delay(8 + Math.floor(Math.random() * 14));
  }

  return {
    x,
    y,
    profile: resolvedProfile,
  };
}

export async function click({
  x,
  y,
  button = "left",
  double = false,
  profile,
}) {
  const resolvedProfile = resolveInteractionProfile(profile);

  if (isHumanPacedProfile(resolvedProfile)) {
    await moveMouse({ x, y, profile: resolvedProfile });
    await waitForProfile(resolvedProfile, "beforeClick");
  }

  const result = await clickRaw({ x, y, button, double });

  if (isHumanPacedProfile(resolvedProfile)) {
    await waitForProfile(resolvedProfile, "afterClick");
  }

  return {
    ...result,
    profile: resolvedProfile,
  };
}

export async function drag({
  fromX,
  fromY,
  toX,
  toY,
  button = "left",
  steps = 16,
  durationMs = 300,
  profile,
}) {
  const resolvedProfile = resolveInteractionProfile(profile);
  const dragOptions = dragOptionsForProfile(resolvedProfile, {
    steps,
    durationMs,
  });

  return await dragRaw({
    fromX,
    fromY,
    toX,
    toY,
    button,
    ...dragOptions,
  });
}

export async function scroll({ x, y, direction = "down", ticks = 3, profile }) {
  const resolvedProfile = resolveInteractionProfile(profile);

  if (
    isHumanPacedProfile(resolvedProfile) &&
    Number.isFinite(x) &&
    Number.isFinite(y)
  ) {
    await moveMouse({ x, y, profile: resolvedProfile });
    await waitForProfile(resolvedProfile, "beforeScroll");
  }

  if (!isHumanPacedProfile(resolvedProfile)) {
    return await scrollRaw({ x, y, direction, ticks });
  }

  let result = null;

  for (let index = 0; index < ticks; index += 1) {
    result = await scrollRaw({ x, y, direction, ticks: 1 });
    await waitForProfile(resolvedProfile, "afterScroll");
  }

  return {
    ...result,
    ticks,
    profile: resolvedProfile,
  };
}

export async function pressKeys({ keys }) {
  return await invokeWindowsAutomation("press-keys", { keys });
}

export async function hotkey({ keys, repeat = 1, delayMs = 80, profile }) {
  const resolvedProfile = resolveInteractionProfile(profile);

  const result = await pressHotkey({
    keys,
    repeat,
    delayMs,
    press: async (translatedKeys) => {
      if (isHumanPacedProfile(resolvedProfile)) {
        await waitForProfile(resolvedProfile, "beforeHotkey");
      }

      return await pressKeys({ keys: translatedKeys });
    },
  });

  return {
    ...result,
    profile: resolvedProfile,
  };
}

export async function typeText({ text, clear = false, delayMs = 80, profile }) {
  const resolvedProfile = resolveInteractionProfile(profile);

  if (clear) {
    await hotkey({
      keys: ["ctrl", "a"],
      profile: resolvedProfile,
    });

    if (delayMs > 0) {
      await delay(delayMs);
    }
  }

  if (!isHumanPacedProfile(resolvedProfile)) {
    const result = await typeTextRaw({ text });

    return {
      ...result,
      clear,
      delayMs,
      profile: resolvedProfile,
    };
  }

  await waitForProfile(resolvedProfile, "beforeType");

  for (const chunk of createTypingChunks(text, resolvedProfile)) {
    await typeTextRaw({ text: chunk });
    const pause = typingDelayForChunk(chunk, resolvedProfile);
    if (pause > 0) {
      await delay(pause);
    }
  }

  return {
    typed: text,
    clear,
    delayMs,
    profile: resolvedProfile,
  };
}

export async function findUiElements({
  hwnd,
  title,
  name,
  automationId,
  className,
  controlType,
  maxResults,
  matchMode,
}) {
  return await invokeWindowsAutomation("ui-find", {
    hwnd,
    title,
    name,
    automationId,
    className,
    controlType,
    maxResults,
    matchMode,
  });
}

export async function clickUiElement({
  hwnd,
  title,
  name,
  automationId,
  className,
  controlType,
  maxResults,
  matchMode,
  profile,
}) {
  const resolvedProfile = resolveInteractionProfile(profile);

  if (isHumanPacedProfile(resolvedProfile)) {
    const result = await findUiElements({
      hwnd,
      title,
      name,
      automationId,
      className,
      controlType,
      maxResults: maxResults ?? 1,
      matchMode,
    });
    const element = result.elements?.[0];

    if (!element?.center) {
      throw new Error("Matched element does not expose clickable coordinates");
    }

    const clickResult = await click({
      x: element.center.x,
      y: element.center.y,
      button: "left",
      double: false,
      profile: resolvedProfile,
    });

    return {
      clickMethod: "human-paced-coordinate",
      element,
      ...clickResult,
    };
  }

  return await invokeWindowsAutomation("ui-click", {
    hwnd,
    title,
    name,
    automationId,
    className,
    controlType,
    maxResults,
    matchMode,
  });
}

export async function seeUi(options = {}) {
  const {
    snapshotId,
    fileName,
    annotatedFileName,
    mode = options.hwnd || options.title ? "window" : "screen",
  } = options;

  if (!["screen", "window"].includes(mode)) {
    throw new Error(`Unsupported see mode '${mode}'`);
  }

  if (mode === "window" && !options.hwnd && !options.title) {
    throw new Error("Window see requires hwnd or title");
  }

  const snapshotPaths = await prepareSnapshotPaths({
    snapshotId,
    fileName,
    annotatedFileName,
  });
  const request = {
    ...options,
    mode,
    path: snapshotPaths.filePath,
    annotatedPath: snapshotPaths.annotatedPath,
  };
  const rawResult = await invokeWindowsAutomation("see-ui", request);
  const result = await attachOcrResult(rawResult);

  await writeSnapshotMetadata({
    snapshotId: snapshotPaths.snapshotId,
    action: "see-ui",
    request,
    result,
    metadataPath: snapshotPaths.metadataPath,
  });

  return {
    snapshotId: snapshotPaths.snapshotId,
    metadataPath: snapshotPaths.metadataPath,
    ...result,
  };
}

export async function clickSnapshotElement({
  snapshotId,
  elementId,
  name,
  matchMode = "exact",
  profile,
}) {
  if (!snapshotId) {
    throw new Error("Missing snapshotId");
  }

  const snapshot = await readSnapshot(snapshotId);
  const element = resolveSnapshotElement(snapshot, {
    elementId,
    name,
    matchMode,
  });
  const rootHwnd = snapshot?.result?.target?.hwnd;
  const resolvedProfile = resolveInteractionProfile(profile);

  if (isHumanPacedProfile(resolvedProfile)) {
    const center = element.center;

    if (!center || !Number.isFinite(center.x) || !Number.isFinite(center.y)) {
      throw new Error("Snapshot element does not expose clickable coordinates");
    }

    const clickResult = await click({
      x: center.x,
      y: center.y,
      button: "left",
      double: false,
      profile: resolvedProfile,
    });

    return {
      snapshotId,
      element,
      resolution: "human-paced-coordinate",
      ...clickResult,
    };
  }

  try {
    const clickResult = await clickUiElement({
      hwnd: rootHwnd,
      name: element.name,
      automationId: element.automationId,
      className: element.className,
      controlType: element.controlType,
      maxResults: 1,
      matchMode: "exact",
    });

    return {
      snapshotId,
      element,
      resolution: "ui-automation",
      ...clickResult,
    };
  } catch (error) {
    const center = element.center;

    if (!center || !Number.isFinite(center.x) || !Number.isFinite(center.y)) {
      throw error;
    }

    const clickResult = await click({
      x: center.x,
      y: center.y,
      button: "left",
      double: false,
    });

    return {
      snapshotId,
      element,
      resolution: "coordinate-fallback",
      ...clickResult,
    };
  }
}

export async function scrollSnapshotElement({
  snapshotId,
  elementId,
  name,
  matchMode = "exact",
  direction = "down",
  ticks = 3,
  profile,
}) {
  if (!snapshotId) {
    throw new Error("Missing snapshotId");
  }

  const snapshot = await readSnapshot(snapshotId);
  const element = resolveSnapshotElement(snapshot, {
    elementId,
    name,
    matchMode,
  });
  const center = element.center;

  if (!center || !Number.isFinite(center.x) || !Number.isFinite(center.y)) {
    throw new Error("Snapshot element does not have scrollable coordinates");
  }

  const scrollResult = await scroll({
    x: center.x,
    y: center.y,
    direction,
    ticks,
    profile,
  });

  return {
    snapshotId,
    element,
    ...scrollResult,
  };
}

export async function clickElementByLabel({
  label,
  snapshotId,
  mode,
  hwnd,
  title,
  screenIndex,
  fileName,
  annotatedFileName,
  matchMode = "contains",
  profile,
}) {
  if (!label) {
    throw new Error("Missing label");
  }

  const resolvedSnapshot = await resolveSnapshotForLabelAction({
    snapshotId,
    mode,
    hwnd,
    title,
    screenIndex,
    fileName,
    annotatedFileName,
    matchMode,
  });

  const snapshot = await readSnapshot(resolvedSnapshot.snapshotId);
  const target = resolveSnapshotLabelTarget(snapshot, {
    name: label,
    matchMode,
  });
  const result = await clickResolvedSnapshotTarget({
    snapshotId: resolvedSnapshot.snapshotId,
    snapshot,
    target,
    profile,
  });

  return {
    ...result,
    label,
    snapshotSource: resolvedSnapshot.source,
  };
}

export async function scrollElementByLabel({
  label,
  snapshotId,
  mode,
  hwnd,
  title,
  screenIndex,
  fileName,
  annotatedFileName,
  matchMode = "contains",
  direction = "down",
  ticks = 3,
  profile,
}) {
  if (!label) {
    throw new Error("Missing label");
  }

  const resolvedSnapshot = await resolveSnapshotForLabelAction({
    snapshotId,
    mode,
    hwnd,
    title,
    screenIndex,
    fileName,
    annotatedFileName,
    matchMode,
  });

  const snapshot = await readSnapshot(resolvedSnapshot.snapshotId);
  const target = resolveSnapshotLabelTarget(snapshot, {
    name: label,
    matchMode,
  });
  const result = await scrollResolvedSnapshotTarget({
    snapshotId: resolvedSnapshot.snapshotId,
    target,
    direction,
    ticks,
    profile,
  });

  return {
    ...result,
    label,
    snapshotSource: resolvedSnapshot.source,
  };
}

export async function listMenuItems({
  hwnd,
  title,
  path,
  maxResults = 40,
  matchMode = "contains",
  profile,
} = {}) {
  if (!hwnd && !title) {
    throw new Error("Menu list requires hwnd or title");
  }

  await resetMenuState({ hwnd, title });

  const pathSegments = parseMenuPath(path);
  if (pathSegments.length > 0) {
    await openMenuPath({
      hwnd,
      title,
      pathSegments,
      matchMode,
      profile,
    });

    await delay(150);
  }

  const result = await findUiElements({
    hwnd,
    title,
    controlType: "menuitem",
    maxResults,
    matchMode: "contains",
  });
  const items = dedupeAndSortMenuItems(result.elements ?? []);

  return {
    hwnd,
    title,
    openedPath: pathSegments,
    items,
    count: items.length,
  };
}

export async function clickMenuPath({
  hwnd,
  title,
  path,
  matchMode = "contains",
  profile,
} = {}) {
  if (!hwnd && !title) {
    throw new Error("Menu click requires hwnd or title");
  }

  const pathSegments = parseMenuPath(path);
  if (pathSegments.length === 0) {
    throw new Error("Missing menu path");
  }

  await resetMenuState({ hwnd, title });

  const steps = await openMenuPath({
    hwnd,
    title,
    pathSegments,
    matchMode,
    profile,
  });

  return {
    hwnd,
    title,
    path: pathSegments,
    steps,
    item: steps.at(-1)?.element ?? null,
  };
}

export async function clickDialogButton({
  hwnd,
  title,
  processId,
  processName,
  button,
  matchMode = "contains",
  profile,
} = {}) {
  if (!button) {
    throw new Error("Missing dialog button");
  }

  const dialog = await resolveDialogWindow({
    hwnd,
    title,
    processId,
    processName,
    matchMode,
  });

  await focusWindow({ hwnd: dialog.hwnd });

  const result = await clickNamedDialogButton({
    hwnd: dialog.hwnd,
    button,
    matchMode,
    profile,
  });

  return {
    dialog,
    button,
    ...result,
  };
}

export async function dragSnapshotElement({
  snapshotId,
  fromElementId,
  toElementId,
  fromName,
  toName,
  matchMode = "exact",
  button = "left",
  steps = 16,
  durationMs = 300,
  profile,
}) {
  if (!snapshotId) {
    throw new Error("Missing snapshotId");
  }

  const snapshot = await readSnapshot(snapshotId);
  const fromElement = resolveSnapshotElement(snapshot, {
    elementId: fromElementId,
    name: fromName,
    matchMode,
  });
  const toElement = resolveSnapshotElement(snapshot, {
    elementId: toElementId,
    name: toName,
    matchMode,
  });

  if (!fromElement.center || !toElement.center) {
    throw new Error("Snapshot elements do not have draggable coordinates");
  }

  const dragResult = await drag({
    fromX: fromElement.center.x,
    fromY: fromElement.center.y,
    toX: toElement.center.x,
    toY: toElement.center.y,
    button,
    steps,
    durationMs,
    profile,
  });

  return {
    snapshotId,
    fromElement,
    toElement,
    ...dragResult,
  };
}

export async function waitForUiElement({
  hwnd,
  title,
  name,
  automationId,
  className,
  controlType,
  maxResults,
  matchMode = "contains",
  timeoutMs = 5000,
  pollMs = 200,
} = {}) {
  if (!name && !automationId && !className && !controlType) {
    throw new Error("UI wait requires at least one selector");
  }

  let attempts = 0;

  return await pollUntil({
    timeoutMs,
    pollMs,
    description: "UI element",
    operation: async () => {
      attempts += 1;
      const result = await findUiElements({
        hwnd,
        title,
        name,
        automationId,
        className,
        controlType,
        maxResults,
        matchMode,
      });

      if (!Array.isArray(result.elements) || result.elements.length === 0) {
        return null;
      }

      return {
        element: result.elements[0],
        elements: result.elements,
        attempts,
        elapsedMs: undefined,
      };
    },
  });
}

export async function captureScreen(options = {}) {
  return await capture("capture-screen", options);
}

export async function captureWindow(options = {}) {
  return await capture("capture-window", options);
}

export async function ocrImage({ path, bounds } = {}) {
  if (!path) {
    throw new Error("Missing image path");
  }

  return await invokeWindowsAutomation("ocr-image", { path, bounds });
}

async function getCursorPositionRaw() {
  return await invokeWindowsAutomation("get-cursor-position");
}

async function moveMouseRaw({ x, y }) {
  return await invokeWindowsAutomation("move-mouse", { x, y });
}

async function clickRaw({ x, y, button = "left", double = false }) {
  return await invokeWindowsAutomation("click", { x, y, button, double });
}

async function dragRaw({
  fromX,
  fromY,
  toX,
  toY,
  button = "left",
  steps = 16,
  durationMs = 300,
}) {
  return await invokeWindowsAutomation("drag", {
    fromX,
    fromY,
    toX,
    toY,
    button,
    steps,
    durationMs,
  });
}

async function scrollRaw({ x, y, direction = "down", ticks = 3 }) {
  return await invokeWindowsAutomation("scroll", { x, y, direction, ticks });
}

async function typeTextRaw({ text }) {
  return await invokeWindowsAutomation("type-text", { text });
}

async function capture(action, options) {
  const snapshotPaths = await prepareSnapshotPaths({
    snapshotId: options.snapshotId,
    fileName: options.fileName,
    annotatedFileName: options.annotatedFileName,
  });
  const request = {
    ...options,
    path: snapshotPaths.filePath,
  };

  const rawResult = await invokeWindowsAutomation(action, request);
  const result = await attachOcrResult(rawResult);
  await writeSnapshotMetadata({
    snapshotId: snapshotPaths.snapshotId,
    action,
    request,
    result,
    metadataPath: snapshotPaths.metadataPath,
  });

  return {
    snapshotId: snapshotPaths.snapshotId,
    metadataPath: snapshotPaths.metadataPath,
    ...result,
  };
}

async function pollUntil({
  operation,
  timeoutMs = 5000,
  pollMs = 200,
  description = "Condition",
}) {
  const startedAt = Date.now();

  while (true) {
    const result = await operation();

    if (result) {
      return {
        ...result,
        elapsedMs: Date.now() - startedAt,
      };
    }

    if (Date.now() - startedAt >= timeoutMs) {
      throw new Error(`${description} not found within ${timeoutMs}ms`);
    }

    await delay(pollMs);
  }
}

async function resolveSnapshotForLabelAction({
  snapshotId,
  mode,
  hwnd,
  title,
  screenIndex,
  fileName,
  annotatedFileName,
  matchMode,
}) {
  const hasCaptureSelector =
    mode !== undefined ||
    hwnd !== undefined ||
    title !== undefined ||
    screenIndex !== undefined ||
    fileName !== undefined ||
    annotatedFileName !== undefined;

  if (snapshotId && hasCaptureSelector) {
    throw new Error(
      "Provide either snapshotId or capture target flags, not both",
    );
  }

  if (hasCaptureSelector) {
    const snapshot = await seeUi({
      mode: mode ?? (hwnd || title ? "window" : "screen"),
      hwnd,
      title,
      screenIndex,
      fileName,
      annotatedFileName,
      matchMode,
    });

    return {
      snapshotId: snapshot.snapshotId,
      source: "fresh-capture",
    };
  }

  const resolvedSnapshotId =
    snapshotId && snapshotId !== "latest"
      ? await resolveSnapshotReference(snapshotId)
      : await resolveLatestLabelSnapshotReference();

  if (!resolvedSnapshotId) {
    throw new Error(
      "No snapshots are available. Provide snapshotId or capture target flags.",
    );
  }

  return {
    snapshotId: resolvedSnapshotId,
    source:
      snapshotId && snapshotId !== "latest"
        ? "stored-snapshot"
        : "latest-snapshot",
  };
}

async function resolveLatestLabelSnapshotReference() {
  const snapshots = await listSnapshots();
  const preferred = snapshots.find(
    (snapshot) =>
      snapshot.valid &&
      snapshot.action === "see-ui" &&
      (Number(snapshot.elementCount) > 0 || Number(snapshot.ocrLineCount) > 0),
  );

  if (preferred) {
    return preferred.snapshotId;
  }

  const fallback = snapshots.find(
    (snapshot) =>
      snapshot.valid &&
      (Number(snapshot.elementCount) > 0 || Number(snapshot.ocrLineCount) > 0),
  );

  return fallback?.snapshotId ?? null;
}

async function clickResolvedSnapshotTarget({
  snapshotId,
  snapshot,
  target,
  profile,
}) {
  if (target.kind === "element") {
    return await clickSnapshotElement({
      snapshotId,
      elementId: target.id,
      profile,
    });
  }

  const center = target.center;

  if (!center || !Number.isFinite(center.x) || !Number.isFinite(center.y)) {
    throw new Error(
      "Snapshot OCR target does not expose clickable coordinates",
    );
  }

  const rootHwnd = snapshot?.result?.target?.hwnd;
  if (rootHwnd) {
    await focusWindow({ hwnd: rootHwnd });
  }

  const clickResult = await click({
    x: center.x,
    y: center.y,
    button: "left",
    double: false,
    profile,
  });

  return {
    snapshotId,
    element: target,
    resolution: target.kind,
    ...clickResult,
  };
}

async function scrollResolvedSnapshotTarget({
  snapshotId,
  target,
  direction,
  ticks,
  profile,
}) {
  if (target.kind === "element") {
    return await scrollSnapshotElement({
      snapshotId,
      elementId: target.id,
      direction,
      ticks,
      profile,
    });
  }

  const center = target.center;

  if (!center || !Number.isFinite(center.x) || !Number.isFinite(center.y)) {
    throw new Error(
      "Snapshot OCR target does not expose scrollable coordinates",
    );
  }

  const scrollResult = await scroll({
    x: center.x,
    y: center.y,
    direction,
    ticks,
    profile,
  });

  return {
    snapshotId,
    element: target,
    resolution: target.kind,
    ...scrollResult,
  };
}

async function attachOcrResult(result) {
  if (!result?.path) {
    return result;
  }

  try {
    const ocr = await ocrImage({
      path: result.path,
      bounds: result.bounds,
    });

    return {
      ...result,
      ocr,
    };
  } catch (error) {
    return {
      ...result,
      ocr: {
        available: false,
        text: "",
        lines: [],
        lineCount: 0,
        wordCount: 0,
        error: error instanceof Error ? error.message : String(error),
      },
    };
  }
}

async function openMenuPath({
  hwnd,
  title,
  pathSegments,
  matchMode = "contains",
  profile,
}) {
  const steps = [];

  for (let index = 0; index < pathSegments.length; index += 1) {
    const segment = pathSegments[index];
    const result = await clickUiElement({
      hwnd,
      title,
      name: segment,
      controlType: "menuitem",
      maxResults: 1,
      matchMode,
      profile,
    });

    steps.push({
      label: segment,
      clickMethod: result.clickMethod ?? result.resolution ?? null,
      element: result.element ?? null,
    });

    if (index < pathSegments.length - 1) {
      await delay(180);
    }
  }

  return steps;
}

async function clickNamedDialogButton({
  hwnd,
  button,
  matchMode = "contains",
  profile,
}) {
  const attempts = [
    {
      name: button,
      controlType: "button",
      maxResults: 1,
      matchMode,
      profile,
    },
    {
      name: button,
      className: "Button",
      maxResults: 1,
      matchMode,
      profile,
    },
    {
      name: button,
      maxResults: 1,
      matchMode,
      profile,
    },
  ];

  let lastError = null;

  for (const attempt of attempts) {
    try {
      return await clickUiElement({
        hwnd,
        ...attempt,
      });
    } catch (error) {
      lastError = error;
    }
  }

  throw (
    lastError ?? new Error(`No matching dialog button '${button}' was found`)
  );
}

async function resolveDialogWindow({
  hwnd,
  title,
  processId,
  processName,
  matchMode = "contains",
}) {
  const result = await listDialogs({
    hwnd,
    title,
    processId,
    processName,
    matchMode,
  });

  if (result.dialogs.length === 0) {
    throw new Error("No matching dialog window found");
  }

  if (result.dialogs.length > 1) {
    throw new Error(
      `Dialog selector matched ${result.dialogs.length} windows; narrow the selector`,
    );
  }

  return result.dialogs[0];
}

async function resetMenuState({ hwnd, title }) {
  await focusWindow({ hwnd, title });
  await delay(80);

  for (let attempt = 0; attempt < 2; attempt += 1) {
    await pressKeys({ keys: "{ESC}" });
    await delay(80);
  }
}

function parseMenuPath(path) {
  if (!path) {
    return [];
  }

  return String(path)
    .split(">")
    .map((segment) => segment.trim())
    .filter(Boolean);
}

function dedupeAndSortMenuItems(elements) {
  const sorted = [...elements].sort((left, right) => {
    const topDelta = (left?.bounds?.top ?? 0) - (right?.bounds?.top ?? 0);
    if (topDelta !== 0) {
      return topDelta;
    }

    const leftDelta = (left?.bounds?.left ?? 0) - (right?.bounds?.left ?? 0);
    if (leftDelta !== 0) {
      return leftDelta;
    }

    return String(left?.name ?? "").localeCompare(String(right?.name ?? ""));
  });

  const seen = new Set();
  const unique = [];

  for (const element of sorted) {
    const key = [
      element?.name ?? "",
      element?.automationId ?? "",
      element?.className ?? "",
      element?.bounds?.left ?? "",
      element?.bounds?.top ?? "",
      element?.bounds?.width ?? "",
      element?.bounds?.height ?? "",
    ].join("|");

    if (seen.has(key)) {
      continue;
    }

    seen.add(key);
    unique.push(element);
  }

  return unique;
}

function isDialogWindow(window) {
  const className = String(window?.className ?? "").toLowerCase();
  return className === "#32770" || className.includes("dialog");
}

function matchesWindowSelector(
  window,
  { hwnd, title, processId, processName, matchMode = "contains" },
) {
  if (hwnd && window.hwnd !== hwnd) {
    return false;
  }

  if (
    processId !== undefined &&
    Number(window.processId) !== Number(processId)
  ) {
    return false;
  }

  if (title && !matchesText(window.title, title, matchMode)) {
    return false;
  }

  if (processName && !matchesText(window.processName, processName, matchMode)) {
    return false;
  }

  return Boolean(hwnd || title || processId !== undefined || processName);
}

function matchesText(candidate, wanted, matchMode) {
  const actual = String(candidate ?? "");
  const expected = String(wanted ?? "");

  if (matchMode === "exact") {
    return actual === expected;
  }

  return actual.toLowerCase().includes(expected.toLowerCase());
}
