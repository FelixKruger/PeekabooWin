import fs from "node:fs/promises";

import { buildSnapshotAiBrief } from "./ai-bridge.js";
import {
  captureScreen,
  captureWindow,
  clickElementByLabel,
  clickUiElement,
  clickSnapshotElement,
  click,
  cleanStoredSnapshots,
  drag,
  dragSnapshotElement,
  clickDialogButton,
  findUiElements,
  focusWindow,
  getStoredSnapshot,
  hotkey,
  launchApp,
  listApps,
  listDialogs,
  listMenuItems,
  listScreens,
  listStoredSnapshots,
  listWindows,
  clickMenuPath,
  moveWindow,
  moveMouse,
  pressKeys,
  quitApp,
  resizeWindow,
  scroll,
  scrollElementByLabel,
  scrollSnapshotElement,
  seeUi,
  setWindowBounds,
  setWindowState,
  switchApp,
  waitForUiElement,
  waitForWindow,
  typeText,
} from "./backend/tools.js";
import { CliUsageError } from "./errors.js";
import { harvestScrollText } from "./harvest.js";
import { listRecipes, planGoal, runGoal, runRecipe } from "./recipes.js";
import { runWorkflowFile } from "./workflow/runner.js";

const helpText = `peekaboo-win

Usage:
  peekaboo-win windows list
  peekaboo-win screens list
  peekaboo-win window focus --hwnd <id>
  peekaboo-win window move (--hwnd <id> | --title <text>) --x <n> --y <n>
  peekaboo-win window resize (--hwnd <id> | --title <text>) --width <n> --height <n>
  peekaboo-win window set-bounds (--hwnd <id> | --title <text>) --x <n> --y <n> --width <n> --height <n>
  peekaboo-win window state (--hwnd <id> | --title <text>) --state restore|maximize|minimize
  peekaboo-win window wait [--hwnd <id> | --title <text> | --process-id <id> | --process-name <text>] [--timeout-ms <n>] [--poll-ms <n>] [--exact]
  peekaboo-win app list
  peekaboo-win app launch --command <program> [--args arg1,arg2]
  peekaboo-win app switch (--process-id <id> | --name <text> | --title <text>) [--exact]
  peekaboo-win app quit (--process-id <id> | --name <text> | --title <text>) [--exact]
  peekaboo-win dialog list [--title <text> | --hwnd <id> | --process-id <id> | --process-name <text>] [--exact]
  peekaboo-win dialog click [--title <text> | --hwnd <id> | --process-id <id> | --process-name <text>] --button <text> [--exact]
  peekaboo-win menu list (--hwnd <id> | --title <text>) [--open "File>Recent"] [--max-results <n>] [--exact]
  peekaboo-win menu click (--hwnd <id> | --title <text>) --path "File>Save" [--exact]
  peekaboo-win screen capture [--screen-index <n>] [--file-name capture.png]
  peekaboo-win window capture (--hwnd <id> | --title <text>) [--file-name capture.png]
  peekaboo-win mouse move --x <n> --y <n>
  peekaboo-win mouse click --x <n> --y <n> [--button left|right] [--double]
  peekaboo-win mouse drag --from-x <n> --from-y <n> --to-x <n> --to-y <n> [--button left|right] [--steps <n>] [--duration-ms <n>]
  peekaboo-win click --on <text> [--snapshot <id|latest> | --mode screen|window [--screen-index <n>] [--hwnd <id> | --title <text>]] [--exact]
  peekaboo-win scroll [--x <n> --y <n> | --on <text> [--snapshot <id|latest> | --mode screen|window [--screen-index <n>] [--hwnd <id> | --title <text>]]] [--direction up|down] [--ticks <n>] [--exact]
  peekaboo-win see [--mode screen|window] [--screen-index <n>] [--hwnd <id> | --title <text>] [--file-name capture.png] [--annotated-file-name annotated.png]
  peekaboo-win snapshot list [--limit <n>]
  peekaboo-win snapshot show --snapshot <id>
  peekaboo-win snapshot click --snapshot <id> [--element-id e1 | --name "Save"] [--exact]
  peekaboo-win snapshot drag --snapshot <id> [--from-element-id e1 | --from-name "Item"] [--to-element-id e2 | --to-name "Folder"] [--button left|right] [--steps <n>] [--duration-ms <n>] [--exact]
  peekaboo-win snapshot scroll --snapshot <id> [--element-id e1 | --name "Items"] [--direction up|down] [--ticks <n>] [--exact]
  peekaboo-win snapshot clean [--snapshot <id> | --all | --older-than-hours <n>]
  peekaboo-win ai brief --snapshot <id|latest>
  peekaboo-win harvest text [--snapshot <id|latest> | --mode screen|window [--hwnd <id> | --title <text>] [--screen-index <n>]] [--label <text>] [--x <n> --y <n>] [--direction up|down] [--ticks <n>] [--max-steps <n>] [--stop-after-stalled-steps <n>] [--overlap-window <n>] [--fuzzy-threshold <ratio>] [--pause-after-scroll-ms <n>] [--output <path>] [--overwrite]
  peekaboo-win recipe list
  peekaboo-win recipe run --id <recipe-id> [--snapshot <id|latest> | --mode screen|window [--hwnd <id> | --title <text>] [--screen-index <n>]] [--label <text>] [--text <text>] [--command <program>] [--output <path>] [--overwrite]
  peekaboo-win goal plan --text "<goal>" [--snapshot <id|latest> | --mode screen|window [--hwnd <id> | --title <text>] [--screen-index <n>]]
  peekaboo-win goal run --text "<goal>" [--snapshot <id|latest> | --mode screen|window [--hwnd <id> | --title <text>] [--screen-index <n>]] [--output <path>] [--overwrite]
  peekaboo-win run --file <workflow.json> [--no-fail-fast]
  peekaboo-win ui find [--title <window>] --name <text> [--control-type button] [--exact]
  peekaboo-win ui click [--title <window>] --name <text> [--control-type button] [--exact]
  peekaboo-win ui wait [--title <window> | --hwnd <id>] --name <text> [--control-type button] [--timeout-ms <n>] [--poll-ms <n>] [--exact]
  peekaboo-win hotkey --keys ctrl,shift,t [--repeat <n>] [--delay-ms <n>]
  peekaboo-win type (--text <text> | --text-file <path>) [--clear] [--delay-ms <n>]
  peekaboo-win press --keys "^l"

Interaction Profiles:
  Add --profile human-paced to interaction commands, or set PEEKABOO_INTERACTION_PROFILE=human-paced.
  This mode adds visible cursor movement and variable pacing for demos/review, not stealth or bot-evasion.
`;

export async function runCli(argv) {
  try {
    const result = await dispatch(argv);
    process.stdout.write(`${JSON.stringify(result, null, 2)}\n`);
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    process.stderr.write(`${message}\n`);
    process.exitCode = 1;
  }
}

async function dispatch(argv) {
  if (argv.length === 0 || argv.includes("--help") || argv.includes("-h")) {
    return { ok: true, help: helpText };
  }

  const [group, action, ...rest] = argv;
  const flagTokens = action?.startsWith("--") ? [action, ...rest] : rest;
  const flags = parseFlags(flagTokens);

  if (group === "windows" && action === "list") {
    return await listWindows();
  }

  if (group === "screens" && action === "list") {
    return await listScreens();
  }

  if (group === "window" && action === "focus") {
    return await focusWindow({
      hwnd: flags.hwnd,
      title: flags.title,
    });
  }

  if (group === "window" && action === "wait") {
    return await waitForWindow({
      hwnd: flags.hwnd,
      title: flags.title,
      processId: parseOptionalNumberFlag(flags["process-id"]),
      processName: flags["process-name"],
      matchMode: flags.exact ? "exact" : "contains",
      timeoutMs: parseOptionalNumberFlag(flags["timeout-ms"]) ?? 5000,
      pollMs: parseOptionalNumberFlag(flags["poll-ms"]) ?? 200,
    });
  }

  if (group === "window" && action === "move") {
    return await moveWindow({
      hwnd: flags.hwnd,
      title: flags.title,
      x: parseNumberFlag(flags.x, "x"),
      y: parseNumberFlag(flags.y, "y"),
    });
  }

  if (group === "window" && action === "resize") {
    return await resizeWindow({
      hwnd: flags.hwnd,
      title: flags.title,
      width: parseNumberFlag(flags.width, "width"),
      height: parseNumberFlag(flags.height, "height"),
    });
  }

  if (group === "window" && action === "set-bounds") {
    return await setWindowBounds({
      hwnd: flags.hwnd,
      title: flags.title,
      x: parseNumberFlag(flags.x, "x"),
      y: parseNumberFlag(flags.y, "y"),
      width: parseNumberFlag(flags.width, "width"),
      height: parseNumberFlag(flags.height, "height"),
    });
  }

  if (group === "window" && action === "state") {
    if (!flags.state) {
      throw new CliUsageError("Missing required flag: --state");
    }

    return await setWindowState({
      hwnd: flags.hwnd,
      title: flags.title,
      state: flags.state,
    });
  }

  if (group === "app" && action === "list") {
    return await listApps();
  }

  if (group === "dialog" && action === "list") {
    return await listDialogs({
      hwnd: flags.hwnd,
      title: flags.title,
      processId: parseOptionalNumberFlag(flags["process-id"]),
      processName: flags["process-name"],
      matchMode: flags.exact ? "exact" : "contains",
    });
  }

  if (group === "dialog" && action === "click") {
    if (!flags.button) {
      throw new CliUsageError("Missing required flag: --button");
    }

    return await clickDialogButton({
      hwnd: flags.hwnd,
      title: flags.title,
      processId: parseOptionalNumberFlag(flags["process-id"]),
      processName: flags["process-name"],
      button: flags.button,
      matchMode: flags.exact ? "exact" : "contains",
      profile: flags.profile,
    });
  }

  if (group === "menu" && action === "list") {
    return await listMenuItems({
      hwnd: flags.hwnd,
      title: flags.title,
      path: flags.open,
      maxResults: parseOptionalNumberFlag(flags["max-results"]) ?? 40,
      matchMode: flags.exact ? "exact" : "contains",
      profile: flags.profile,
    });
  }

  if (group === "menu" && action === "click") {
    if (!flags.path) {
      throw new CliUsageError("Missing required flag: --path");
    }

    return await clickMenuPath({
      hwnd: flags.hwnd,
      title: flags.title,
      path: flags.path,
      matchMode: flags.exact ? "exact" : "contains",
      profile: flags.profile,
    });
  }

  if (group === "window" && action === "capture") {
    return await captureWindow({
      hwnd: flags.hwnd,
      title: flags.title,
      fileName: flags["file-name"],
    });
  }

  if (group === "screen" && action === "capture") {
    return await captureScreen({
      screenIndex: parseOptionalNumberFlag(flags["screen-index"]),
      fileName: flags["file-name"],
    });
  }

  if (group === "mouse" && action === "move") {
    return await moveMouse({
      x: parseNumberFlag(flags.x, "x"),
      y: parseNumberFlag(flags.y, "y"),
      profile: flags.profile,
    });
  }

  if (group === "mouse" && action === "click") {
    return await click({
      x: parseNumberFlag(flags.x, "x"),
      y: parseNumberFlag(flags.y, "y"),
      button: flags.button ?? "left",
      double: Boolean(flags.double),
      profile: flags.profile,
    });
  }

  if (group === "mouse" && action === "drag") {
    return await drag({
      fromX: parseNumberFlag(flags["from-x"], "from-x"),
      fromY: parseNumberFlag(flags["from-y"], "from-y"),
      toX: parseNumberFlag(flags["to-x"], "to-x"),
      toY: parseNumberFlag(flags["to-y"], "to-y"),
      button: flags.button ?? "left",
      steps: parseOptionalNumberFlag(flags.steps) ?? 16,
      durationMs: parseOptionalNumberFlag(flags["duration-ms"]) ?? 300,
      profile: flags.profile,
    });
  }

  if (group === "click") {
    if (!flags.on) {
      throw new CliUsageError("Missing required flag: --on");
    }

    return await clickElementByLabel({
      label: flags.on,
      snapshotId: flags.snapshot,
      mode: flags.mode,
      hwnd: flags.hwnd,
      title: flags.title,
      screenIndex: parseOptionalNumberFlag(flags["screen-index"]),
      fileName: flags["file-name"],
      annotatedFileName: flags["annotated-file-name"],
      matchMode: flags.exact ? "exact" : "contains",
      profile: flags.profile,
    });
  }

  if (group === "scroll") {
    if (flags.on) {
      return await scrollElementByLabel({
        label: flags.on,
        snapshotId: flags.snapshot,
        mode: flags.mode,
        hwnd: flags.hwnd,
        title: flags.title,
        screenIndex: parseOptionalNumberFlag(flags["screen-index"]),
        fileName: flags["file-name"],
        annotatedFileName: flags["annotated-file-name"],
        direction: flags.direction ?? "down",
        ticks: parseOptionalNumberFlag(flags.ticks) ?? 3,
        matchMode: flags.exact ? "exact" : "contains",
        profile: flags.profile,
      });
    }

    return await scroll({
      x: parseOptionalNumberFlag(flags.x),
      y: parseOptionalNumberFlag(flags.y),
      direction: flags.direction ?? "down",
      ticks: parseOptionalNumberFlag(flags.ticks) ?? 3,
      profile: flags.profile,
    });
  }

  if (group === "see") {
    const mode =
      flags.mode ?? (flags.hwnd || flags.title ? "window" : "screen");

    if (!["screen", "window"].includes(mode)) {
      throw new CliUsageError(
        "Invalid --mode value. Expected screen or window",
      );
    }

    if (mode === "window" && !flags.hwnd && !flags.title) {
      throw new CliUsageError("Window see requires --hwnd or --title");
    }

    return await seeUi({
      mode,
      hwnd: flags.hwnd,
      title: flags.title,
      screenIndex: parseOptionalNumberFlag(flags["screen-index"]),
      fileName: flags["file-name"],
      annotatedFileName: flags["annotated-file-name"],
      snapshotId: flags.snapshot,
      maxResults: parseOptionalNumberFlag(flags["max-results"]),
      matchMode: flags.exact ? "exact" : "contains",
    });
  }

  if (group === "snapshot" && action === "list") {
    return await listStoredSnapshots({
      limit: parseOptionalNumberFlag(flags.limit),
    });
  }

  if (group === "snapshot" && action === "show") {
    if (!flags.snapshot) {
      throw new CliUsageError("Missing required flag: --snapshot");
    }

    return await getStoredSnapshot({
      snapshotId: flags.snapshot,
    });
  }

  if (group === "snapshot" && action === "click") {
    if (!flags.snapshot) {
      throw new CliUsageError("Missing required flag: --snapshot");
    }

    if (!flags["element-id"] && !flags.name) {
      throw new CliUsageError("Missing required flag: --element-id or --name");
    }

    return await clickSnapshotElement({
      snapshotId: flags.snapshot,
      elementId: flags["element-id"],
      name: flags.name,
      matchMode: flags.exact ? "exact" : "contains",
      profile: flags.profile,
    });
  }

  if (group === "snapshot" && action === "clean") {
    return await cleanStoredSnapshots({
      snapshotId: flags.snapshot,
      all: Boolean(flags.all),
      olderThanHours: parseOptionalNumberFlag(flags["older-than-hours"]),
    });
  }

  if (group === "ai" && action === "brief") {
    if (!flags.snapshot) {
      throw new CliUsageError("Missing required flag: --snapshot");
    }

    let snapshotId = flags.snapshot;

    if (snapshotId === "latest") {
      const latest = await listStoredSnapshots({ limit: 1 });
      snapshotId = latest.snapshots?.[0]?.snapshotId;
    }

    if (!snapshotId) {
      throw new CliUsageError("No snapshots are available");
    }

    return buildSnapshotAiBrief(
      await getStoredSnapshot({
        snapshotId,
      }),
    );
  }

  if (group === "harvest" && action === "text") {
    return await harvestScrollText(buildAutomationIntentArgs(flags));
  }

  if (group === "recipe" && action === "list") {
    return listRecipes();
  }

  if (group === "recipe" && action === "run") {
    if (!flags.id) {
      throw new CliUsageError("Missing required flag: --id");
    }

    return await runRecipe({
      recipeId: flags.id,
      continueOnError: flags["no-fail-fast"] ? true : undefined,
      ...buildAutomationIntentArgs(flags),
    });
  }

  if (group === "goal" && action === "plan") {
    if (!flags.text) {
      throw new CliUsageError("Missing required flag: --text");
    }

    return planGoal({
      goal: flags.text,
      ...buildAutomationIntentArgs(flags),
    });
  }

  if (group === "goal" && action === "run") {
    if (!flags.text) {
      throw new CliUsageError("Missing required flag: --text");
    }

    return await runGoal({
      goal: flags.text,
      continueOnError: flags["no-fail-fast"] ? true : undefined,
      ...buildAutomationIntentArgs(flags),
    });
  }

  if (group === "snapshot" && action === "drag") {
    if (!flags.snapshot) {
      throw new CliUsageError("Missing required flag: --snapshot");
    }

    if (!flags["from-element-id"] && !flags["from-name"]) {
      throw new CliUsageError(
        "Missing required flag: --from-element-id or --from-name",
      );
    }

    if (!flags["to-element-id"] && !flags["to-name"]) {
      throw new CliUsageError(
        "Missing required flag: --to-element-id or --to-name",
      );
    }

    return await dragSnapshotElement({
      snapshotId: flags.snapshot,
      fromElementId: flags["from-element-id"],
      toElementId: flags["to-element-id"],
      fromName: flags["from-name"],
      toName: flags["to-name"],
      button: flags.button ?? "left",
      steps: parseOptionalNumberFlag(flags.steps) ?? 16,
      durationMs: parseOptionalNumberFlag(flags["duration-ms"]) ?? 300,
      matchMode: flags.exact ? "exact" : "contains",
      profile: flags.profile,
    });
  }

  if (group === "snapshot" && action === "scroll") {
    if (!flags.snapshot) {
      throw new CliUsageError("Missing required flag: --snapshot");
    }

    if (!flags["element-id"] && !flags.name) {
      throw new CliUsageError("Missing required flag: --element-id or --name");
    }

    return await scrollSnapshotElement({
      snapshotId: flags.snapshot,
      elementId: flags["element-id"],
      name: flags.name,
      direction: flags.direction ?? "down",
      ticks: parseOptionalNumberFlag(flags.ticks) ?? 3,
      matchMode: flags.exact ? "exact" : "contains",
      profile: flags.profile,
    });
  }

  if (group === "run") {
    if (!flags.file) {
      throw new CliUsageError("Missing required flag: --file");
    }

    return await runWorkflowFile({
      filePath: flags.file,
      continueOnError: flags["no-fail-fast"] ? true : undefined,
      profile: flags.profile,
    });
  }

  if (group === "hotkey") {
    if (!flags.keys) {
      throw new CliUsageError("Missing required flag: --keys");
    }

    return await hotkey({
      keys: String(flags.keys).split(","),
      repeat: parseOptionalNumberFlag(flags.repeat) ?? 1,
      delayMs: parseOptionalNumberFlag(flags["delay-ms"]) ?? 80,
      profile: flags.profile,
    });
  }

  if (group === "ui" && action === "find") {
    return await findUiElements({
      hwnd: flags.hwnd,
      title: flags.title,
      name: flags.name,
      automationId: flags["automation-id"],
      className: flags["class-name"],
      controlType: flags["control-type"],
      maxResults: parseOptionalNumberFlag(flags["max-results"]),
      matchMode: flags.exact ? "exact" : "contains",
      profile: flags.profile,
    });
  }

  if (group === "ui" && action === "wait") {
    if (!flags.name) {
      throw new CliUsageError("Missing required flag: --name");
    }

    return await waitForUiElement({
      hwnd: flags.hwnd,
      title: flags.title,
      name: flags.name,
      automationId: flags["automation-id"],
      className: flags["class-name"],
      controlType: flags["control-type"],
      maxResults: parseOptionalNumberFlag(flags["max-results"]),
      matchMode: flags.exact ? "exact" : "contains",
      timeoutMs: parseOptionalNumberFlag(flags["timeout-ms"]) ?? 5000,
      pollMs: parseOptionalNumberFlag(flags["poll-ms"]) ?? 200,
    });
  }

  if (group === "ui" && action === "click") {
    if (!flags.name) {
      throw new CliUsageError("Missing required flag: --name");
    }

    return await clickUiElement({
      hwnd: flags.hwnd,
      title: flags.title,
      name: flags.name,
      automationId: flags["automation-id"],
      className: flags["class-name"],
      controlType: flags["control-type"],
      maxResults: parseOptionalNumberFlag(flags["max-results"]),
      matchMode: flags.exact ? "exact" : "contains",
    });
  }

  if (group === "type") {
    if (!flags.text && !flags["text-file"]) {
      throw new CliUsageError("Missing required flag: --text or --text-file");
    }

    return await typeText({
      text: flags["text-file"]
        ? await fs.readFile(String(flags["text-file"]), "utf8")
        : flags.text,
      clear: Boolean(flags.clear),
      delayMs: parseOptionalNumberFlag(flags["delay-ms"]) ?? 80,
      profile: flags.profile,
    });
  }

  if (group === "press") {
    if (!flags.keys) {
      throw new CliUsageError("Missing required flag: --keys");
    }

    return await pressKeys({ keys: flags.keys });
  }

  if (group === "app" && action === "launch") {
    if (!flags.command) {
      throw new CliUsageError("Missing required flag: --command");
    }

    return await launchApp({
      command: flags.command,
      args: flags.args ? String(flags.args).split(",").filter(Boolean) : [],
    });
  }

  if (group === "app" && action === "switch") {
    return await switchApp({
      processId: parseOptionalNumberFlag(flags["process-id"]),
      name: flags.name,
      title: flags.title,
      matchMode: flags.exact ? "exact" : "contains",
    });
  }

  if (group === "app" && action === "quit") {
    return await quitApp({
      processId: parseOptionalNumberFlag(flags["process-id"]),
      name: flags.name,
      title: flags.title,
      matchMode: flags.exact ? "exact" : "contains",
    });
  }

  throw new CliUsageError(`Unknown command.\n\n${helpText}`);
}

function parseFlags(argv) {
  const flags = {};

  for (let index = 0; index < argv.length; index += 1) {
    const token = argv[index];

    if (!token.startsWith("--")) {
      continue;
    }

    const key = token.slice(2);
    const next = argv[index + 1];

    if (!next || next.startsWith("--")) {
      flags[key] = true;
      continue;
    }

    flags[key] = next;
    index += 1;
  }

  return flags;
}

function parseNumberFlag(value, key) {
  const numericValue = Number(value);

  if (!Number.isFinite(numericValue)) {
    throw new CliUsageError(`Missing or invalid --${key} value`);
  }

  return numericValue;
}

function parseOptionalNumberFlag(value) {
  if (value === undefined) {
    return undefined;
  }

  const numericValue = Number(value);
  return Number.isFinite(numericValue) ? numericValue : undefined;
}

function buildAutomationIntentArgs(flags) {
  return compactObject({
    snapshotId: flags.snapshot,
    mode: flags.mode,
    hwnd: flags.hwnd,
    title: flags.title,
    processId: parseOptionalNumberFlag(flags["process-id"]),
    processName: flags["process-name"],
    screenIndex: parseOptionalNumberFlag(flags["screen-index"]),
    label: flags.label,
    scrollLabel: flags.label,
    text: flags["text-file"] ? undefined : flags.text,
    clear: Boolean(flags.clear),
    command: flags.command,
    args: parseOptionalListFlag(flags.args),
    x: parseOptionalNumberFlag(flags.x),
    y: parseOptionalNumberFlag(flags.y),
    direction: flags.direction,
    ticks: parseOptionalNumberFlag(flags.ticks),
    maxSteps: parseOptionalNumberFlag(flags["max-steps"]),
    stopAfterStalledSteps: parseOptionalNumberFlag(
      flags["stop-after-stalled-steps"],
    ),
    overlapWindow: parseOptionalNumberFlag(flags["overlap-window"]),
    fuzzyThreshold: parseOptionalNumberFlag(flags["fuzzy-threshold"]),
    pauseAfterScrollMs: parseOptionalNumberFlag(flags["pause-after-scroll-ms"]),
    outputPath: flags.output,
    overwrite: Boolean(flags.overwrite),
    timeoutMs: parseOptionalNumberFlag(flags["timeout-ms"]),
    matchMode: flags.exact ? "exact" : "contains",
    profile: flags.profile,
  });
}

function parseOptionalListFlag(value) {
  if (value === undefined) {
    return undefined;
  }

  return String(value)
    .split(",")
    .map((entry) => entry.trim())
    .filter(Boolean);
}

function compactObject(value) {
  return Object.fromEntries(
    Object.entries(value).filter(
      ([, entry]) => entry !== undefined && entry !== null && entry !== "",
    ),
  );
}
