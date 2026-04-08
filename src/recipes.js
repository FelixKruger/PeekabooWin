import { buildSnapshotAiBrief } from "./ai-bridge.js";
import { getStoredSnapshot } from "./backend/tools.js";
import { harvestScrollText } from "./harvest.js";
import { resolveSnapshotReference } from "./snapshots.js";
import { runWorkflowDefinition } from "./workflow/runner.js";

const RECIPE_DEFINITIONS = [
  {
    id: "inspect-app",
    title: "Inspect app",
    description:
      "Take a fresh snapshot of a window or screen so PeekabooWin can see controls and text.",
    requiresTarget: true,
    examples: ["inspect this app", "take a snapshot", "look at this window"],
  },
  {
    id: "read-screen-text",
    title: "Read screen text",
    description:
      "Capture or reuse a snapshot and return the text Windows OCR can read from it.",
    requiresTarget: true,
    supportsSnapshot: true,
    examples: [
      "read the text on this screen",
      "what does this app say?",
      "ocr this window",
    ],
  },
  {
    id: "harvest-scroll-text",
    title: "Harvest scrolling text",
    description:
      "Capture, scroll, and collect a long OCR transcript from the current app or screen.",
    requiresTarget: true,
    supportsSnapshot: true,
    examples: [
      "scrape this thread",
      "scroll and collect the messages",
      "harvest the visible conversation",
    ],
  },
  {
    id: "handoff-to-ai",
    title: "Prepare AI handoff",
    description:
      "Capture or reuse a snapshot and build a ready-to-paste AI brief.",
    requiresTarget: true,
    supportsSnapshot: true,
    examples: [
      "summarize this app for AI",
      "prepare an AI handoff",
      "copy a summary for chatgpt",
    ],
  },
  {
    id: "click-visible-label",
    title: "Click visible label",
    description: "Click a visible UI label or OCR-recognized text target.",
    requiresTarget: true,
    supportsSnapshot: true,
    requiredArgs: ["label"],
    examples: ["click Save", 'press "Continue"', "select OK"],
  },
  {
    id: "type-into-app",
    title: "Type into app",
    description: "Focus a target app window and type text into it.",
    requiresTarget: true,
    requiredArgs: ["text"],
    examples: [
      'type "hello world"',
      "enter my email address",
      "write this into the app",
    ],
  },
  {
    id: "open-and-inspect",
    title: "Open and inspect app",
    description:
      "Launch a known app or command, wait for its window, and take a fresh snapshot.",
    requiredArgs: ["command"],
    examples: [
      "open notepad",
      "launch calculator and inspect it",
      "open paint",
    ],
  },
];

const APP_ALIASES = {
  notepad: "notepad.exe",
  calculator: "calc.exe",
  calc: "calc.exe",
  paint: "mspaint.exe",
};

export function listRecipes() {
  return {
    recipes: RECIPE_DEFINITIONS,
  };
}

export function planGoal({
  goal,
  snapshotId,
  hwnd,
  title,
  processId,
  processName,
  mode,
  screenIndex,
} = {}) {
  if (!goal || !String(goal).trim()) {
    throw new Error("Missing goal text");
  }

  const normalizedGoal = String(goal).trim();
  const lowerGoal = normalizedGoal.toLowerCase();
  const selectors = compactObject({
    snapshotId,
    hwnd,
    title,
    processId,
    processName,
    mode,
    screenIndex,
  });
  const quotedValues = extractQuotedValues(normalizedGoal);
  const missing = [];
  let recipeId = "inspect-app";
  let reason = "Defaulting to a fresh inspection step.";
  let confidence = 0.55;
  const args = { ...selectors };

  const command = extractCommandFromGoal(lowerGoal);
  const label = quotedValues[0] ?? extractLabelFromGoal(normalizedGoal);
  const text = quotedValues[0] ?? extractTextFromGoal(normalizedGoal);

  if (command) {
    recipeId = "open-and-inspect";
    args.command = command;
    reason = "The goal starts by opening a known app.";
    confidence = 0.9;
  } else if (
    containsAny(lowerGoal, [
      "handoff",
      "summary",
      "summarize",
      "brief",
      "for ai",
      "chatgpt",
      "claude",
      "codex",
    ])
  ) {
    recipeId = "handoff-to-ai";
    reason = "The goal asks for an AI-ready summary.";
    confidence = 0.86;
  } else if (
    containsAny(lowerGoal, [
      "harvest",
      "scrape",
      "scroll and",
      "collect messages",
      "collect the messages",
      "read hundreds",
      "extract conversation",
      "capture the thread",
    ])
  ) {
    recipeId = "harvest-scroll-text";
    reason = "The goal asks for a longer scroll-and-read collection pass.";
    confidence = 0.88;
  } else if (
    containsAny(lowerGoal, [
      "read text",
      "ocr",
      "what does",
      "what can you read",
      "text on",
    ])
  ) {
    recipeId = "read-screen-text";
    reason = "The goal focuses on text recognition rather than action.";
    confidence = 0.86;
  } else if (containsAny(lowerGoal, ["click", "press", "select", "choose"])) {
    recipeId = "click-visible-label";
    if (label) {
      args.label = label;
    } else {
      missing.push("label");
    }
    reason =
      "The goal asks PeekabooWin to activate something visible on screen.";
    confidence = label ? 0.84 : 0.62;
  } else if (containsAny(lowerGoal, ["type", "enter", "write", "fill"])) {
    recipeId = "type-into-app";
    if (text) {
      args.text = text;
    } else {
      missing.push("text");
    }
    reason = "The goal asks PeekabooWin to send text into an app.";
    confidence = text ? 0.84 : 0.62;
  } else if (
    containsAny(lowerGoal, ["inspect", "snapshot", "capture", "look at", "see"])
  ) {
    recipeId = "inspect-app";
    reason = "The goal asks for a fresh capture or inspection.";
    confidence = 0.78;
  }

  const definition = getRecipeDefinition(recipeId);

  if (definition.requiresTarget && !hasTargetContext(args)) {
    missing.push("target");
  }

  const uniqueMissing = [...new Set(missing)];
  const ok = uniqueMissing.length === 0;

  return {
    ok,
    goal: normalizedGoal,
    recipeId,
    recipeTitle: definition.title,
    confidence,
    reason,
    missing: uniqueMissing,
    args,
    preview: buildRecipePreview(recipeId, args),
  };
}

export async function runGoal({
  goal,
  actionHandlers,
  continueOnError,
  profile,
  ...context
} = {}) {
  const plan = planGoal({
    goal,
    ...context,
  });

  if (!plan.ok) {
    throw new Error(
      `Goal is missing required information: ${plan.missing.join(", ")}`,
    );
  }

  const execution = await runRecipe({
    recipeId: plan.recipeId,
    actionHandlers,
    continueOnError,
    profile,
    goal,
    ...context,
    ...plan.args,
  });

  return {
    goal: plan.goal,
    plan,
    execution,
  };
}

export async function runRecipe({
  recipeId,
  actionHandlers,
  continueOnError,
  profile,
  ...args
} = {}) {
  if (!recipeId) {
    throw new Error("Missing recipeId");
  }

  const definition = getRecipeDefinition(recipeId);
  let result;

  switch (recipeId) {
    case "inspect-app":
      result = await runInspectRecipe({
        args,
        actionHandlers,
        continueOnError,
        profile,
      });
      break;
    case "read-screen-text":
      result = await runReadTextRecipe({
        args,
        actionHandlers,
        continueOnError,
        profile,
      });
      break;
    case "harvest-scroll-text":
      result = await runHarvestTextRecipe({ args, profile });
      break;
    case "handoff-to-ai":
      result = await runAiHandoffRecipe({
        args,
        actionHandlers,
        continueOnError,
        profile,
      });
      break;
    case "click-visible-label":
      result = await runClickLabelRecipe({
        args,
        actionHandlers,
        continueOnError,
        profile,
      });
      break;
    case "type-into-app":
      result = await runTypeIntoAppRecipe({
        args,
        actionHandlers,
        continueOnError,
        profile,
      });
      break;
    case "open-and-inspect":
      result = await runOpenAndInspectRecipe({
        args,
        actionHandlers,
        continueOnError,
        profile,
      });
      break;
    default:
      throw new Error(`Unsupported recipe '${recipeId}'`);
  }

  return {
    recipeId: definition.id,
    recipeTitle: definition.title,
    ...result,
  };
}

function getRecipeDefinition(recipeId) {
  const recipe = RECIPE_DEFINITIONS.find((entry) => entry.id === recipeId);

  if (!recipe) {
    throw new Error(`Unknown recipe '${recipeId}'`);
  }

  return recipe;
}

async function runInspectRecipe({
  args,
  actionHandlers,
  continueOnError,
  profile,
}) {
  ensureTargetContext(args, "Inspect app");

  const workflow = await runWorkflowDefinition({
    definition: {
      steps: [
        {
          action: "see",
          saveAs: "capture",
          ...selectCaptureArgs(args),
        },
      ],
    },
    actionHandlers,
    continueOnError,
    profile,
  });

  return await buildSnapshotRecipeResult({
    workflow,
    snapshotId: workflow.context.capture?.snapshotId,
    summary: "Fresh snapshot ready.",
  });
}

async function runReadTextRecipe({
  args,
  actionHandlers,
  continueOnError,
  profile,
}) {
  const snapshot = await resolveRecipeSnapshot({
    args,
    actionHandlers,
    continueOnError,
    profile,
  });
  const ocr = snapshot.snapshot?.result?.ocr ?? {};

  return {
    ...snapshot,
    ocrText: ocr.text ?? "",
    ocrLines: Array.isArray(ocr.lines) ? ocr.lines : [],
    summary: ocr.text
      ? `Read ${Array.isArray(ocr.lines) ? ocr.lines.length : 0} text lines from the snapshot.`
      : "No readable text was found in the snapshot.",
  };
}

async function runAiHandoffRecipe({
  args,
  actionHandlers,
  continueOnError,
  profile,
}) {
  const snapshot = await resolveRecipeSnapshot({
    args,
    actionHandlers,
    continueOnError,
    profile,
  });
  const brief = buildSnapshotAiBrief(snapshot.snapshot);

  return {
    ...snapshot,
    brief,
    summary: "AI handoff brief is ready.",
  };
}

async function runHarvestTextRecipe({ args, profile }) {
  ensureTargetContext(args, "Harvest scrolling text");

  const result = await harvestScrollText({
    ...args,
    profile,
  });

  return {
    ...result,
    summary: `Collected ${result.totalLines} OCR lines across ${result.snapshotCount} snapshots.`,
  };
}

async function runClickLabelRecipe({
  args,
  actionHandlers,
  continueOnError,
  profile,
}) {
  if (!args.label) {
    throw new Error("Click visible label recipe requires a label");
  }

  ensureTargetContext(args, "Click visible label");

  const workflow = await runWorkflowDefinition({
    definition: {
      steps: [
        {
          action: "element.click",
          saveAs: "click",
          label: args.label,
          ...selectElementArgs(args),
        },
      ],
    },
    actionHandlers,
    continueOnError,
    profile,
  });

  return {
    workflow,
    actionResult: workflow.context.click ?? null,
    summary: `Clicked "${args.label}".`,
  };
}

async function runTypeIntoAppRecipe({
  args,
  actionHandlers,
  continueOnError,
  profile,
}) {
  if (!args.text) {
    throw new Error("Type into app recipe requires text");
  }

  ensureTargetContext(args, "Type into app");

  const steps = [];
  const focusStep = buildFocusStep(args);

  if (focusStep) {
    steps.push(focusStep);
  }

  steps.push({
    action: "type",
    saveAs: "typed",
    text: args.text,
    clear: args.clear === true,
  });

  const workflow = await runWorkflowDefinition({
    definition: {
      steps,
    },
    actionHandlers,
    continueOnError,
    profile,
  });

  return {
    workflow,
    actionResult: workflow.context.typed ?? null,
    summary:
      args.clear === true
        ? "Replaced the selected text."
        : "Typed text into the target app.",
  };
}

async function runOpenAndInspectRecipe({
  args,
  actionHandlers,
  continueOnError,
  profile,
}) {
  if (!args.command) {
    throw new Error("Open and inspect recipe requires a command");
  }

  const workflow = await runWorkflowDefinition({
    definition: {
      steps: [
        {
          action: "app.launch",
          command: args.command,
          args: Array.isArray(args.args) ? args.args : [],
          saveAs: "launch",
        },
        {
          action: "window.wait",
          processId: "$launch.processId",
          timeoutMs: args.timeoutMs ?? 5000,
        },
        {
          action: "see",
          hwnd: "$launch.hwnds.0",
          mode: "window",
          saveAs: "capture",
        },
      ],
    },
    actionHandlers,
    continueOnError,
    profile,
  });

  return await buildSnapshotRecipeResult({
    workflow,
    snapshotId: workflow.context.capture?.snapshotId,
    summary: `Opened ${args.command} and captured it.`,
  });
}

async function resolveRecipeSnapshot({
  args,
  actionHandlers,
  continueOnError,
  profile,
}) {
  if (args.snapshotId) {
    const resolvedSnapshotId = await resolveSnapshotReference(args.snapshotId);

    if (!resolvedSnapshotId) {
      throw new Error("No matching snapshot is available");
    }

    const snapshot = await getStoredSnapshot({
      snapshotId: resolvedSnapshotId,
    });

    return {
      workflow: null,
      snapshotId: resolvedSnapshotId,
      snapshot,
    };
  }

  ensureTargetContext(args, "Recipe");

  return await runInspectRecipe({
    args,
    actionHandlers,
    continueOnError,
    profile,
  });
}

async function buildSnapshotRecipeResult({ workflow, snapshotId, summary }) {
  if (!snapshotId) {
    return {
      workflow,
      snapshotId: null,
      snapshot: null,
      summary,
    };
  }

  const snapshot = await getStoredSnapshot({ snapshotId });
  const brief = buildSnapshotAiBrief(snapshot);

  return {
    workflow,
    snapshotId,
    snapshot,
    brief,
    ocrText: snapshot.result?.ocr?.text ?? "",
    summary,
  };
}

function ensureTargetContext(args, name) {
  if (!hasTargetContext(args)) {
    throw new Error(
      `${name} requires a snapshot, window selector, app selector, or screen mode`,
    );
  }
}

function hasTargetContext(args) {
  return Boolean(
    args.snapshotId ||
    args.hwnd ||
    args.title ||
    args.processId !== undefined ||
    args.processName ||
    args.mode ||
    args.screenIndex !== undefined,
  );
}

function selectCaptureArgs(args) {
  return compactObject({
    mode: args.mode,
    hwnd: args.hwnd,
    title: args.title,
    screenIndex: args.screenIndex,
  });
}

function selectElementArgs(args) {
  if (args.snapshotId) {
    return {
      snapshotId: args.snapshotId,
    };
  }

  return selectCaptureArgs(args);
}

function buildFocusStep(args) {
  if (args.hwnd || args.title) {
    return {
      action: "window.focus",
      hwnd: args.hwnd,
      title: args.title,
    };
  }

  if (args.processId !== undefined || args.processName) {
    return {
      action: "app.switch",
      processId: args.processId,
      name: args.processName,
    };
  }

  return null;
}

function buildRecipePreview(recipeId, args) {
  switch (recipeId) {
    case "inspect-app":
      return ["Take a fresh snapshot of the selected app or screen."];
    case "read-screen-text":
      return args.snapshotId
        ? ["Reuse the selected snapshot.", "Read recognized text from it."]
        : ["Take a fresh snapshot.", "Read recognized text from it."];
    case "harvest-scroll-text":
      return [
        args.snapshotId
          ? "Reuse the selected snapshot."
          : "Take a fresh snapshot.",
        "Scroll through the target while collecting OCR text.",
        "Save the full transcript to a local text file.",
      ];
    case "handoff-to-ai":
      return args.snapshotId
        ? ["Reuse the selected snapshot.", "Build an AI-ready summary."]
        : ["Take a fresh snapshot.", "Build an AI-ready summary."];
    case "click-visible-label":
      return [
        `Find "${args.label ?? "the requested label"}" in the target app.`,
        "Click it.",
      ];
    case "type-into-app":
      return [
        "Bring the target app forward.",
        args.clear ? "Replace the current text." : "Type the requested text.",
      ];
    case "open-and-inspect":
      return [
        `Open ${args.command ?? "the requested app"}.`,
        "Wait for its window.",
        "Take a fresh snapshot.",
      ];
    default:
      return [];
  }
}

function extractQuotedValues(text) {
  const matches = [];
  const pattern = /"([^"]+)"|'([^']+)'/g;
  let match = null;

  while ((match = pattern.exec(text)) !== null) {
    matches.push(match[1] ?? match[2] ?? "");
  }

  return matches.filter(Boolean);
}

function extractCommandFromGoal(lowerGoal) {
  const appMatch = /\b(open|launch)\s+([a-z0-9 .+-]+)/i.exec(lowerGoal);

  if (!appMatch) {
    return null;
  }

  const rawTarget = appMatch[2]
    .replace(/\band inspect\b.*$/i, "")
    .replace(/\band capture\b.*$/i, "")
    .replace(/\band then\b.*$/i, "")
    .trim();

  return APP_ALIASES[rawTarget] ?? null;
}

function extractLabelFromGoal(goal) {
  const match =
    /\b(?:click|press|select|choose)\s+(.+?)(?:\s+(?:in|on|for)\b.*)?$/i.exec(
      goal,
    );
  return match?.[1] ? cleanupExtractedText(match[1]) : null;
}

function extractTextFromGoal(goal) {
  const match =
    /\b(?:type|enter|write|fill)\s+(.+?)(?:\s+(?:into|in|on)\b.*)?$/i.exec(
      goal,
    );
  return match?.[1] ? cleanupExtractedText(match[1]) : null;
}

function cleanupExtractedText(text) {
  return String(text)
    .replace(/[.?!]+$/, "")
    .trim();
}

function compactObject(value) {
  return Object.fromEntries(
    Object.entries(value).filter(
      ([, entry]) => entry !== undefined && entry !== null && entry !== "",
    ),
  );
}

function containsAny(text, patterns) {
  return patterns.some((pattern) => text.includes(pattern));
}
