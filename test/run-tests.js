import assert from "node:assert/strict";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { PassThrough } from "node:stream";

import { buildSnapshotAiBrief } from "../src/ai-bridge.js";
import {
  harvestScrollText,
  levenshteinRatio,
  mergeHarvestLines,
  normalizeHarvestText,
} from "../src/harvest.js";
import { runMcpServer } from "../src/mcp/server.js";
import { translateHotkey } from "../src/hotkeys.js";
import { listRecipes, planGoal, runGoal } from "../src/recipes.js";
import {
  prepareSnapshotPaths,
  readSnapshot,
  resolveSnapshotElement,
  resolveSnapshotLabelTarget,
  resolveSnapshotReference,
  writeSnapshotMetadata,
} from "../src/snapshots.js";
import {
  runWorkflowDefinition,
  runWorkflowFile,
} from "../src/workflow/runner.js";

async function main() {
  await testInitialize();
  await testToolsList();
  await testSnapshotHelpers();
  testAiBridge();
  await testHarvestModule();
  testLevenshteinRatio();
  testFuzzyMerge();
  await testHarvestAutoFallback();
  await testRecipes();
  await testWorkflowRunner();
  testHotkeys();
  process.stdout.write("All tests passed.\n");
}

async function testInitialize() {
  const input = new PassThrough();
  const output = new PassThrough();
  let raw = "";

  output.on("data", (chunk) => {
    raw += chunk.toString("utf8");
  });

  const serverPromise = runMcpServer({
    input,
    output,
    error: new PassThrough(),
  });
  input.write(
    '{"jsonrpc":"2.0","id":1,"method":"initialize","params":{"protocolVersion":"2024-11-05","capabilities":{},"clientInfo":{"name":"test","version":"0.0.0"}}}\n',
  );
  input.end();
  await serverPromise;

  const lines = raw.trim().split("\n");
  const message = JSON.parse(lines[0]);

  assert.equal(message.id, 1);
  assert.equal(message.result.serverInfo.name, "peekaboo-windows");
  assert.equal(message.result.protocolVersion, "2024-11-05");
}

async function testToolsList() {
  const input = new PassThrough();
  const output = new PassThrough();
  let raw = "";

  output.on("data", (chunk) => {
    raw += chunk.toString("utf8");
  });

  const serverPromise = runMcpServer({
    input,
    output,
    error: new PassThrough(),
  });
  input.write('{"jsonrpc":"2.0","id":1,"method":"tools/list","params":{}}\n');
  input.end();
  await serverPromise;

  const message = JSON.parse(raw.trim());
  assert.equal(message.id, 1);
  assert.ok(Array.isArray(message.result.tools));
  assert.ok(
    message.result.tools.some((tool) => tool.name === "screen_capture"),
  );
  assert.ok(message.result.tools.some((tool) => tool.name === "screens_list"));
  assert.ok(message.result.tools.some((tool) => tool.name === "window_move"));
  assert.ok(message.result.tools.some((tool) => tool.name === "window_resize"));
  assert.ok(
    message.result.tools.some((tool) => tool.name === "window_set_bounds"),
  );
  assert.ok(
    message.result.tools.some((tool) => tool.name === "window_set_state"),
  );
  assert.ok(message.result.tools.some((tool) => tool.name === "window_wait"));
  assert.ok(message.result.tools.some((tool) => tool.name === "ui_snapshot"));
  assert.ok(message.result.tools.some((tool) => tool.name === "snapshot_list"));
  assert.ok(message.result.tools.some((tool) => tool.name === "snapshot_get"));
  assert.ok(message.result.tools.some((tool) => tool.name === "dialog_list"));
  assert.ok(message.result.tools.some((tool) => tool.name === "dialog_click"));
  assert.ok(message.result.tools.some((tool) => tool.name === "menu_list"));
  assert.ok(message.result.tools.some((tool) => tool.name === "menu_click"));
  assert.ok(
    message.result.tools.some((tool) => tool.name === "snapshot_click"),
  );
  assert.ok(message.result.tools.some((tool) => tool.name === "element_click"));
  assert.ok(
    message.result.tools.some((tool) => tool.name === "snapshot_clean"),
  );
  assert.ok(message.result.tools.some((tool) => tool.name === "mouse_drag"));
  assert.ok(message.result.tools.some((tool) => tool.name === "snapshot_drag"));
  assert.ok(message.result.tools.some((tool) => tool.name === "workflow_run"));
  assert.ok(
    message.result.tools.some((tool) => tool.name === "harvest_scroll_text"),
  );
  assert.ok(message.result.tools.some((tool) => tool.name === "recipe_list"));
  assert.ok(message.result.tools.some((tool) => tool.name === "recipe_run"));
  assert.ok(message.result.tools.some((tool) => tool.name === "goal_plan"));
  assert.ok(message.result.tools.some((tool) => tool.name === "goal_run"));
  assert.ok(message.result.tools.some((tool) => tool.name === "hotkey_press"));
  assert.ok(message.result.tools.some((tool) => tool.name === "ui_wait"));
  assert.ok(message.result.tools.some((tool) => tool.name === "scroll"));
  assert.ok(
    message.result.tools.some((tool) => tool.name === "snapshot_scroll"),
  );
  assert.ok(
    message.result.tools.some((tool) => tool.name === "element_scroll"),
  );
  assert.ok(message.result.tools.some((tool) => tool.name === "app_list"));
  assert.ok(message.result.tools.some((tool) => tool.name === "app_switch"));
  assert.ok(message.result.tools.some((tool) => tool.name === "app_quit"));
}

async function testHarvestModule() {
  assert.equal(normalizeHarvestText("  Hello   WORLD "), "hello world");

  const mergedDown = mergeHarvestLines(
    [
      { text: "A", normalizedText: "a" },
      { text: "B", normalizedText: "b" },
      { text: "C", normalizedText: "c" },
    ],
    [
      { text: "C", normalizedText: "c" },
      { text: "D", normalizedText: "d" },
    ],
    { direction: "down" },
  );
  assert.deepEqual(
    mergedDown.lines.map((line) => line.text),
    ["A", "B", "C", "D"],
  );

  const mergedUp = mergeHarvestLines(
    [
      { text: "C", normalizedText: "c" },
      { text: "D", normalizedText: "d" },
    ],
    [
      { text: "A", normalizedText: "a" },
      { text: "B", normalizedText: "b" },
      { text: "C", normalizedText: "c" },
    ],
    { direction: "up" },
  );
  assert.deepEqual(
    mergedUp.lines.map((line) => line.text),
    ["A", "B", "C", "D"],
  );

  const tempDir = await fs.mkdtemp(
    path.join(os.tmpdir(), "peekaboo-harvest-test-"),
  );
  const snapshots = [
    makeSnapshot("snap-1", ["Line 1", "Line 2", "Line 3"]),
    makeSnapshot("snap-2", ["Line 3", "Line 4", "Line 5"]),
    makeSnapshot("snap-3", ["Line 3", "Line 4", "Line 5"]),
  ];
  const scrollCalls = [];
  let captureIndex = 0;

  try {
    const result = await harvestScrollText(
      {
        mode: "window",
        hwnd: "0x123",
        maxSteps: 5,
        stopAfterStalledSteps: 1,
        outputPath: path.join(tempDir, "thread.txt"),
      },
      {
        capture: async () => snapshots[captureIndex++],
        scrollAt: async (args) => {
          scrollCalls.push(args);
          return args;
        },
        focusTarget: async () => ({ ok: true }),
      },
    );

    assert.equal(result.stopReason, "stalled");
    assert.equal(result.totalLines, 5);
    assert.equal(result.snapshotCount, 3);
    assert.equal(scrollCalls.length, 2);

    const outputText = await fs.readFile(
      path.join(tempDir, "thread.txt"),
      "utf8",
    );
    assert.ok(outputText.includes("Line 1"));
    assert.ok(outputText.includes("Line 5"));

    await assert.rejects(
      () =>
        harvestScrollText(
          {
            mode: "window",
            hwnd: "0x123",
            maxSteps: 1,
            outputPath: path.join(tempDir, "thread.txt"),
          },
          {
            capture: async () => makeSnapshot("snap-9", ["Only line"]),
            focusTarget: async () => ({ ok: true }),
          },
        ),
      /already exists/i,
    );

    const overwriteResult = await harvestScrollText(
      {
        mode: "window",
        hwnd: "0x123",
        maxSteps: 1,
        outputPath: path.join(tempDir, "thread.txt"),
        overwrite: true,
      },
      {
        capture: async () => makeSnapshot("snap-10", ["Overwrite line"]),
        focusTarget: async () => ({ ok: true }),
      },
    );

    assert.equal(overwriteResult.totalLines, 1);
  } finally {
    await fs.rm(tempDir, { recursive: true, force: true });
  }
}

function testLevenshteinRatio() {
  assert.equal(levenshteinRatio("hello", "hello"), 1.0);
  assert.equal(levenshteinRatio("", ""), 1.0);
  assert.equal(levenshteinRatio("abc", ""), 0.0);
  assert.equal(levenshteinRatio("", "abc"), 0.0);

  const closeMatch = levenshteinRatio(
    "message 42 from user",
    "message 42 from usar",
  );
  assert.ok(closeMatch > 0.9, `Expected > 0.9, got ${closeMatch}`);

  const distantMatch = levenshteinRatio("hello world", "completely different");
  assert.ok(distantMatch < 0.5, `Expected < 0.5, got ${distantMatch}`);
}

function testFuzzyMerge() {
  const existing = [
    { text: "line a", normalizedText: "line a" },
    { text: "line b", normalizedText: "line b" },
    { text: "line c", normalizedText: "line c" },
  ];
  const incoming = [
    { text: "lina c", normalizedText: "lina c" },
    { text: "line d", normalizedText: "line d" },
  ];

  const exactResult = mergeHarvestLines(existing, incoming, {
    direction: "down",
  });
  assert.equal(exactResult.overlapCount, 0);
  assert.equal(exactResult.lines.length, 5);

  const fuzzyResult = mergeHarvestLines(existing, incoming, {
    direction: "down",
    fuzzyThreshold: 0.8,
  });
  assert.equal(fuzzyResult.overlapCount, 1);
  assert.equal(fuzzyResult.lines.length, 4);
}

async function testHarvestAutoFallback() {
  const tempDir = await fs.mkdtemp(
    path.join(os.tmpdir(), "peekaboo-fallback-test-"),
  );

  const emptySnapshot = {
    snapshotId: "empty-snap",
    request: { mode: "window", hwnd: "0x123" },
    result: {
      bounds: { left: 100, top: 200, width: 800, height: 600 },
      target: { hwnd: "0x123", title: "Electron App" },
      ocr: { lines: [] },
    },
  };

  const screenSnapshot = makeSnapshot("screen-snap", [
    "Recovered line 1",
    "Recovered line 2",
  ]);

  let captureCallCount = 0;

  try {
    const result = await harvestScrollText(
      {
        mode: "window",
        hwnd: "0x123",
        maxSteps: 1,
        outputPath: path.join(tempDir, "fallback.txt"),
      },
      {
        capture: async (args) => {
          captureCallCount++;
          if (captureCallCount === 1) return emptySnapshot;
          return screenSnapshot;
        },
        focusTarget: async () => ({ ok: true }),
      },
    );

    assert.ok(
      captureCallCount >= 2,
      `Expected >= 2 captures, got ${captureCallCount}`,
    );
    assert.equal(result.totalLines, 2);
    assert.ok(result.ok);
  } finally {
    await fs.rm(tempDir, { recursive: true, force: true });
  }
}

async function testSnapshotHelpers() {
  const tempHome = await fs.mkdtemp(
    path.join(os.tmpdir(), "peekaboo-win-test-"),
  );
  const previousHome = process.env.PEEKABOO_WINDOWS_HOME;

  process.env.PEEKABOO_WINDOWS_HOME = tempHome;

  try {
    const snapshotPaths = await prepareSnapshotPaths({ snapshotId: "snap-1" });
    await writeSnapshotMetadata({
      snapshotId: "snap-1",
      action: "see-ui",
      request: { mode: "window" },
      metadataPath: snapshotPaths.metadataPath,
      result: {
        target: { hwnd: "0x123" },
        elements: [
          {
            id: "e1",
            name: "Save",
            automationId: "SaveButton",
            className: "Button",
            controlType: "button",
            center: { x: 100, y: 200 },
          },
          {
            id: "e2",
            name: "Cancel",
            automationId: "CancelButton",
            className: "Button",
            controlType: "button",
            center: { x: 140, y: 200 },
          },
        ],
      },
    });

    const snapshot = await readSnapshot("snap-1");
    assert.equal(snapshot.snapshotId, "snap-1");
    assert.equal(snapshot.result.elements.length, 2);
    assert.equal(
      resolveSnapshotElement(snapshot, { elementId: "e1" }).name,
      "Save",
    );
    assert.equal(
      resolveSnapshotElement(snapshot, {
        name: "cancel",
        matchMode: "contains",
      }).id,
      "e2",
    );
    assert.equal(
      resolveSnapshotLabelTarget(snapshot, {
        name: "save",
        matchMode: "contains",
      }).source,
      "ui-element",
    );
    assert.equal(await resolveSnapshotReference("snap-1"), "snap-1");
    assert.equal(await resolveSnapshotReference("latest"), "snap-1");

    await new Promise((resolve) => setTimeout(resolve, 10));

    const rawSnapshotPaths = await prepareSnapshotPaths({
      snapshotId: "snap-2",
    });
    await writeSnapshotMetadata({
      snapshotId: "snap-2",
      action: "capture-window",
      request: { hwnd: "0x123" },
      metadataPath: rawSnapshotPaths.metadataPath,
      result: {
        target: { hwnd: "0x123" },
      },
    });

    assert.equal(await resolveSnapshotReference("latest"), "snap-2");
    assert.equal(
      await resolveSnapshotReference("latest", { requireElements: true }),
      "snap-1",
    );

    const ocrSnapshotPaths = await prepareSnapshotPaths({
      snapshotId: "snap-3",
    });
    await writeSnapshotMetadata({
      snapshotId: "snap-3",
      action: "capture-window",
      request: { hwnd: "0x999" },
      metadataPath: ocrSnapshotPaths.metadataPath,
      result: {
        target: { hwnd: "0x999" },
        ocr: {
          lines: [
            {
              text: "OCR TEST 123",
              bounds: { left: 10, top: 20, width: 140, height: 24 },
              center: { x: 80, y: 32 },
              words: [
                {
                  text: "OCR",
                  bounds: { left: 10, top: 20, width: 36, height: 24 },
                  center: { x: 28, y: 32 },
                },
              ],
            },
          ],
        },
      },
    });

    const ocrSnapshot = await readSnapshot("snap-3");
    assert.equal(
      resolveSnapshotLabelTarget(ocrSnapshot, {
        name: "OCR TEST 123",
        matchMode: "exact",
      }).source,
      "ocr-line",
    );
    assert.equal(
      resolveSnapshotLabelTarget(ocrSnapshot, {
        name: "OCR",
        matchMode: "exact",
      }).source,
      "ocr-word",
    );
  } finally {
    if (previousHome === undefined) {
      delete process.env.PEEKABOO_WINDOWS_HOME;
    } else {
      process.env.PEEKABOO_WINDOWS_HOME = previousHome;
    }

    await fs.rm(tempHome, { recursive: true, force: true });
  }
}

async function testWorkflowRunner() {
  const calls = [];
  const handlers = {
    "app.launch": async ({ command }) => {
      calls.push(["app.launch", command]);
      return { processId: 42, hwnds: ["0xABC"], command };
    },
    "window.state": async ({ hwnd, state }) => {
      calls.push(["window.state", hwnd, state]);
      return { hwnd, state };
    },
    fail: async () => {
      calls.push(["fail"]);
      throw new Error("boom");
    },
  };

  const definition = {
    steps: [
      { action: "app.launch", command: "notepad.exe", saveAs: "launch" },
      { action: "window.state", hwnd: "$launch.hwnds.0", state: "maximize" },
    ],
  };

  const result = await runWorkflowDefinition({
    definition,
    actionHandlers: handlers,
  });
  assert.equal(result.ok, true);
  assert.equal(result.context.launch.processId, 42);
  assert.deepEqual(calls, [
    ["app.launch", "notepad.exe"],
    ["window.state", "0xABC", "maximize"],
  ]);

  const continueResult = await runWorkflowDefinition({
    definition: {
      continueOnError: true,
      steps: [
        { action: "fail" },
        { action: "app.launch", command: "calc.exe", saveAs: "launch" },
      ],
    },
    actionHandlers: handlers,
  });

  assert.equal(continueResult.ok, false);
  assert.equal(continueResult.steps.length, 2);
  assert.equal(continueResult.steps[0].ok, false);
  assert.equal(continueResult.context.launch.command, "calc.exe");

  const profileCalls = [];
  await runWorkflowDefinition({
    definition: {
      profile: "human-paced",
      steps: [{ action: "window.state", hwnd: "0xABC", state: "restore" }],
    },
    actionHandlers: {
      "window.state": async ({ profile }) => {
        profileCalls.push(profile);
        return { profile };
      },
    },
  });

  assert.deepEqual(profileCalls, ["human-paced"]);

  const tempDir = await fs.mkdtemp(
    path.join(os.tmpdir(), "peekaboo-workflow-test-"),
  );

  try {
    const workflowPath = path.join(tempDir, "workflow.json");
    await fs.writeFile(
      workflowPath,
      `${JSON.stringify(definition, null, 2)}\n`,
      "utf8",
    );

    const fileResult = await runWorkflowFile({
      filePath: workflowPath,
      actionHandlers: handlers,
    });

    assert.equal(fileResult.workflowPath, workflowPath);
    assert.equal(fileResult.ok, true);
  } finally {
    await fs.rm(tempDir, { recursive: true, force: true });
  }
}

async function testRecipes() {
  const recipes = listRecipes();
  assert.ok(
    recipes.recipes.some((recipe) => recipe.id === "harvest-scroll-text"),
  );

  const plan = planGoal({
    goal: "scrape this thread",
    title: "Discord",
  });
  assert.equal(plan.ok, true);
  assert.equal(plan.recipeId, "harvest-scroll-text");

  const goalResult = await runGoal({
    goal: 'type "hello world"',
    title: "Notepad",
    actionHandlers: {
      "window.focus": async ({ title }) => ({ title }),
      type: async ({ text }) => ({ text }),
    },
  });

  assert.equal(goalResult.plan.recipeId, "type-into-app");
  assert.equal(goalResult.execution.actionResult.text, "hello world");
}

function testHotkeys() {
  assert.equal(translateHotkey("ctrl,shift,t"), "^+t");
  assert.equal(translateHotkey(["alt", "tab"]), "%{TAB}");
  assert.equal(translateHotkey("ctrl,enter"), "^{ENTER}");
}

function testAiBridge() {
  const brief = buildSnapshotAiBrief({
    snapshotId: "snap-1",
    request: {
      mode: "window",
      path: "C:\\captures\\capture.png",
      annotatedPath: "C:\\captures\\annotated.png",
    },
    result: {
      mode: "window",
      path: "C:\\captures\\capture.png",
      annotatedPath: "C:\\captures\\annotated.png",
      target: {
        hwnd: "0x123",
        title: "Calculator",
      },
      ocr: {
        available: true,
        lines: [
          {
            text: "123 + 9",
            center: { x: 740, y: 220 },
          },
        ],
      },
      elements: [
        {
          id: "e1",
          controlType: "button",
          name: "Equals",
          automationId: "equalButton",
          center: { x: 801, y: 622 },
        },
      ],
    },
  });

  assert.equal(brief.snapshotId, "snap-1");
  assert.equal(brief.targetLabel, "window | Calculator | HWND 0x123");
  assert.ok(brief.text.includes("Snapshot ID: snap-1"));
  assert.ok(brief.text.includes('e1 | button | "Equals"'));
  assert.ok(brief.text.includes("Recognized text:"));
  assert.ok(brief.text.includes('"123 + 9" | center=740,220'));
  assert.ok(brief.text.includes("Start the MCP server:"));
  assert.ok(
    brief.clickExample.includes(
      "snapshot click --snapshot snap-1 --element-id e1",
    ),
  );

  const ocrOnlyBrief = buildSnapshotAiBrief({
    snapshotId: "snap-2",
    request: {
      mode: "window",
      path: "C:\\captures\\capture.png",
    },
    result: {
      mode: "window",
      path: "C:\\captures\\capture.png",
      target: {
        hwnd: "0x456",
        title: "Notepad",
      },
      ocr: {
        available: true,
        lines: [
          {
            text: "OCR ONLY",
            center: { x: 400, y: 300 },
          },
        ],
      },
      elements: [],
    },
  });

  assert.ok(
    ocrOnlyBrief.clickExample.includes(
      'click --on "OCR ONLY" --snapshot snap-2 --exact',
    ),
  );
}

function makeSnapshot(snapshotId, lines) {
  return {
    snapshotId,
    request: {
      mode: "window",
      hwnd: "0x123",
    },
    result: {
      bounds: {
        left: 100,
        top: 200,
        width: 800,
        height: 600,
      },
      target: {
        hwnd: "0x123",
        title: "Thread viewer",
      },
      ocr: {
        lines: lines.map((text, index) => ({
          text,
          center: { x: 120, y: 220 + index * 24 },
        })),
      },
    },
  };
}

await main();
