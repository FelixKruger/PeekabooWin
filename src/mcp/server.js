import readline from "node:readline";

import {
  captureScreen,
  captureWindow,
  clickDialogButton,
  clickElementByLabel,
  clickUiElement,
  clickSnapshotElement,
  click,
  cleanStoredSnapshots,
  drag,
  dragSnapshotElement,
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
} from "../backend/tools.js";
import { harvestScrollText } from "../harvest.js";
import { listRecipes, planGoal, runGoal, runRecipe } from "../recipes.js";
import { runWorkflowFile } from "../workflow/runner.js";

const serverInfo = {
  name: "peekaboo-windows",
  version: "0.1.0",
};

const toolDefinitions = [
  {
    name: "windows_list",
    title: "List Windows",
    description:
      "List visible top-level windows with title, bounds, and process information.",
    inputSchema: {
      type: "object",
      additionalProperties: false,
      properties: {},
    },
  },
  {
    name: "screens_list",
    title: "List Screens",
    description: "List connected displays with their bounds and working areas.",
    inputSchema: {
      type: "object",
      additionalProperties: false,
      properties: {},
    },
  },
  {
    name: "window_focus",
    title: "Focus Window",
    description:
      "Bring a window to the foreground using either its HWND or a title substring.",
    inputSchema: {
      type: "object",
      additionalProperties: false,
      properties: {
        hwnd: {
          type: "string",
          description: "Window handle in decimal or hex string form.",
        },
        title: {
          type: "string",
          description: "Case-insensitive title substring.",
        },
      },
    },
  },
  {
    name: "window_move",
    title: "Move Window",
    description:
      "Move a top-level window to absolute screen coordinates while keeping its current size.",
    inputSchema: {
      type: "object",
      additionalProperties: false,
      required: ["x", "y"],
      properties: {
        hwnd: { type: "string" },
        title: { type: "string" },
        x: { type: "number" },
        y: { type: "number" },
      },
    },
  },
  {
    name: "window_resize",
    title: "Resize Window",
    description:
      "Resize a top-level window while keeping its current position.",
    inputSchema: {
      type: "object",
      additionalProperties: false,
      required: ["width", "height"],
      properties: {
        hwnd: { type: "string" },
        title: { type: "string" },
        width: { type: "number" },
        height: { type: "number" },
      },
    },
  },
  {
    name: "window_set_bounds",
    title: "Set Window Bounds",
    description: "Set the full bounds of a top-level window.",
    inputSchema: {
      type: "object",
      additionalProperties: false,
      required: ["x", "y", "width", "height"],
      properties: {
        hwnd: { type: "string" },
        title: { type: "string" },
        x: { type: "number" },
        y: { type: "number" },
        width: { type: "number" },
        height: { type: "number" },
      },
    },
  },
  {
    name: "window_set_state",
    title: "Set Window State",
    description: "Restore, maximize, or minimize a top-level window.",
    inputSchema: {
      type: "object",
      additionalProperties: false,
      required: ["state"],
      properties: {
        hwnd: { type: "string" },
        title: { type: "string" },
        state: { type: "string", enum: ["restore", "maximize", "minimize"] },
      },
    },
  },
  {
    name: "window_wait",
    title: "Wait For Window",
    description: "Poll until a window matching the selector appears.",
    inputSchema: {
      type: "object",
      additionalProperties: false,
      properties: {
        hwnd: { type: "string" },
        title: { type: "string" },
        processId: { type: "number" },
        processName: { type: "string" },
        matchMode: { type: "string", enum: ["contains", "exact"] },
        timeoutMs: { type: "number" },
        pollMs: { type: "number" },
      },
    },
  },
  {
    name: "app_list",
    title: "List Apps",
    description:
      "List visible desktop apps grouped by process with their top-level window titles.",
    inputSchema: {
      type: "object",
      additionalProperties: false,
      properties: {},
    },
  },
  {
    name: "app_switch",
    title: "Switch App",
    description:
      "Bring an app to the foreground using its process ID, process name, or window title.",
    inputSchema: {
      type: "object",
      additionalProperties: false,
      properties: {
        processId: { type: "number" },
        name: { type: "string" },
        title: { type: "string" },
        matchMode: { type: "string", enum: ["contains", "exact"] },
      },
    },
  },
  {
    name: "app_quit",
    title: "Quit App",
    description:
      "Send a graceful close request to an app using its process ID, process name, or window title.",
    inputSchema: {
      type: "object",
      additionalProperties: false,
      properties: {
        processId: { type: "number" },
        name: { type: "string" },
        title: { type: "string" },
        matchMode: { type: "string", enum: ["contains", "exact"] },
      },
    },
  },
  {
    name: "dialog_list",
    title: "List Dialogs",
    description:
      "List visible standard dialog windows, optionally filtered by HWND, title, or process.",
    inputSchema: {
      type: "object",
      additionalProperties: false,
      properties: {
        hwnd: { type: "string" },
        title: { type: "string" },
        processId: { type: "number" },
        processName: { type: "string" },
        matchMode: { type: "string", enum: ["contains", "exact"] },
      },
    },
  },
  {
    name: "dialog_click",
    title: "Click Dialog Button",
    description: "Find a visible standard dialog and click a button inside it.",
    inputSchema: {
      type: "object",
      additionalProperties: false,
      required: ["button"],
      properties: {
        hwnd: { type: "string" },
        title: { type: "string" },
        processId: { type: "number" },
        processName: { type: "string" },
        button: { type: "string" },
        matchMode: { type: "string", enum: ["contains", "exact"] },
        profile: { type: "string", enum: ["default", "human-paced"] },
      },
    },
  },
  {
    name: "menu_list",
    title: "List Menu Items",
    description:
      "List visible menu items for a window, optionally opening a menu path first.",
    inputSchema: {
      type: "object",
      additionalProperties: false,
      properties: {
        hwnd: { type: "string" },
        title: { type: "string" },
        path: {
          type: "string",
          description:
            "Optional menu path to open first, for example File>Recent.",
        },
        maxResults: { type: "number" },
        matchMode: { type: "string", enum: ["contains", "exact"] },
        profile: { type: "string", enum: ["default", "human-paced"] },
      },
    },
  },
  {
    name: "menu_click",
    title: "Click Menu Path",
    description:
      "Open and activate a window menu path such as File>Save or Edit>Replace.",
    inputSchema: {
      type: "object",
      additionalProperties: false,
      required: ["path"],
      properties: {
        hwnd: { type: "string" },
        title: { type: "string" },
        path: { type: "string" },
        matchMode: { type: "string", enum: ["contains", "exact"] },
        profile: { type: "string", enum: ["default", "human-paced"] },
      },
    },
  },
  {
    name: "screen_capture",
    title: "Capture Screen",
    description:
      "Capture the full virtual desktop to a PNG file and return the saved path.",
    inputSchema: {
      type: "object",
      additionalProperties: false,
      properties: {
        screenIndex: { type: "number" },
        fileName: { type: "string" },
        snapshotId: { type: "string" },
      },
    },
  },
  {
    name: "window_capture",
    title: "Capture Window",
    description:
      "Capture a single top-level window to a PNG file using its HWND or title.",
    inputSchema: {
      type: "object",
      additionalProperties: false,
      properties: {
        hwnd: { type: "string" },
        title: { type: "string" },
        fileName: { type: "string" },
        snapshotId: { type: "string" },
      },
    },
  },
  {
    name: "ui_find",
    title: "Find UI Elements",
    description:
      "Search Windows UI Automation elements by accessible name and optional control type.",
    inputSchema: {
      type: "object",
      additionalProperties: false,
      properties: {
        hwnd: { type: "string" },
        title: { type: "string" },
        name: { type: "string" },
        automationId: { type: "string" },
        className: { type: "string" },
        controlType: { type: "string" },
        maxResults: { type: "number" },
        matchMode: { type: "string", enum: ["contains", "exact"] },
        profile: { type: "string", enum: ["default", "human-paced"] },
      },
    },
  },
  {
    name: "ui_wait",
    title: "Wait For UI Element",
    description:
      "Poll until a UI Automation element matching the selector appears.",
    inputSchema: {
      type: "object",
      additionalProperties: false,
      properties: {
        hwnd: { type: "string" },
        title: { type: "string" },
        name: { type: "string" },
        automationId: { type: "string" },
        className: { type: "string" },
        controlType: { type: "string" },
        maxResults: { type: "number" },
        matchMode: { type: "string", enum: ["contains", "exact"] },
        timeoutMs: { type: "number" },
        pollMs: { type: "number" },
      },
    },
  },
  {
    name: "ui_click",
    title: "Click UI Element",
    description:
      "Find and activate a Windows UI Automation element by name inside a window or globally.",
    inputSchema: {
      type: "object",
      additionalProperties: false,
      required: ["name"],
      properties: {
        hwnd: { type: "string" },
        title: { type: "string" },
        name: { type: "string" },
        automationId: { type: "string" },
        className: { type: "string" },
        controlType: { type: "string" },
        maxResults: { type: "number" },
        matchMode: { type: "string", enum: ["contains", "exact"] },
      },
    },
  },
  {
    name: "ui_snapshot",
    title: "See UI Snapshot",
    description:
      "Capture a window or screen, index visible UI Automation elements, and save an annotated snapshot.",
    inputSchema: {
      type: "object",
      additionalProperties: false,
      properties: {
        mode: { type: "string", enum: ["screen", "window"] },
        screenIndex: { type: "number" },
        hwnd: { type: "string" },
        title: { type: "string" },
        fileName: { type: "string" },
        annotatedFileName: { type: "string" },
        snapshotId: { type: "string" },
        maxResults: { type: "number" },
        matchMode: { type: "string", enum: ["contains", "exact"] },
      },
    },
  },
  {
    name: "snapshot_list",
    title: "List Snapshots",
    description: "List saved snapshots and their metadata summaries.",
    inputSchema: {
      type: "object",
      additionalProperties: false,
      properties: {
        limit: { type: "number" },
      },
    },
  },
  {
    name: "snapshot_get",
    title: "Get Snapshot",
    description: "Read a saved snapshot metadata file by snapshot ID.",
    inputSchema: {
      type: "object",
      additionalProperties: false,
      required: ["snapshotId"],
      properties: {
        snapshotId: { type: "string" },
      },
    },
  },
  {
    name: "snapshot_click",
    title: "Click Snapshot Element",
    description:
      "Click an indexed element from a previously captured UI snapshot by element ID or unique name.",
    inputSchema: {
      type: "object",
      additionalProperties: false,
      required: ["snapshotId"],
      properties: {
        snapshotId: { type: "string" },
        elementId: { type: "string" },
        name: { type: "string" },
        matchMode: { type: "string", enum: ["contains", "exact"] },
        profile: { type: "string", enum: ["default", "human-paced"] },
      },
    },
  },
  {
    name: "element_click",
    title: "Click Element By Label",
    description:
      "Click a visible labeled element using either a stored snapshot or a fresh UI snapshot captured from a screen or window selector.",
    inputSchema: {
      type: "object",
      additionalProperties: false,
      required: ["label"],
      properties: {
        label: { type: "string" },
        snapshotId: { type: "string" },
        mode: { type: "string", enum: ["screen", "window"] },
        screenIndex: { type: "number" },
        hwnd: { type: "string" },
        title: { type: "string" },
        fileName: { type: "string" },
        annotatedFileName: { type: "string" },
        matchMode: { type: "string", enum: ["contains", "exact"] },
        profile: { type: "string", enum: ["default", "human-paced"] },
      },
    },
  },
  {
    name: "snapshot_clean",
    title: "Clean Snapshots",
    description:
      "Delete one snapshot, all snapshots, or snapshots older than a threshold.",
    inputSchema: {
      type: "object",
      additionalProperties: false,
      properties: {
        snapshotId: { type: "string" },
        all: { type: "boolean" },
        olderThanHours: { type: "number" },
      },
    },
  },
  {
    name: "mouse_drag",
    title: "Drag Mouse",
    description:
      "Drag the mouse from one absolute screen coordinate to another.",
    inputSchema: {
      type: "object",
      additionalProperties: false,
      required: ["fromX", "fromY", "toX", "toY"],
      properties: {
        fromX: { type: "number" },
        fromY: { type: "number" },
        toX: { type: "number" },
        toY: { type: "number" },
        button: { type: "string", enum: ["left", "right"] },
        steps: { type: "number" },
        durationMs: { type: "number" },
        profile: { type: "string", enum: ["default", "human-paced"] },
      },
    },
  },
  {
    name: "snapshot_drag",
    title: "Drag Snapshot Element",
    description:
      "Drag from one indexed snapshot element to another indexed snapshot element.",
    inputSchema: {
      type: "object",
      additionalProperties: false,
      required: ["snapshotId"],
      properties: {
        snapshotId: { type: "string" },
        fromElementId: { type: "string" },
        toElementId: { type: "string" },
        fromName: { type: "string" },
        toName: { type: "string" },
        matchMode: { type: "string", enum: ["contains", "exact"] },
        button: { type: "string", enum: ["left", "right"] },
        steps: { type: "number" },
        durationMs: { type: "number" },
        profile: { type: "string", enum: ["default", "human-paced"] },
      },
    },
  },
  {
    name: "workflow_run",
    title: "Run Workflow",
    description:
      "Run a local Peekaboo Windows workflow JSON file that chains multiple automation steps.",
    inputSchema: {
      type: "object",
      additionalProperties: false,
      required: ["filePath"],
      properties: {
        filePath: { type: "string" },
        continueOnError: { type: "boolean" },
        profile: { type: "string", enum: ["default", "human-paced"] },
      },
    },
  },
  {
    name: "harvest_scroll_text",
    title: "Harvest Scrolling Text",
    description:
      "Capture, scroll, and collect OCR text across multiple snapshots, saving the full transcript to a local file.",
    inputSchema: {
      type: "object",
      additionalProperties: false,
      properties: {
        snapshotId: { type: "string" },
        mode: { type: "string", enum: ["screen", "window"] },
        hwnd: { type: "string" },
        title: { type: "string" },
        processId: { type: "number" },
        processName: { type: "string" },
        screenIndex: { type: "number" },
        scrollLabel: {
          type: "string",
          description: "Optional visible label or OCR text to scroll over.",
        },
        x: { type: "number" },
        y: { type: "number" },
        direction: { type: "string", enum: ["up", "down"] },
        ticks: { type: "number" },
        maxSteps: { type: "number" },
        stopAfterStalledSteps: { type: "number" },
        overlapWindow: { type: "number" },
        fuzzyThreshold: {
          type: "number",
          description:
            "Levenshtein similarity ratio (0.0-1.0) for fuzzy overlap matching. Useful when OCR text varies slightly between passes, such as screen-mode captures. Leave unset for exact matching.",
        },
        pauseAfterScrollMs: { type: "number" },
        outputPath: {
          type: "string",
          description:
            "Optional text file path for the full harvested transcript.",
        },
        overwrite: {
          type: "boolean",
          description:
            "Allow overwrite when the target output files already exist.",
        },
        matchMode: { type: "string", enum: ["contains", "exact"] },
        profile: { type: "string", enum: ["default", "human-paced"] },
      },
    },
  },
  {
    name: "recipe_list",
    title: "List Recipes",
    description:
      "List high-level PeekabooWin recipes that wrap common desktop tasks.",
    inputSchema: {
      type: "object",
      additionalProperties: false,
      properties: {},
    },
  },
  {
    name: "recipe_run",
    title: "Run Recipe",
    description:
      "Run a high-level PeekabooWin recipe such as inspect, read text, handoff to AI, or harvest scrolling text.",
    inputSchema: {
      type: "object",
      additionalProperties: false,
      required: ["recipeId"],
      properties: {
        recipeId: { type: "string" },
        snapshotId: { type: "string" },
        mode: { type: "string", enum: ["screen", "window"] },
        hwnd: { type: "string" },
        title: { type: "string" },
        processId: { type: "number" },
        processName: { type: "string" },
        screenIndex: { type: "number" },
        label: { type: "string" },
        text: { type: "string" },
        clear: { type: "boolean" },
        command: { type: "string" },
        args: {
          type: "array",
          items: { type: "string" },
        },
        x: { type: "number" },
        y: { type: "number" },
        direction: { type: "string", enum: ["up", "down"] },
        ticks: { type: "number" },
        maxSteps: { type: "number" },
        stopAfterStalledSteps: { type: "number" },
        overlapWindow: { type: "number" },
        pauseAfterScrollMs: { type: "number" },
        outputPath: { type: "string" },
        overwrite: { type: "boolean" },
        timeoutMs: { type: "number" },
        continueOnError: { type: "boolean" },
        matchMode: { type: "string", enum: ["contains", "exact"] },
        profile: { type: "string", enum: ["default", "human-paced"] },
      },
    },
  },
  {
    name: "goal_plan",
    title: "Plan Goal",
    description:
      "Turn a plain-language desktop task into the best matching PeekabooWin recipe and required arguments.",
    inputSchema: {
      type: "object",
      additionalProperties: false,
      required: ["goal"],
      properties: {
        goal: { type: "string" },
        snapshotId: { type: "string" },
        mode: { type: "string", enum: ["screen", "window"] },
        hwnd: { type: "string" },
        title: { type: "string" },
        processId: { type: "number" },
        processName: { type: "string" },
        screenIndex: { type: "number" },
      },
    },
  },
  {
    name: "goal_run",
    title: "Run Goal",
    description:
      "Execute a plain-language desktop goal through the best matching PeekabooWin recipe.",
    inputSchema: {
      type: "object",
      additionalProperties: false,
      required: ["goal"],
      properties: {
        goal: { type: "string" },
        snapshotId: { type: "string" },
        mode: { type: "string", enum: ["screen", "window"] },
        hwnd: { type: "string" },
        title: { type: "string" },
        processId: { type: "number" },
        processName: { type: "string" },
        screenIndex: { type: "number" },
        label: { type: "string" },
        x: { type: "number" },
        y: { type: "number" },
        direction: { type: "string", enum: ["up", "down"] },
        ticks: { type: "number" },
        maxSteps: { type: "number" },
        stopAfterStalledSteps: { type: "number" },
        overlapWindow: { type: "number" },
        pauseAfterScrollMs: { type: "number" },
        outputPath: { type: "string" },
        overwrite: { type: "boolean" },
        continueOnError: { type: "boolean" },
        profile: { type: "string", enum: ["default", "human-paced"] },
      },
    },
  },
  {
    name: "hotkey_press",
    title: "Press Hotkey",
    description: "Press a human-friendly hotkey combo like ctrl,shift,t.",
    inputSchema: {
      type: "object",
      additionalProperties: false,
      required: ["keys"],
      properties: {
        keys: {
          oneOf: [
            { type: "string" },
            { type: "array", items: { type: "string" } },
          ],
        },
        repeat: { type: "number" },
        delayMs: { type: "number" },
        profile: { type: "string", enum: ["default", "human-paced"] },
      },
    },
  },
  {
    name: "scroll",
    title: "Scroll",
    description:
      "Scroll the mouse wheel at the current cursor or at explicit screen coordinates.",
    inputSchema: {
      type: "object",
      additionalProperties: false,
      properties: {
        x: { type: "number" },
        y: { type: "number" },
        direction: { type: "string", enum: ["up", "down"] },
        ticks: { type: "number" },
        profile: { type: "string", enum: ["default", "human-paced"] },
      },
    },
  },
  {
    name: "snapshot_scroll",
    title: "Scroll Snapshot Element",
    description:
      "Scroll over an indexed element from a previously captured UI snapshot.",
    inputSchema: {
      type: "object",
      additionalProperties: false,
      required: ["snapshotId"],
      properties: {
        snapshotId: { type: "string" },
        elementId: { type: "string" },
        name: { type: "string" },
        matchMode: { type: "string", enum: ["contains", "exact"] },
        direction: { type: "string", enum: ["up", "down"] },
        ticks: { type: "number" },
        profile: { type: "string", enum: ["default", "human-paced"] },
      },
    },
  },
  {
    name: "element_scroll",
    title: "Scroll Element By Label",
    description:
      "Scroll over a visible labeled element using either a stored snapshot or a fresh UI snapshot captured from a screen or window selector.",
    inputSchema: {
      type: "object",
      additionalProperties: false,
      required: ["label"],
      properties: {
        label: { type: "string" },
        snapshotId: { type: "string" },
        mode: { type: "string", enum: ["screen", "window"] },
        screenIndex: { type: "number" },
        hwnd: { type: "string" },
        title: { type: "string" },
        fileName: { type: "string" },
        annotatedFileName: { type: "string" },
        direction: { type: "string", enum: ["up", "down"] },
        ticks: { type: "number" },
        matchMode: { type: "string", enum: ["contains", "exact"] },
        profile: { type: "string", enum: ["default", "human-paced"] },
      },
    },
  },
  {
    name: "mouse_move",
    title: "Move Mouse",
    description: "Move the mouse cursor to absolute screen coordinates.",
    inputSchema: {
      type: "object",
      additionalProperties: false,
      required: ["x", "y"],
      properties: {
        x: { type: "number" },
        y: { type: "number" },
        profile: { type: "string", enum: ["default", "human-paced"] },
      },
    },
  },
  {
    name: "mouse_click",
    title: "Click Mouse",
    description:
      "Move the mouse to coordinates and click with the requested button.",
    inputSchema: {
      type: "object",
      additionalProperties: false,
      required: ["x", "y"],
      properties: {
        x: { type: "number" },
        y: { type: "number" },
        button: { type: "string", enum: ["left", "right"] },
        double: { type: "boolean" },
        profile: { type: "string", enum: ["default", "human-paced"] },
      },
    },
  },
  {
    name: "type_text",
    title: "Type Text",
    description: "Type Unicode text into the currently focused window.",
    inputSchema: {
      type: "object",
      additionalProperties: false,
      required: ["text"],
      properties: {
        text: { type: "string" },
        clear: { type: "boolean" },
        delayMs: { type: "number" },
        profile: { type: "string", enum: ["default", "human-paced"] },
      },
    },
  },
  {
    name: "press_keys",
    title: "Press Keys",
    description:
      "Send a key chord like ^l or %{F4} to the currently focused window.",
    inputSchema: {
      type: "object",
      additionalProperties: false,
      required: ["keys"],
      properties: {
        keys: { type: "string" },
      },
    },
  },
  {
    name: "app_launch",
    title: "Launch App",
    description: "Launch a Windows application or shell command.",
    inputSchema: {
      type: "object",
      additionalProperties: false,
      required: ["command"],
      properties: {
        command: { type: "string" },
        args: {
          type: "array",
          items: { type: "string" },
        },
      },
    },
  },
];

const toolHandlers = {
  windows_list: async () => await listWindows(),
  screens_list: async () => await listScreens(),
  window_focus: async (args) => await focusWindow(args ?? {}),
  window_move: async (args) => await moveWindow(args ?? {}),
  window_resize: async (args) => await resizeWindow(args ?? {}),
  window_set_bounds: async (args) => await setWindowBounds(args ?? {}),
  window_set_state: async (args) => await setWindowState(args ?? {}),
  window_wait: async (args) => await waitForWindow(args ?? {}),
  app_list: async () => await listApps(),
  app_switch: async (args) => await switchApp(args ?? {}),
  app_quit: async (args) => await quitApp(args ?? {}),
  dialog_list: async (args) => await listDialogs(args ?? {}),
  dialog_click: async (args) => await clickDialogButton(args ?? {}),
  menu_list: async (args) => await listMenuItems(args ?? {}),
  menu_click: async (args) => await clickMenuPath(args ?? {}),
  screen_capture: async (args) => await captureScreen(args ?? {}),
  window_capture: async (args) => await captureWindow(args ?? {}),
  ui_find: async (args) => await findUiElements(args ?? {}),
  ui_wait: async (args) => await waitForUiElement(args ?? {}),
  ui_click: async (args) => await clickUiElement(args ?? {}),
  ui_snapshot: async (args) => await seeUi(args ?? {}),
  snapshot_list: async (args) => await listStoredSnapshots(args ?? {}),
  snapshot_get: async (args) => await getStoredSnapshot(args ?? {}),
  snapshot_click: async (args) => await clickSnapshotElement(args ?? {}),
  element_click: async (args) => await clickElementByLabel(args ?? {}),
  snapshot_clean: async (args) => await cleanStoredSnapshots(args ?? {}),
  mouse_drag: async (args) => await drag(args ?? {}),
  snapshot_drag: async (args) => await dragSnapshotElement(args ?? {}),
  workflow_run: async (args) =>
    await runWorkflowFile({
      filePath: args?.filePath,
      continueOnError: args?.continueOnError,
      profile: args?.profile,
    }),
  harvest_scroll_text: async (args) => await harvestScrollText(args ?? {}),
  recipe_list: async () => listRecipes(),
  recipe_run: async (args) =>
    await runRecipe({
      recipeId: args?.recipeId,
      ...omitUndefined({
        snapshotId: args?.snapshotId,
        mode: args?.mode,
        hwnd: args?.hwnd,
        title: args?.title,
        processId: args?.processId,
        processName: args?.processName,
        screenIndex: args?.screenIndex,
        label: args?.label,
        scrollLabel: args?.label,
        text: args?.text,
        clear: args?.clear,
        command: args?.command,
        args: args?.args,
        x: args?.x,
        y: args?.y,
        direction: args?.direction,
        ticks: args?.ticks,
        maxSteps: args?.maxSteps,
        stopAfterStalledSteps: args?.stopAfterStalledSteps,
        overlapWindow: args?.overlapWindow,
        pauseAfterScrollMs: args?.pauseAfterScrollMs,
        outputPath: args?.outputPath,
        overwrite: args?.overwrite,
        timeoutMs: args?.timeoutMs,
        continueOnError: args?.continueOnError,
        matchMode: args?.matchMode,
        profile: args?.profile,
      }),
    }),
  goal_plan: async (args) =>
    planGoal({
      goal: args?.goal,
      snapshotId: args?.snapshotId,
      mode: args?.mode,
      hwnd: args?.hwnd,
      title: args?.title,
      processId: args?.processId,
      processName: args?.processName,
      screenIndex: args?.screenIndex,
    }),
  goal_run: async (args) =>
    await runGoal({
      goal: args?.goal,
      snapshotId: args?.snapshotId,
      mode: args?.mode,
      hwnd: args?.hwnd,
      title: args?.title,
      processId: args?.processId,
      processName: args?.processName,
      screenIndex: args?.screenIndex,
      label: args?.label,
      scrollLabel: args?.label,
      x: args?.x,
      y: args?.y,
      direction: args?.direction,
      ticks: args?.ticks,
      maxSteps: args?.maxSteps,
      stopAfterStalledSteps: args?.stopAfterStalledSteps,
      overlapWindow: args?.overlapWindow,
      pauseAfterScrollMs: args?.pauseAfterScrollMs,
      outputPath: args?.outputPath,
      overwrite: args?.overwrite,
      continueOnError: args?.continueOnError,
      profile: args?.profile,
    }),
  hotkey_press: async (args) => await hotkey(args ?? {}),
  scroll: async (args) => await scroll(args ?? {}),
  snapshot_scroll: async (args) => await scrollSnapshotElement(args ?? {}),
  element_scroll: async (args) => await scrollElementByLabel(args ?? {}),
  mouse_move: async (args) => await moveMouse(args ?? {}),
  mouse_click: async (args) => await click(args ?? {}),
  type_text: async (args) => await typeText(args ?? {}),
  press_keys: async (args) => await pressKeys(args ?? {}),
  app_launch: async (args) => await launchApp(args ?? {}),
};

export async function runMcpServer({
  input = process.stdin,
  output = process.stdout,
  error = process.stderr,
} = {}) {
  const rl = readline.createInterface({ input, crlfDelay: Infinity });

  for await (const line of rl) {
    const trimmed = line.trim();

    if (!trimmed) {
      continue;
    }

    let message;

    try {
      message = JSON.parse(trimmed);
    } catch {
      write(output, makeError(null, -32700, "Parse error"));
      continue;
    }

    try {
      await handleMessage(message, { output });
    } catch (handlerError) {
      const messageText =
        handlerError instanceof Error
          ? handlerError.message
          : String(handlerError);
      error.write(`${messageText}\n`);

      if (Object.prototype.hasOwnProperty.call(message, "id")) {
        write(output, makeError(message.id, -32603, messageText));
      }
    }
  }
}

async function handleMessage(message, { output }) {
  if (
    !message ||
    message.jsonrpc !== "2.0" ||
    typeof message.method !== "string"
  ) {
    write(output, makeError(message?.id ?? null, -32600, "Invalid Request"));
    return;
  }

  const { id, method, params } = message;

  if (method === "initialize") {
    write(output, {
      jsonrpc: "2.0",
      id,
      result: {
        protocolVersion: "2024-11-05",
        capabilities: {
          tools: {
            listChanged: false,
          },
        },
        serverInfo,
      },
    });
    return;
  }

  if (method === "notifications/initialized") {
    return;
  }

  if (method === "ping") {
    write(output, {
      jsonrpc: "2.0",
      id,
      result: {},
    });
    return;
  }

  if (method === "tools/list") {
    write(output, {
      jsonrpc: "2.0",
      id,
      result: {
        tools: toolDefinitions,
      },
    });
    return;
  }

  if (method === "tools/call") {
    const toolName = params?.name;
    const handler = toolHandlers[toolName];

    if (!handler) {
      write(output, makeError(id, -32601, `Unknown tool: ${toolName}`));
      return;
    }

    try {
      const result = await handler(params?.arguments ?? {});
      write(output, {
        jsonrpc: "2.0",
        id,
        result: {
          content: [
            {
              type: "text",
              text: JSON.stringify(result, null, 2),
            },
          ],
          structuredContent: result,
          isError: false,
        },
      });
    } catch (toolError) {
      const messageText =
        toolError instanceof Error ? toolError.message : String(toolError);
      write(output, {
        jsonrpc: "2.0",
        id,
        result: {
          content: [
            {
              type: "text",
              text: messageText,
            },
          ],
          isError: true,
        },
      });
    }

    return;
  }

  write(output, makeError(id ?? null, -32601, `Method not found: ${method}`));
}

function makeError(id, code, message) {
  return {
    jsonrpc: "2.0",
    id,
    error: {
      code,
      message,
    },
  };
}

function write(output, payload) {
  output.write(`${JSON.stringify(payload)}\n`);
}

function omitUndefined(value) {
  return Object.fromEntries(
    Object.entries(value).filter(([, entry]) => entry !== undefined),
  );
}
