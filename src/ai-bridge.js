import path from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const appRoot = path.resolve(__dirname, "..");
const cliPath = path.join(appRoot, "bin", "peekaboo-win.js");
const mcpPath = path.join(appRoot, "bin", "peekaboo-win-mcp.js");

export function buildSnapshotAiBrief(
  snapshot,
  { maxElements = 12, maxOcrLines = 8 } = {},
) {
  if (!snapshot || typeof snapshot !== "object") {
    throw new Error("Snapshot metadata is required");
  }

  const result = snapshot.result ?? {};
  const target = result.target ?? {};
  const elements = Array.isArray(result.elements) ? result.elements : [];
  const ocr = result.ocr ?? {};
  const ocrLines = Array.isArray(ocr.lines) ? ocr.lines : [];
  const trimmedOcrLines = ocrLines.slice(0, maxOcrLines);
  const trimmedElements = elements.slice(0, maxElements);
  const targetLabel = buildTargetLabel(
    target,
    result.mode ?? snapshot.request?.mode,
  );
  const clickExampleElement =
    trimmedElements.find((element) => element.id) ?? null;
  const clickExampleOcrLine = !clickExampleElement
    ? trimmedOcrLines.find((line) => line.text)
    : null;
  const clickExample = clickExampleElement
    ? `${getCliCommand()} snapshot click --snapshot ${snapshot.snapshotId} --element-id ${clickExampleElement.id}`
    : clickExampleOcrLine
      ? `${getCliCommand()} click --on ${quoteCliArg(clickExampleOcrLine.text)} --snapshot ${snapshot.snapshotId} --exact`
      : null;

  const lines = [
    "You can control this Windows desktop with PeekabooWin.",
    "",
    "Current desktop context:",
    `- Snapshot ID: ${snapshot.snapshotId}`,
    `- Target: ${targetLabel}`,
    `- Capture image: ${result.path ?? snapshot.request?.path ?? "Not available"}`,
    `- Annotated image: ${result.annotatedPath ?? snapshot.request?.annotatedPath ?? "Not available"}`,
    `- Visible indexed elements: ${elements.length}`,
    `- OCR text lines: ${ocrLines.length}`,
    "",
    "Indexed controls:",
  ];

  if (trimmedElements.length === 0) {
    lines.push(
      "- No indexed UI Automation elements were found in this snapshot.",
    );
  } else {
    for (const element of trimmedElements) {
      lines.push(`- ${formatElementLine(element)}`);
    }
  }

  if (elements.length > trimmedElements.length) {
    lines.push(
      `- ...and ${elements.length - trimmedElements.length} more elements.`,
    );
  }

  lines.push("");
  lines.push("Recognized text:");

  if (!ocr.available) {
    lines.push(`- OCR not available${ocr.error ? `: ${ocr.error}` : "."}`);
  } else if (trimmedOcrLines.length === 0) {
    lines.push("- No OCR text was recognized in this capture.");
  } else {
    for (const line of trimmedOcrLines) {
      lines.push(`- ${formatOcrLine(line)}`);
    }
  }

  if (ocrLines.length > trimmedOcrLines.length) {
    lines.push(
      `- ...and ${ocrLines.length - trimmedOcrLines.length} more OCR lines.`,
    );
  }

  lines.push("");
  lines.push("How to connect:");
  lines.push(`- Start the MCP server: ${getMcpCommand()}`);
  lines.push(`- Run a one-off CLI command: ${getCliCommand()} windows list`);

  if (clickExample) {
    lines.push(`- Click example: ${clickExample}`);
  }

  lines.push("- Prefer snapshot element IDs when they are available.");

  return {
    snapshotId: snapshot.snapshotId,
    targetLabel,
    cliCommand: getCliCommand(),
    mcpCommand: getMcpCommand(),
    clickExample,
    text: lines.join("\n"),
  };
}

export function getCliCommand() {
  return `node "${cliPath}"`;
}

export function getMcpCommand() {
  return `node "${mcpPath}"`;
}

function buildTargetLabel(target, mode) {
  const pieces = [];

  if (mode) {
    pieces.push(mode);
  }

  if (target.title) {
    pieces.push(target.title);
  }

  if (target.hwnd) {
    pieces.push(`HWND ${target.hwnd}`);
  }

  return pieces.length > 0
    ? pieces.join(" | ")
    : "Desktop target not specified";
}

function formatElementLine(element) {
  const parts = [];

  parts.push(element.id ?? "unknown");
  parts.push(element.controlType ?? "control");
  parts.push(element.name ? `"${element.name}"` : "(unnamed)");

  if (element.automationId) {
    parts.push(`automationId=${element.automationId}`);
  }

  if (
    element.center &&
    Number.isFinite(element.center.x) &&
    Number.isFinite(element.center.y)
  ) {
    parts.push(`center=${element.center.x},${element.center.y}`);
  }

  return parts.join(" | ");
}

function formatOcrLine(line) {
  const parts = [line.text ? `"${line.text}"` : "(blank)"];

  if (
    line.center &&
    Number.isFinite(line.center.x) &&
    Number.isFinite(line.center.y)
  ) {
    parts.push(`center=${line.center.x},${line.center.y}`);
  }

  return parts.join(" | ");
}

function quoteCliArg(value) {
  return JSON.stringify(String(value ?? ""));
}
