#!/usr/bin/env node

import { spawn } from "node:child_process";
import net from "node:net";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const mcpPath = path.join(__dirname, "peekaboo-win-mcp.js");
const defaultPipeName = "peekaboo-win-codex";

export function buildPipePath(pipeName = defaultPipeName) {
  if (!/^[a-z0-9._-]+$/i.test(pipeName)) {
    throw new Error("Pipe name may contain only letters, numbers, dots, underscores, and hyphens");
  }

  return `\\\\.\\pipe\\${pipeName}`;
}

export async function startPipeSidecar({
  pipeName = defaultPipeName,
  createChild = createPeekabooChild,
} = {}) {
  const pipePath = buildPipePath(pipeName);
  let activeSocket = null;
  let activeChild = null;

  const server = net.createServer((socket) => {
    if (activeSocket) {
      socket.destroy(new Error("PeekabooWin bridge already has an active client"));
      return;
    }

    activeSocket = socket;
    activeChild = createChild();
    socket.pipe(activeChild.stdin);
    activeChild.stdout.pipe(socket);

    if (activeChild.stderr) {
      activeChild.stderr.pipe(process.stderr, { end: false });
    }

    const release = () => {
      if (activeSocket === socket) {
        activeSocket = null;
        activeChild = null;
      }
    };

    socket.once("close", () => {
      activeChild?.kill?.();
      release();
    });
    activeChild.once?.("exit", () => {
      socket.end();
      release();
    });
    activeChild.once?.("error", (error) => {
      process.stderr.write(`PeekabooWin MCP child failed: ${error.message}\n`);
      socket.destroy(error);
      release();
    });
  });

  await new Promise((resolve, reject) => {
    server.once("error", reject);
    server.listen(pipePath, resolve);
  });

  return {
    pipePath,
    close: () =>
      new Promise((resolve, reject) => {
        activeSocket?.destroy();
        activeChild?.kill?.();
        server.close((error) => (error ? reject(error) : resolve()));
      }),
  };
}

export async function runPipeRelay({
  pipeName = defaultPipeName,
  input = process.stdin,
  output = process.stdout,
} = {}) {
  const pipePath = buildPipePath(pipeName);

  await new Promise((resolve, reject) => {
    const socket = net.createConnection(pipePath);
    socket.once("connect", () => input.pipe(socket));
    socket.pipe(output, { end: false });
    socket.once("close", resolve);
    socket.once("error", reject);
    input.once("error", reject);
    output.once("error", reject);
  });
}

function createPeekabooChild() {
  const appHome =
    process.env.PEEKABOO_WINDOWS_HOME ??
    path.join(process.env.CC_ROOT ?? os.homedir(), "state", "peekaboo-win");

  return spawn(process.execPath, [mcpPath], {
    env: {
      ...process.env,
      PEEKABOO_WINDOWS_HOME: appHome,
    },
    stdio: ["pipe", "pipe", "pipe"],
    windowsHide: true,
  });
}

async function main() {
  const mode = process.argv[2];
  const pipeName = process.env.PEEKABOO_WINDOWS_PIPE ?? defaultPipeName;

  if (mode === "sidecar") {
    const sidecar = await startPipeSidecar({ pipeName });
    process.stderr.write(`PeekabooWin sidecar listening on ${sidecar.pipePath}\n`);

    const stop = async () => {
      await sidecar.close();
      process.exit(0);
    };
    process.once("SIGINT", stop);
    process.once("SIGTERM", stop);
    return;
  }

  if (mode === "relay") {
    await runPipeRelay({ pipeName });
    return;
  }

  throw new Error("Usage: peekaboo-win-bridge.js <sidecar|relay>");
}

const isDirectRun =
  process.argv[1] &&
  path.resolve(process.argv[1]).toLowerCase() ===
    path.resolve(fileURLToPath(import.meta.url)).toLowerCase();

if (isDirectRun) {
  await main();
}
