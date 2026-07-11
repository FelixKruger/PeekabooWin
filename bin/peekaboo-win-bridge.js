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
  let activeConnection = null;

  const server = net.createServer({ allowHalfOpen: true }, (socket) => {
    if (activeConnection) {
      socket.destroy();
      return;
    }

    const child = createChild();
    const connection = { socket, child };
    activeConnection = connection;
    socket.pipe(child.stdin);
    child.stdout.pipe(socket);

    child.stderr?.resume?.();

    socket.once("error", () => {});
    socket.once("close", () => {
      child.kill?.();
      if (activeConnection === connection) {
        activeConnection = null;
      }
    });
    child.once?.("exit", () => {
      socket.end();
    });
    child.once?.("error", (error) => {
      process.stderr.write(`PeekabooWin MCP child failed: ${error.message}\n`);
      socket.destroy(error);
    });
  });

  await new Promise((resolve, reject) => {
    server.once("error", reject);
    server.listen(
      {
        path: pipePath,
        readableAll: false,
        writableAll: false,
      },
      resolve,
    );
  });

  return {
    pipePath,
    close: () =>
      new Promise((resolve, reject) => {
        activeConnection?.socket.destroy();
        activeConnection?.child.kill?.();
        activeConnection = null;
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
    const socket = net.createConnection({
      path: pipePath,
      allowHalfOpen: true,
    });
    let settled = false;
    const cleanup = () => {
      input.off("error", onError);
      output.off("error", onError);
    };
    const finish = (error) => {
      if (settled) return;
      settled = true;
      cleanup();
      if (error) reject(error);
      else resolve();
    };
    const onError = (error) => finish(error);

    socket.once("connect", () => input.pipe(socket));
    socket.pipe(output, { end: false });
    socket.once("close", () => finish());
    socket.once("error", onError);
    input.once("error", onError);
    output.once("error", onError);
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
