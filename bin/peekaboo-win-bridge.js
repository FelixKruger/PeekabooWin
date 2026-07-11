#!/usr/bin/env node

import { spawn } from "node:child_process";
import { appendFileSync, mkdirSync } from "node:fs";
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
  appHome,
  createChild,
  log = () => {},
} = {}) {
  const pipePath = buildPipePath(pipeName);
  let activeConnection = null;
  const spawnChild = createChild ?? (() => createPeekabooChild(appHome));

  const server = net.createServer({ allowHalfOpen: true }, (socket) => {
    if (activeConnection) {
      log("Rejected client while another MCP client was active");
      socket.destroy();
      return;
    }

    const child = spawnChild();
    const connection = { socket, child };
    activeConnection = connection;
    log(`Accepted MCP client; child PID ${child.pid ?? "test"}`);
    socket.pipe(child.stdin);
    child.stdout.pipe(socket);

    child.stderr?.resume?.();

    socket.once("error", (error) => {
      log(`Client socket error: ${error.message}`);
    });
    socket.once("close", () => {
      child.kill?.();
      if (activeConnection === connection) {
        activeConnection = null;
      }
      log("MCP client disconnected");
    });
    child.once?.("exit", (code, signal) => {
      log(`MCP child exited code=${code ?? "null"} signal=${signal ?? "null"}`);
      socket.end();
    });
    child.once?.("error", (error) => {
      log(`MCP child failed: ${error.message}`);
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
    const socket = net.createConnection(pipePath);
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

function createPeekabooChild(configuredAppHome) {
  const appHome =
    configuredAppHome ??
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
  const pipeName =
    readOption("--pipe-name") ??
    process.env.PEEKABOO_WINDOWS_PIPE ??
    defaultPipeName;

  if (mode === "sidecar") {
    const log = createLifecycleLogger(readOption("--log"));
    const sidecar = await startPipeSidecar({
      pipeName,
      appHome: readOption("--home"),
      log,
    });
    log(`Sidecar PID ${process.pid} listening on ${sidecar.pipePath}`);

    const stop = async () => {
      log(`Sidecar PID ${process.pid} stopping`);
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

function readOption(name) {
  const index = process.argv.indexOf(name);
  if (index < 0) return undefined;
  const value = process.argv[index + 1];
  if (!value || value.startsWith("--")) {
    throw new Error(`${name} requires a value`);
  }
  return value;
}

function createLifecycleLogger(logPath) {
  if (!logPath) {
    return (message) => process.stderr.write(`${message}\n`);
  }

  mkdirSync(path.dirname(logPath), { recursive: true });
  return (message) => {
    appendFileSync(logPath, `[${new Date().toISOString()}] ${message}\n`, "utf8");
  };
}

const isDirectRun =
  process.argv[1] &&
  path.resolve(process.argv[1]).toLowerCase() ===
    path.resolve(fileURLToPath(import.meta.url)).toLowerCase();

if (isDirectRun) {
  await main();
}
