import { spawn } from "node:child_process";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { BackendError } from "../errors.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const scriptPath = path.resolve(
  __dirname,
  "../../scripts/windows-automation.ps1",
);

export async function invokeWindowsAutomation(action, payload = {}) {
  const request = { action, payload };
  const stdoutChunks = [];
  const stderrChunks = [];

  return await new Promise((resolve, reject) => {
    const child = spawn(
      "powershell.exe",
      [
        "-NoLogo",
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        scriptPath,
        "-RequestJson",
        JSON.stringify(request),
      ],
      {
        windowsHide: true,
      },
    );

    child.stdout.on("data", (chunk) => {
      stdoutChunks.push(chunk);
    });

    child.stderr.on("data", (chunk) => {
      stderrChunks.push(chunk);
    });

    child.on("error", (error) => {
      reject(
        new BackendError("Failed to launch PowerShell backend", {
          cause: error,
        }),
      );
    });

    child.on("close", (code) => {
      const stdout = Buffer.concat(stdoutChunks).toString("utf8").trim();
      const stderr = Buffer.concat(stderrChunks).toString("utf8").trim();
      let parsed = null;

      if (!stdout) {
        reject(
          new BackendError("PowerShell backend returned no output", {
            code,
            stderr,
          }),
        );
        return;
      }

      try {
        parsed = JSON.parse(stdout);
      } catch (error) {
        if (code !== 0) {
          reject(
            new BackendError(
              `PowerShell backend exited with an error${stderr ? `: ${stderr}` : ""}${stdout ? ` ${stdout}` : ""}`,
              {
                code,
                stderr,
                stdout,
                cause: error,
              },
            ),
          );
          return;
        }

        reject(
          new BackendError("Failed to parse PowerShell backend output", {
            stdout,
            stderr,
            cause: error,
          }),
        );
        return;
      }

      if (parsed?.ok === false) {
        reject(
          new BackendError(
            parsed.error ?? "Windows automation command failed",
            {
              action,
              payload,
              code,
              stderr,
              data: parsed,
            },
          ),
        );
        return;
      }

      if (code !== 0) {
        reject(
          new BackendError("PowerShell backend exited with an error", {
            code,
            stderr,
            stdout,
          }),
        );
        return;
      }

      resolve(parsed.result);
    });
  });
}
