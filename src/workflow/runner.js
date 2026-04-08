import fs from "node:fs/promises";
import path from "node:path";

import { defaultWorkflowActionHandlers } from "./actions.js";

export async function runWorkflowFile({
  filePath,
  actionHandlers = defaultWorkflowActionHandlers,
  continueOnError,
  profile,
}) {
  if (!filePath) {
    throw new Error("Missing workflow file path");
  }

  const resolvedPath = path.resolve(filePath);
  const raw = await fs.readFile(resolvedPath, "utf8");
  const definition = JSON.parse(stripBom(raw));
  const result = await runWorkflowDefinition({
    definition,
    actionHandlers,
    continueOnError,
    profile,
  });

  return {
    workflowPath: resolvedPath,
    ...result,
  };
}

export async function runWorkflowDefinition({
  definition,
  actionHandlers = defaultWorkflowActionHandlers,
  continueOnError,
  profile,
}) {
  if (!definition || typeof definition !== "object") {
    throw new Error("Workflow definition must be an object");
  }

  if (!Array.isArray(definition.steps)) {
    throw new Error("Workflow definition must include a steps array");
  }

  const context = resolveWorkflowValue(definition.vars ?? {}, {});
  const workflowProfile = profile ?? definition.profile;
  const shouldContinueOnError =
    continueOnError ?? definition.continueOnError === true;
  const stepResults = [];

  for (let index = 0; index < definition.steps.length; index += 1) {
    const step = definition.steps[index];

    if (!step || typeof step !== "object") {
      throw new Error(`Workflow step ${index + 1} must be an object`);
    }

    if (!step.action || typeof step.action !== "string") {
      throw new Error(`Workflow step ${index + 1} is missing an action`);
    }

    const handler = actionHandlers[step.action];

    if (!handler) {
      throw new Error(`Unknown workflow action '${step.action}'`);
    }

    const stepInput = resolveStepInput(step, context, workflowProfile);

    try {
      const result = await handler(stepInput, {
        context,
        step,
        index,
      });

      if (step.saveAs) {
        context[step.saveAs] = result;
      }

      stepResults.push({
        index,
        name: step.name ?? null,
        action: step.action,
        saveAs: step.saveAs ?? null,
        ok: true,
        result,
      });
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      stepResults.push({
        index,
        name: step.name ?? null,
        action: step.action,
        saveAs: step.saveAs ?? null,
        ok: false,
        error: message,
      });

      if (!(shouldContinueOnError || step.continueOnError === true)) {
        return {
          ok: false,
          continueOnError: shouldContinueOnError,
          steps: stepResults,
          context,
        };
      }
    }
  }

  return {
    ok: stepResults.every((step) => step.ok),
    continueOnError: shouldContinueOnError,
    steps: stepResults,
    context,
  };
}

export function resolveWorkflowValue(value, context) {
  if (typeof value === "string") {
    if (value.startsWith("$$")) {
      return value.slice(1);
    }

    if (value.startsWith("$")) {
      return readContextPath(context, value.slice(1));
    }

    return value;
  }

  if (Array.isArray(value)) {
    return value.map((entry) => resolveWorkflowValue(entry, context));
  }

  if (value && typeof value === "object") {
    return Object.fromEntries(
      Object.entries(value).map(([key, entry]) => [
        key,
        resolveWorkflowValue(entry, context),
      ]),
    );
  }

  return value;
}

function resolveStepInput(step, context, workflowProfile) {
  const input = {};

  for (const [key, value] of Object.entries(step)) {
    if (["action", "saveAs", "name", "continueOnError"].includes(key)) {
      continue;
    }

    input[key] = resolveWorkflowValue(value, context);
  }

  if (workflowProfile && input.profile === undefined) {
    input.profile = workflowProfile;
  }

  return input;
}

function readContextPath(context, pathText) {
  if (!pathText) {
    return context;
  }

  const segments = pathText.split(".");
  let current = context;

  for (const segment of segments) {
    if (current === null || current === undefined) {
      throw new Error(
        `Workflow reference '$${pathText}' could not be resolved`,
      );
    }

    if (Array.isArray(current)) {
      const index = Number(segment);

      if (!Number.isInteger(index) || index < 0 || index >= current.length) {
        throw new Error(
          `Workflow reference '$${pathText}' could not be resolved`,
        );
      }

      current = current[index];
      continue;
    }

    if (!Object.prototype.hasOwnProperty.call(current, segment)) {
      throw new Error(
        `Workflow reference '$${pathText}' could not be resolved`,
      );
    }

    current = current[segment];
  }

  return current;
}

function stripBom(text) {
  return text.charCodeAt(0) === 0xfeff ? text.slice(1) : text;
}
