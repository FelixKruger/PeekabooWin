import { setTimeout as delay } from "node:timers/promises";

const DEFAULT_PROFILE = "default";
const HUMAN_PACED_PROFILE = "human-paced";

export function resolveInteractionProfile(profile) {
  const value = String(
    profile ?? process.env.PEEKABOO_INTERACTION_PROFILE ?? DEFAULT_PROFILE,
  )
    .trim()
    .toLowerCase();

  if (value === "human" || value === "human-paced") {
    return HUMAN_PACED_PROFILE;
  }

  if (value === "default") {
    return DEFAULT_PROFILE;
  }

  throw new Error(
    `Unsupported interaction profile '${profile}'. Expected 'default' or 'human-paced'`,
  );
}

export function isHumanPacedProfile(profile) {
  return resolveInteractionProfile(profile) === HUMAN_PACED_PROFILE;
}

export async function waitForProfile(profile, stage = "step") {
  if (!isHumanPacedProfile(profile)) {
    return;
  }

  const [min, max] = stageDelayMap[stage] ?? stageDelayMap.step;
  await delay(randomInt(min, max));
}

export function createMousePath(
  from,
  to,
  profile,
  { minSteps = 8, maxSteps = 28 } = {},
) {
  if (!isHumanPacedProfile(profile)) {
    return [to];
  }

  const dx = to.x - from.x;
  const dy = to.y - from.y;
  const distance = Math.hypot(dx, dy);
  const steps = clamp(Math.round(distance / 28), minSteps, maxSteps);
  const path = [];

  for (let index = 1; index <= steps; index += 1) {
    const progress = index / steps;
    const eased = 1 - Math.pow(1 - progress, 3);
    const jitterScale = 1 - progress;
    const jitterX = randomInt(-2, 2) * jitterScale;
    const jitterY = randomInt(-2, 2) * jitterScale;

    path.push({
      x: Math.round(from.x + dx * eased + jitterX),
      y: Math.round(from.y + dy * eased + jitterY),
    });
  }

  path[path.length - 1] = {
    x: to.x,
    y: to.y,
  };

  return dedupePoints(path);
}

export function createTypingChunks(text, profile) {
  if (!isHumanPacedProfile(profile)) {
    return [text];
  }

  const chunks = [];
  let buffer = "";

  for (const char of String(text)) {
    buffer += char;

    if (
      /\s/.test(char) ||
      /[.,;:!?]/.test(char) ||
      buffer.length >= randomInt(1, 3)
    ) {
      chunks.push(buffer);
      buffer = "";
    }
  }

  if (buffer) {
    chunks.push(buffer);
  }

  return chunks;
}

export function typingDelayForChunk(chunk, profile) {
  if (!isHumanPacedProfile(profile)) {
    return 0;
  }

  if (/[.!?]$/.test(chunk)) {
    return randomInt(180, 320);
  }

  if (/[,\s]$/.test(chunk)) {
    return randomInt(90, 180);
  }

  return randomInt(45, 120);
}

export function dragOptionsForProfile(profile, { steps, durationMs } = {}) {
  if (!isHumanPacedProfile(profile)) {
    return {
      steps: steps ?? 16,
      durationMs: durationMs ?? 300,
    };
  }

  return {
    steps: Math.max(steps ?? 0, 24) || 24,
    durationMs: Math.max(durationMs ?? 0, 550) || 550,
  };
}

const stageDelayMap = {
  step: [80, 180],
  beforeClick: [60, 160],
  afterClick: [50, 140],
  beforeScroll: [50, 130],
  afterScroll: [40, 110],
  beforeType: [80, 200],
  beforeHotkey: [60, 150],
};

function dedupePoints(points) {
  const result = [];

  for (const point of points) {
    const previous = result[result.length - 1];

    if (!previous || previous.x !== point.x || previous.y !== point.y) {
      result.push(point);
    }
  }

  return result;
}

function randomInt(min, max) {
  return Math.floor(Math.random() * (max - min + 1)) + min;
}

function clamp(value, min, max) {
  return Math.min(max, Math.max(min, value));
}
