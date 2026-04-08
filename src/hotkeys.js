import { setTimeout as delay } from "node:timers/promises";

const modifierMap = {
  ctrl: "^",
  control: "^",
  shift: "+",
  alt: "%",
};

const keyMap = {
  tab: "{TAB}",
  enter: "{ENTER}",
  return: "{ENTER}",
  esc: "{ESC}",
  escape: "{ESC}",
  space: "{SPACE}",
  up: "{UP}",
  down: "{DOWN}",
  left: "{LEFT}",
  right: "{RIGHT}",
  home: "{HOME}",
  end: "{END}",
  pgup: "{PGUP}",
  pageup: "{PGUP}",
  pgdn: "{PGDN}",
  pagedown: "{PGDN}",
  insert: "{INSERT}",
  delete: "{DEL}",
  del: "{DEL}",
  backspace: "{BACKSPACE}",
};

export function translateHotkey(keys) {
  const tokens = normalizeHotkeyTokens(keys);
  const modifiers = [];
  const nonModifiers = [];

  for (const token of tokens) {
    if (modifierMap[token]) {
      modifiers.push(modifierMap[token]);
    } else {
      nonModifiers.push(token);
    }
  }

  if (nonModifiers.length !== 1) {
    throw new Error("Hotkey requires exactly one non-modifier key");
  }

  return `${modifiers.join("")}${translateTargetKey(nonModifiers[0])}`;
}

export async function pressHotkey({ keys, repeat = 1, delayMs = 80, press }) {
  if (typeof press !== "function") {
    throw new Error("Hotkey press function is required");
  }

  const translatedKeys = translateHotkey(keys);
  const count = Math.max(
    1,
    Number.isFinite(Number(repeat)) ? Number(repeat) : 1,
  );

  for (let index = 0; index < count; index += 1) {
    await press(translatedKeys);

    if (index < count - 1 && delayMs > 0) {
      await delay(delayMs);
    }
  }

  return {
    keys: Array.isArray(keys) ? keys : normalizeHotkeyTokens(keys),
    sendKeys: translatedKeys,
    repeat: count,
  };
}

function normalizeHotkeyTokens(keys) {
  if (Array.isArray(keys)) {
    return keys.map((token) => normalizeToken(token));
  }

  if (typeof keys === "string") {
    return keys
      .split(",")
      .map((token) => normalizeToken(token))
      .filter(Boolean);
  }

  throw new Error("Hotkey keys must be a comma-separated string or array");
}

function normalizeToken(token) {
  return String(token ?? "")
    .trim()
    .toLowerCase();
}

function translateTargetKey(token) {
  if (keyMap[token]) {
    return keyMap[token];
  }

  if (/^f\d{1,2}$/i.test(token)) {
    return `{${token.toUpperCase()}}`;
  }

  if (token.length === 1 && /[a-z0-9]/i.test(token)) {
    return token;
  }

  throw new Error(`Unsupported hotkey key '${token}'`);
}
