import os from "node:os";
import path from "node:path";

export const APP_DIR_NAME = "outlook-classic-mcp";
export const DEFAULT_CONFIRM_TIMEOUT = 180;
export const MIN_CONFIRM_TIMEOUT = 15;
export const MAX_CONFIRM_TIMEOUT = 600;

const isOff = (v) => v === "0" || v.toLowerCase() === "false";
const isOn = (v) => v === "1" || v.toLowerCase() === "true";

/** Folders a non-technical user keeps files in, including OneDrive-redirected ones. */
function defaultAttachRoots(env, home) {
  const roots = ["Documents", "Desktop", "Downloads"].map((d) => path.join(home, d));
  for (const key of ["OneDrive", "OneDriveCommercial", "OneDriveConsumer"]) {
    const base = (env[key] || "").trim();
    if (base) {
      for (const d of ["Documents", "Desktop", "ドキュメント", "デスクトップ"]) {
        roots.push(path.join(base, d));
      }
    }
  }
  return [...new Set(roots)];
}

/**
 * Settings. Everything has a safe default, so no configuration is required.
 * @param {Record<string, string | undefined>} env
 */
export function loadConfig(env = process.env) {
  // The extension host substitutes settings into the environment; a setting that was left
  // empty can arrive as an empty string or as an unreplaced "${...}" placeholder.
  const get = (name) => {
    const value = (env[name] ?? "").trim();
    return value === "" || value.startsWith("${") ? null : value;
  };

  const home = (env.USERPROFILE || "").trim() || os.homedir();
  const config = {
    allowSend: true,
    diagnostics: false,
    confirmTimeout: DEFAULT_CONFIRM_TIMEOUT,
    saveDir: path.join(home, "Downloads", APP_DIR_NAME),
    attachRoots: defaultAttachRoots(env, home),
    auditDir: path.join((env.LOCALAPPDATA || "").trim() || path.join(home, ".local"), APP_DIR_NAME),
    lang: "auto", // "ja" | "en" | "auto"
  };

  // Sending is on unless it is switched off explicitly.
  const allow = get("OUTLOOK_MCP_ALLOW_SEND");
  if (allow !== null && isOff(allow)) config.allowSend = false;

  const diagnostics = get("OUTLOOK_MCP_DIAGNOSTICS");
  if (diagnostics !== null && isOn(diagnostics)) config.diagnostics = true;

  const timeout = get("OUTLOOK_MCP_CONFIRM_TIMEOUT");
  if (timeout !== null && /^-?\d+$/.test(timeout)) {
    config.confirmTimeout = Math.min(MAX_CONFIRM_TIMEOUT, Math.max(MIN_CONFIRM_TIMEOUT, Number(timeout)));
  }

  const saveDir = get("OUTLOOK_MCP_SAVE_DIR");
  if (saveDir !== null) config.saveDir = saveDir;

  const attachDirs = get("OUTLOOK_MCP_ATTACH_DIRS");
  if (attachDirs !== null) {
    const roots = attachDirs.split(";").map((s) => s.trim()).filter(Boolean);
    if (roots.length > 0) config.attachRoots = roots;
  }

  const auditDir = get("OUTLOOK_MCP_AUDIT_DIR");
  if (auditDir !== null) config.auditDir = auditDir;

  const lang = get("OUTLOOK_MCP_LANG")?.toLowerCase();
  if (lang === "ja" || lang === "en") config.lang = lang;

  return config;
}
