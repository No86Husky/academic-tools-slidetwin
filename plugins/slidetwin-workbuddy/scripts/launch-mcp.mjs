#!/usr/bin/env node

import { spawn } from "node:child_process";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import readline from "node:readline";
import { fileURLToPath } from "node:url";

const launcherDirectory = path.dirname(fileURLToPath(import.meta.url));
const pluginRoot = path.dirname(launcherDirectory);
const requiredScripts = [
  "mcp-server.mjs",
  "prepare_image_scene.py",
  "build_powerpoint_from_scene_v1.ps1",
  "inspect_powerpoint_v1.ps1",
  "inspect_pptx.py",
  "compare_slide_images.py",
  "apply_powerpoint_plan_v4.ps1",
  "probe_powerpoint_v4.ps1",
];

function normalizeRuntimeRoot(candidate) {
  if (!candidate || typeof candidate !== "string") {
    return null;
  }
  const resolved = path.resolve(candidate);
  if (path.basename(resolved).toLowerCase() === "scripts") {
    return path.dirname(resolved);
  }
  return resolved;
}

function isUsableRuntime(candidate) {
  const root = normalizeRuntimeRoot(candidate);
  if (!root) {
    return false;
  }
  const scriptsDirectory = path.join(root, "scripts");
  return requiredScripts.every((name) => fs.existsSync(path.join(scriptsDirectory, name)));
}

function runtimeCandidates() {
  const home = os.homedir();
  return [
    process.env.SLIDETWIN_RUNTIME_ROOT,
    path.resolve(pluginRoot, "..", "ppt-visual-reconstructor"),
    path.join(home, ".slidetwin", "academic-tools-slidetwin", "plugins", "ppt-visual-reconstructor"),
    path.join(home, "academic-tools-slidetwin", "plugins", "ppt-visual-reconstructor"),
  ]
    .map(normalizeRuntimeRoot)
    .filter(Boolean);
}

function genericDescription(value) {
  if (typeof value !== "string") {
    return value;
  }
  return value
    .replaceAll("a Codex-authored", "an AI-agent-authored")
    .replaceAll("Codex-authored", "AI-agent-authored")
    .replaceAll("after Codex analyzes", "after the host AI agent analyzes")
    .replaceAll("Codex tools", "agent tools")
    .replaceAll("Codex", "the host AI agent");
}

function adaptResponse(message) {
  if (message?.result?.serverInfo?.name === "slidetwin-tools") {
    message.result.serverInfo.name = "slidetwin-workbuddy-tools";
  }
  if (Array.isArray(message?.result?.tools)) {
    message.result.tools = message.result.tools.map((tool) => ({
      ...tool,
      description: genericDescription(tool.description),
    }));
  }
  return message;
}

const runtimeRoot = runtimeCandidates().find(isUsableRuntime);
if (!runtimeRoot) {
  const checked = runtimeCandidates().map((item) => `  - ${item}`).join("\n");
  process.stderr.write(
    [
      "SlideTwin runtime was not found.",
      "Run install-workbuddy.ps1 from the SlideTwin repository, or set SLIDETWIN_RUNTIME_ROOT",
      "to the existing plugins/ppt-visual-reconstructor directory.",
      "Checked:",
      checked || "  - no candidate paths",
      "",
    ].join("\n"),
  );
  process.exit(1);
}

const serverPath = path.join(runtimeRoot, "scripts", "mcp-server.mjs");
const child = spawn(process.execPath, [serverPath], {
  cwd: runtimeRoot,
  env: {
    ...process.env,
    SLIDETWIN_RUNTIME_ROOT: runtimeRoot,
  },
  stdio: ["pipe", "pipe", "pipe"],
  windowsHide: true,
});

process.stdin.pipe(child.stdin);
child.stderr.pipe(process.stderr);

const childOutput = readline.createInterface({ input: child.stdout, crlfDelay: Infinity });
childOutput.on("line", (line) => {
  try {
    const message = adaptResponse(JSON.parse(line));
    process.stdout.write(`${JSON.stringify(message)}\n`);
  } catch {
    process.stdout.write(`${line}\n`);
  }
});

child.on("error", (error) => {
  process.stderr.write(`Could not start the SlideTwin MCP runtime: ${error.message}\n`);
  process.exit(1);
});

child.on("exit", (code, signal) => {
  if (signal) {
    process.kill(process.pid, signal);
    return;
  }
  process.exit(code ?? 1);
});

for (const signal of ["SIGINT", "SIGTERM"]) {
  process.on(signal, () => {
    if (!child.killed) {
      child.kill(signal);
    }
  });
}
