#!/usr/bin/env node

import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import path from "node:path";
import readline from "node:readline";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const serverPath = path.join(
  root,
  "plugins",
  "ppt-visual-reconstructor",
  "scripts",
  "mcp-server.mjs",
);

const child = spawn(process.execPath, [serverPath], {
  cwd: root,
  stdio: ["pipe", "pipe", "pipe"],
});
const lines = readline.createInterface({ input: child.stdout, crlfDelay: Infinity });
const pending = new Map();
let nextId = 1;
let stderr = "";

child.stderr.on("data", (chunk) => {
  stderr += chunk.toString("utf8");
});

lines.on("line", (line) => {
  const message = JSON.parse(line);
  const waiter = pending.get(message.id);
  if (waiter) {
    pending.delete(message.id);
    waiter.resolve(message);
  }
});

function request(method, params = {}) {
  const id = nextId++;
  return new Promise((resolve, reject) => {
    const timeout = setTimeout(() => {
      pending.delete(id);
      reject(new Error(`Timed out waiting for MCP response ${id}. stderr: ${stderr}`));
    }, 10_000);
    pending.set(id, {
      resolve: (value) => {
        clearTimeout(timeout);
        resolve(value);
      },
    });
    child.stdin.write(`${JSON.stringify({ jsonrpc: "2.0", id, method, params })}\n`);
  });
}

try {
  const initialized = await request("initialize", {
    protocolVersion: "2025-06-18",
    capabilities: {},
    clientInfo: { name: "repository-test", version: "1.0.0" },
  });
  assert.equal(initialized.result.serverInfo.name, "ppt-visual-tools");
  assert.equal(initialized.result.serverInfo.version, "0.4.0");

  child.stdin.write(
    `${JSON.stringify({ jsonrpc: "2.0", method: "notifications/initialized", params: {} })}\n`,
  );

  const listed = await request("tools/list");
  const names = listed.result.tools.map((tool) => tool.name);
  assert.deepEqual(names, [
    "ppt_environment_status",
    "ppt_probe_powerpoint",
    "ppt_prepare_image_scene",
    "ppt_build_editable_slide",
    "ppt_inspect_powerpoint",
    "ppt_inspect_pptx_structure",
    "ppt_compare_slide",
    "ppt_apply_correction_plan",
  ]);

  const statusResponse = await request("tools/call", {
    name: "ppt_environment_status",
    arguments: {},
  });
  const status = JSON.parse(statusResponse.result.content[0].text);
  assert.equal(status.success, true);
  assert.equal(status.server, "ppt-visual-tools");
  assert.equal(status.scripts["prepare_image_scene.py"], true);

  const invalidPathResponse = await request("tools/call", {
    name: "ppt_prepare_image_scene",
    arguments: {
      reference_image_path: "relative-reference.png",
      scene_plan_path: "relative-scene.json",
      output_directory: "relative-output",
    },
  });
  assert.equal(invalidPathResponse.result.isError, true);
  const invalidPath = JSON.parse(invalidPathResponse.result.content[0].text);
  assert.match(invalidPath.error, /absolute path/);

  console.log("MCP server protocol test passed");
} finally {
  child.stdin.end();
}
