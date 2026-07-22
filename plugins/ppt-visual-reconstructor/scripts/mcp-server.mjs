#!/usr/bin/env node

import { spawn } from "node:child_process";
import fs from "node:fs";
import path from "node:path";
import readline from "node:readline";
import { fileURLToPath } from "node:url";

const SERVER_NAME = "slidetwin-tools";
const SERVER_VERSION = "0.5.0";
const SUPPORTED_PROTOCOLS = new Set(["2024-11-05", "2025-03-26", "2025-06-18"]);
const DEFAULT_PROTOCOL = "2025-06-18";
const MAX_CAPTURE_BYTES = 4 * 1024 * 1024;
const DEFAULT_TIMEOUT_MS = 15 * 60 * 1000;
const scriptDirectory = path.dirname(fileURLToPath(import.meta.url));
const pluginRoot = path.dirname(scriptDirectory);

const TOOLS = [
  {
    name: "ppt_environment_status",
    description:
      "Check the installed SlideTwin runtime and explain which tools are available. Use this first on a new machine. It does not launch PowerPoint or modify files.",
    inputSchema: {
      type: "object",
      properties: {},
      additionalProperties: false,
    },
  },
  {
    name: "ppt_probe_powerpoint",
    description:
      "On Windows, launch a hidden desktop PowerPoint COM test, create Chinese text and a shape, save a PPTX, and export a PNG. Use once before the first reconstruction on a machine.",
    inputSchema: {
      type: "object",
      properties: {
        output_directory: {
          type: "string",
          description: "Absolute directory for the probe PPTX, PNG, and JSON report.",
        },
      },
      required: ["output_directory"],
      additionalProperties: false,
    },
  },
  {
    name: "ppt_prepare_image_scene",
    description:
      "Validate a Codex-authored image scene plan, extract protected artwork regions as separate PNG assets, and generate a semantic SVG preview. This is the first executable step after Codex analyzes the uploaded slide image and writes the scene-plan JSON.",
    inputSchema: {
      type: "object",
      properties: {
        reference_image_path: {
          type: "string",
          description: "Absolute path to the source PNG or JPEG slide image.",
        },
        scene_plan_path: {
          type: "string",
          description: "Absolute path to the semantic scene-plan JSON created from the image.",
        },
        output_directory: {
          type: "string",
          description: "Absolute output directory for the resolved plan, assets, SVG, and report.",
        },
      },
      required: ["reference_image_path", "scene_plan_path", "output_directory"],
      additionalProperties: false,
    },
  },
  {
    name: "ppt_build_editable_slide",
    description:
      "On Windows, create a new editable PowerPoint slide from a reference image and semantic scene plan. Text and simple geometry become native PowerPoint objects; only explicitly protected artwork stays as separate pictures. The tool also exports a native PowerPoint PNG render.",
    inputSchema: {
      type: "object",
      properties: {
        reference_image_path: {
          type: "string",
          description: "Absolute path to the source PNG or JPEG slide image.",
        },
        scene_plan_path: {
          type: "string",
          description: "Absolute path to the semantic scene-plan JSON.",
        },
        output_directory: {
          type: "string",
          description: "Absolute output directory for the editable PPTX, render, SVG, assets, and reports.",
        },
        render_width: {
          type: "integer",
          minimum: 800,
          maximum: 7680,
          default: 1600,
          description: "Width in pixels of the native PowerPoint PNG export.",
        },
      },
      required: ["reference_image_path", "scene_plan_path", "output_directory"],
      additionalProperties: false,
    },
  },
  {
    name: "ppt_inspect_powerpoint",
    description:
      "On Windows, open an existing PPTX read-only in desktop PowerPoint, inventory native shapes and grouped objects, and export every slide with the real PowerPoint renderer. Use before producing a correction plan.",
    inputSchema: {
      type: "object",
      properties: {
        presentation_path: {
          type: "string",
          description: "Absolute path to an existing PPTX.",
        },
        output_directory: {
          type: "string",
          description: "Absolute output directory for the inspection JSON and slide renders.",
        },
        render_width: {
          type: "integer",
          minimum: 800,
          maximum: 7680,
          default: 1600,
          description: "Width in pixels of each native PowerPoint slide render.",
        },
      },
      required: ["presentation_path", "output_directory"],
      additionalProperties: false,
    },
  },
  {
    name: "ppt_inspect_pptx_structure",
    description:
      "Read a PPTX package without launching PowerPoint and write a structural JSON inventory. This cross-platform fallback reports slide sizes, object types, text, fonts, images, and possible overflow.",
    inputSchema: {
      type: "object",
      properties: {
        presentation_path: {
          type: "string",
          description: "Absolute path to an existing PPTX.",
        },
        output_directory: {
          type: "string",
          description: "Absolute directory for pptx-structure.json.",
        },
      },
      required: ["presentation_path", "output_directory"],
      additionalProperties: false,
    },
  },
  {
    name: "ppt_compare_slide",
    description:
      "Compare a native PowerPoint slide render against the reference image, compute pixel, SSIM, edge, and composite metrics, and create an overlay plus difference heatmap. Use after every build or correction pass.",
    inputSchema: {
      type: "object",
      properties: {
        reference_image_path: {
          type: "string",
          description: "Absolute path to the reference slide image.",
        },
        candidate_image_path: {
          type: "string",
          description: "Absolute path to the PowerPoint-rendered candidate PNG.",
        },
        output_directory: {
          type: "string",
          description: "Absolute output directory for metrics, overlay, aligned reference, and heatmap.",
        },
        fit: {
          type: "string",
          enum: ["contain", "cover", "stretch"],
          default: "contain",
          description: "How to align a differently sized reference to the candidate canvas.",
        },
      },
      required: ["reference_image_path", "candidate_image_path", "output_directory"],
      additionalProperties: false,
    },
  },
  {
    name: "ppt_apply_correction_plan",
    description:
      "On Windows, copy an existing PPTX, apply a reviewable JSON correction plan, protect designated pictures, save a new editable PPTX, and export new PowerPoint renders. Never modifies the source file.",
    inputSchema: {
      type: "object",
      properties: {
        presentation_path: {
          type: "string",
          description: "Absolute path to the existing PPTX to copy and correct.",
        },
        correction_plan_path: {
          type: "string",
          description: "Absolute path to the explicit JSON correction plan.",
        },
        output_directory: {
          type: "string",
          description: "Absolute directory for reconstructed.pptx, renders, and application-result.json.",
        },
        render_width: {
          type: "integer",
          minimum: 800,
          maximum: 7680,
          default: 1600,
          description: "Width in pixels of each corrected native PowerPoint render.",
        },
      },
      required: ["presentation_path", "correction_plan_path", "output_directory"],
      additionalProperties: false,
    },
  },
];

function send(message) {
  process.stdout.write(`${JSON.stringify(message)}\n`);
}

function textResult(payload, isError = false) {
  return {
    content: [{ type: "text", text: JSON.stringify(payload, null, 2) }],
    ...(isError ? { isError: true } : {}),
  };
}

function absolutePath(value, label) {
  if (typeof value !== "string" || value.trim() === "") {
    throw new Error(`${label} is required.`);
  }
  if (!path.isAbsolute(value)) {
    throw new Error(`${label} must be an absolute path: ${value}`);
  }
  return path.normalize(value);
}

function existingFile(value, label) {
  const resolved = absolutePath(value, label);
  if (!fs.existsSync(resolved) || !fs.statSync(resolved).isFile()) {
    throw new Error(`${label} does not exist or is not a file: ${resolved}`);
  }
  return resolved;
}

function outputDirectory(value) {
  const resolved = absolutePath(value, "output_directory");
  fs.mkdirSync(resolved, { recursive: true });
  return resolved;
}

function integerInRange(value, fallback, minimum, maximum, label) {
  const resolved = value === undefined ? fallback : value;
  if (!Number.isInteger(resolved) || resolved < minimum || resolved > maximum) {
    throw new Error(`${label} must be an integer from ${minimum} through ${maximum}.`);
  }
  return resolved;
}

function requireWindows(toolName) {
  if (process.platform !== "win32") {
    throw new Error(`${toolName} requires 64-bit Windows with desktop PowerPoint installed.`);
  }
}

function appendLimited(current, chunk) {
  if (current.length >= MAX_CAPTURE_BYTES) {
    return current;
  }
  const next = Buffer.from(chunk);
  const remaining = MAX_CAPTURE_BYTES - current.length;
  return Buffer.concat([current, next.subarray(0, remaining)]);
}

function runProcess(command, args, options = {}) {
  return new Promise((resolve, reject) => {
    const child = spawn(command, args, {
      cwd: options.cwd ?? pluginRoot,
      windowsHide: false,
      shell: false,
      env: { ...process.env, PYTHONIOENCODING: "utf-8" },
    });
    let stdout = Buffer.alloc(0);
    let stderr = Buffer.alloc(0);
    let timedOut = false;
    const timeout = setTimeout(() => {
      timedOut = true;
      child.kill();
    }, options.timeoutMs ?? DEFAULT_TIMEOUT_MS);

    child.stdout.on("data", (chunk) => {
      stdout = appendLimited(stdout, chunk);
    });
    child.stderr.on("data", (chunk) => {
      stderr = appendLimited(stderr, chunk);
    });
    child.on("error", (error) => {
      clearTimeout(timeout);
      reject(new Error(`Could not start ${command}: ${error.message}`));
    });
    child.on("close", (exitCode, signal) => {
      clearTimeout(timeout);
      resolve({
        command,
        arguments: args,
        exit_code: exitCode,
        signal,
        timed_out: timedOut,
        stdout: stdout.toString("utf8").trim(),
        stderr: stderr.toString("utf8").trim(),
        output_truncated:
          stdout.length >= MAX_CAPTURE_BYTES || stderr.length >= MAX_CAPTURE_BYTES,
      });
    });
  });
}

function readJsonIfPresent(filePath) {
  if (!filePath || !fs.existsSync(filePath)) {
    return null;
  }
  const text = fs.readFileSync(filePath, "utf8").replace(/^\uFEFF/, "");
  return JSON.parse(text);
}

function pythonCommand() {
  return process.env.PYTHON || (process.platform === "win32" ? "python" : "python3");
}

function powershellCommand() {
  return process.env.POWERSHELL || "powershell.exe";
}

function powershellArgs(scriptName, namedArguments) {
  const args = [
    "-NoLogo",
    "-NoProfile",
    "-NonInteractive",
    "-ExecutionPolicy",
    "Bypass",
    "-File",
    path.join(scriptDirectory, scriptName),
  ];
  for (const [name, value] of namedArguments) {
    args.push(`-${name}`, String(value));
  }
  return args;
}

async function runWithReport(command, args, reportPath) {
  const execution = await runProcess(command, args);
  let report = null;
  let report_error = null;
  try {
    report = readJsonIfPresent(reportPath);
  } catch (error) {
    report_error = `Could not parse report JSON: ${error.message}`;
  }
  const payload = {
    success: execution.exit_code === 0 && !execution.timed_out,
    report_path: reportPath,
    report,
    report_error,
    execution,
  };
  return textResult(payload, !payload.success);
}

async function probeCommand(command, args) {
  try {
    const result = await runProcess(command, args, { timeoutMs: 5_000 });
    return {
      available: result.exit_code === 0,
      command,
      version: (result.stdout || result.stderr).split(/\r?\n/, 1)[0] || null,
    };
  } catch (error) {
    return {
      available: false,
      command,
      error: error instanceof Error ? error.message : String(error),
    };
  }
}

async function callTool(name, input) {
  const args = input && typeof input === "object" && !Array.isArray(input) ? input : {};

  if (name === "ppt_environment_status") {
    const requiredScripts = [
      "prepare_image_scene.py",
      "build_powerpoint_from_scene_v1.ps1",
      "inspect_powerpoint_v1.ps1",
      "inspect_pptx.py",
      "compare_slide_images.py",
      "apply_powerpoint_plan_v4.ps1",
      "probe_powerpoint_v4.ps1",
    ];
    const scripts = Object.fromEntries(
      requiredScripts.map((item) => [item, fs.existsSync(path.join(scriptDirectory, item))]),
    );
    const python = await probeCommand(pythonCommand(), ["--version"]);
    const powershell =
      process.platform === "win32"
        ? await probeCommand(powershellCommand(), ["-NoProfile", "-Command", "$PSVersionTable.PSVersion.ToString()"])
        : { available: false, required: false, reason: "Only required by Windows PowerPoint tools." };
    const runtimeReady =
      Object.values(scripts).every(Boolean) &&
      python.available &&
      (process.platform !== "win32" || powershell.available);
    return textResult({
      success: runtimeReady,
      server: SERVER_NAME,
      version: SERVER_VERSION,
      platform: process.platform,
      architecture: process.arch,
      node_version: process.version,
      plugin_root: pluginRoot,
      python,
      powershell,
      scripts,
      cross_platform_tools: [
        "ppt_prepare_image_scene",
        "ppt_inspect_pptx_structure",
        "ppt_compare_slide",
      ],
      windows_powerpoint_tools: [
        "ppt_probe_powerpoint",
        "ppt_build_editable_slide",
        "ppt_inspect_powerpoint",
        "ppt_apply_correction_plan",
      ],
      next_step:
        process.platform === "win32"
          ? "Run ppt_probe_powerpoint once before the first editable-slide build."
          : "Scene preparation and comparison are available; native PPTX construction requires the Windows PowerPoint bridge.",
    });
  }

  if (name === "ppt_probe_powerpoint") {
    requireWindows(name);
    const output = outputDirectory(args.output_directory);
    return runWithReport(
      powershellCommand(),
      powershellArgs("probe_powerpoint_v4.ps1", [["OutputDirectory", output]]),
      path.join(output, "powerpoint-probe.json"),
    );
  }

  if (name === "ppt_prepare_image_scene") {
    const reference = existingFile(args.reference_image_path, "reference_image_path");
    const plan = existingFile(args.scene_plan_path, "scene_plan_path");
    const output = outputDirectory(args.output_directory);
    return runWithReport(
      pythonCommand(),
      [
        path.join(scriptDirectory, "prepare_image_scene.py"),
        "--reference",
        reference,
        "--plan",
        plan,
        "--output-dir",
        output,
      ],
      path.join(output, "scene-preparation-result.json"),
    );
  }

  if (name === "ppt_build_editable_slide") {
    requireWindows(name);
    const reference = existingFile(args.reference_image_path, "reference_image_path");
    const plan = existingFile(args.scene_plan_path, "scene_plan_path");
    const output = outputDirectory(args.output_directory);
    const width = integerInRange(args.render_width, 1600, 800, 7680, "render_width");
    return runWithReport(
      powershellCommand(),
      powershellArgs("build_powerpoint_from_scene_v1.ps1", [
        ["ReferenceImagePath", reference],
        ["ScenePlanPath", plan],
        ["OutputDirectory", output],
        ["RenderWidth", width],
      ]),
      path.join(output, "image-only-build-result.json"),
    );
  }

  if (name === "ppt_inspect_powerpoint") {
    requireWindows(name);
    const presentation = existingFile(args.presentation_path, "presentation_path");
    const output = outputDirectory(args.output_directory);
    const width = integerInRange(args.render_width, 1600, 800, 7680, "render_width");
    return runWithReport(
      powershellCommand(),
      powershellArgs("inspect_powerpoint_v1.ps1", [
        ["PresentationPath", presentation],
        ["OutputDirectory", output],
        ["RenderWidth", width],
      ]),
      path.join(output, "powerpoint-inspection.json"),
    );
  }

  if (name === "ppt_inspect_pptx_structure") {
    const presentation = existingFile(args.presentation_path, "presentation_path");
    const output = outputDirectory(args.output_directory);
    const reportPath = path.join(output, "pptx-structure.json");
    return runWithReport(
      pythonCommand(),
      [path.join(scriptDirectory, "inspect_pptx.py"), presentation, "--output", reportPath],
      reportPath,
    );
  }

  if (name === "ppt_compare_slide") {
    const reference = existingFile(args.reference_image_path, "reference_image_path");
    const candidate = existingFile(args.candidate_image_path, "candidate_image_path");
    const output = outputDirectory(args.output_directory);
    const fit = args.fit ?? "contain";
    if (!["contain", "cover", "stretch"].includes(fit)) {
      throw new Error("fit must be contain, cover, or stretch.");
    }
    return runWithReport(
      pythonCommand(),
      [
        path.join(scriptDirectory, "compare_slide_images.py"),
        "--reference",
        reference,
        "--candidate",
        candidate,
        "--output-dir",
        output,
        "--fit",
        fit,
      ],
      path.join(output, "comparison.json"),
    );
  }

  if (name === "ppt_apply_correction_plan") {
    requireWindows(name);
    const presentation = existingFile(args.presentation_path, "presentation_path");
    const plan = existingFile(args.correction_plan_path, "correction_plan_path");
    const output = outputDirectory(args.output_directory);
    const width = integerInRange(args.render_width, 1600, 800, 7680, "render_width");
    return runWithReport(
      powershellCommand(),
      powershellArgs("apply_powerpoint_plan_v4.ps1", [
        ["PresentationPath", presentation],
        ["PlanPath", plan],
        ["OutputDirectory", output],
        ["RenderWidth", width],
      ]),
      path.join(output, "application-result.json"),
    );
  }

  throw new Error(`Unknown tool: ${name}`);
}

async function handleRequest(message) {
  const { id, method, params } = message;
  if (method === "initialize") {
    const requested = params?.protocolVersion;
    const protocolVersion = SUPPORTED_PROTOCOLS.has(requested) ? requested : DEFAULT_PROTOCOL;
    return {
      jsonrpc: "2.0",
      id,
      result: {
        protocolVersion,
        capabilities: { tools: { listChanged: false } },
        serverInfo: { name: SERVER_NAME, version: SERVER_VERSION },
      },
    };
  }
  if (method === "tools/list") {
    return { jsonrpc: "2.0", id, result: { tools: TOOLS } };
  }
  if (method === "ping") {
    return { jsonrpc: "2.0", id, result: {} };
  }
  if (method === "tools/call") {
    try {
      const result = await callTool(params?.name, params?.arguments ?? {});
      return { jsonrpc: "2.0", id, result };
    } catch (error) {
      return {
        jsonrpc: "2.0",
        id,
        result: textResult(
          {
            success: false,
            error: error instanceof Error ? error.message : String(error),
            error_type: error instanceof Error ? error.name : "Error",
          },
          true,
        ),
      };
    }
  }
  return {
    jsonrpc: "2.0",
    id,
    error: { code: -32601, message: `Method not found: ${method}` },
  };
}

const input = readline.createInterface({ input: process.stdin, crlfDelay: Infinity });
input.on("line", async (line) => {
  if (line.trim() === "") {
    return;
  }
  let message;
  try {
    message = JSON.parse(line);
  } catch (error) {
    send({ jsonrpc: "2.0", id: null, error: { code: -32700, message: error.message } });
    return;
  }
  if (message.method?.startsWith("notifications/")) {
    return;
  }
  if (message.id === undefined) {
    return;
  }
  send(await handleRequest(message));
});
