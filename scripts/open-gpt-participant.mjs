#!/usr/bin/env node

import { spawn } from "node:child_process";
import { mkdirSync, readFileSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { setTimeout as delay } from "node:timers/promises";
import {
  buildDedicatedBrowserLaunch,
  isDedicatedBrowserEndpoint,
} from "./dedicated-browser-runtime.mjs";

const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const options = {
  autoPrepare: false,
  browser: "chrome",
  browserPath: process.env.MEETING_COPILOT_BROWSER_PATH || process.env.MEETING_COPILOT_CHROME_PATH || "",
  cdpPort: process.env.MEETING_COPILOT_CDP_PORT || "9223",
  dryRun: false,
  join: false,
  joinDelay: process.env.MEETING_COPILOT_JOIN_DELAY || "2",
  microphoneDevice: "",
  name: process.env.MEETING_COPILOT_NAME || "GPT-Live",
  profileDir: process.env.MEETING_COPILOT_PROFILE_DIR || "",
  speakerDevice: "",
  url: "",
};

function usage(stream = process.stdout) {
  stream.write(`Usage: node scripts/open-gpt-participant.mjs [options] MEETING_URL

Cross-platform launcher for Google Meet and Microsoft Teams Web.

Options:
  --browser chrome|edge       Browser to launch (default: chrome)
  --browser-path PATH         Override the browser executable path
  --profile-dir PATH          Override the dedicated profile directory
  --cdp-port PORT             Local automation port (default: 9223)
  --name NAME                 Participant display name (default: GPT-Live)
  --microphone-device NAME    Virtual microphone name used by meeting preparation
  --speaker-device NAME       Virtual speaker name used by meeting preparation
  --auto-prepare              Prepare the pre-join screen without joining
  --join                      Prepare and request admission
  --join-delay SEC            Wait before requesting admission (default: 2)
  --dry-run                   Print the launch plan without starting the browser
  -h, --help                  Show this help
`);
}

const args = process.argv.slice(2);
for (let index = 0; index < args.length; index += 1) {
  const argument = args[index];
  switch (argument) {
    case "--browser": options.browser = args[++index] || ""; break;
    case "--browser-path": options.browserPath = args[++index] || ""; break;
    case "--profile-dir": options.profileDir = args[++index] || ""; break;
    case "--cdp-port": options.cdpPort = args[++index] || ""; break;
    case "--name": options.name = args[++index] || ""; break;
    case "--microphone-device": options.microphoneDevice = args[++index] || ""; break;
    case "--speaker-device": options.speakerDevice = args[++index] || ""; break;
    case "--auto-prepare": options.autoPrepare = true; break;
    case "--join": options.autoPrepare = true; options.join = true; break;
    case "--join-delay": options.joinDelay = args[++index] || ""; break;
    case "--dry-run": options.dryRun = true; break;
    case "-h":
    case "--help": usage(); process.exit(0); break;
    default:
      if (argument.startsWith("--")) {
        process.stderr.write(`Unknown option: ${argument}\n`);
        usage(process.stderr);
        process.exit(2);
      }
      if (options.url) {
        process.stderr.write("Only one meeting URL may be supplied.\n");
        process.exit(2);
      }
      options.url = argument;
  }
}

if (!options.url) {
  process.stderr.write("A meeting URL is required.\n");
  usage(process.stderr);
  process.exit(2);
}
if (!Number.isFinite(Number(options.joinDelay)) || Number(options.joinDelay) < 0) {
  process.stderr.write("--join-delay must be a non-negative number.\n");
  process.exit(2);
}

let launch;
try {
  launch = buildDedicatedBrowserLaunch({
    meetingUrl: options.url,
    browser: options.browser,
    browserPath: options.browserPath,
    profileDir: options.profileDir,
    cdpPort: options.cdpPort,
  });
} catch (error) {
  process.stderr.write(`${error.message}\n`);
  process.exit(2);
}

process.stdout.write(`Provider:     ${launch.providerLabel}\n`);
process.stdout.write(`Browser:      ${launch.browserLabel} (${launch.executable})\n`);
process.stdout.write(`Profile data: ${launch.profileDir}\n`);
process.stdout.write(`Automation:   ${launch.cdpEndpoint} (local only)\n`);

if (options.dryRun) {
  process.stdout.write(`${JSON.stringify({ executable: launch.executable, args: launch.args }, null, 2)}\n`);
  process.exit(0);
}

mkdirSync(launch.profileDir, { recursive: true });

async function endpointReady() {
  try {
    let activePort = "";
    try {
      activePort = readFileSync(resolve(launch.profileDir, "DevToolsActivePort"), "utf8")
        .split(/\r?\n/, 1)[0];
    } catch {
      // Edge may omit Chromium's profile marker; validate its CDP product below.
    }
    const response = await fetch(`${launch.cdpEndpoint}/json/version`, {
      signal: AbortSignal.timeout(750),
    });
    const body = await response.json();
    return response.ok && isDedicatedBrowserEndpoint({
      activePort,
      browser: launch.browser,
      cdpPort: launch.cdpPort,
      version: body,
    });
  } catch {
    return false;
  }
}

if (!await endpointReady()) {
  const child = spawn(launch.executable, launch.args, {
    detached: true,
    stdio: "ignore",
    windowsHide: false,
  });
  child.on("error", (error) => {
    process.stderr.write(`Browser launch failed: ${error.message}\n`);
  });
  child.unref();

  let ready = false;
  for (let attempt = 0; attempt < 40; attempt += 1) {
    await delay(250);
    if (await endpointReady()) {
      ready = true;
      break;
    }
  }
  if (!ready) {
    process.stderr.write(`Browser automation endpoint did not start at ${launch.cdpEndpoint}.\n`);
    process.exit(1);
  }
} else {
  process.stdout.write("[INFO] Reusing the running dedicated browser profile.\n");
  if (!options.autoPrepare) {
    const child = spawn(launch.executable, launch.args, {
      detached: true,
      stdio: "ignore",
      windowsHide: false,
    });
    child.unref();
  }
}

if (options.autoPrepare) {
  const prepareScript = launch.provider === "microsoft-teams" ? "prepare-teams.mjs" : "prepare-meet.mjs";
  const prepareArgs = [
    resolve(repoRoot, "scripts", prepareScript),
    "--cdp", launch.cdpEndpoint,
    "--name", options.name,
    "--url", launch.url,
  ];
  if (options.microphoneDevice) prepareArgs.push("--microphone-device", options.microphoneDevice);
  if (options.speakerDevice) prepareArgs.push("--speaker-device", options.speakerDevice);
  if (options.join) prepareArgs.push("--join", "--join-delay", String(options.joinDelay));

  const preparation = spawn(process.execPath, prepareArgs, { stdio: "inherit" });
  const exitCode = await new Promise((resolveExit) => preparation.once("exit", resolveExit));
  if (exitCode !== 0) process.exit(exitCode ?? 1);
}

process.stdout.write("Dedicated meeting browser is ready.\n");
