import { existsSync } from "node:fs";
import path from "node:path";
import { normalizeMeetingJoinUrl } from "./meeting-provider.mjs";

const BROWSERS = Object.freeze({
  chrome: Object.freeze({ label: "Google Chrome" }),
  edge: Object.freeze({ label: "Microsoft Edge" }),
});

function platformPath(platform) {
  return platform === "win32" ? path.win32 : path.posix;
}

function compact(values) {
  return [...new Set(values.filter(Boolean))];
}

export function browserExecutableCandidates({
  browser = "chrome",
  platform = process.platform,
  env = process.env,
} = {}) {
  if (!BROWSERS[browser]) {
    throw new Error(`Unsupported browser: ${browser}`);
  }

  const paths = platformPath(platform);
  if (platform === "win32") {
    const roots = compact([
      env.PROGRAMFILES,
      env["PROGRAMFILES(X86)"],
      env.LOCALAPPDATA,
    ]);
    const relative = browser === "edge"
      ? ["Microsoft", "Edge", "Application", "msedge.exe"]
      : ["Google", "Chrome", "Application", "chrome.exe"];
    return roots.map((root) => paths.join(root, ...relative));
  }

  if (platform === "darwin") {
    const application = browser === "edge" ? "Microsoft Edge" : "Google Chrome";
    return compact([
      `/Applications/${application}.app/Contents/MacOS/${application}`,
      env.HOME
        ? paths.join(env.HOME, "Applications", `${application}.app`, "Contents", "MacOS", application)
        : "",
    ]);
  }

  return browser === "edge"
    ? ["/usr/bin/microsoft-edge", "/usr/bin/microsoft-edge-stable"]
    : ["/usr/bin/google-chrome", "/usr/bin/google-chrome-stable", "/usr/bin/chromium"];
}

export function defaultDedicatedProfileDir({
  platform = process.platform,
  env = process.env,
} = {}) {
  const paths = platformPath(platform);
  if (platform === "win32") {
    const root = env.LOCALAPPDATA || env.APPDATA;
    if (!root) throw new Error("LOCALAPPDATA or APPDATA is required on Windows.");
    return paths.join(root, "Meetron", "GPTParticipantChromium");
  }
  if (platform === "darwin") {
    if (!env.HOME) throw new Error("HOME is required on macOS.");
    return paths.join(
      env.HOME,
      "Library",
      "Application Support",
      "MeetingCopilot",
      "GPTParticipantChrome",
    );
  }
  const root = env.XDG_STATE_HOME || (env.HOME ? paths.join(env.HOME, ".local", "state") : "");
  if (!root) throw new Error("XDG_STATE_HOME or HOME is required.");
  return paths.join(root, "meetron", "gpt-participant-chromium");
}

export function resolveBrowserExecutable({
  browser = "chrome",
  browserPath = "",
  platform = process.platform,
  env = process.env,
  pathExists = existsSync,
} = {}) {
  const candidates = compact([
    browserPath,
    ...browserExecutableCandidates({ browser, platform, env }),
  ]);
  const executable = candidates.find((candidate) => pathExists(candidate));
  if (!executable) {
    throw new Error(
      `${BROWSERS[browser]?.label || browser} was not found. ` +
      "Use --browser-path to specify its executable.",
    );
  }
  return executable;
}

export function buildDedicatedBrowserLaunch({
  meetingUrl,
  browser = "chrome",
  browserPath = "",
  profileDir = "",
  cdpPort = 9223,
  platform = process.platform,
  env = process.env,
  pathExists = existsSync,
} = {}) {
  const normalized = normalizeMeetingJoinUrl(meetingUrl);
  const numericPort = Number(cdpPort);
  if (!Number.isInteger(numericPort) || numericPort < 1024 || numericPort > 65_535) {
    throw new Error("--cdp-port must be an integer between 1024 and 65535.");
  }
  const selectedProfileDir = profileDir || defaultDedicatedProfileDir({ platform, env });
  const executable = resolveBrowserExecutable({
    browser,
    browserPath,
    platform,
    env,
    pathExists,
  });
  return {
    browser,
    browserLabel: BROWSERS[browser].label,
    cdpEndpoint: `http://127.0.0.1:${numericPort}`,
    cdpPort: numericPort,
    executable,
    profileDir: selectedProfileDir,
    provider: normalized.provider,
    providerLabel: normalized.providerLabel,
    url: normalized.url,
    args: [
      "--remote-debugging-address=127.0.0.1",
      `--remote-debugging-port=${numericPort}`,
      `--user-data-dir=${selectedProfileDir}`,
      "--no-first-run",
      "--new-window",
      normalized.url,
    ],
  };
}
