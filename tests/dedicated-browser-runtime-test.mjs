#!/usr/bin/env node

import assert from "node:assert/strict";
import {
  browserExecutableCandidates,
  buildDedicatedBrowserLaunch,
  defaultDedicatedProfileDir,
  resolveBrowserExecutable,
} from "../scripts/dedicated-browser-runtime.mjs";

const windowsEnv = {
  LOCALAPPDATA: "C:\\Users\\Tester\\AppData\\Local",
  PROGRAMFILES: "C:\\Program Files",
  "PROGRAMFILES(X86)": "C:\\Program Files (x86)",
};

assert.deepEqual(
  browserExecutableCandidates({ browser: "edge", platform: "win32", env: windowsEnv }),
  [
    "C:\\Program Files\\Microsoft\\Edge\\Application\\msedge.exe",
    "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe",
    "C:\\Users\\Tester\\AppData\\Local\\Microsoft\\Edge\\Application\\msedge.exe",
  ],
);
assert.equal(
  defaultDedicatedProfileDir({ platform: "win32", env: windowsEnv }),
  "C:\\Users\\Tester\\AppData\\Local\\Meetron\\GPTParticipantChromium",
);

const expectedChrome = "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe";
assert.equal(resolveBrowserExecutable({
  browser: "chrome",
  platform: "win32",
  env: windowsEnv,
  pathExists: (candidate) => candidate === expectedChrome,
}), expectedChrome);

const launch = buildDedicatedBrowserLaunch({
  browser: "edge",
  cdpPort: 9333,
  env: windowsEnv,
  meetingUrl: "https://teams.microsoft.com/meet/1234567890123?p=example#ignored",
  pathExists: (candidate) => candidate.endsWith("msedge.exe"),
  platform: "win32",
});
assert.equal(launch.executable, "C:\\Program Files\\Microsoft\\Edge\\Application\\msedge.exe");
assert.equal(launch.provider, "microsoft-teams");
assert.equal(launch.cdpEndpoint, "http://127.0.0.1:9333");
assert(launch.args.includes("--remote-debugging-address=127.0.0.1"));
assert(launch.args.includes("--user-data-dir=C:\\Users\\Tester\\AppData\\Local\\Meetron\\GPTParticipantChromium"));
assert(!launch.url.includes("#"));

for (const port of [80, 65_536, "invalid"]) {
  assert.throws(() => buildDedicatedBrowserLaunch({
    browser: "edge",
    cdpPort: port,
    env: windowsEnv,
    meetingUrl: "https://teams.microsoft.com/meet/1234567890123?p=example",
    pathExists: () => true,
    platform: "win32",
  }), /cdp-port/);
}

assert.throws(() => buildDedicatedBrowserLaunch({
  browser: "edge",
  env: windowsEnv,
  meetingUrl: "https://teams.microsoft.com.evil.example/meet/123",
  pathExists: () => true,
  platform: "win32",
}), /会議URL/);

process.stdout.write("Dedicated browser runtime tests passed.\n");
