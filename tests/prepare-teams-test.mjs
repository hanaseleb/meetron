#!/usr/bin/env node

import { execFile, spawn } from "node:child_process";
import { existsSync } from "node:fs";
import { mkdtemp, rm } from "node:fs/promises";
import net from "node:net";
import { tmpdir } from "node:os";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";
import { connectToChromeOverCDP } from "../scripts/playwright-cdp.mjs";

const execFileAsync = promisify(execFile);
const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const executablePath = [
  process.env.MEETRON_TEST_CHROME,
  "/usr/bin/google-chrome",
  "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
].find((candidate) => candidate && existsSync(candidate));
if (!executablePath) {
  process.stdout.write("Teams preparation browser test skipped: Chrome not installed.\n");
  process.exit(0);
}

const profileDir = await mkdtemp(resolve(tmpdir(), "meetron-prepare-teams-"));
const port = await new Promise((resolvePort, reject) => {
  const server = net.createServer();
  server.on("error", reject);
  server.listen(0, "127.0.0.1", () => {
    const allocated = server.address().port;
    server.close(() => resolvePort(allocated));
  });
});

const chrome = spawn(executablePath, [
  "--headless=new",
  "--no-sandbox",
  "--remote-debugging-address=127.0.0.1",
  `--remote-debugging-port=${port}`,
  `--user-data-dir=${profileDir}`,
  "--no-first-run",
  "--disable-background-networking",
  "about:blank",
], { stdio: "ignore" });

let browser;
try {
  let endpointReady = false;
  for (let attempt = 0; attempt < 80; attempt += 1) {
    endpointReady = await fetch(`http://127.0.0.1:${port}/json/version`)
      .then((response) => response.ok)
      .catch(() => false);
    if (endpointReady) break;
    await new Promise((resolveDelay) => setTimeout(resolveDelay, 100));
  }
  if (!endpointReady) throw new Error("Chrome CDP endpoint did not start.");

  browser = await connectToChromeOverCDP(`http://127.0.0.1:${port}`);
  const context = browser.contexts()[0];
  await context.route("https://teams.microsoft.com/**", (route) =>
    route.fulfill({
      contentType: "text/html",
      body: `<!doctype html><html><body>
        <button aria-label="Continue on this browser" onclick="this.remove()">Continue</button>
        <input aria-label="Type your name">
        <button aria-label="Device settings" onclick="document.querySelector('#devices').hidden = false">Devices</button>
        <div id="devices" hidden>
          <label>Microphone<select aria-label="Microphone">
            <option>Physical microphone</option>
            <option>Meetron: AI to Meeting</option>
          </select></label>
          <label>Speaker<select aria-label="Speaker">
            <option>Physical speaker</option>
            <option>Meetron: Meeting to AI</option>
          </select></label>
        </div>
        <button id="mic" aria-label="Mute" onclick="this.setAttribute('aria-label', 'Unmute')">Mic</button>
        <button aria-label="Turn camera on">Camera</button>
        <button aria-label="Join now" onclick="this.remove(); const leave = document.createElement('button'); leave.setAttribute('aria-label', 'Leave'); document.body.append(leave)">Join</button>
      </body></html>`,
    }),
  );

  const meetingUrl = "https://teams.microsoft.com/l/meetup-join/19%3ameeting_example%40thread.v2/0?context=%7B%7D";
  const { stdout } = await execFileAsync(process.execPath, [
    resolve(repoRoot, "scripts/prepare-teams.mjs"),
    "--cdp",
    `http://127.0.0.1:${port}`,
    "--url",
    meetingUrl,
    "--microphone-device",
    "Meetron: AI to Meeting",
    "--speaker-device",
    "Meetron: Meeting to AI",
    "--join",
    "--join-delay",
    "0",
  ], { cwd: repoRoot, timeout: 30_000 });
  const result = JSON.parse(stdout);
  if (
    result.provider !== "microsoft-teams" ||
    result.continuedInBrowser !== true ||
    result.participantNameFilled !== true ||
    result.microphoneMuted !== true ||
    result.cameraDisabled !== true ||
    result.microphoneDevice !== "Meetron: AI to Meeting" ||
    result.speakerDevice !== "Meetron: Meeting to AI" ||
    result.joinStatus !== "joined"
  ) {
    throw new Error(`Teams preparation failed: ${stdout}`);
  }
} finally {
  await browser?.close().catch(() => {});
  chrome.kill();
  if (chrome.exitCode === null) {
    await Promise.race([
      new Promise((resolveExit) => chrome.once("exit", resolveExit)),
      new Promise((resolveDelay) => setTimeout(resolveDelay, 2_000)),
    ]);
  }
  await rm(profileDir, { recursive: true, force: true, maxRetries: 5, retryDelay: 100 });
}

process.stdout.write("Teams pre-join preparation and admission passed.\n");
