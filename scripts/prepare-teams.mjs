#!/usr/bin/env node

import { getAudioStatus } from "./audio-backend.mjs";
import {
  isMeetingProviderPage,
  normalizeMeetingJoinUrl,
} from "./meeting-provider.mjs";
import { connectToChromeOverCDP } from "./playwright-cdp.mjs";

const args = process.argv.slice(2);
const options = {
  cdp: "http://127.0.0.1:9223",
  join: false,
  joinDelay: 2,
  microphoneDevice: "",
  name: "GPT-Live",
  speakerDevice: "",
  url: "",
};

function usage() {
  process.stdout.write(`Usage: node scripts/prepare-teams.mjs [options]\n\nOptions:\n  --cdp URL                Chrome DevTools endpoint (default: ${options.cdp})\n  --name NAME              Teams participant name (default: ${options.name})\n  --url URL                Expected Microsoft Teams Web meeting URL\n  --microphone-device NAME Override the selected virtual microphone\n  --speaker-device NAME    Override the selected virtual speaker\n  --join                   Join the meeting after preparing the pre-join screen\n  --join-delay SEC         Wait before joining (default: ${options.joinDelay})\n  -h, --help               Show this help\n`);
}

for (let index = 0; index < args.length; index += 1) {
  const argument = args[index];
  switch (argument) {
    case "--cdp":
      options.cdp = args[++index] || "";
      break;
    case "--name":
      options.name = args[++index] || "";
      break;
    case "--microphone-device":
      options.microphoneDevice = args[++index] || "";
      break;
    case "--speaker-device":
      options.speakerDevice = args[++index] || "";
      break;
    case "--url":
      options.url = args[++index] || "";
      break;
    case "--join":
      options.join = true;
      break;
    case "--join-delay":
      options.joinDelay = Number(args[++index]);
      break;
    case "-h":
    case "--help":
      usage();
      process.exit(0);
      break;
    default:
      process.stderr.write(`Unknown argument: ${argument}\n`);
      usage();
      process.exit(2);
  }
}

const normalized = normalizeMeetingJoinUrl(options.url, {
  allowedProviders: ["microsoft-teams"],
});
options.url = normalized.url;

if (!Number.isFinite(options.joinDelay) || options.joinDelay < 0) {
  process.stderr.write("--join-delay must be a non-negative number.\n");
  process.exit(2);
}

if (!options.microphoneDevice || !options.speakerDevice) {
  const audio = await getAudioStatus();
  options.microphoneDevice ||= audio.routing.meetingMicrophone.name;
  options.speakerDevice ||= audio.routing.meetingSpeaker.name;
}

function exactDevicePattern(name) {
  return new RegExp(`^${name.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}$`, "i");
}

async function locatorIsVisible(locator) {
  try {
    return (await locator.count()) > 0 && (await locator.first().isVisible());
  } catch {
    return false;
  }
}

async function clickFirst(locators, timeout = 2_000) {
  for (const locator of locators) {
    try {
      if (await locatorIsVisible(locator)) {
        await locator.first().click({ timeout });
        return true;
      }
    } catch {
      // Teams changes frequently; try the next accessible representation.
    }
  }
  return false;
}

const browser = await connectToChromeOverCDP(options.cdp);
const context = browser.contexts()[0];
if (!context) {
  throw new Error("Chrome did not expose a browser context.");
}

await context.grantPermissions(["microphone"], {
  origin: "https://teams.microsoft.com",
});

const teamsPages = context
  .pages()
  .filter((candidate) => isMeetingProviderPage(candidate.url(), "microsoft-teams"));
let page = teamsPages.find((candidate) => candidate.url().startsWith(options.url));
await Promise.all(
  teamsPages.filter((candidate) => candidate !== page).map((candidate) => candidate.close()),
);

if (!page) {
  page = await context.newPage();
  await page.goto(options.url, { waitUntil: "domcontentloaded" });
}

await page.bringToFront();
await page.waitForLoadState("domcontentloaded");
page.setDefaultTimeout(5_000);

const continuedInBrowser = await clickFirst([
  page.getByRole("button", { name: /このブラウザーで続ける|このブラウザで続ける|continue on this browser/i }),
  page.getByRole("link", { name: /このブラウザーで続ける|このブラウザで続ける|continue on this browser/i }),
]);
if (continuedInBrowser) {
  await page.waitForTimeout(500);
}

let participantNameFilled = false;
for (const field of [
  page.getByLabel(/名前を入力|名前|type your name|enter your name/i),
  page.getByPlaceholder(/名前を入力|名前|type your name|enter your name/i),
]) {
  try {
    if ((await field.count()) === 1 && (await field.first().isVisible())) {
      await field.first().fill(options.name);
      participantNameFilled = true;
      break;
    }
  } catch {
    // Signed-in meetings do not show the guest name field.
  }
}

const openedDeviceSettings = await clickFirst([
  page.getByRole("button", { name: /デバイスの設定|オーディオの設定|device settings|audio settings/i }),
  page.getByLabel(/デバイスの設定|オーディオの設定|device settings|audio settings/i),
]);
if (!openedDeviceSettings) {
  throw new Error("Teamsのデバイス設定を開けませんでした");
}

async function selectDevice(label, targetName) {
  const pattern = exactDevicePattern(targetName);
  const candidates = [
    page.getByRole("combobox", { name: label }),
    page.getByLabel(label),
  ];
  let control;
  for (const candidate of candidates) {
    if (await locatorIsVisible(candidate)) {
      control = candidate.first();
      break;
    }
  }
  if (!control) {
    throw new Error(`Teamsの音声デバイス欄が見つかりません: ${label}`);
  }

  const tagName = await control.evaluate((element) => element.tagName.toLowerCase());
  if (tagName === "select") {
    await control.selectOption({ label: targetName });
  } else {
    const current = [
      await control.getAttribute("aria-label"),
      await control.getAttribute("value"),
      await control.textContent(),
    ].filter(Boolean).join(" ");
    if (!pattern.test(current.trim())) {
      await control.click();
      const option = page.getByRole("option", { name: pattern });
      const menuItem = page.getByRole("menuitemradio", { name: pattern });
      if (await locatorIsVisible(option)) {
        await option.first().click();
      } else if (await locatorIsVisible(menuItem)) {
        await menuItem.first().click();
      } else {
        throw new Error(`Teamsの音声デバイスを選択できません: ${targetName}`);
      }
    }
  }

  const selected = tagName === "select"
    ? await control.locator("option:checked").textContent()
    : [
        await control.getAttribute("aria-label"),
        await control.getAttribute("value"),
        await control.textContent(),
      ].filter(Boolean).join(" ");
  if (!pattern.test(String(selected || "").trim())) {
    throw new Error(`Teams did not select the required audio device: ${targetName}`);
  }
  return String(selected).trim();
}

const microphoneDevice = await selectDevice(/^(マイク|Microphone)$/i, options.microphoneDevice);
const speakerDevice = await selectDevice(/^(スピーカー|Speaker)$/i, options.speakerDevice);

const turnMicrophoneOn = page.getByRole("button", {
  name: /^(ミュート解除|マイクをオン|unmute|turn microphone on)(?:\s|$)/i,
});
const turnMicrophoneOff = page.getByRole("button", {
  name: /^(ミュート|マイクをオフ|mute|turn microphone off)(?:\s|$)/i,
});
let microphoneMuted = await locatorIsVisible(turnMicrophoneOn);
if (!microphoneMuted && await locatorIsVisible(turnMicrophoneOff)) {
  await turnMicrophoneOff.first().click();
  await page.waitForTimeout(300);
  microphoneMuted = await locatorIsVisible(turnMicrophoneOn);
}
if (!microphoneMuted) {
  throw new Error("Teams microphone could not be verified as muted before admission.");
}

const turnCameraOn = page.getByRole("button", {
  name: /^(カメラをオン|ビデオを開始|turn camera on|start video)(?:\s|$)/i,
});
const turnCameraOff = page.getByRole("button", {
  name: /^(カメラをオフ|ビデオを停止|turn camera off|stop video)(?:\s|$)/i,
});
let cameraDisabled = await locatorIsVisible(turnCameraOn);
if (!cameraDisabled && await locatorIsVisible(turnCameraOff)) {
  await turnCameraOff.first().click();
  await page.waitForTimeout(300);
  cameraDisabled = await locatorIsVisible(turnCameraOn);
}
if (!cameraDisabled) {
  throw new Error("Teams camera could not be verified as disabled before admission.");
}

let joinStatus = "not-requested";
if (options.join) {
  const joinButton = page.getByRole("button", {
    name: /今すぐ参加|参加をリクエスト|join now|ask to join|join meeting/i,
  });
  await joinButton.first().waitFor({ state: "visible", timeout: 15_000 });
  await page.waitForTimeout(options.joinDelay * 1_000);
  await joinButton.first().click({ force: true, timeout: 5_000 });
  await page.waitForTimeout(500);

  const leave = page.getByRole("button", {
    name: /退出|通話から退出|leave|hang up/i,
  });
  const bodyText = await page.locator("body").innerText().catch(() => "");
  if (await locatorIsVisible(leave)) {
    joinStatus = "joined";
  } else if (/ロビー|参加を許可|waiting.*(?:admit|lobby)|someone.*let you in/i.test(bodyText)) {
    joinStatus = "waiting-for-admission";
  } else if (/参加できません|拒否|can't join|cannot join|denied/i.test(bodyText)) {
    joinStatus = "rejected";
  } else {
    joinStatus = "requested-status-unknown";
  }
}

process.stdout.write(`${JSON.stringify({
  url: page.url(),
  provider: "microsoft-teams",
  continuedInBrowser,
  participantNameFilled,
  microphoneMuted,
  cameraDisabled,
  microphoneDevice,
  speakerDevice,
  joinStatus,
  title: await page.title(),
}, null, 2)}\n`);

if (joinStatus === "rejected") process.exit(14);
if (joinStatus === "requested-status-unknown") process.exit(15);
process.exit(0);
