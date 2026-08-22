#!/usr/bin/env node

import {
  isMeetingProviderPage,
  meetingProviderForUrl,
  normalizeMeetingJoinUrl,
} from "../scripts/meeting-provider.mjs";

const cases = [
  ["https://meet.google.com/abc-defg-hij", "google-meet"],
  [
    "https://teams.microsoft.com/l/meetup-join/19%3ameeting_example%40thread.v2/0?context=%7B%7D",
    "microsoft-teams",
  ],
  [
    "https://teams.microsoft.com/v2/l/meetup-join/19%3ameeting_example%40thread.v2/0?context=%7B%7D",
    "microsoft-teams",
  ],
  [
    "https://teams.microsoft.com/l/meetup-join/19%3ameeting_example%40thread.v2",
    "microsoft-teams",
  ],
  ["https://teams.microsoft.com/meet/1234567890123?p=example", "microsoft-teams"],
];

for (const [value, expected] of cases) {
  const provider = meetingProviderForUrl(value);
  const normalized = normalizeMeetingJoinUrl(`${value}#ignored`);
  if (provider?.id !== expected || normalized.provider !== expected || normalized.url.includes("#")) {
    throw new Error(`Meeting URL was not normalized: ${value}`);
  }
}

for (const value of [
  "http://meet.google.com/abc-defg-hij",
  "https://user@meet.google.com/abc-defg-hij",
  "https://meet.google.com.evil.example/abc-defg-hij",
  "https://teams.microsoft.com/v2/",
  "https://teams.microsoft.com/l/chat/0/0",
  "https://example.com/meeting",
]) {
  if (meetingProviderForUrl(value)) {
    throw new Error(`Unsafe or unsupported URL was accepted: ${value}`);
  }
}

try {
  normalizeMeetingJoinUrl(cases[1][0], { allowedProviders: ["google-meet"] });
  throw new Error("Provider allow-list was not enforced.");
} catch (error) {
  if (!error.message.includes("meet.google.com")) throw error;
}

if (
  !isMeetingProviderPage("https://teams.microsoft.com/v2/", "microsoft-teams") ||
  isMeetingProviderPage("https://teams.microsoft.com.evil.example/v2/", "microsoft-teams")
) {
  throw new Error("Provider page origin validation failed.");
}

process.stdout.write("Meeting provider URL validation passed.\n");
