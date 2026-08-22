const PROVIDERS = Object.freeze([
  Object.freeze({
    id: "google-meet",
    label: "Google Meet",
    hostname: "meet.google.com",
    joinPath: /^\/[a-z]{3}-[a-z]{4}-[a-z]{3}\/?$/i,
    example: "https://meet.google.com/xxx-xxxx-xxx",
  }),
  Object.freeze({
    id: "microsoft-teams",
    label: "Microsoft Teams Web",
    hostname: "teams.microsoft.com",
    joinPath: /^(?:\/(?:v2\/)?l\/meetup-join\/[^/]+(?:\/\d+)?\/?|\/meet\/\d+(?:-\d+)*\/?)$/i,
    example: "https://teams.microsoft.com/l/meetup-join/...",
  }),
]);

export const MEETING_PROVIDERS = Object.freeze(
  Object.fromEntries(PROVIDERS.map((provider) => [provider.id, provider])),
);

function safeUrl(value) {
  try {
    return new URL(String(value || "").trim());
  } catch {
    return null;
  }
}

function hasSafeAuthority(url) {
  return (
    url.protocol === "https:" &&
    !url.port &&
    !url.username &&
    !url.password
  );
}

export function meetingProviderForUrl(value) {
  const url = safeUrl(value);
  if (!url || !hasSafeAuthority(url)) {
    return null;
  }
  return PROVIDERS.find(
    (provider) =>
      url.hostname === provider.hostname && provider.joinPath.test(url.pathname),
  ) || null;
}

export function normalizeMeetingJoinUrl(value, { allowedProviders } = {}) {
  const url = safeUrl(value);
  const provider = meetingProviderForUrl(value);
  const allowed = allowedProviders ? new Set(allowedProviders) : null;
  if (!url || !provider || (allowed && !allowed.has(provider.id))) {
    const examples = PROVIDERS
      .filter((candidate) => !allowed || allowed.has(candidate.id))
      .map((candidate) => candidate.example)
      .join(" または ");
    throw new Error(`${examples} 形式の会議URLを入力してください`);
  }
  url.hash = "";
  return {
    provider: provider.id,
    providerLabel: provider.label,
    url: url.toString(),
  };
}

export function isMeetingProviderPage(value, providerId) {
  const url = safeUrl(value);
  const provider = MEETING_PROVIDERS[providerId];
  return Boolean(
    url && provider && hasSafeAuthority(url) && url.hostname === provider.hostname,
  );
}
