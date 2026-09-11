const IM_HOST = "im.kaspersky.com";
const IM_APP_PROTOCOL = "squadus";
const IM_REMINDER_GREETING = "Привет, посмотри, пожалуйста, мой PR";
export const HEXA_UI_CONTRIBUTE_LABEL = "Hexa UI Contribute";
const HEXA_UI_CONTRIBUTE_PATH = "channel/hexa-ui-contribute";

/**
 * Deep link десктопного Squadus. Схема из приложения:
 * `squadus://room?host=…&path=direct/{username}` или `path=channel/{name}`.
 *
 * @param {string} path
 */
export function buildImDesktopRoomLink(path) {
  const normalizedPath = String(path ?? "").replace(/^\/+/, "").trim();

  if (!normalizedPath) {
    throw new Error("Не передан путь чата IM.");
  }

  const params = new URLSearchParams({
    host: IM_HOST,
    path: normalizedPath,
  });

  return `${IM_APP_PROTOCOL}://room?${params.toString()}`;
}

/**
 * @param {string} username
 */
export function buildImDesktopDeepLink(username) {
  const normalizedUsername = String(username ?? "").trim();

  if (!normalizedUsername) {
    throw new Error("Не передан логин IM.");
  }

  return buildImDesktopRoomLink(`direct/${normalizedUsername}`);
}

/**
 * @param {string} path
 */
export function openImDesktopRoom(path) {
  const anchor = document.createElement("a");
  anchor.href = buildImDesktopRoomLink(path);
  anchor.rel = "noreferrer";
  document.body.append(anchor);
  anchor.click();
  anchor.remove();
}

/**
 * Открывает личку в приложении Squadus. Текст в композер нативного окна
 * из расширения подставить нельзя — черновик копируется в буфер в popup.
 *
 * @param {string} username
 */
export function openImDesktopChat(username) {
  const normalizedUsername = String(username ?? "").trim();

  if (!normalizedUsername) {
    throw new Error("Не передан логин IM.");
  }

  openImDesktopRoom(`direct/${normalizedUsername}`);
}

export function openImHexaUiContributeChannel() {
  openImDesktopRoom(HEXA_UI_CONTRIBUTE_PATH);
}

/**
 * @param {any} item
 * @returns {{ url: string, linkLabel: string, title: string, line: string }}
 */
function buildImPullRequestLine(item) {
  const url = typeof item?.url === "string" ? item.url.trim() : "";
  const title = typeof item?.title === "string" ? item.title.trim() : "";
  const id = String(item?.id ?? "").trim();
  const linkLabel = id ? `Pull Request ${id}:` : "Pull Request:";
  const line = url
    ? `[${linkLabel}](${url})${title ? ` ${title}` : ""}`
    : `${linkLabel}${title ? ` ${title}` : ""}`;

  return { url, linkLabel, title, line };
}

/**
 * @param {any} item
 * @returns {{
 *   greeting: string,
 *   url: string,
 *   linkLabel: string,
 *   title: string,
 *   text: string,
 * }}
 */
export function buildImReminderDraft(item) {
  const { url, linkLabel, title, line } = buildImPullRequestLine(item);

  return {
    greeting: IM_REMINDER_GREETING,
    url,
    linkLabel,
    title,
    text: `${IM_REMINDER_GREETING}\n${line}`.trim(),
  };
}

/**
 * Черновик в канал Hexa UI Contribute: ссылка на PR, группы без финального апрува
 * и @логины людей со статусом Waiting for the author.
 *
 * @param {any} item
 * @returns {{ text: string, url: string, linkLabel: string, title: string }}
 */
export function buildImContributeDraft(item) {
  const { url, linkLabel, title, line } = buildImPullRequestLine(item);
  const groups = Array.isArray(item?.pendingReviewerGroupNames)
    ? item.pendingReviewerGroupNames.map((name) => String(name ?? "").trim()).filter(Boolean)
    : [];
  const mentions = uniqueNonEmpty(
    Array.isArray(item?.waitingForAuthorReviewers)
      ? item.waitingForAuthorReviewers.map(formatImMentionLine)
      : [],
  );

  return {
    url,
    linkLabel,
    title,
    text: [line, ...groups, ...mentions].filter(Boolean).join("\n").trim(),
  };
}

/**
 * @param {{ imUsername?: string, displayName?: string }} person
 */
function formatImMentionLine(person) {
  const username = typeof person?.imUsername === "string" ? person.imUsername.trim() : "";

  if (username) {
    return `@${username}`;
  }

  return String(person?.displayName ?? "").trim();
}

/**
 * @param {string[]} values
 */
function uniqueNonEmpty(values) {
  const seen = new Set();
  const result = [];

  for (const value of values) {
    const normalized = String(value ?? "").trim();

    if (!normalized || seen.has(normalized)) {
      continue;
    }

    seen.add(normalized);
    result.push(normalized);
  }

  return result;
}
