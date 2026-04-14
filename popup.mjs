const STORAGE_KEY = "prState";
const ADO_CONFIG_KEY = "adoConfig";
const REFRESH_MESSAGE_TYPE = "manual-refresh";
const APPROVE_MESSAGE_TYPE = "approve-pull-request";
const DEFAULT_STATE = {
  items: [],
  count: 0,
  lastCheckedAt: null,
  lastError: null,
};

const countBadge = document.querySelector("#count-badge");
const lastUpdated = document.querySelector("#last-updated");
const messageBox = document.querySelector("#message-box");
const emptyState = document.querySelector("#empty-state");
const emptySetupHint = document.querySelector("#empty-setup-hint");
const emptySetupLink = document.querySelector("#empty-setup-link");
const itemsList = document.querySelector("#items-list");
const refreshButton = document.querySelector("#refresh-button");
const optionsLink = document.querySelector("#options-link");

let isRefreshing = false;
let transientMessage = "";
let transientMessageTone = "error";
let transientMessageTimer = null;
const approvingPullRequestIds = new Set();
let currentState = { ...DEFAULT_STATE };
let hasConfiguredGroups = false;

void init();

function applyPopupMaxHeight() {
  const availableHeight = window.screen?.availHeight || window.innerHeight || 0;

  if (!availableHeight) {
    return;
  }

  document.documentElement.style.setProperty(
    "--popup-max-height",
    `${Math.floor(availableHeight * 0.5)}px`,
  );
}

async function init() {
  applyPopupMaxHeight();
  currentState = await loadState();
  hasConfiguredGroups = await loadHasConfiguredGroups();
  render();
  refreshButton.addEventListener("click", () => {
    void refreshNow();
  });
  optionsLink?.addEventListener("click", () => {
    void chrome.runtime.openOptionsPage();
  });
  emptySetupLink?.addEventListener("click", () => {
    void chrome.runtime.openOptionsPage();
  });
  window.addEventListener("resize", applyPopupMaxHeight);

  chrome.storage.onChanged.addListener((changes, areaName) => {
    if (areaName !== "local") {
      return;
    }

    if (changes[STORAGE_KEY]) {
      currentState = {
        ...DEFAULT_STATE,
        ...(changes[STORAGE_KEY].newValue ?? {}),
      };
    }

    if (changes[ADO_CONFIG_KEY]) {
      hasConfiguredGroups = hasAnyConfiguredGroups(changes[ADO_CONFIG_KEY].newValue);
    }

    render();
  });
}

async function loadState() {
  const stored = await chrome.storage.local.get(STORAGE_KEY);

  return {
    ...DEFAULT_STATE,
    ...(stored[STORAGE_KEY] ?? {}),
  };
}

function render() {
  const hasError = !!currentState.lastError;

  if (hasError) {
    countBadge.classList.add("hidden");
  } else {
    countBadge.classList.remove("hidden");
    countBadge.textContent = String(currentState.count ?? 0);
  }

  lastUpdated.textContent = `Последняя проверка: ${formatTimestamp(currentState.lastCheckedAt)}`;
  refreshButton.disabled = isRefreshing;
  refreshButton.setAttribute(
    "aria-label",
    isRefreshing ? "Обновление выполняется" : "Обновить сейчас",
  );
  refreshButton.setAttribute(
    "title",
    isRefreshing ? "Обновление выполняется" : "Обновить сейчас",
  );

  const hasTransientMessage = Boolean(transientMessage);
  const message = transientMessage || (
    currentState.lastError
      ? `Последняя проверка завершилась ошибкой: ${currentState.lastError}`
      : ""
  );
  const messageTone = hasTransientMessage ? transientMessageTone : "error";

  if (message) {
    messageBox.textContent = message;
    messageBox.classList.toggle("popup__message--success", messageTone === "success");
    messageBox.classList.remove("hidden");
  } else {
    messageBox.textContent = "";
    messageBox.classList.remove("popup__message--success");
    messageBox.classList.add("hidden");
  }

  itemsList.textContent = "";

  if (!currentState.items.length) {
    emptyState.classList.remove("hidden");
    emptySetupHint?.classList.toggle("hidden", hasConfiguredGroups);
    return;
  }

  emptyState.classList.add("hidden");
  emptySetupHint?.classList.add("hidden");

  for (const item of currentState.items) {
    itemsList.append(createItemElement(item));
  }
}

async function loadHasConfiguredGroups() {
  const stored = await chrome.storage.local.get(ADO_CONFIG_KEY);
  return hasAnyConfiguredGroups(stored[ADO_CONFIG_KEY]);
}

function hasAnyConfiguredGroups(config) {
  if (!config || typeof config !== "object") {
    return false;
  }

  const selectedGroupIds = Array.isArray(config.selectedGroupIds)
    ? config.selectedGroupIds
    : [];

  if (selectedGroupIds.some((id) => String(id ?? "").trim())) {
    return true;
  }

  const selectedTeamIds = Array.isArray(config.selectedTeamIds)
    ? config.selectedTeamIds
    : [];

  if (selectedTeamIds.some((id) => String(id ?? "").trim())) {
    return true;
  }

  return String(config.teamReviewerIds ?? "").trim().length > 0;
}

async function refreshNow() {
  return refreshState({
    clearTransientMessage: true,
    errorPrefix: "Ручное обновление завершилось ошибкой",
  });
}

async function refreshState({ clearTransientMessage, errorPrefix }) {
  if (isRefreshing) {
    return false;
  }

  isRefreshing = true;
  if (clearTransientMessage) {
    transientMessage = "";
    transientMessageTone = "error";
  }
  render();

  try {
    const response = await chrome.runtime.sendMessage({
      type: REFRESH_MESSAGE_TYPE,
    });

    if (!response?.ok) {
      throw new Error(response?.error || "Не удалось выполнить ручное обновление.");
    }

    return true;
  } catch (error) {
    showTransientMessage(
      `${errorPrefix}: ${error instanceof Error ? error.message : String(error)}`,
      "error",
    );
    return false;
  } finally {
    isRefreshing = false;
    currentState = await loadState();
    render();
  }
}

function createItemElement(item) {
  const listItem = document.createElement("li");
  listItem.className = "popup__item";
  const isTechPR = isTechPullRequest(item.description);

  const stale = isStalePullRequest(item, currentState.lastCheckedAt);

  const itemMain = document.createElement("div");
  itemMain.className = "popup__item-main";

  if (item.avatarUrl) {
    const avatar = document.createElement("img");
    avatar.className = "popup__avatar";
    avatar.src = item.avatarUrl;
    avatar.alt = item.author || "Автор";
    avatar.loading = "lazy";
    itemMain.append(avatar);
  }

  const itemContent = document.createElement("div");
  itemContent.className = "popup__item-content";

  const itemHeader = document.createElement("div");
  itemHeader.className = "popup__item-header";

  const button = document.createElement("button");
  button.className = "popup__link";
  button.type = "button";
  button.textContent = item.title;
  button.addEventListener("click", async () => {
    await chrome.tabs.create({ url: item.url });
    window.close();
  });

  const authorRow = document.createElement("div");
  authorRow.className = "popup__author-row";

  const author = document.createElement("p");
  author.className = "popup__author";
  fillAuthorMetaParagraph(author, item, currentState.lastCheckedAt, stale);

  authorRow.append(author);

  /** @type {{ icon: HTMLElement, section: HTMLElement } | null} */
  let descriptionUi = null;

  if (item.description) {
    descriptionUi = createDescriptionBlock(item.description, item.id);
    authorRow.append(descriptionUi.icon);
  }

  if (isTechPR) {
    const badge = document.createElement("span");
    badge.className = "popup__badge";
    badge.textContent = "ТЕХ ПР";
    authorRow.append(badge);
  }

  if (isTechPR) {
    itemHeader.append(button, createApproveButton(item));
  } else {
    itemHeader.append(button);
  }

  itemContent.append(itemHeader, authorRow);

  if (descriptionUi) {
    itemContent.append(descriptionUi.section);
  }

  itemMain.append(itemContent);
  listItem.append(itemMain);

  return listItem;
}

const DESC_ICON_SVG = `<svg viewBox="0 0 16 16" xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="none" aria-hidden="true">
  <path d="M1.5 3L1.5 13L6.293 13L8 14.707L9.707 13L14.5 13L14.5 3L1.5 3ZM2.5 4L13.5 4L13.5 12L9.293 12L8 13.293L6.707 12L2.5 12L2.5 4ZM4.5 5.5L4.5 6.5L11.5 6.5L11.5 5.5L4.5 5.5ZM4.5 7.5L4.5 8.5L11.5 8.5L11.5 7.5L4.5 7.5ZM4.5 9.5L4.5 10.5L9.5 10.5L9.5 9.5L4.5 9.5Z" fill="currentColor" fill-rule="nonzero" />
</svg>`;

/**
 * Иконка в строке автора (не &lt;button&gt;), панель ниже.
 *
 * @param {string} description
 * @param {string | number} itemId
 * @returns {{ icon: HTMLSpanElement, section: HTMLDivElement }}
 */
function createDescriptionBlock(description, itemId) {
  const section = document.createElement("div");
  section.className = "popup__item-desc";

  const panel = document.createElement("div");
  panel.className = "popup__description-panel";
  panel.hidden = true;
  panel.setAttribute("role", "region");
  panel.id = `pr-desc-${String(itemId).replace(/[^\w-]/g, "_")}`;

  const inner = document.createElement("div");
  inner.className = "popup__description-panel-inner popup__markdown";
  inner.innerHTML = renderMarkdown(description);

  panel.append(inner);
  section.append(panel);

  const icon = document.createElement("span");
  icon.className = "popup__info";
  icon.setAttribute("role", "button");
  icon.tabIndex = 0;
  icon.setAttribute("aria-expanded", "false");
  icon.setAttribute("aria-controls", panel.id);
  icon.setAttribute("aria-label", "Показать или скрыть описание PR");
  icon.innerHTML = DESC_ICON_SVG;

  const toggle = () => {
    const open = panel.hidden;
    panel.hidden = !open;
    icon.setAttribute("aria-expanded", open ? "true" : "false");
    section.classList.toggle("popup__item-desc--open", open);
  };

  icon.addEventListener("click", (event) => {
    event.preventDefault();
    toggle();
  });

  icon.addEventListener("keydown", (event) => {
    if (event.key === "Enter" || event.key === " ") {
      event.preventDefault();
      toggle();
    }
  });

  return { icon, section };
}

function createApproveButton(item) {
  const approveButton = document.createElement("button");
  const isApproving = approvingPullRequestIds.has(item.id);
  const label = isApproving ? `Approve выполняется для PR #${item.id}` : `Approve PR #${item.id}`;

  approveButton.className = "popup__approve";
  approveButton.type = "button";
  approveButton.disabled = isApproving;
  approveButton.setAttribute("aria-label", label);
  approveButton.setAttribute("title", label);
  approveButton.append(createApproveIcon());
  approveButton.addEventListener("click", () => {
    void approvePullRequest(item);
  });

  return approveButton;
}

function createApproveIcon() {
  const icon = document.createElementNS("http://www.w3.org/2000/svg", "svg");
  icon.classList.add("popup__approve-icon");
  icon.setAttribute("viewBox", "0 0 16 16");
  icon.setAttribute("aria-hidden", "true");

  const path = document.createElementNS("http://www.w3.org/2000/svg", "path");
  path.setAttribute("d", "M3.5 8.5 6.5 11.5 12.5 4.5");

  icon.append(path);
  return icon;
}

async function approvePullRequest(item) {
  if (!item?.id || approvingPullRequestIds.has(item.id)) {
    return;
  }

  approvingPullRequestIds.add(item.id);
  transientMessage = "";
  transientMessageTone = "error";
  render();

  try {
    const response = await chrome.runtime.sendMessage({
      type: APPROVE_MESSAGE_TYPE,
      pullRequestId: item.id,
    });

    if (!response?.ok) {
      throw new Error(response?.error || "Не удалось лайкнуть PR ((.");
    }

    showTransientMessage(`PR #${item.id} одобрен. Обновляю список…`, "success", 2000);

    const refreshOk = await refreshState({
      clearTransientMessage: false,
      errorPrefix: `PR #${item.id} одобрен, но обновление списка завершилось ошибкой`,
    });

    if (refreshOk) {
      showTransientMessage(`PR #${item.id} одобрен.`, "success", 2000);
    }
  } catch (error) {
    showTransientMessage(
      `Approve завершился ошибкой для PR #${item.id}: ${
        error instanceof Error ? error.message : String(error)
      }`,
      "error",
    );
  } finally {
    approvingPullRequestIds.delete(item.id);
    currentState = await loadState();
    render();
  }
}

function showTransientMessage(message, tone = "error", autoHideMs = 0) {
  transientMessage = String(message ?? "");
  transientMessageTone = tone === "success" ? "success" : "error";

  if (transientMessageTimer) {
    clearTimeout(transientMessageTimer);
    transientMessageTimer = null;
  }

  if (autoHideMs > 0) {
    transientMessageTimer = setTimeout(() => {
      transientMessage = "";
      transientMessageTone = "error";
      transientMessageTimer = null;
      render();
    }, autoHideMs);
  }

  render();
}

function formatTimestamp(timestamp) {
  if (!timestamp) {
    return "ещё не выполнялась";
  }

  const date = new Date(timestamp);

  return new Intl.DateTimeFormat("ru-RU", {
    dateStyle: "short",
    timeStyle: "short",
  }).format(date);
}

/**
 * @param {HTMLParagraphElement} el
 * @param {boolean} stale
 */
function fillAuthorMetaParagraph(el, item, checkedAt, stale) {
  const authorName = item.author || "Автор не определён";
  const relativeTime = formatElapsedSince(item.updatedAt ?? item.createdAt, checkedAt);

  el.replaceChildren();

  if (!relativeTime) {
    el.textContent = authorName;
    return;
  }

  const nameSpan = document.createElement("span");
  nameSpan.className = "popup__author-name";
  nameSpan.textContent = authorName;

  const sep = document.createElement("span");
  sep.className = "popup__author-sep";
  sep.textContent = " · ";

  const timeSpan = document.createElement("span");
  timeSpan.className = "popup__author-time";
  if (stale) {
    timeSpan.classList.add("popup__author-time--stale");
  }
  timeSpan.textContent = relativeTime;

  el.append(nameSpan, sep, timeSpan);
}

function formatElapsedSince(createdAt, checkedAt) {
  const totalMinutes = getElapsedMinutes(createdAt, checkedAt);

  if (totalMinutes === null) {
    return "";
  }

  if (totalMinutes < 60) {
    return `${totalMinutes} мин`;
  }

  const hours = Math.floor(totalMinutes / 60);
  const minutes = totalMinutes % 60;

  if (totalMinutes >= 10 * 60) {
    return `${hours} ч`;
  }

  return `${hours} ч ${minutes} мин`;
}

function isStalePullRequest(item, checkedAt) {
  const totalMinutes = getElapsedMinutes(item.updatedAt ?? item.createdAt, checkedAt);

  if (totalMinutes === null) {
    return false;
  }

  return totalMinutes > 48 * 60;
}

function getElapsedMinutes(createdAt, checkedAt) {
  if (!createdAt || !checkedAt) {
    return null;
  }

  const createdAtDate = new Date(createdAt);
  const checkedAtDate = new Date(checkedAt);

  if (Number.isNaN(createdAtDate.getTime()) || Number.isNaN(checkedAtDate.getTime())) {
    return null;
  }

  const diffMs = checkedAtDate.getTime() - createdAtDate.getTime();

  if (diffMs < 0) {
    return null;
  }

  return Math.floor(diffMs / 60000);
}

function escapeHtml(s) {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function decodeBasicEntities(s) {
  return String(s)
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"');
}

/**
 * Только http(s); иначе ссылка не создаётся.
 *
 * @param {string} href
 */
function sanitizeMarkdownUrl(href) {
  const raw = decodeBasicEntities(String(href).trim());

  if (!/^https?:\/\//i.test(raw)) {
    return null;
  }

  try {
    const parsed = new URL(raw);

    if (parsed.protocol !== "http:" && parsed.protocol !== "https:") {
      return null;
    }

    return parsed.href;
  } catch (_error) {
    return null;
  }
}

/**
 * Текст внутри пары скобок, начиная с `(` на индексе openIdx (глубина по вложенным `(` / `)`).
 *
 * @param {string} text
 * @param {number} openIdx
 * @returns {{ content: string, nextIndex: number } | null}
 */
function extractBalancedParenContent(text, openIdx) {
  if (String(text[openIdx] ?? "") !== "(") {
    return null;
  }

  let depth = 1;
  let i = openIdx + 1;

  while (i < text.length && depth > 0) {
    const c = text[i];
    if (c === "(") {
      depth += 1;
    } else if (c === ")") {
      depth -= 1;
    }
    i += 1;
  }

  if (depth !== 0) {
    return null;
  }

  return {
    content: text.slice(openIdx + 1, i - 1),
    nextIndex: i,
  };
}

/**
 * Склеивает переносы и лишние пробелы внутри markdown-ссылок и картинок `[]()` / `![]()`,
 * чтобы `text.split(/\n\n+/)` не рвал URL и чтобы `)` корректно закрывала адрес.
 *
 * @param {string} raw
 */
function preprocessMarkdown(raw) {
  let t = String(raw ?? "");
  t = t.replace(/\]\s*\r?\n\s*\(/g, "](");

  t = t
    .split("\n")
    .map((line) => {
      const trimmed = line.trim();

      if (!trimmed.startsWith("(")) {
        return line;
      }

      const ext = extractBalancedParenContent(trimmed, 0);

      if (!ext) {
        return line;
      }

      const after = trimmed.slice(ext.nextIndex).trim();

      if (after !== "") {
        return line;
      }

      const inner = ext.content.trim();

      if (!/^https?:\/\//i.test(inner)) {
        return line;
      }

      const s = sanitizeMarkdownUrl(inner);

      if (!s) {
        return line;
      }

      return `[${inner}](${s})`;
    })
    .join("\n");

  const collapseResourceParens = (input) => {
    let out = input;
    let changed = true;

    while (changed) {
      changed = false;
      const re = /!?\[[^\]]*\]\(/g;
      let m;

      while ((m = re.exec(out)) !== null) {
        const openIdx = m.index + m[0].length - 1;
        const ext = extractBalancedParenContent(out, openIdx);

        if (!ext) {
          continue;
        }

        const collapsed = ext.content.replace(/\r?\n/g, "").replace(/\s{2,}/g, " ").trim();

        if (collapsed === ext.content || !/^https?:\/\//i.test(collapsed)) {
          continue;
        }

        out = `${out.slice(0, m.index)}${m[0]}${collapsed})${out.slice(ext.nextIndex)}`;
        changed = true;
        break;
      }
    }

    return out;
  };

  let prev;
  do {
    prev = t;
    t = collapseResourceParens(t);
  } while (t !== prev);

  return t;
}

function replaceMarkdownImages(t) {
  let out = "";
  let cur = 0;

  while (cur < t.length) {
    const i = t.indexOf("![", cur);

    if (i === -1) {
      out += t.slice(cur);
      break;
    }

    out += t.slice(cur, i);

    const closeBracket = t.indexOf("]", i + 2);

    if (closeBracket === -1) {
      out += "![";
      cur = i + 2;
      continue;
    }

    if (t.slice(closeBracket, closeBracket + 2) !== "](") {
      out += "![";
      cur = i + 2;
      continue;
    }

    const alt = t.slice(i + 2, closeBracket);
    const openParenIdx = closeBracket + 1;
    const ext = extractBalancedParenContent(t, openParenIdx);

    if (!ext) {
      out += "![";
      cur = i + 2;
      continue;
    }

    const url = sanitizeMarkdownUrl(ext.content);

    if (!url) {
      out += t.slice(i, ext.nextIndex);
    } else {
      const altText = alt.trim() || "Изображение";
      out += `<img class="popup__md-img" src="${escapeHtml(url)}" alt="${escapeHtml(altText)}" loading="lazy" referrerpolicy="no-referrer">`;
    }

    cur = ext.nextIndex;
  }

  return out;
}

function replaceMarkdownLinks(t) {
  let out = "";
  let cur = 0;

  while (cur < t.length) {
    const i = t.indexOf("[", cur);

    if (i === -1) {
      out += t.slice(cur);
      break;
    }

    if (t[i + 1] === "[") {
      out += t.slice(cur, i + 1);
      cur = i + 1;
      continue;
    }

    if (i > 0 && t[i - 1] === "!") {
      out += t.slice(cur, i + 1);
      cur = i + 1;
      continue;
    }

    const closeBracket = t.indexOf("]", i + 1);

    if (closeBracket === -1) {
      out += t.slice(cur, i + 1);
      cur = i + 1;
      continue;
    }

    if (t.slice(closeBracket, closeBracket + 2) !== "](") {
      out += t.slice(cur, i + 1);
      cur = i + 1;
      continue;
    }

    const label = t.slice(i + 1, closeBracket);

    if (!label.trim()) {
      out += t.slice(cur, i + 1);
      cur = i + 1;
      continue;
    }

    const openParenIdx = closeBracket + 1;
    const ext = extractBalancedParenContent(t, openParenIdx);

    if (!ext) {
      out += t.slice(cur, i + 1);
      cur = i + 1;
      continue;
    }

    const url = sanitizeMarkdownUrl(ext.content);

    out += t.slice(cur, i);

    if (!url) {
      out += t.slice(i, ext.nextIndex);
    } else {
      out += `<a href="${escapeHtml(url)}" target="_blank" rel="noopener noreferrer">${label}</a>`;
    }

    cur = ext.nextIndex;
  }

  return out;
}

/**
 * @param {string} s
 */
function applyInlineMarkdown(s) {
  let t = String(s);

  t = replaceMarkdownImages(t);
  t = replaceMarkdownLinks(t);

  t = t.replace(/`([^`]+)`/g, "<code>$1</code>");
  t = t.replace(/\*\*([^*]+)\*\*/g, "<strong>$1</strong>");
  t = t.replace(/\*([^*]+)\*/g, "<em>$1</em>");
  return t;
}

/**
 * @param {string | undefined | null} raw
 */
function renderMarkdown(raw) {
  const text = escapeHtml(preprocessMarkdown(raw ?? ""));
  const blocks = text.split(/\n\n+/);
  const out = [];

  for (const block of blocks) {
    const b = block.trim();

    if (!b) {
      continue;
    }

    if (/^(-{3,}|\*{3,})$/.test(b)) {
      out.push("<hr class=\"popup__md-hr\">");
      continue;
    }

    const lines = b.split("\n");
    const first = lines[0] ?? "";

    if (/^#{1,6}\s+/.test(first)) {
      const level = first.match(/^#+/)?.[0].length ?? 1;
      const content = first.replace(/^#{1,6}\s+/, "");
      const tag = level <= 2 ? "h3" : "h4";
      out.push(`<${tag} class="popup__md-heading">${applyInlineMarkdown(content)}</${tag}>`);

      if (lines.length > 1) {
        const rest = lines.slice(1).join("\n");
        out.push(
          `<p class="popup__md-p">${applyInlineMarkdown(rest).replace(/\n/g, "<br>")}</p>`,
        );
      }

      continue;
    }

    const listLines = lines.filter((ln) => ln.trim() !== "");

    if (
      listLines.length > 0
      && listLines.every((ln) => /^[-*]\s+/.test(ln.trim()))
    ) {
      const items = listLines.map((ln) => {
        const body = ln.trim().replace(/^[-*]\s+/, "");
        return `<li class="popup__md-li">${applyInlineMarkdown(body)}</li>`;
      });
      out.push(`<ul class="popup__md-ul">${items.join("")}</ul>`);
      continue;
    }

    const numbered = listLines.every((ln) => /^\d+\.\s+/.test(ln.trim()));

    if (listLines.length > 0 && numbered) {
      const items = listLines.map((ln) => {
        const body = ln.trim().replace(/^\d+\.\s+/, "");
        return `<li class="popup__md-li">${applyInlineMarkdown(body)}</li>`;
      });
      out.push(`<ol class="popup__md-ol">${items.join("")}</ol>`);
      continue;
    }

    out.push(`<p class="popup__md-p">${applyInlineMarkdown(b).replace(/\n/g, "<br>")}</p>`);
  }

  return out.join("");
}

function isTechPullRequest(description) {
  if (typeof description !== "string") {
    return false;
  }

  return /те[хx]\s*п[рp]/i.test(description);
}
