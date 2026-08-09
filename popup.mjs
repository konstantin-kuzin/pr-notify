import {
  BADGE_STYLES,
  getItemWorkingTimeFrom,
  getItemWorkingTimeUrgency,
  getWorkingElapsedMinutes,
  hasUpdatesAfterLastGroupComment,
  sortPullRequestsOldestFirst,
} from "./working-time.mjs";

const STORAGE_KEY = "prState";
const UPDATE_STATE_KEY = "prUpdateState";
const ADO_CONFIG_KEY = "adoConfig";
const ACTIVE_TAB_KEY = "popupActiveTab";
const REFRESH_MESSAGE_TYPE = "manual-refresh";
const APPROVE_MESSAGE_TYPE = "approve-pull-request";
const LOAD_MY_COMPLETED_MESSAGE_TYPE = "load-my-completed-pull-requests";
const MY_COMPLETED_PAGE_SIZE = 20;
const TAB_REVIEW = "review";
const TAB_MY = "my";
const STORYBOOK_BASE_URL = "https://storybook.s1.ksc-web.avp.ru/hexa-ui";
const STORYBOOK_EXPIRED_REASON_RE = /\[OSMP\]\s*Storybook Hexa UI deploy for Review expired/i;
const DEFAULT_STATE = {
  items: [],
  count: 0,
  myItems: [],
  myCount: 0,
  lastCheckedAt: null,
  lastError: null,
};

const countBadge = document.querySelector("#count-badge");
const lastUpdated = document.querySelector("#last-updated");
const messageBox = document.querySelector("#message-box");
const emptyState = document.querySelector("#empty-state");
const emptyStateText = document.querySelector("#empty-state-text");
const emptySetupHint = document.querySelector("#empty-setup-hint");
const emptySetupLink = document.querySelector("#empty-setup-link");
const itemsList = document.querySelector("#items-list");
const refreshButton = document.querySelector("#refresh-button");
const optionsLink = document.querySelector("#options-link");
const updateChip = document.querySelector("#update-chip");
const tabReview = document.querySelector("#tab-review");
const tabMy = document.querySelector("#tab-my");

let isRefreshing = false;
let transientMessage = "";
let transientMessageTone = "error";
let transientMessageTimer = null;
const approvingPullRequestIds = new Set();
let currentState = { ...DEFAULT_STATE };
let updateState = { hasUpdate: false, latestVersion: "" };
let hasConfiguredGroups = false;
/** @type {"review" | "my"} */
let activeTab = TAB_REVIEW;
/** @type {any[]} */
let myCompletedItems = [];
let myCompletedHasMore = false;
let myCompletedLoaded = false;
let isLoadingMyCompleted = false;
let myCompletedError = "";

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
  updateState = await loadUpdateState();
  hasConfiguredGroups = await loadHasConfiguredGroups();
  activeTab = await loadActiveTab();
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
  tabReview?.addEventListener("click", () => {
    void setActiveTab(TAB_REVIEW);
  });
  tabMy?.addEventListener("click", () => {
    void setActiveTab(TAB_MY);
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
      resetMyCompletedState();
    }

    if (changes[ADO_CONFIG_KEY]) {
      hasConfiguredGroups = hasAnyConfiguredGroups(changes[ADO_CONFIG_KEY].newValue);
    }

    if (changes[UPDATE_STATE_KEY]) {
      updateState = normalizeUpdateState(changes[UPDATE_STATE_KEY].newValue);
      renderUpdateChip();
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

async function loadUpdateState() {
  const stored = await chrome.storage.local.get(UPDATE_STATE_KEY);
  return normalizeUpdateState(stored[UPDATE_STATE_KEY]);
}

function normalizeUpdateState(rawState) {
  return {
    hasUpdate: Boolean(rawState?.hasUpdate),
    latestVersion:
      typeof rawState?.latestVersion === "string" ? rawState.latestVersion : "",
  };
}

function render() {
  const hasError = !!currentState.lastError;
  const isMyTab = activeTab === TAB_MY;
  const visibleItems = isMyTab
    ? (Array.isArray(currentState.myItems) ? currentState.myItems : [])
    : (Array.isArray(currentState.items) ? currentState.items : []);
  const visibleCount = isMyTab
    ? (currentState.myCount ?? visibleItems.length)
    : (currentState.count ?? visibleItems.length);
  const completedItems = isMyTab ? myCompletedItems : [];
  const hasListContent = visibleItems.length > 0 || completedItems.length > 0;
  const waitingCompleted = isMyTab
    && !hasListContent
    && (isLoadingMyCompleted || !myCompletedLoaded);

  renderTabs();

  if (isMyTab) {
    void ensureMyCompletedLoaded();
  }

  if (hasError) {
    countBadge.classList.add("hidden");
  } else {
    countBadge.classList.remove("hidden");
    countBadge.textContent = String(visibleCount ?? 0);
    applyCountBadgeStyle("gray");
  }

  lastUpdated.textContent = formatLastCheckedAt(currentState.lastCheckedAt);
  refreshButton.disabled = isRefreshing;
  refreshButton.setAttribute(
    "aria-label",
    isRefreshing ? "Обновление выполняется" : "Обновить сейчас",
  );
  refreshButton.setAttribute(
    "title",
    isRefreshing ? "Обновление выполняется" : "Обновить сейчас",
  );
  renderUpdateChip();

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

  if (!hasListContent && !waitingCompleted) {
    emptyState.classList.remove("hidden");
    if (emptyStateText) {
      emptyStateText.textContent = isMyTab
        ? (myCompletedError || "Нет ваших активных pull requests")
        : "Нет pull requests для ревью";
    }
    emptySetupHint?.classList.toggle("hidden", isMyTab || hasConfiguredGroups);
    return;
  }

  emptyState.classList.add("hidden");
  emptySetupHint?.classList.add("hidden");

  const orderedItems = isMyTab
    ? visibleItems
    : sortPullRequestsOldestFirst(visibleItems);

  for (const item of orderedItems) {
    itemsList.append(createItemElement(item, { mode: activeTab }));
  }

  if (!isMyTab) {
    return;
  }

  for (const item of completedItems) {
    itemsList.append(createItemElement(item, { mode: TAB_MY }));
  }

  if (waitingCompleted || (isLoadingMyCompleted && completedItems.length === 0)) {
    itemsList.append(createMyCompletedStatusRow("Загрузка…"));
  } else if (myCompletedError && completedItems.length === 0) {
    itemsList.append(createMyCompletedStatusRow(myCompletedError, { isError: true }));
  } else if (myCompletedHasMore) {
    itemsList.append(createShowMoreCompletedButton());
  }
}

function resetMyCompletedState() {
  myCompletedItems = [];
  myCompletedHasMore = false;
  myCompletedLoaded = false;
  isLoadingMyCompleted = false;
  myCompletedError = "";
}

async function ensureMyCompletedLoaded() {
  if (myCompletedLoaded || isLoadingMyCompleted) {
    return;
  }

  await loadMoreMyCompleted({ reset: true });
}

/**
 * @param {{ reset?: boolean }} [options]
 */
async function loadMoreMyCompleted(options = {}) {
  if (isLoadingMyCompleted) {
    return;
  }

  const reset = options.reset === true;
  isLoadingMyCompleted = true;
  myCompletedError = "";
  render();

  try {
    const skip = reset ? 0 : myCompletedItems.length;
    const response = await chrome.runtime.sendMessage({
      type: LOAD_MY_COMPLETED_MESSAGE_TYPE,
      skip,
      top: MY_COMPLETED_PAGE_SIZE,
    });

    if (!response?.ok) {
      throw new Error(response?.error || "Не удалось загрузить завершённые pull requests.");
    }

    const nextItems = Array.isArray(response.items) ? response.items : [];
    myCompletedItems = reset ? nextItems : [...myCompletedItems, ...nextItems];
    myCompletedHasMore = Boolean(response.hasMore);
    myCompletedLoaded = true;
  } catch (error) {
    myCompletedError = error instanceof Error ? error.message : String(error);
    myCompletedLoaded = true;
    myCompletedHasMore = false;
  } finally {
    isLoadingMyCompleted = false;
    render();
  }
}

function createShowMoreCompletedButton() {
  const row = document.createElement("li");
  row.className = "popup__load-more";

  const button = document.createElement("button");
  button.className = "popup__load-more-button";
  button.type = "button";
  button.textContent = "Показать ещё";
  button.disabled = isLoadingMyCompleted;
  button.addEventListener("click", () => {
    void loadMoreMyCompleted();
  });

  row.append(button);
  return row;
}

/**
 * @param {string} text
 * @param {{ isError?: boolean }} [options]
 */
function createMyCompletedStatusRow(text, options = {}) {
  const row = document.createElement("li");
  row.className = "popup__load-more";

  const status = document.createElement("p");
  status.className = options.isError
    ? "popup__load-more-status popup__load-more-status--error"
    : "popup__load-more-status";
  status.textContent = text;
  row.append(status);
  return row;
}

function renderTabs() {
  const isMyTab = activeTab === TAB_MY;

  tabReview?.classList.toggle("popup__tab--active", !isMyTab);
  tabMy?.classList.toggle("popup__tab--active", isMyTab);
  tabReview?.setAttribute("aria-selected", isMyTab ? "false" : "true");
  tabMy?.setAttribute("aria-selected", isMyTab ? "true" : "false");
}

async function loadActiveTab() {
  try {
    const stored = await chrome.storage.session?.get?.(ACTIVE_TAB_KEY);
    return normalizeTab(stored?.[ACTIVE_TAB_KEY]);
  } catch (_error) {
    return TAB_REVIEW;
  }
}

async function setActiveTab(tab) {
  activeTab = normalizeTab(tab);

  try {
    await chrome.storage.session?.set?.({
      [ACTIVE_TAB_KEY]: activeTab,
    });
  } catch (_error) {
    // session storage может быть недоступен — вкладка живёт до закрытия popup
  }

  render();
}

function normalizeTab(value) {
  return value === TAB_MY ? TAB_MY : TAB_REVIEW;
}

function renderUpdateChip() {
  if (!updateChip) {
    return;
  }

  const latestVersion = updateState.latestVersion.trim();
  const isVisible = updateState.hasUpdate && latestVersion;

  updateChip.classList.toggle("hidden", !isVisible);
  updateChip.textContent = isVisible ? `Новая версия - ${latestVersion}` : "";
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

  return selectedGroupIds.some((id) => String(id ?? "").trim());
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

/**
 * @param {any} item
 * @param {{ mode?: "review" | "my" }} [options]
 */
function createItemElement(item, options = {}) {
  const mode = options.mode === TAB_MY ? TAB_MY : TAB_REVIEW;
  const listItem = document.createElement("li");
  listItem.className = "popup__item";
  const isCompleted = mode === TAB_MY && item?.status === "completed";
  const isTechPR = mode === TAB_REVIEW && isTechPullRequest(item.description);
  const hasNoUpdates = mode === TAB_REVIEW && !hasUpdatesAfterLastGroupComment(item);

  if (hasNoUpdates) {
    listItem.classList.add("popup__item--no-updates");
  }

  const timeUrgency = mode === TAB_REVIEW
    ? getItemWorkingTimeUrgency(item, currentState.lastCheckedAt)
    : null;

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

  if (mode === TAB_MY) {
    fillMyMetaParagraph(author, item);
  } else {
    fillAuthorMetaParagraph(author, item, currentState.lastCheckedAt, timeUrgency);
  }

  authorRow.append(author);

  if (item?.id && shouldShowStorybookLink(item)) {
    authorRow.append(createStorybookButton(item));
  }

  /** @type {{ icon: HTMLElement, section: HTMLElement } | null} */
  let descriptionUi = null;

  if (item.description) {
    descriptionUi = createDescriptionBlock(item.description, item.id);
    authorRow.append(descriptionUi.icon);
  }

  /** @type {{ icon: HTMLElement, section: HTMLElement } | null} */
  let blockersUi = null;

  if (mode === TAB_REVIEW) {
    blockersUi = createBlockingReasonsToggle(item);
    if (blockersUi) {
      authorRow.append(blockersUi.icon);
    }
  }

  /** @type {{ icon: HTMLElement, section: HTMLElement } | null} */
  const conflictUi = isCompleted ? null : createConflictToggle(item);
  if (conflictUi) {
    authorRow.append(conflictUi.icon);
  }

  if (isCompleted) {
    const badge = document.createElement("span");
    badge.className = "popup__badge popup__badge--complete";
    badge.textContent = "Complete";
    authorRow.append(badge);
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

  if (mode === TAB_MY && !isCompleted) {
    itemContent.append(createBlockingReasonsBlock(item));
  }

  if (blockersUi) {
    itemContent.append(blockersUi.section);
  }

  if (conflictUi) {
    itemContent.append(conflictUi.section);
  }

  if (descriptionUi) {
    itemContent.append(descriptionUi.section);
  }

  itemMain.append(itemContent);
  listItem.append(itemMain);

  return listItem;
}

/**
 * @param {HTMLParagraphElement} el
 * @param {any} item
 */
function fillMyMetaParagraph(el, item) {
  el.replaceChildren();

  const branch = typeof item.targetBranch === "string" ? item.targetBranch.trim() : "";
  const dateSource = item?.status === "completed" && item.closedAt
    ? item.closedAt
    : item.createdAt;
  const createdLabel = dateSource ? formatTimestamp(dateSource) : "";

  if (branch) {
    const branchSpan = document.createElement("span");
    branchSpan.className = "popup__author-name";
    branchSpan.textContent = `→ ${branch}`;
    el.append(branchSpan);
  }

  if (createdLabel) {
    if (el.childNodes.length > 0) {
      const sep = document.createElement("span");
      sep.className = "popup__author-sep";
      sep.textContent = " · ";
      el.append(sep);
    }

    const timeSpan = document.createElement("span");
    timeSpan.className = "popup__author-time";
    timeSpan.textContent = createdLabel;
    el.append(timeSpan);
  }

  if (!el.childNodes.length) {
    el.textContent = item.author || "Ваш pull request";
  }
}

/**
 * @param {any} item
 * @returns {{ required: string[], optional: string[] }}
 */
function getPolicyReasonLists(item) {
  const required = Array.isArray(item?.blockingReasons)
    ? item.blockingReasons.map((reason) => String(reason ?? "").trim()).filter(Boolean)
    : [];
  const optional = Array.isArray(item?.optionalPolicyReasons)
    ? item.optionalPolicyReasons.map((reason) => String(reason ?? "").trim()).filter(Boolean)
    : [];

  return { required, optional };
}

/**
 * @param {any} item
 */
function createBlockingReasonsBlock(item) {
  if (item.blockingReasons === null) {
    const failed = document.createElement("p");
    failed.className = "popup__blocker";
    failed.textContent = "Не удалось загрузить политики Complete";
    return failed;
  }

  const { required, optional } = getPolicyReasonLists(item);
  const wrap = document.createElement("div");
  wrap.className = "popup__policies";

  if (required.length === 0) {
    const ready = document.createElement("p");
    ready.className = "popup__ready";
    ready.textContent = "Готов к Complete";
    wrap.append(ready);
  } else {
    wrap.append(createBlockingReasonsList(required));
  }

  if (optional.length > 0) {
    const heading = document.createElement("p");
    heading.className = "popup__policies-heading";
    heading.textContent = "Optional";
    wrap.append(heading, createBlockingReasonsList(optional, { optional: true }));
  }

  return wrap;
}

/**
 * @param {string[]} reasons
 * @param {{ optional?: boolean }} [options]
 */
function createBlockingReasonsList(reasons, options = {}) {
  const list = document.createElement("ul");
  list.className = options.optional
    ? "popup__blockers popup__blockers--optional"
    : "popup__blockers";
  list.setAttribute(
    "aria-label",
    options.optional
      ? "Optional policies"
      : "Причины, блокирующие Complete",
  );

  for (const reason of reasons) {
    const row = document.createElement("li");
    row.className = "popup__blocker";

    const mark = document.createElement("span");
    mark.className = "popup__blocker-mark";
    mark.setAttribute("aria-hidden", "true");
    mark.textContent = "×";

    const text = document.createElement("span");
    text.className = "popup__blocker-text";
    text.textContent = reason;

    row.append(mark, text);
    list.append(row);
  }

  return list;
}

/**
 * @param {string[]} required
 * @param {string[]} optional
 */
function createPolicyReasonsBlock(required, optional) {
  const wrap = document.createElement("div");
  wrap.className = "popup__policies";

  if (required.length > 0) {
    wrap.append(createBlockingReasonsList(required));
  }

  if (optional.length > 0) {
    const heading = document.createElement("p");
    heading.className = "popup__policies-heading";
    heading.textContent = "Optional";
    wrap.append(heading, createBlockingReasonsList(optional, { optional: true }));
  }

  return wrap;
}

/** Набор причин, при котором бейдж показывает «Votes check» вместо числа. */
const VOTES_CHECK_ONLY_REASONS = new Set([
  "0 of 1 reviewers approved",
  "Votes check",
  "Required reviewers have not approved",
]);

/**
 * @param {string[]} reasons
 */
function isVotesCheckOnlyReasons(reasons) {
  if (reasons.length !== VOTES_CHECK_ONLY_REASONS.size) {
    return false;
  }

  return reasons.every((reason) => VOTES_CHECK_ONLY_REASONS.has(reason));
}

/**
 * Бейдж с числом замечаний справа от описания: раскрывает список проблемных политик.
 *
 * @param {any} item
 * @returns {{ icon: HTMLSpanElement, section: HTMLDivElement } | null}
 */
function createBlockingReasonsToggle(item) {
  if (item.blockingReasons === null) {
    return null;
  }

  const { required, optional } = getPolicyReasonLists(item);
  const totalCount = required.length + optional.length;

  if (totalCount === 0) {
    return null;
  }

  const section = document.createElement("div");
  section.className = "popup__item-blockers";

  const panel = document.createElement("div");
  panel.className = "popup__blockers-panel";
  panel.hidden = true;
  panel.setAttribute("role", "region");
  panel.id = `pr-blockers-${String(item.id).replace(/[^\w-]/g, "_")}`;
  panel.append(createPolicyReasonsBlock(required, optional));
  section.append(panel);

  const votesCheckOnly = optional.length === 0 && isVotesCheckOnlyReasons(required);
  const badgeLabel = votesCheckOnly ? "Votes check" : String(totalCount);

  const badge = document.createElement("span");
  badge.className = "popup__blockers-badge";
  badge.setAttribute("role", "button");
  badge.tabIndex = 0;
  badge.setAttribute("aria-expanded", "false");
  badge.setAttribute("aria-controls", panel.id);
  badge.setAttribute(
    "aria-label",
    `Показать или скрыть проблемы политик Complete (${badgeLabel})`,
  );
  badge.textContent = badgeLabel;

  const toggle = () => {
    const open = panel.hidden;
    panel.hidden = !open;
    badge.setAttribute("aria-expanded", open ? "true" : "false");
    section.classList.toggle("popup__item-blockers--open", open);
  };

  badge.addEventListener("click", (event) => {
    event.preventDefault();
    event.stopPropagation();
    toggle();
  });
  badge.addEventListener("keydown", (event) => {
    if (event.key === "Enter" || event.key === " ") {
      event.preventDefault();
      event.stopPropagation();
      toggle();
    }
  });

  return { icon: badge, section };
}

/**
 * Бейдж «конфликт» справа от замечаний: раскрывает текст конфликта слияния.
 *
 * @param {any} item
 * @returns {{ icon: HTMLSpanElement, section: HTMLDivElement } | null}
 */
function createConflictToggle(item) {
  const conflictText = typeof item.conflictText === "string" ? item.conflictText.trim() : "";

  if (!conflictText) {
    return null;
  }

  const section = document.createElement("div");
  section.className = "popup__item-conflict";

  const panel = document.createElement("div");
  panel.className = "popup__conflict-panel";
  panel.hidden = true;
  panel.setAttribute("role", "region");
  panel.id = `pr-conflict-${String(item.id).replace(/[^\w-]/g, "_")}`;

  const body = document.createElement("p");
  body.className = "popup__conflict-text";
  body.textContent = conflictText;
  panel.append(body);
  section.append(panel);

  const badge = document.createElement("span");
  badge.className = "popup__conflict-badge";
  badge.setAttribute("role", "button");
  badge.tabIndex = 0;
  badge.setAttribute("aria-expanded", "false");
  badge.setAttribute("aria-controls", panel.id);
  badge.setAttribute("aria-label", "Показать или скрыть конфликт слияния");
  badge.textContent = "CONFLICT";

  const toggle = () => {
    const open = panel.hidden;
    panel.hidden = !open;
    badge.setAttribute("aria-expanded", open ? "true" : "false");
    section.classList.toggle("popup__item-conflict--open", open);
  };

  badge.addEventListener("click", (event) => {
    event.preventDefault();
    toggle();
  });

  badge.addEventListener("keydown", (event) => {
    if (event.key === "Enter" || event.key === " ") {
      event.preventDefault();
      toggle();
    }
  });

  return { icon: badge, section };
}

const DESC_ICON_SVG = `<svg viewBox="0 0 16 16" xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="none" aria-hidden="true">
  <path d="M1.5 3L1.5 13L6.293 13L8 14.707L9.707 13L14.5 13L14.5 3L1.5 3ZM2.5 4L13.5 4L13.5 12L9.293 12L8 13.293L6.707 12L2.5 12L2.5 4ZM4.5 5.5L4.5 6.5L11.5 6.5L11.5 5.5L4.5 5.5ZM4.5 7.5L4.5 8.5L11.5 8.5L11.5 7.5L4.5 7.5ZM4.5 9.5L4.5 10.5L9.5 10.5L9.5 9.5L4.5 9.5Z" fill="currentColor" fill-rule="nonzero" />
</svg>`;

/**
 * Скрываем Storybook при конфликтах слияния или истёкшем OSMP Storybook deploy.
 *
 * @param {any} item
 * @returns {boolean}
 */
function shouldShowStorybookLink(item) {
  const conflictText = typeof item?.conflictText === "string" ? item.conflictText.trim() : "";

  if (conflictText) {
    return false;
  }

  const policyReasons = [
    ...(Array.isArray(item?.blockingReasons) ? item.blockingReasons : []),
    ...(Array.isArray(item?.optionalPolicyReasons) ? item.optionalPolicyReasons : []),
  ];

  return !policyReasons.some((reason) => STORYBOOK_EXPIRED_REASON_RE.test(String(reason ?? "")));
}

/**
 * Ссылка на Storybook превью для PR: /hexa-ui/<id>/.
 *
 * @param {{ id: string | number }} item
 * @returns {HTMLButtonElement}
 */
function createStorybookButton(item) {
  const button = document.createElement("button");
  const label = `Открыть Storybook для PR #${item.id}`;

  button.className = "popup__storybook";
  button.type = "button";
  button.setAttribute("aria-label", label);
  button.setAttribute("title", label);

  const icon = document.createElement("img");
  icon.className = "popup__storybook-icon";
  icon.src = "icons/storybook.svg";
  icon.width = 18;
  icon.height = 18;
  icon.alt = "";
  icon.setAttribute("aria-hidden", "true");
  button.append(icon);

  button.addEventListener("click", async (event) => {
    event.preventDefault();
    event.stopPropagation();
    await chrome.tabs.create({ url: `${STORYBOOK_BASE_URL}/${item.id}/` });
    window.close();
  });

  return button;
}

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
 * Время последней проверки: сегодня — только время, иначе дата и время.
 *
 * @param {string | null | undefined} timestamp
 */
function formatLastCheckedAt(timestamp) {
  if (!timestamp) {
    return "ещё не выполнялась";
  }

  const date = new Date(timestamp);

  if (Number.isNaN(date.getTime())) {
    return "ещё не выполнялась";
  }

  const now = new Date();
  const isToday = date.getFullYear() === now.getFullYear()
    && date.getMonth() === now.getMonth()
    && date.getDate() === now.getDate();

  if (isToday) {
    return new Intl.DateTimeFormat("ru-RU", {
      timeStyle: "short",
    }).format(date);
  }

  return new Intl.DateTimeFormat("ru-RU", {
    dateStyle: "short",
    timeStyle: "short",
  }).format(date);
}

/**
 * @param {HTMLParagraphElement} el
 * @param {"yellow"|"orange"|"red"|null} timeUrgency
 */
function fillAuthorMetaParagraph(el, item, checkedAt, timeUrgency) {
  const authorName = item.author || "Автор не определён";
  const hasNoUpdates = !hasUpdatesAfterLastGroupComment(item);
  const relativeTime = hasNoUpdates
    ? null
    : formatElapsedSince(getItemWorkingTimeFrom(item), checkedAt);

  el.replaceChildren();

  const nameSpan = document.createElement("span");
  nameSpan.className = "popup__author-name";
  nameSpan.textContent = authorName;

  if (!relativeTime && !hasNoUpdates) {
    el.append(nameSpan);
    return;
  }

  const sep = document.createElement("span");
  sep.className = "popup__author-sep";
  sep.textContent = " · ";

  const timeSpan = document.createElement("span");
  timeSpan.className = "popup__author-time";

  if (hasNoUpdates) {
    timeSpan.textContent = "Нет обновлений";
  } else {
    if (timeUrgency) {
      timeSpan.classList.add(`popup__author-time--${timeUrgency}`);
    }
    timeSpan.textContent = relativeTime;
  }

  el.append(nameSpan, sep, timeSpan);
}

function formatElapsedSince(createdAt, checkedAt) {
  const totalMinutes = getWorkingElapsedMinutes(createdAt, checkedAt);

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

function applyCountBadgeStyle(urgency) {
  const style = BADGE_STYLES[urgency] ?? BADGE_STYLES.gray;
  countBadge.style.backgroundColor = style.background;
  countBadge.style.borderColor = style.background;
  countBadge.style.color = style.text;
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
 * @param {number} level
 */
function markdownHeadingTag(level) {
  if (level <= 2) {
    return "h3";
  }

  if (level <= 4) {
    return "h4";
  }

  return "h5";
}

/**
 * @param {string} line
 */
function parseMarkdownHeadingLine(line) {
  const trimmed = line.trim();

  if (!/^#{1,6}\s+/.test(trimmed)) {
    return null;
  }

  const level = trimmed.match(/^#+/)?.[0].length ?? 1;
  const content = trimmed.replace(/^#{1,6}\s+/, "");
  const tag = markdownHeadingTag(level);

  return {
    tag,
    level,
    html: `<${tag} class="popup__md-heading popup__md-heading--l${level}">${applyInlineMarkdown(content)}</${tag}>`,
  };
}

/**
 * @param {string[]} lines
 */
function renderMarkdownBlockLines(lines) {
  const out = [];
  let i = 0;

  while (i < lines.length) {
    const trimmed = lines[i].trim();

    if (!trimmed) {
      i += 1;
      continue;
    }

    const heading = parseMarkdownHeadingLine(trimmed);

    if (heading) {
      out.push(heading.html);
      i += 1;
      continue;
    }

    if (/^[-*]\s+/.test(trimmed)) {
      const items = [];

      while (i < lines.length && /^[-*]\s+/.test(lines[i].trim())) {
        const body = lines[i].trim().replace(/^[-*]\s+/, "");
        items.push(`<li class="popup__md-li">${applyInlineMarkdown(body)}</li>`);
        i += 1;
      }

      out.push(`<ul class="popup__md-ul">${items.join("")}</ul>`);
      continue;
    }

    if (/^\d+\.\s+/.test(trimmed)) {
      const items = [];

      while (i < lines.length && /^\d+\.\s+/.test(lines[i].trim())) {
        const body = lines[i].trim().replace(/^\d+\.\s+/, "");
        items.push(`<li class="popup__md-li">${applyInlineMarkdown(body)}</li>`);
        i += 1;
      }

      out.push(`<ol class="popup__md-ol">${items.join("")}</ol>`);
      continue;
    }

    const paraLines = [];

    while (i < lines.length) {
      const lineTrimmed = lines[i].trim();

      if (!lineTrimmed) {
        break;
      }

      if (parseMarkdownHeadingLine(lineTrimmed)) {
        break;
      }

      if (/^[-*]\s+/.test(lineTrimmed) || /^\d+\.\s+/.test(lineTrimmed)) {
        break;
      }

      paraLines.push(lines[i]);
      i += 1;
    }

    if (paraLines.length > 0) {
      const content = paraLines.join("\n").trim();
      out.push(
        `<p class="popup__md-p">${applyInlineMarkdown(content).replace(/\n/g, "<br>")}</p>`,
      );
    }
  }

  return out.join("");
}

/**
 * @param {string} block
 */
function renderMarkdownBlock(block) {
  const b = block.trim();

  if (!b) {
    return "";
  }

  if (/^(-{3,}|\*{3,})$/.test(b)) {
    return "<hr class=\"popup__md-hr\">";
  }

  const lines = b.split("\n");
  const listLines = lines.filter((ln) => ln.trim() !== "");

  if (
    listLines.length > 0
    && listLines.every((ln) => /^[-*]\s+/.test(ln.trim()))
  ) {
    const items = listLines.map((ln) => {
      const body = ln.trim().replace(/^[-*]\s+/, "");
      return `<li class="popup__md-li">${applyInlineMarkdown(body)}</li>`;
    });
    return `<ul class="popup__md-ul">${items.join("")}</ul>`;
  }

  if (listLines.length > 0 && listLines.every((ln) => /^\d+\.\s+/.test(ln.trim()))) {
    const items = listLines.map((ln) => {
      const body = ln.trim().replace(/^\d+\.\s+/, "");
      return `<li class="popup__md-li">${applyInlineMarkdown(body)}</li>`;
    });
    return `<ol class="popup__md-ol">${items.join("")}</ol>`;
  }

  return renderMarkdownBlockLines(lines);
}

/**
 * @param {string | undefined | null} raw
 */
function renderMarkdown(raw) {
  const text = escapeHtml(preprocessMarkdown(raw ?? ""));
  const blocks = text.split(/\n\n+/);

  return blocks
    .map((block) => renderMarkdownBlock(block))
    .filter(Boolean)
    .join("");
}

function isTechPullRequest(description) {
  if (typeof description !== "string") {
    return false;
  }

  return /(те[хx]\s*п[рp]|технический\s*п[рp])/i.test(description);
}
