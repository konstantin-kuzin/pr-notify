const STORAGE_KEY = "prState";
const REFRESH_MESSAGE_TYPE = "manual-refresh";
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
const itemsList = document.querySelector("#items-list");
const refreshButton = document.querySelector("#refresh-button");

let isRefreshing = false;
let currentState = { ...DEFAULT_STATE };

void init();

async function init() {
  applyPopupMaxHeight();
  currentState = await loadState();
  render();
  refreshButton.addEventListener("click", () => {
    void refreshNow();
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
  countBadge.textContent = String(currentState.count ?? 0);
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

  if (currentState.lastError) {
    messageBox.textContent = `Последняя проверка завершилась ошибкой: ${currentState.lastError}`;
    messageBox.classList.remove("hidden");
  } else {
    messageBox.textContent = "";
    messageBox.classList.add("hidden");
  }

  itemsList.textContent = "";

  if (!currentState.items.length) {
    emptyState.classList.remove("hidden");
    return;
  }

  emptyState.classList.add("hidden");

  for (const item of currentState.items) {
    itemsList.append(createItemElement(item));
  }
}

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

async function refreshNow() {
  if (isRefreshing) {
    return;
  }

  isRefreshing = true;
  render();

  try {
    const response = await chrome.runtime.sendMessage({
      type: REFRESH_MESSAGE_TYPE,
    });

    if (!response?.ok) {
      throw new Error(response?.error || "Не удалось выполнить ручное обновление.");
    }
  } catch (error) {
    messageBox.textContent = `Ручное обновление завершилось ошибкой: ${
      error instanceof Error ? error.message : String(error)
    }`;
    messageBox.classList.remove("hidden");
  } finally {
    isRefreshing = false;
    currentState = await loadState();
    render();
  }
}

function createItemElement(item) {
  const listItem = document.createElement("li");
  listItem.className = "popup__item";

  if (isStalePullRequest(item, currentState.lastCheckedAt)) {
    listItem.classList.add("popup__item--stale");
  }

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
  author.textContent = buildAuthorMeta(item, currentState.lastCheckedAt);

  authorRow.append(author);

  if (item.description) {
    const { infoButton, isTechPR } = createInfoTooltip(item.description);
    authorRow.append(infoButton);

    if (isTechPR) {
      const badge = document.createElement("span");
      badge.className = "popup__badge";
      badge.textContent = "ТЕХ ПР";
      authorRow.append(badge);
    }
  }

  itemContent.append(button, authorRow);
  itemMain.append(itemContent);
  listItem.append(itemMain);

  return listItem;
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

function buildAuthorMeta(item, checkedAt) {
  const author = item.author || "Автор не определён";
  const relativeTime = formatElapsedSince(item.createdAt, checkedAt);

  if (!relativeTime) {
    return author;
  }

  return `${author} · ${relativeTime}`;
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

  return `${hours} ч ${minutes} мин`;
}

function isStalePullRequest(item, checkedAt) {
  const totalMinutes = getElapsedMinutes(item.createdAt, checkedAt);

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

function createInfoTooltip(description) {
  const wrapper = document.createElement("div");
  wrapper.className = "popup__info-wrapper";

  const infoButton = document.createElement("button");
  infoButton.className = "popup__info";
  infoButton.type = "button";
  infoButton.setAttribute("aria-label", "Показать описание");
  infoButton.innerHTML = `<svg viewBox="0 0 16 16" xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="none">
    <path d="M1.5 3L1.5 13L6.293 13L8 14.707L9.707 13L14.5 13L14.5 3L1.5 3ZM2.5 4L13.5 4L13.5 12L9.293 12L8 13.293L6.707 12L2.5 12L2.5 4ZM4.5 5.5L4.5 6.5L11.5 6.5L11.5 5.5L4.5 5.5ZM4.5 7.5L4.5 8.5L11.5 8.5L11.5 7.5L4.5 7.5ZM4.5 9.5L4.5 10.5L9.5 10.5L9.5 9.5L4.5 9.5Z" fill="currentColor" fill-rule="nonzero" />
  </svg>`;

  const tooltip = document.createElement("div");
  tooltip.className = "popup__tooltip popup__tooltip--markdown";
  tooltip.innerHTML = renderMarkdown(description);
  tooltip.setAttribute("role", "tooltip");

  const isTechPR = /те[хx]\s*п[рp]/i.test(description);

  let isVisible = false;

  const showTooltip = () => {
    if (!isVisible) {
      isVisible = true;
      tooltip.classList.add("popup__tooltip--visible");
      tooltip.style.top = "";
      tooltip.style.bottom = "";
      requestAnimationFrame(() => {
        const rect = tooltip.getBoundingClientRect();
        if (rect.top < 0) {
          tooltip.style.bottom = "auto";
          tooltip.style.top = "100%";
          tooltip.style.marginTop = "6px";
          tooltip.classList.remove("popup__tooltip--top");
          tooltip.classList.add("popup__tooltip--bottom");
        } else {
          tooltip.style.bottom = "calc(100% + 6px)";
          tooltip.style.top = "auto";
          tooltip.style.marginTop = "";
          tooltip.classList.remove("popup__tooltip--bottom");
          tooltip.classList.add("popup__tooltip--top");
        }
      });
      infoButton.setAttribute("aria-expanded", "true");
    }
  };

  const hideTooltip = () => {
    if (isVisible) {
      isVisible = false;
      tooltip.classList.remove("popup__tooltip--visible");
      infoButton.setAttribute("aria-expanded", "false");
    }
  };

  infoButton.addEventListener("mouseenter", showTooltip);
  infoButton.addEventListener("mouseleave", hideTooltip);
  tooltip.addEventListener("mouseenter", showTooltip);
  tooltip.addEventListener("mouseleave", hideTooltip);

  wrapper.append(infoButton, tooltip);
  return { infoButton: wrapper, isTechPR };
}

function renderMarkdown(text) {
  return text
    .replace(/\*\*(.+?)\*\*/g, "<strong>$1</strong>")
    .replace(/\*(.+?)\*/g, "<em>$1</em>")
    .replace(/`(.+?)`/g, "<code>$1</code>");
}
