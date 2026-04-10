import { parsePullRequests, TARGET_PAGE_URL } from "./parser.mjs";

const ALARM_NAME = "refresh-pull-requests";
const CHECK_INTERVAL_MINUTES = 10;
const REFRESH_MESSAGE_TYPE = "manual-refresh";
const APPROVE_MESSAGE_TYPE = "approve-pull-request";
const STORAGE_KEY = "prState";
const APPROVE_TAB_TIMEOUT_MS = 45_000;
const APPROVE_POLL_INTERVAL_MS = 750;
const APPROVE_CONFIRM_TIMEOUT_MS = 15_000;
const APPROVE_REFRESH_TIMEOUT_MS = 15_000;
const APPROVE_REFRESH_INTERVAL_MS = 2_000;
const DEFAULT_STATE = {
  items: [],
  count: 0,
  matchedSectionTitle: null,
  lastCheckedAt: null,
  lastSuccessAt: null,
  lastTrigger: null,
  lastError: null,
  previousItemIds: [],
};

chrome.runtime.onInstalled.addListener(() => {
  void bootstrap({ refresh: true, trigger: "install" });
});

chrome.runtime.onStartup.addListener(() => {
  void bootstrap({ refresh: true, trigger: "startup" });
});

chrome.alarms.onAlarm.addListener((alarm) => {
  if (alarm.name !== ALARM_NAME) {
    return;
  }

  void restoreBadgeFromState().then(() => refreshPullRequests("alarm"));
});

chrome.runtime.onMessage.addListener((message, _sender, sendResponse) => {
  if (message?.type === REFRESH_MESSAGE_TYPE) {
    void restoreBadgeFromState();

    void refreshPullRequests("manual")
      .then((state) => {
        sendResponse({ ok: true, state });
      })
      .catch((error) => {
        sendResponse({
          ok: false,
          error: error instanceof Error ? error.message : String(error),
        });
      });

    return true;
  }

  if (message?.type === APPROVE_MESSAGE_TYPE) {
    void approvePullRequest(message?.url, message?.pullRequestId)
      .then((result) => {
        sendResponse({ ok: true, result });
      })
      .catch((error) => {
        sendResponse({
          ok: false,
          error: error instanceof Error ? error.message : String(error),
        });
      });

    return true;
  }

  return undefined;
});

void bootstrap({ refresh: false, trigger: "service-worker-load" });

async function restoreBadgeFromState() {
  const state = await getStoredState();
  await updateBadge(state.count, !!state.lastError);
}

async function bootstrap({ refresh, trigger }) {
  await ensureAlarm();

  const state = await getStoredState();
  const hasError = !!state.lastError;
  await updateBadge(hasError ? 0 : state.count, hasError);

  if (refresh || !state.lastSuccessAt) {
    await refreshPullRequests(trigger);
  }
}

async function ensureAlarm() {
  const alarm = await chrome.alarms.get(ALARM_NAME);

  if (alarm) {
    return;
  }

  await chrome.alarms.create(ALARM_NAME, {
    periodInMinutes: CHECK_INTERVAL_MINUTES,
  });
}

async function refreshPullRequests(trigger) {
  const previousState = await getStoredState();
  const checkedAt = new Date().toISOString();

  try {
    const response = await fetch(TARGET_PAGE_URL, {
      credentials: "include",
      cache: "no-store",
    });

    if (!response.ok) {
      throw new Error(`HTTP ${response.status}`);
    }

    const html = await response.text();
    const parsed = parsePullRequests(html);

    if (!parsed.sectionFound) {
      const nextState = {
        ...previousState,
        count: 0,
        lastCheckedAt: checkedAt,
        lastTrigger: trigger,
        lastError:
          "Не найден блок Assigned to my teams/Assigned to me на странице pull requests.",
      };

      await updateBadge(0, true);
      await saveState(nextState);
      return nextState;
    }

    const nextState = {
      items: parsed.items,
      count: parsed.items.length,
      matchedSectionTitle: parsed.matchedSectionTitle,
      lastCheckedAt: checkedAt,
      lastSuccessAt: checkedAt,
      lastTrigger: trigger,
      lastError: null,
      previousItemIds: parsed.items.map((item) => item.id),
    };

    await saveState(nextState);
    await updateBadge(nextState.count, false);

    const newItems = parsed.items.filter(
      (item) => !previousState.previousItemIds?.includes(item.id),
    );

    if (newItems.length > 0) {
      void showNotification(newItems);
    }

    return nextState;
  } catch (error) {
    const nextState = {
      ...previousState,
      count: 0,
      lastCheckedAt: checkedAt,
      lastTrigger: trigger,
      lastError: error instanceof Error ? error.message : String(error),
    };

    await updateBadge(0, true);
    await saveState(nextState);
    return nextState;
  }
}

async function getStoredState() {
  const stored = await chrome.storage.local.get(STORAGE_KEY);
  return {
    ...DEFAULT_STATE,
    ...(stored[STORAGE_KEY] ?? {}),
  };
}

async function saveState(state) {
  await chrome.storage.local.set({
    [STORAGE_KEY]: state,
  });
}

async function updateBadge(count, isError) {
  const text = isError ? "" : (count > 0 ? String(count) : "");

  await chrome.action.setBadgeBackgroundColor({ color: isError ? "#a00000" : "#ca2c2c" });
  await chrome.action.setBadgeText({ text });

  if (chrome.action.setBadgeTextColor) {
    await chrome.action.setBadgeTextColor({ color: "#ffffff" });
  }

  if (isError) {
    await chrome.action.setIcon({
      path: {
        16: "icons/icon-16-error.png",
        32: "icons/icon-32-error.png",
      },
    });
  } else {
    await chrome.action.setIcon({
      path: {
        16: "icons/icon-16.png",
        32: "icons/icon-32.png",
      },
    });
  }
}

async function showNotification(newItems) {
  const count = newItems.length;
  const title = count === 1
    ? `Новый pull request`
    : `Новых pull requests: ${count}`;

  const messages = newItems
    .slice(0, 3)
    .map((item) => `#${item.id} ${item.title}`);

  if (newItems.length > 3) {
    messages.push(`…и ещё ${newItems.length - 3}`);
  }

  const message = messages.join("\n");

  await chrome.notifications.create({
    type: "basic",
    iconUrl: "icons/icon-128.png",
    title,
    message,
    priority: 1,
    requireInteraction: false,
  });
}

async function approvePullRequest(url, pullRequestId) {
  const normalizedUrl = normalizePullRequestUrl(url);
  const normalizedPullRequestId = normalizePullRequestId(pullRequestId);
  const { tabId, shouldClose } = await getApproveTab(normalizedUrl);

  if (typeof tabId !== "number") {
    throw new Error("Не удалось создать вкладку для approve.");
  }

  try {
    await waitForTabComplete(tabId, APPROVE_TAB_TIMEOUT_MS);
    await clickApproveButton(tabId, APPROVE_TAB_TIMEOUT_MS);
    await waitForApproveConfirmation(tabId, APPROVE_CONFIRM_TIMEOUT_MS);
    const state = normalizedPullRequestId
      ? await forceRefreshAfterApprove(normalizedPullRequestId)
      : await refreshPullRequests("approve");

    return {
      approved: true,
      state,
      url: normalizedUrl,
    };
  } finally {
    if (shouldClose) {
      await chrome.tabs.remove(tabId).catch(() => {});
    }
  }
}

function normalizePullRequestId(pullRequestId) {
  if (pullRequestId === null || pullRequestId === undefined) {
    return null;
  }

  const normalizedPullRequestId = String(pullRequestId).trim();
  return normalizedPullRequestId || null;
}

function normalizePullRequestUrl(url) {
  if (typeof url !== "string" || !url.trim()) {
    throw new Error("Не передан URL pull request.");
  }

  const normalizedUrl = new URL(url);

  if (normalizedUrl.origin !== "https://hqrndtfs.avp.ru") {
    throw new Error("Approve доступен только для Azure DevOps PR.");
  }

  return normalizedUrl.href;
}

function waitForTabComplete(tabId, timeoutMs) {
  return new Promise((resolve, reject) => {
    let timeoutId = null;

    const finish = (callback) => {
      if (timeoutId) {
        clearTimeout(timeoutId);
      }

      chrome.tabs.onUpdated.removeListener(handleUpdated);
      callback();
    };

    const handleUpdated = (updatedTabId, changeInfo) => {
      if (updatedTabId !== tabId || changeInfo.status !== "complete") {
        return;
      }

      finish(resolve);
    };

    timeoutId = setTimeout(() => {
      finish(() => reject(new Error("Страница PR не загрузилась вовремя.")));
    }, timeoutMs);

    chrome.tabs.onUpdated.addListener(handleUpdated);

    void chrome.tabs.get(tabId).then((tab) => {
      if (tab.status === "complete") {
        finish(resolve);
      }
    }).catch(() => {});
  });
}

async function getApproveTab(url) {
  const fallbackCreateProperties = {
    url,
    active: false,
  };

  const targetWindow = await chrome.windows.getLastFocused({
    populate: false,
    windowTypes: ["normal"],
  }).catch(() => null);

  const createdTab = await chrome.tabs.create(
    targetWindow?.id
      ? {
          ...fallbackCreateProperties,
          windowId: targetWindow.id,
        }
      : fallbackCreateProperties,
  );

  return {
    tabId: createdTab.id,
    shouldClose: true,
  };
}

async function forceRefreshAfterApprove(pullRequestId) {
  const startedAt = Date.now();
  let latestState = await refreshPullRequests("approve");

  while (
    hasPullRequest(latestState, pullRequestId)
    && Date.now() - startedAt < APPROVE_REFRESH_TIMEOUT_MS
  ) {
    await delay(APPROVE_REFRESH_INTERVAL_MS);
    latestState = await refreshPullRequests("approve");
  }

  return latestState;
}

function hasPullRequest(state, pullRequestId) {
  return Array.isArray(state?.items)
    && state.items.some((item) => String(item?.id ?? "") === pullRequestId);
}

async function clickApproveButton(tabId, timeoutMs) {
  const startedAt = Date.now();

  while (Date.now() - startedAt < timeoutMs) {
    const result = await executeApproveScript(tabId, inspectAndClickApproveButton);

    if (result?.status === "clicked") {
      return result;
    }

    if (result?.status === "already-approved") {
      throw new Error("Этот PR уже не находится в состоянии Approve.");
    }

    if (result?.status === "unexpected-label") {
      throw new Error(`Невозможно выполнить Approve: кнопка находится в состоянии "${result.label}".`);
    }

    await delay(APPROVE_POLL_INTERVAL_MS);
  }

  throw new Error("Не удалось найти кнопку Approve на странице PR.");
}

async function waitForApproveConfirmation(tabId, timeoutMs) {
  const startedAt = Date.now();

  while (Date.now() - startedAt < timeoutMs) {
    const result = await executeApproveScript(tabId, inspectApproveState);

    if (result?.status === "approved") {
      return result;
    }

    if (result?.status === "unexpected-label") {
      return result;
    }

    await delay(APPROVE_POLL_INTERVAL_MS);
  }

  throw new Error("Клик по Approve выполнен, но подтверждение изменения статуса не получено.");
}

async function executeApproveScript(tabId, func) {
  const [injectionResult] = await chrome.scripting.executeScript({
    target: { tabId },
    func,
  });

  return injectionResult?.result ?? null;
}

function inspectAndClickApproveButton() {
  const button = document.querySelector("#pull-request-vote-button");

  if (!(button instanceof HTMLButtonElement)) {
    return {
      status: "waiting",
    };
  }

  const label = (button.innerText || button.textContent || "")
    .replace(/\s+/g, " ")
    .trim();

  if (!label) {
    return {
      status: "waiting",
    };
  }

  if (button.disabled) {
    return {
      status: "waiting",
      label,
    };
  }

  if (/^approve$/i.test(label)) {
    button.click();
    return {
      status: "clicked",
      label,
    };
  }

  if (/^approved$/i.test(label)) {
    return {
      status: "already-approved",
      label,
    };
  }

  return {
    status: "unexpected-label",
    label,
  };
}

function inspectApproveState() {
  const button = document.querySelector("#pull-request-vote-button");

  if (!(button instanceof HTMLButtonElement)) {
    return {
      status: "waiting",
    };
  }

  const label = (button.innerText || button.textContent || "")
    .replace(/\s+/g, " ")
    .trim();

  if (/^approved$/i.test(label)) {
    return {
      status: "approved",
      label,
    };
  }

  if (/^approve$/i.test(label)) {
    return {
      status: "waiting",
      label,
    };
  }

  return {
    status: "unexpected-label",
    label,
  };
}

function delay(ms) {
  return new Promise((resolve) => {
    setTimeout(resolve, ms);
  });
}
