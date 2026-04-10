import { parsePullRequests, TARGET_PAGE_URL } from "./parser.mjs";

const ALARM_NAME = "refresh-pull-requests";
const CHECK_INTERVAL_MINUTES = 10;
const REFRESH_MESSAGE_TYPE = "manual-refresh";
const STORAGE_KEY = "prState";
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

  void refreshPullRequests("alarm");
});

chrome.runtime.onMessage.addListener((message, _sender, sendResponse) => {
  if (message?.type !== REFRESH_MESSAGE_TYPE) {
    return undefined;
  }

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
});

void bootstrap({ refresh: false, trigger: "service-worker-load" });

async function bootstrap({ refresh, trigger }) {
  await ensureAlarm();

  const state = await getStoredState();
  await updateBadge(state.count);

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
        lastCheckedAt: checkedAt,
        lastTrigger: trigger,
        lastError:
          "Не найден блок Assigned to my teams/Assigned to me на странице pull requests.",
      };

      await saveState(nextState);
      await updateBadge(nextState.count);
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
    await updateBadge(nextState.count);

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
      lastCheckedAt: checkedAt,
      lastTrigger: trigger,
      lastError: error instanceof Error ? error.message : String(error),
    };

    await saveState(nextState);
    await updateBadge(nextState.count);
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

async function updateBadge(count) {
  const text = count > 0 ? String(count) : "";

  await chrome.action.setBadgeBackgroundColor({ color: "#ca2c2c" });
  await chrome.action.setBadgeText({ text });

  if (chrome.action.setBadgeTextColor) {
    await chrome.action.setBadgeTextColor({ color: "#ffffff" });
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
