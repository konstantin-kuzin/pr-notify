import {
  ADO_CONFIG_KEY,
  loadAdoConfig,
  validateAdoConfig,
} from "./ado-config.mjs";
import {
  attachPullRequestLastCommitTimes,
  fetchConnectionIdentity,
  filterPullRequestsForExtension,
  getExtensionReviewerContext,
  listActivePullRequestsForAllowedReviewers,
  logAdoError,
  mapPullRequestToItem,
  setReviewerVoteApprove,
  sortPullRequestsOldestFirst,
} from "./ado-api.mjs";

const ALARM_NAME = "refresh-pull-requests";
const CHECK_INTERVAL_MINUTES = 10;
const REFRESH_MESSAGE_TYPE = "manual-refresh";
const APPROVE_MESSAGE_TYPE = "approve-pull-request";
const STORAGE_KEY = "prState";
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

chrome.storage.onChanged.addListener((changes, areaName) => {
  if (areaName !== "local" || !changes[ADO_CONFIG_KEY]) {
    return;
  }

  void refreshPullRequests("config-change");
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
    void approvePullRequest(message?.pullRequestId)
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
  const config = await loadAdoConfig();
  const validationErrors = validateAdoConfig(config);

  if (validationErrors.length > 0) {
    const nextState = {
      ...previousState,
      count: 0,
      lastCheckedAt: checkedAt,
      lastTrigger: trigger,
      lastError: `${validationErrors.join(" ")} Откройте настройки расширения.`,
    };

    logAdoError("config", new Error(nextState.lastError));
    await updateBadge(0, true);
    await saveState(nextState);
    return nextState;
  }

  try {
    const identity = await fetchConnectionIdentity(config);
    const { allowedReviewerIds } = getExtensionReviewerContext(config, identity.id);
    const rawPullRequests = await listActivePullRequestsForAllowedReviewers(
      config,
      allowedReviewerIds,
    );
    const { filtered, matchedSectionTitle } = await filterPullRequestsForExtension(
      config,
      rawPullRequests,
      identity.id,
    );
    const enrichedPullRequests = await attachPullRequestLastCommitTimes(config, filtered);

    const items = sortPullRequestsOldestFirst(
      enrichedPullRequests
        .map((pr) => mapPullRequestToItem(pr, config))
        .filter(Boolean),
    );

    const nextState = {
      items,
      count: items.length,
      matchedSectionTitle,
      lastCheckedAt: checkedAt,
      lastSuccessAt: checkedAt,
      lastTrigger: trigger,
      lastError: null,
      previousItemIds: items.map((item) => item.id),
    };

    await saveState(nextState);
    await updateBadge(nextState.count, false);

    const newItems = items.filter(
      (item) => !previousState.previousItemIds?.includes(item.id),
    );

    if (newItems.length > 0) {
      void showNotification(newItems);
    }

    return nextState;
  } catch (error) {
    logAdoError("refreshPullRequests", error);
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

async function approvePullRequest(pullRequestId) {
  const normalizedPullRequestId = normalizePullRequestId(pullRequestId);

  if (!normalizedPullRequestId) {
    throw new Error("Не передан идентификатор pull request.");
  }

  const config = await loadAdoConfig();
  const validationErrors = validateAdoConfig(config);

  if (validationErrors.length > 0) {
    throw new Error(`${validationErrors.join(" ")} Откройте настройки расширения.`);
  }

  const identity = await fetchConnectionIdentity(config);
  await setReviewerVoteApprove(config, normalizedPullRequestId, identity.id);

  const state = await forceRefreshAfterApprove(normalizedPullRequestId);

  return {
    approved: true,
    state,
    pullRequestId: normalizedPullRequestId,
  };
}

function normalizePullRequestId(pullRequestId) {
  if (pullRequestId === null || pullRequestId === undefined) {
    return null;
  }

  const normalizedPullRequestId = String(pullRequestId).trim();
  return normalizedPullRequestId || null;
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

function delay(ms) {
  return new Promise((resolve) => {
    setTimeout(resolve, ms);
  });
}
