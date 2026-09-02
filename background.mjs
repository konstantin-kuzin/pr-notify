import {
  ADO_CONFIG_KEY,
  loadAdoConfig,
  validateAdoConfig,
} from "./ado-config.mjs";
import {
  attachPullRequestBlockingReasons,
  attachPullRequestConflictInfo,
  attachPullRequestLastCommitTimes,
  fetchConnectionIdentity,
  filterMyPullRequests,
  filterPullRequestsForExtension,
  getExtensionReviewerContext,
  listActivePullRequestsByCreator,
  listActivePullRequestsForAllowedReviewers,
  listCompletedPullRequestsByCreator,
  logAdoError,
  mapPullRequestToItem,
  MY_TAB_CREATOR_OVERRIDE_ID,
  resolveConfiguredGroupMemberIds,
  setReviewerVoteApprove,
  sortMyPullRequestsNewestFirst,
  sortPullRequestsOldestFirst,
} from "./ado-api.mjs";
import { BADGE_STYLES, getBadgeUrgencyFromItems } from "./working-time.mjs";

const ALARM_NAME = "refresh-pull-requests";
const CHECK_INTERVAL_MINUTES = 10;
const REFRESH_MESSAGE_TYPE = "manual-refresh";
const APPROVE_MESSAGE_TYPE = "approve-pull-request";
const LOAD_MY_COMPLETED_MESSAGE_TYPE = "load-my-completed-pull-requests";
const MY_COMPLETED_PAGE_SIZE = 10;
const STORAGE_KEY = "prState";
const UPDATE_STATE_KEY = "prUpdateState";
const GITHUB_MANIFEST_URL = "https://raw.githubusercontent.com/konstantin-kuzin/pr-notify/main/manifest.json";
const APPROVE_REFRESH_TIMEOUT_MS = 15_000;
const APPROVE_REFRESH_INTERVAL_MS = 2_000;

const DEFAULT_STATE = {
  items: [],
  count: 0,
  approvedItems: [],
  myItems: [],
  myCount: 0,
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

  if (message?.type === LOAD_MY_COMPLETED_MESSAGE_TYPE) {
    void loadMyCompletedPullRequests({
      skip: message?.skip,
      top: message?.top,
    })
      .then((result) => {
        sendResponse({ ok: true, ...result });
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
  await updateBadge(state.count, {
    isError: !!state.lastError,
    items: state.items,
    checkedAt: state.lastCheckedAt,
  });
}

async function bootstrap({ refresh, trigger }) {
  await ensureAlarm();
  void checkForUpdates();

  const state = await getStoredState();
  const hasError = !!state.lastError;
  await updateBadge(hasError ? 0 : state.count, {
    isError: hasError,
    items: state.items,
    checkedAt: state.lastCheckedAt,
  });

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
    await updateBadge(0, { isError: true });
    await saveState(nextState);
    return nextState;
  }

  try {
    const identity = await fetchConnectionIdentity(config);
    // TEMP: вкладка My показывает PR тестового пользователя, не текущего.
    const myCreatorId = MY_TAB_CREATOR_OVERRIDE_ID || identity.id;
    const { allowedReviewerIds } = getExtensionReviewerContext(config, identity.id);
    const [rawPullRequests, rawMyPullRequests] = await Promise.all([
      listActivePullRequestsForAllowedReviewers(config, allowedReviewerIds),
      listActivePullRequestsByCreator(config, myCreatorId),
    ]);
    const { filtered, approved } = await filterPullRequestsForExtension(
      config,
      rawPullRequests,
      identity.id,
    );
    const myFiltered = filterMyPullRequests(rawMyPullRequests);
    const groupMemberIds = await resolveConfiguredGroupMemberIds(config);
    const enrichedPullRequests = await attachPullRequestLastCommitTimes(
      config,
      filtered,
      groupMemberIds,
    );
    const [reviewWithReasons, myWithReasons] = await Promise.all([
      attachPullRequestBlockingReasons(config, enrichedPullRequests),
      attachPullRequestBlockingReasons(config, myFiltered),
    ]);
    const [reviewWithConflicts, myWithConflicts] = await Promise.all([
      attachPullRequestConflictInfo(config, reviewWithReasons),
      attachPullRequestConflictInfo(config, myWithReasons),
    ]);

    const items = sortPullRequestsOldestFirst(
      reviewWithConflicts
        .map((pr) => mapPullRequestToItem(pr, config))
        .filter(Boolean),
    );
    const approvedItems = sortMyPullRequestsNewestFirst(
      approved
        .map((pr) => mapPullRequestToItem(pr, config))
        .filter(Boolean),
    );
    const myItems = sortMyPullRequestsNewestFirst(
      myWithConflicts
        .map((pr) => mapPullRequestToItem(pr, config))
        .filter(Boolean),
    );

    const nextState = {
      items,
      count: items.length,
      approvedItems,
      myItems,
      myCount: myItems.length,
      lastCheckedAt: checkedAt,
      lastSuccessAt: checkedAt,
      lastTrigger: trigger,
      lastError: null,
      previousItemIds: items.map((item) => item.id),
    };

    await saveState(nextState);
    await updateBadge(nextState.count, {
      items: nextState.items,
      checkedAt: nextState.lastCheckedAt,
    });

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

    await updateBadge(0, { isError: true });
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

async function checkForUpdates() {
  const localVersion = chrome.runtime.getManifest().version;
  const checkedAt = new Date().toISOString();
  const abortController = new AbortController();
  const timeoutId = setTimeout(() => abortController.abort(), 3000);

  try {
    const response = await fetch(GITHUB_MANIFEST_URL, {
      cache: "no-store",
      signal: abortController.signal,
    });

    if (!response.ok) {
      throw new Error(`GitHub manifest HTTP ${response.status}`);
    }

    const remoteManifest = await response.json();
    const latestVersion = typeof remoteManifest?.version === "string"
      ? remoteManifest.version.trim()
      : "";

    if (!latestVersion) {
      throw new Error("GitHub manifest does not contain version");
    }

    await chrome.storage.local.set({
      [UPDATE_STATE_KEY]: {
        checkedAt,
        localVersion,
        latestVersion,
        hasUpdate: isRemoteVersionNewer(latestVersion, localVersion),
        error: null,
      },
    });
  } catch (error) {
    await chrome.storage.local.set({
      [UPDATE_STATE_KEY]: {
        checkedAt,
        localVersion,
        latestVersion: "",
        hasUpdate: false,
        error: error instanceof Error ? error.message : String(error),
      },
    });
  } finally {
    clearTimeout(timeoutId);
  }
}

function isRemoteVersionNewer(remoteVersion, localVersion) {
  const remoteParts = parseVersionParts(remoteVersion);
  const localParts = parseVersionParts(localVersion);
  const maxLength = Math.max(remoteParts.length, localParts.length);

  for (let index = 0; index < maxLength; index += 1) {
    const remotePart = remoteParts[index] ?? 0;
    const localPart = localParts[index] ?? 0;

    if (remotePart > localPart) {
      return true;
    }

    if (remotePart < localPart) {
      return false;
    }
  }

  return false;
}

function parseVersionParts(version) {
  return String(version ?? "")
    .split(".")
    .map((part) => Number.parseInt(part, 10))
    .map((part) => (Number.isFinite(part) ? part : 0));
}

const ICON_PATHS = {
  default: {
    16: "icons/icon-16.png",
    32: "icons/icon-32.png",
  },
  orange: {
    16: "icons/icon-16-orange.png",
    32: "icons/icon-32-orange.png",
  },
  red: {
    16: "icons/icon-16-red.png",
    32: "icons/icon-32-red.png",
  },
  error: {
    16: "icons/icon-16-error.png",
    32: "icons/icon-32-error.png",
  },
};

function getIconPathsFromItems(items, checkedAt) {
  const urgency = getBadgeUrgencyFromItems(items, checkedAt);

  if (urgency === "red") {
    return ICON_PATHS.red;
  }

  if (urgency === "orange") {
    return ICON_PATHS.orange;
  }

  return ICON_PATHS.default;
}

async function updateBadge(count, { isError = false, items = [], checkedAt = null } = {}) {
  const text = isError || count <= 0 ? "" : String(count);

  if (isError) {
    await chrome.action.setIcon({ path: ICON_PATHS.error });
    await chrome.action.setBadgeText({ text: "" });
    return;
  }

  await chrome.action.setIcon({ path: getIconPathsFromItems(items, checkedAt) });

  if (count <= 0) {
    await chrome.action.setBadgeText({ text: "" });
    return;
  }

  const urgency = getBadgeUrgencyFromItems(items, checkedAt);
  const style = BADGE_STYLES[urgency] ?? BADGE_STYLES.gray;

  await chrome.action.setBadgeBackgroundColor({ color: style.background });
  await chrome.action.setBadgeText({ text });

  if (chrome.action.setBadgeTextColor) {
    await chrome.action.setBadgeTextColor({ color: style.text });
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

/**
 * Страница завершённых PR текущего пользователя для вкладки My.
 *
 * @param {{ skip?: unknown, top?: unknown }} [options]
 */
async function loadMyCompletedPullRequests(options = {}) {
  const config = await loadAdoConfig();
  const validationErrors = validateAdoConfig(config);

  if (validationErrors.length > 0) {
    throw new Error(`${validationErrors.join(" ")} Откройте настройки расширения.`);
  }

  const skip = Math.max(0, Number(options?.skip) || 0);
  const top = Math.max(1, Number(options?.top) || MY_COMPLETED_PAGE_SIZE);
  const identity = await fetchConnectionIdentity(config);
  const myCreatorId = MY_TAB_CREATOR_OVERRIDE_ID || identity.id;
  const { pullRequests, hasMore } = await listCompletedPullRequestsByCreator(
    config,
    myCreatorId,
    { skip, top },
  );
  const items = sortMyPullRequestsNewestFirst(
    pullRequests
      .map((pr) => mapPullRequestToItem(pr, config))
      .filter(Boolean),
  );

  return {
    items,
    hasMore,
    nextSkip: skip + pullRequests.length,
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
