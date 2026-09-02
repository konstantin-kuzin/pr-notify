import {
  normalizeApiRoot,
  resolveApiVersion,
} from "./ado-config.mjs";

const PAGE_SIZE = 100;
const MAX_RETRIES = 2;
const RETRY_BASE_MS = 900;
const IDENTITY_BATCH_SIZE = 40;
const REVIEWER_GROUPS_CACHE_TTL_MS = 5 * 60 * 1000;
const GROUP_MEMBER_IDS_CACHE_TTL_MS = 5 * 60 * 1000;
const reviewerGroupsCache = new Map();
const groupMemberIdsCache = new Map();
const commitTimestampCache = new Map();
const pushTimestampCache = new Map();

/**
 * @param {import("./ado-config.mjs").DEFAULT_ADO_CONFIG} config
 */
function buildAuthHeaders(config) {
  const headers = {
    Accept: "application/json",
  };

  if (config.authMode === "pat" && config.pat?.trim()) {
    const token = config.pat.trim();
    const basic = btoa(`:${token}`);
    headers.Authorization = `Basic ${basic}`;
  }

  return headers;
}

/**
 * @param {import("./ado-config.mjs").DEFAULT_ADO_CONFIG} config
 * @param {string} pathAndQuery
 */
export async function adoFetch(config, pathAndQuery, init = {}) {
  const root = normalizeApiRoot(config.apiRoot);
  const url = `${root}/${pathAndQuery.replace(/^\//, "")}`;
  const isWrite = ["POST", "PUT", "PATCH", "DELETE"].includes(
    (init.method ?? "GET").toUpperCase(),
  );

  const headers = {
    ...buildAuthHeaders(config),
    ...(init.headers ?? {}),
  };

  if (isWrite && !headers["Content-Type"]) {
    headers["Content-Type"] = "application/json";
  }

  let lastError = null;

  for (let attempt = 0; attempt <= MAX_RETRIES; attempt++) {
    const response = await fetch(url, {
      ...init,
      headers,
      credentials: config.authMode === "session" ? "include" : "omit",
      cache: "no-store",
    });

    if (response.status === 429 && attempt < MAX_RETRIES) {
      await delay(RETRY_BASE_MS * 2 ** attempt);
      continue;
    }

    if (!response.ok) {
      lastError = await buildAdoHttpError(response);
      break;
    }

    if (response.status === 204) {
      return null;
    }

    const text = await response.text();

    if (!text) {
      return null;
    }

    try {
      return JSON.parse(text);
    } catch (_error) {
      lastError = new Error("Ответ API не является JSON.");
      break;
    }
  }

  throw lastError ?? new Error("Запрос к Azure DevOps не выполнен.");
}

async function buildAdoHttpError(response) {
  let detail = "";

  try {
    const text = await response.text();
    if (text) {
      const parsed = JSON.parse(text);
      detail = parsed?.message || parsed?.Message || "";
    }
  } catch (_error) {
    // ignore
  }

  const status = response.status;
  let base = mapStatusToMessage(status);

  if (status === 400 && detail) {
    if (/preview flag must be supplied|-preview/i.test(detail)) {
      base = "Для этой версии API сервер требует суффикс -preview (например 6.0-preview). Укажите это в настройках расширения.";
    } else if (/out of range|REST API version|api version/i.test(detail)) {
      base = "Версия REST API не подходит серверу. В настройках укажите поддерживаемый api-version (для on-prem часто 6.0-preview или 6.0).";
    }
  }

  if (detail && !looksSensitive(detail)) {
    return new Error(`${base} ${detail}`.trim());
  }

  return new Error(base);
}

function looksSensitive(text) {
  return /pat|password|token|authorization|bearer/i.test(text);
}

function mapStatusToMessage(status) {
  if (status === 401) {
    return "Доступ запрещён (401): войдите в Azure DevOps в браузере или укажите PAT в настройках.";
  }

  if (status === 403) {
    return "Недостаточно прав (403): проверьте права на репозиторий или PAT.";
  }

  if (status === 404) {
    return "Ресурс не найден (404): проверьте project, repository и корень API.";
  }

  if (status === 429) {
    return "Слишком много запросов (429): повторите позже.";
  }

  if (status >= 500) {
    return `Ошибка сервера Azure DevOps (${status}).`;
  }

  return `Ошибка HTTP ${status}.`;
}

function delay(ms) {
  return new Promise((resolve) => {
    setTimeout(resolve, ms);
  });
}

/**
 * @param {import("./ado-config.mjs").DEFAULT_ADO_CONFIG} config
 */
export async function fetchConnectionIdentity(config) {
  const query = new URLSearchParams({
    connectOptions: "1",
    lastChangeId: "-1",
    lastChangeId64: "-1",
    "api-version": resolveApiVersion(config),
  });

  const data = await adoFetch(config, `_apis/connectionData?${query.toString()}`);
  const id = data?.authenticatedUser?.id;

  if (!id) {
    throw new Error("Не удалось определить текущего пользователя (connectionData).");
  }

  return {
    id: String(id),
    displayName: data?.authenticatedUser?.displayName ?? "",
  };
}

/**
 * Одна страница PR по searchCriteria (Git REST API).
 *
 * @param {import("./ado-config.mjs").DEFAULT_ADO_CONFIG} config
 * @param {{
 *   status?: string,
 *   reviewerId?: string | null,
 *   creatorId?: string | null,
 *   top?: number,
 *   skip?: number,
 * }} [criteria]
 */
async function listPullRequestsPage(config, criteria = {}) {
  const status = String(criteria?.status ?? "active").trim() || "active";
  const reviewerId = criteria?.reviewerId ?? null;
  const creatorId = criteria?.creatorId ?? null;
  const top = Math.max(1, Number(criteria?.top) || PAGE_SIZE);
  const skip = Math.max(0, Number(criteria?.skip) || 0);
  const project = encodeURIComponent(config.project.trim());
  const repo = encodeURIComponent(config.repositoryId.trim());
  const basePath = `${project}/_apis/git/repositories/${repo}/pullrequests`;

  const query = new URLSearchParams({
    "searchCriteria.status": status,
    "api-version": resolveApiVersion(config),
    "$top": String(top),
    "$skip": String(skip),
  });

  if (reviewerId) {
    query.set("searchCriteria.reviewerId", String(reviewerId).trim());
  }

  if (creatorId) {
    query.set("searchCriteria.creatorId", String(creatorId).trim());
  }

  const data = await adoFetch(config, `${basePath}?${query.toString()}`);
  return Array.isArray(data?.value) ? data.value : [];
}

/**
 * Список активных PR. При переданном reviewerId / creatorId сервер фильтрует через
 * searchCriteria (см. Git REST API), без обхода всех активных PR репозитория.
 *
 * @param {import("./ado-config.mjs").DEFAULT_ADO_CONFIG} config
 * @param {{ reviewerId?: string | null, creatorId?: string | null }} [criteria]
 */
async function listActivePullRequests(config, criteria = {}) {
  const all = [];
  let skip = 0;

  for (;;) {
    const batch = await listPullRequestsPage(config, {
      status: "active",
      reviewerId: criteria?.reviewerId ?? null,
      creatorId: criteria?.creatorId ?? null,
      top: PAGE_SIZE,
      skip,
    });

    all.push(...batch);

    if (batch.length < PAGE_SIZE) {
      break;
    }

    skip += PAGE_SIZE;
  }

  return all;
}

function dedupePullRequestsById(pullRequests) {
  const byId = new Map();

  for (const pr of pullRequests) {
    const pid = pr?.pullRequestId;

    if (pid == null) {
      continue;
    }

    if (!byId.has(pid)) {
      byId.set(pid, pr);
    }
  }

  return [...byId.values()];
}

/**
 * Активные PR, где указанные identities числятся ревьюерами. Запросы по ревьюерам идут параллельно, результат объединяется.
 *
 * @param {import("./ado-config.mjs").DEFAULT_ADO_CONFIG} config
 * @param {string[]} reviewerIds
 */
export async function listActivePullRequestsForAllowedReviewers(config, reviewerIds) {
  const ids = [...new Set(reviewerIds.map(String).map((id) => id.trim()).filter(Boolean))];

  if (ids.length === 0) {
    return [];
  }

  const batches = await Promise.all(
    ids.map((id) => listActivePullRequests(config, { reviewerId: id })),
  );
  return dedupePullRequestsById(batches.flat());
}

/**
 * Активные PR, созданные указанным пользователем.
 *
 * @param {import("./ado-config.mjs").DEFAULT_ADO_CONFIG} config
 * @param {string} creatorId
 */
export async function listActivePullRequestsByCreator(config, creatorId) {
  const id = String(creatorId ?? "").trim();

  if (!id) {
    return [];
  }

  return listActivePullRequests(config, { creatorId: id });
}

/**
 * Страница завершённых (Complete) PR, созданных указанным пользователем.
 *
 * @param {import("./ado-config.mjs").DEFAULT_ADO_CONFIG} config
 * @param {string} creatorId
 * @param {{ top?: number, skip?: number }} [options]
 * @returns {Promise<{ pullRequests: Array<any>, hasMore: boolean }>}
 */
export async function listCompletedPullRequestsByCreator(config, creatorId, options = {}) {
  const id = String(creatorId ?? "").trim();
  const top = Math.max(1, Number(options?.top) || 10);
  const skip = Math.max(0, Number(options?.skip) || 0);

  if (!id) {
    return { pullRequests: [], hasMore: false };
  }

  const pullRequests = await listPullRequestsPage(config, {
    status: "completed",
    creatorId: id,
    top,
    skip,
  });

  return {
    pullRequests,
    hasMore: pullRequests.length >= top,
  };
}

/**
 * TEMP для тестирования вкладки My: identity id вместо текущего пользователя.
 * После QA вернуть `null`.
 */
export const MY_TAB_CREATOR_OVERRIDE_ID = null;

const POLICY_TYPE_MINIMUM_REVIEWERS = "fa4e907d-c16b-4a4c-9dfa-4906e5d171dd";
const POLICY_TYPE_REQUIRED_REVIEWERS = "fd2167ab-b0be-447a-8ec8-39368250530e";
const POLICY_TYPE_COMMENT_REQUIREMENTS = "c6a1889d-b943-4856-b76f-9e46bb6b0df2";
const POLICY_TYPE_WORK_ITEM_LINKING = "40e92b44-2fe1-4dd6-b3d8-74a9c21d0c6e";
const POLICY_TYPE_BUILD = "0609b952-1397-4640-95ec-e00a01b2c241";
const POLICY_TYPE_STATUS = "cbdc66da-9728-4af8-aada-9a5a32e4a226";
/** В overview Policies ADO обычно не показывает — только при Complete. */
const POLICY_TYPE_MERGE_STRATEGY = "fa4e907d-c16b-4a4c-9dfa-4916e5d171ab";
const projectIdCache = new Map();

function resolvePolicyEvaluationsApiVersion(config) {
  const ver = resolveApiVersion(config);

  if (/preview\.\d+$/i.test(ver)) {
    return ver;
  }

  if (/preview$/i.test(ver)) {
    return `${ver}.1`;
  }

  return `${ver}-preview.1`;
}

function isGuidString(value) {
  return /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(
    String(value ?? "").trim(),
  );
}

function projectIdCacheKey(config) {
  return [
    normalizeApiRoot(config.apiRoot),
    resolveApiVersion(config),
    String(config.project ?? "").trim().toLowerCase(),
  ].join("|");
}

/**
 * GUID проекта для artifactId policy evaluations.
 *
 * @param {import("./ado-config.mjs").DEFAULT_ADO_CONFIG} config
 * @param {any} [pullRequest]
 */
async function resolveProjectId(config, pullRequest = null) {
  const fromPr = normalizePlainText(
    pullRequest?.repository?.project?.id
      ?? pullRequest?.repository?.project?.Id
      ?? "",
  );

  if (isGuidString(fromPr)) {
    return fromPr;
  }

  const cacheKey = projectIdCacheKey(config);

  if (projectIdCache.has(cacheKey)) {
    return projectIdCache.get(cacheKey);
  }

  const projectSeg = encodeURIComponent(config.project.trim());
  const query = new URLSearchParams({
    "api-version": resolveApiVersion(config),
  });
  const data = await adoFetch(config, `_apis/projects/${projectSeg}?${query.toString()}`);
  const projectId = normalizePlainText(data?.id ?? data?.Id ?? "");

  if (!isGuidString(projectId)) {
    throw new Error("Не удалось определить GUID проекта для policy evaluations.");
  }

  projectIdCache.set(cacheKey, projectId);
  return projectId;
}

function buildPullRequestPolicyArtifactId(projectId, pullRequestId) {
  return `vstfs:///CodeReview/CodeReviewId/${projectId}/${pullRequestId}`;
}

/**
 * @param {import("./ado-config.mjs").DEFAULT_ADO_CONFIG} config
 * @param {any} pullRequest
 */
async function fetchPullRequestPolicyEvaluations(config, pullRequest) {
  const prId = pullRequest?.pullRequestId;

  if (prId == null) {
    return [];
  }

  const projectId = await resolveProjectId(config, pullRequest);
  const projectSeg = encodeURIComponent(
    normalizePlainText(
      pullRequest?.repository?.project?.name
        ?? pullRequest?.repository?.project?.Name
        ?? config.project,
    ),
  );
  const artifactId = buildPullRequestPolicyArtifactId(projectId, prId);
  const query = new URLSearchParams({
    artifactId,
    "api-version": resolvePolicyEvaluationsApiVersion(config),
  });
  const data = await adoFetch(
    config,
    `${projectSeg}/_apis/policy/evaluations?${query.toString()}`,
  );

  return Array.isArray(data?.value) ? data.value : Array.isArray(data) ? data : [];
}

/**
 * Статусы PR (для текстов Status policy, как в ADO UI).
 *
 * @param {import("./ado-config.mjs").DEFAULT_ADO_CONFIG} config
 * @param {any} pullRequest
 */
async function fetchPullRequestStatuses(config, pullRequest) {
  const prId = pullRequest?.pullRequestId;

  if (prId == null) {
    return [];
  }

  const projectSeg = encodeURIComponent(
    normalizePlainText(
      pullRequest?.repository?.project?.name
        ?? pullRequest?.repository?.project?.Name
        ?? config.project,
    ),
  );
  const repoSeg = encodeURIComponent(
    normalizePlainText(
      pullRequest?.repository?.id
        ?? pullRequest?.repository?.name
        ?? config.repositoryId,
    ),
  );
  const prIdSeg = encodeURIComponent(String(prId).trim());
  const query = new URLSearchParams({
    "api-version": resolveApiVersion(config),
  });
  const data = await adoFetch(
    config,
    `${projectSeg}/_apis/git/repositories/${repoSeg}/pullRequests/${prIdSeg}/statuses?${query.toString()}`,
  );

  return Array.isArray(data?.value) ? data.value : [];
}

/**
 * @param {Array<any>} statuses
 * @returns {Map<number, any>}
 */
function indexPullRequestStatusesById(statuses) {
  /** @type {Map<number, any>} */
  const byId = new Map();

  for (const status of statuses) {
    const id = Number(status?.id);

    if (!Number.isFinite(id)) {
      continue;
    }

    byId.set(id, status);
  }

  return byId;
}

/**
 * Политика видна в overview Policies ADO (Required или Optional):
 * включена, не удалена, не merge strategy, статус не approved/notApplicable.
 *
 * @param {any} evaluation
 */
function isOverviewPolicyEvaluation(evaluation) {
  const configuration = evaluation?.configuration;

  if (!configuration || typeof configuration !== "object") {
    return false;
  }

  if (configuration.isEnabled === false || configuration.isDeleted === true) {
    return false;
  }

  const typeId = normalizePlainText(configuration?.type?.id).toLowerCase();

  // Merge strategy в overview Policies не выводится — только при Complete.
  if (typeId === POLICY_TYPE_MERGE_STRATEGY) {
    return false;
  }

  const status = String(evaluation?.status ?? "").toLowerCase();

  return status !== "approved" && status !== "notapplicable";
}

/**
 * Required-политика (`isBlocking`), блокирующая Complete.
 *
 * @param {any} evaluation
 */
function isRequiredPolicyEvaluation(evaluation) {
  return evaluation?.configuration?.isBlocking === true;
}

function countBlockingIndividualReviewers(pullRequest) {
  const reviewers = Array.isArray(pullRequest?.reviewers) ? pullRequest.reviewers : [];

  return reviewers.filter((reviewer) => {
    if (reviewer?.isContainer === true) {
      return false;
    }

    return Number(reviewer?.vote ?? 0) < 0;
  }).length;
}

function countApprovedReviewers(pullRequest) {
  const reviewers = Array.isArray(pullRequest?.reviewers) ? pullRequest.reviewers : [];

  return reviewers.filter((reviewer) => {
    if (reviewer?.isContainer === true) {
      return false;
    }

    return Number(reviewer?.vote ?? 0) >= 10;
  }).length;
}

/**
 * Человекочитаемая причина в духе блока Policies в ADO (Required / Optional).
 *
 * @param {any} evaluation
 * @param {any} [pullRequest]
 * @param {Map<number, any>} [statusById]
 * @returns {string | null}
 */
function formatPolicyBlockingReason(evaluation, pullRequest = null, statusById = null) {
  const configuration = evaluation?.configuration ?? {};
  const settings = configuration?.settings && typeof configuration.settings === "object"
    ? configuration.settings
    : {};
  const context = evaluation?.context && typeof evaluation.context === "object"
    ? evaluation.context
    : {};
  const type = configuration?.type && typeof configuration.type === "object"
    ? configuration.type
    : {};
  const typeId = normalizePlainText(type.id).toLowerCase();
  const typeName = normalizePlainText(type.displayName);
  const status = String(evaluation?.status ?? "").toLowerCase();

  if (
    typeId === POLICY_TYPE_MINIMUM_REVIEWERS
    || /minimum (number of )?reviewers|minimum approval count|approval count/i.test(typeName)
  ) {
    const allowDownvotes = settings.allowDownvotes === true;
    const blockers = countBlockingIndividualReviewers(pullRequest);

    if (!allowDownvotes && blockers > 0) {
      return blockers === 1
        ? "1 reviewer is blocking"
        : `${blockers} reviewers are blocking`;
    }

    const minimum = Number(
      context.minimumApproverCount ?? settings.minimumApproverCount ?? NaN,
    );
    const actual = Number.isFinite(Number(
      context.approverCount
        ?? context.actualApproverCount
        ?? context.approvedCount
        ?? context.approveCount,
    ))
      ? Number(
        context.approverCount
          ?? context.actualApproverCount
          ?? context.approvedCount
          ?? context.approveCount,
      )
      : countApprovedReviewers(pullRequest);

    if (Number.isFinite(minimum) && minimum > 0) {
      return `${Number.isFinite(actual) ? actual : 0} of ${minimum} reviewers approved`;
    }

    return typeName || "Minimum number of reviewers";
  }

  if (
    typeId === POLICY_TYPE_REQUIRED_REVIEWERS
    || /^required reviewers$/i.test(typeName)
  ) {
    return "Required reviewers have not approved";
  }

  if (
    typeId === POLICY_TYPE_COMMENT_REQUIREMENTS
    || /^comment requirements$/i.test(typeName)
  ) {
    return "Not all comments resolved";
  }

  if (
    typeId === POLICY_TYPE_WORK_ITEM_LINKING
    || /work item linking/i.test(typeName)
  ) {
    return "Work items not linked";
  }

  if (typeId === POLICY_TYPE_BUILD || /^build$/i.test(typeName)) {
    const name = normalizePlainText(settings.displayName) || typeName || "Build";

    if (context.isExpired === true) {
      return `${name} expired`;
    }

    if (status === "running") {
      return `${name} — выполняется`;
    }

    if (status === "broken") {
      return `${name} — ошибка проверки`;
    }

    return name;
  }

  if (typeId === POLICY_TYPE_STATUS || /^status$/i.test(typeName)) {
    const latestStatusId = Number(context.latestStatusId);
    const linkedStatus = Number.isFinite(latestStatusId) && statusById instanceof Map
      ? statusById.get(latestStatusId)
      : null;
    const fromStatus = normalizePlainText(linkedStatus?.description);
    const name = fromStatus
      || normalizePlainText(settings.defaultDisplayName)
      || normalizePlainText(settings.displayName)
      || typeName
      || "Status";

    if (context.isExpired === true && !/\bexpired\b/i.test(name)) {
      return `${name} expired`;
    }

    return name;
  }

  const reason = normalizePlainText(settings.displayName)
    || normalizePlainText(settings.defaultDisplayName)
    || normalizePlainText(
      context.message
        ?? context.statusMessage
        ?? context.errorMessage
        ?? context.displayName
        ?? "",
    )
    || typeName;

  return reason || null;
}

/**
 * Для каждого PR подмешивает проблемы Policies из overview ADO:
 * - `blockingReasons` — Required (`isBlocking`), из‑за которых недоступен Complete;
 * - `optionalPolicyReasons` — Optional (не blocking), красные пункты в Optional.
 *
 * @param {import("./ado-config.mjs").DEFAULT_ADO_CONFIG} config
 * @param {Array<any>} pullRequests
 */
export async function attachPullRequestBlockingReasons(config, pullRequests) {
  return Promise.all(
    pullRequests.map(async (pullRequest) => {
      const prId = pullRequest?.pullRequestId;

      if (prId == null) {
        return pullRequest;
      }

      try {
        const [evaluations, statuses] = await Promise.all([
          fetchPullRequestPolicyEvaluations(config, pullRequest),
          fetchPullRequestStatuses(config, pullRequest).catch((error) => {
            logAdoError(`fetchPullRequestStatuses ${prId}`, error);
            return [];
          }),
        ]);
        const statusById = indexPullRequestStatusesById(statuses);
        const blockingReasons = [];
        const optionalPolicyReasons = [];
        const seen = new Set();

        for (const evaluation of evaluations) {
          if (!isOverviewPolicyEvaluation(evaluation)) {
            continue;
          }

          const reason = formatPolicyBlockingReason(evaluation, pullRequest, statusById);

          if (!reason || seen.has(reason)) {
            continue;
          }

          seen.add(reason);

          if (isRequiredPolicyEvaluation(evaluation)) {
            blockingReasons.push(reason);
          } else {
            optionalPolicyReasons.push(reason);
          }
        }

        return {
          ...pullRequest,
          blockingReasons,
          optionalPolicyReasons,
        };
      } catch (error) {
        logAdoError(`fetchPullRequestPolicyEvaluations ${prId}`, error);
        return {
          ...pullRequest,
          blockingReasons: null,
          optionalPolicyReasons: null,
        };
      }
    }),
  );
}

function hasPullRequestMergeConflicts(pullRequest) {
  return String(pullRequest?.mergeStatus ?? "").toLowerCase() === "conflicts";
}

/**
 * api-version для Git Conflicts (preview-ресурс).
 *
 * @param {import("./ado-config.mjs").DEFAULT_ADO_CONFIG} config
 */
function resolveConflictsApiVersion(config) {
  const ver = resolveApiVersion(config);

  if (/preview/i.test(ver)) {
    return ver;
  }

  return `${ver}-preview`;
}

/**
 * @param {string} conflictType
 */
function formatConflictTypeLabel(conflictType) {
  switch (String(conflictType ?? "").toLowerCase()) {
    case "addadd":
      return "Added in both";
    case "addrename":
      return "Added in source, renamed in target";
    case "deleteedit":
      return "Deleted in source, edited in target";
    case "deleterename":
      return "Deleted in source, renamed in target";
    case "directoryfile":
      return "Directory in source, file in target";
    case "filedirectory":
      return "File in source, directory in target";
    case "editdelete":
      return "Edited in source, deleted in target";
    case "editedit":
      return "Edited in both";
    case "renameadd":
      return "Renamed in source, added in target";
    case "renamedelete":
      return "Renamed in source, deleted in target";
    case "renamerename":
      return "Renamed in both";
    default:
      return normalizePlainText(conflictType) || "Conflict";
  }
}

/**
 * @param {number} count
 */
function formatConflictSummaryMessage(count) {
  if (count === 1) {
    return "1 conflict prevents automatic merging";
  }

  if (count > 1) {
    return `${count} conflicts prevent automatic merging`;
  }

  return "Conflict prevents automatic merging";
}

/**
 * @param {import("./ado-config.mjs").DEFAULT_ADO_CONFIG} config
 * @param {any} pullRequest
 */
async function fetchPullRequestConflicts(config, pullRequest) {
  const prId = pullRequest?.pullRequestId;

  if (prId == null) {
    return [];
  }

  const projectSeg = encodeURIComponent(
    normalizePlainText(
      pullRequest?.repository?.project?.name
        ?? pullRequest?.repository?.project?.Name
        ?? config.project,
    ),
  );
  const repoSeg = encodeURIComponent(
    normalizePlainText(
      pullRequest?.repository?.id
        ?? pullRequest?.repository?.name
        ?? config.repositoryId,
    ),
  );
  const prIdSeg = encodeURIComponent(String(prId).trim());
  const query = new URLSearchParams({
    "api-version": resolveConflictsApiVersion(config),
  });
  const data = await adoFetch(
    config,
    `${projectSeg}/_apis/git/repositories/${repoSeg}/pullRequests/${prIdSeg}/conflicts?${query.toString()}`,
  );

  return Array.isArray(data?.value) ? data.value : [];
}

/**
 * @param {Array<any>} conflicts
 */
function formatConflictText(conflicts) {
  const rows = Array.isArray(conflicts) ? conflicts : [];
  const fileLines = [];

  for (const conflict of rows) {
    const path = normalizePlainText(conflict?.conflictPath).replace(/^\//, "");
    const typeLabel = formatConflictTypeLabel(conflict?.conflictType);

    if (!path) {
      continue;
    }

    fileLines.push(`${path} — ${typeLabel}`);
  }

  const summary = formatConflictSummaryMessage(fileLines.length || rows.length);
  return fileLines.length > 0 ? `${summary}\n\n${fileLines.join("\n")}` : summary;
}

/**
 * Для PR с `mergeStatus: conflicts` подмешивает `conflictText` (как баннер Conflicts в ADO).
 *
 * @param {import("./ado-config.mjs").DEFAULT_ADO_CONFIG} config
 * @param {Array<any>} pullRequests
 */
export async function attachPullRequestConflictInfo(config, pullRequests) {
  return Promise.all(
    pullRequests.map(async (pullRequest) => {
      if (!hasPullRequestMergeConflicts(pullRequest)) {
        return pullRequest;
      }

      const prId = pullRequest?.pullRequestId;

      try {
        const conflicts = await fetchPullRequestConflicts(config, pullRequest);
        return {
          ...pullRequest,
          conflictText: formatConflictText(conflicts),
        };
      } catch (error) {
        logAdoError(`fetchPullRequestConflicts ${prId}`, error);
        return {
          ...pullRequest,
          conflictText: formatConflictSummaryMessage(0),
        };
      }
    }),
  );
}

/**
 * Активные не-draft PR текущего пользователя (создателя).
 *
 * @param {Array<any>} pullRequests
 */
export function filterMyPullRequests(pullRequests) {
  return pullRequests.filter((pullRequest) => isVisiblePullRequestForExtension(pullRequest));
}

/**
 * @param {Array<{ createdAt?: string | null, id?: string }>} items
 */
export function sortMyPullRequestsNewestFirst(items) {
  return [...items].sort((left, right) => {
    const leftTs = Date.parse(left?.createdAt ?? "") || 0;
    const rightTs = Date.parse(right?.createdAt ?? "") || 0;

    if (rightTs !== leftTs) {
      return rightTs - leftTs;
    }

    return String(right?.id ?? "").localeCompare(String(left?.id ?? ""), undefined, {
      numeric: true,
    });
  });
}

function commitTimestampCacheKey(config, project, repo, commitId) {
  return [
    normalizeApiRoot(config.apiRoot),
    resolveApiVersion(config),
    String(project),
    String(repo),
    String(commitId),
  ].join("|");
}

function pushTimestampCacheKey(config, project, repo, pushId) {
  return [
    normalizeApiRoot(config.apiRoot),
    resolveApiVersion(config),
    String(project),
    String(repo),
    String(pushId),
  ].join("|");
}

function pickPushMeta(ref) {
  const push = ref?.push ?? ref?.Push;

  if (!push || typeof push !== "object") {
    return { date: null, pushId: null };
  }

  const date = normalizeIsoDate(push.date ?? push.Date ?? "");
  const rawPushId = push.pushId ?? push.PushId;
  const pushId = rawPushId == null || rawPushId === "" ? null : rawPushId;

  return { date, pushId };
}

/** Дата коммита: committer ближе к моменту push (amend/rebase сохраняют старый author). */
function pickCommitterAuthorTimestamp(ref) {
  if (!ref || typeof ref !== "object") {
    return "";
  }

  const author = ref.author ?? ref.Author;
  const committer = ref.committer ?? ref.Committer;

  return normalizeIsoDate(
    committer?.date ?? committer?.Date ?? author?.date ?? author?.Date ?? "",
  ) ?? "";
}

async function fetchPushTimestamp(config, project, repo, pushId) {
  const normalizedPushId = String(pushId ?? "").trim();

  if (!normalizedPushId) {
    return "";
  }

  const cacheKey = pushTimestampCacheKey(config, project, repo, normalizedPushId);

  if (pushTimestampCache.has(cacheKey)) {
    return pushTimestampCache.get(cacheKey);
  }

  const projectSeg = encodeURIComponent(String(project).trim());
  const repoSeg = encodeURIComponent(String(repo).trim());
  const pushSeg = encodeURIComponent(normalizedPushId);
  const apiVersion = encodeURIComponent(resolveApiVersion(config));
  const path = `${projectSeg}/_apis/git/repositories/${repoSeg}/pushes/${pushSeg}?api-version=${apiVersion}`;
  const push = await adoFetch(config, path);
  const timestamp = normalizeIsoDate(push?.date ?? push?.Date ?? "") ?? "";

  pushTimestampCache.set(cacheKey, timestamp);
  return timestamp;
}

/** Время пуша: inline `push.date`, иначе `GET .../pushes/{pushId}`. Без fallback на author. */
async function resolvePushTimestampOnly(config, project, repo, ref) {
  const { date, pushId } = pickPushMeta(ref);

  if (date) {
    return date;
  }

  if (pushId == null) {
    return "";
  }

  try {
    return await fetchPushTimestamp(config, project, repo, pushId);
  } catch (error) {
    logAdoError(`fetchPushTimestamp ${pushId}`, error);
    return "";
  }
}

async function fetchCommitTimestamp(config, project, repo, commitId) {
  const cacheKey = commitTimestampCacheKey(config, project, repo, commitId);

  if (commitTimestampCache.has(cacheKey)) {
    return commitTimestampCache.get(cacheKey);
  }

  const projectSeg = encodeURIComponent(String(project).trim());
  const repoSeg = encodeURIComponent(String(repo).trim());
  const commitSeg = encodeURIComponent(String(commitId).trim());
  const apiVersion = encodeURIComponent(resolveApiVersion(config));
  const path = `${projectSeg}/_apis/git/repositories/${repoSeg}/commits/${commitSeg}?api-version=${apiVersion}`;
  const commit = await adoFetch(config, path);
  const pushTimestamp = await resolvePushTimestampOnly(config, project, repo, commit);
  const timestamp = pushTimestamp || pickCommitterAuthorTimestamp(commit);

  commitTimestampCache.set(cacheKey, timestamp);
  return timestamp;
}

/**
 * Один PR с полным description и (при includeCommits) списком коммитов с датами.
 *
 * @param {import("./ado-config.mjs").DEFAULT_ADO_CONFIG} config
 * @param {string} project
 * @param {string} repo
 * @param {number | string} pullRequestId
 */
async function fetchGitPullRequestById(config, project, repo, pullRequestId) {
  const projectSeg = encodeURIComponent(String(project).trim());
  const repoSeg = encodeURIComponent(String(repo).trim());
  const prIdSeg = encodeURIComponent(String(pullRequestId).trim());
  const query = new URLSearchParams({
    includeCommits: "true",
    "api-version": resolveApiVersion(config),
  });
  const path = `${projectSeg}/_apis/git/repositories/${repoSeg}/pullrequests/${prIdSeg}?${query.toString()}`;
  return adoFetch(config, path);
}

/**
 * `createdDate` последней итерации PR — момент последнего push в ветку PR.
 *
 * @param {import("./ado-config.mjs").DEFAULT_ADO_CONFIG} config
 * @param {string} project
 * @param {string} repo
 * @param {number | string} pullRequestId
 */
async function fetchLatestPullRequestIterationDate(config, project, repo, pullRequestId) {
  const projectSeg = encodeURIComponent(String(project).trim());
  const repoSeg = encodeURIComponent(String(repo).trim());
  const prIdSeg = encodeURIComponent(String(pullRequestId).trim());
  const query = new URLSearchParams({
    "api-version": resolveApiVersion(config),
  });
  const path = `${projectSeg}/_apis/git/repositories/${repoSeg}/pullRequests/${prIdSeg}/iterations?${query.toString()}`;
  const data = await adoFetch(config, path);
  const iterations = Array.isArray(data?.value) ? data.value : [];
  let latest = "";
  let latestTimestamp = Number.NEGATIVE_INFINITY;

  for (const iteration of iterations) {
    const createdAt = normalizeIsoDate(
      iteration?.createdDate ?? iteration?.CreatedDate ?? "",
    );

    if (!createdAt) {
      continue;
    }

    const timestamp = Date.parse(createdAt);

    if (!Number.isFinite(timestamp) || timestamp <= latestTimestamp) {
      continue;
    }

    latest = createdAt;
    latestTimestamp = timestamp;
  }

  return latest;
}

/**
 * Время пуша source: push (inline / GET push), GET commit, в конце — committer/author.
 */
async function resolveLastCommitAtFromPrDetail(config, detail, listPr, project, repo) {
  const sourceId = String(
    detail?.lastMergeSourceCommit?.commitId
      ?? detail?.lastMergeSourceCommit?.CommitId
      ?? listPr?.lastMergeSourceCommit?.commitId
      ?? listPr?.lastMergeSourceCommit?.CommitId
      ?? "",
  ).trim();

  if (!sourceId) {
    return "";
  }

  const commits = Array.isArray(detail?.commits) ? detail.commits : [];
  const hit = commits.find(
    (c) => String(c?.commitId ?? c?.CommitId ?? "").toLowerCase() === sourceId.toLowerCase(),
  );
  const commitRefs = [hit, detail?.lastMergeSourceCommit].filter(
    (ref) => ref && typeof ref === "object",
  );

  for (const ref of commitRefs) {
    const pushTimestamp = await resolvePushTimestampOnly(config, project, repo, ref);

    if (pushTimestamp) {
      return pushTimestamp;
    }
  }

  try {
    const fromCommit = await fetchCommitTimestamp(config, project, repo, sourceId);

    if (fromCommit) {
      return fromCommit;
    }
  } catch (_error) {
    // fallback на даты из PR ниже
  }

  for (const ref of commitRefs) {
    const commitTimestamp = pickCommitterAuthorTimestamp(ref);

    if (commitTimestamp) {
      return commitTimestamp;
    }
  }

  return "";
}

function groupMemberIdsCacheKey(config, groupIds) {
  return [
    normalizeApiRoot(config.apiRoot),
    resolveApiVersion(config),
    [...groupIds].sort().join(","),
  ].join("|");
}

function isIdentityUser(identity) {
  if (!identity || typeof identity !== "object") {
    return false;
  }

  if (identity.isContainer === true) {
    return false;
  }

  const schemaClassName = getIdentityProperty(identity, "SchemaClassName");

  if (schemaClassName) {
    return schemaClassName === "User";
  }

  return !identity.isContainer;
}

/**
 * User id всех участников выбранных reviewer-групп (ExpandedDown через Identities API).
 *
 * @param {import("./ado-config.mjs").DEFAULT_ADO_CONFIG} config
 * @returns {Promise<Set<string>>}
 */
export async function resolveConfiguredGroupMemberIds(config) {
  const groupIds = normalizeConfiguredGroupIds(config);

  if (groupIds.length === 0) {
    return new Set();
  }

  const cacheKey = groupMemberIdsCacheKey(config, groupIds);
  const cached = groupMemberIdsCache.get(cacheKey);

  if (cached && Date.now() - cached.updatedAt <= GROUP_MEMBER_IDS_CACHE_TTL_MS) {
    return cached.memberIds;
  }

  const memberIds = new Set();

  for (const groupId of groupIds) {
    try {
      const query = new URLSearchParams({
        identityIds: groupId,
        queryMembership: "ExpandedDown",
        "api-version": resolveApiVersion(config),
      });
      const data = await adoFetchIdentitiesApi(config, query);
      const identities = Array.isArray(data?.value) ? data.value : [];
      let foundUsers = false;

      for (const identity of identities) {
        if (!isIdentityUser(identity)) {
          continue;
        }

        const id = normalizePlainText(identity?.id);

        if (id) {
          memberIds.add(id);
          foundUsers = true;
        }
      }

      if (foundUsers) {
        continue;
      }

      const groupIdentity = identities.find(
        (identity) => normalizePlainText(identity?.id) === groupId,
      );
      const nestedMemberIds = [
        ...(Array.isArray(groupIdentity?.memberIds) ? groupIdentity.memberIds : []),
        ...(Array.isArray(groupIdentity?.members) ? groupIdentity.members : [])
          .flatMap((entry) => getDescriptorCandidates(entry)),
      ]
        .map((value) => normalizePlainText(value))
        .filter(Boolean);

      if (nestedMemberIds.length === 0) {
        continue;
      }

      const resolvedMembers = await resolveIdentityDescriptors(config, nestedMemberIds);

      for (const identity of resolvedMembers) {
        if (!isIdentityUser(identity)) {
          continue;
        }

        const id = normalizePlainText(identity?.id);

        if (id) {
          memberIds.add(id);
        }
      }
    } catch (error) {
      logAdoError(`resolveConfiguredGroupMemberIds ${groupId}`, error);
    }
  }

  groupMemberIdsCache.set(cacheKey, {
    memberIds,
    updatedAt: Date.now(),
  });

  return memberIds;
}

/**
 * @param {import("./ado-config.mjs").DEFAULT_ADO_CONFIG} config
 * @param {string} project
 * @param {string} repo
 * @param {number | string} pullRequestId
 */
async function fetchGitPullRequestThreads(config, project, repo, pullRequestId) {
  const projectSeg = encodeURIComponent(String(project).trim());
  const repoSeg = encodeURIComponent(String(repo).trim());
  const prIdSeg = encodeURIComponent(String(pullRequestId).trim());
  const query = new URLSearchParams({
    "api-version": resolveApiVersion(config),
  });
  const path = `${projectSeg}/_apis/git/repositories/${repoSeg}/pullRequests/${prIdSeg}/threads?${query.toString()}`;
  const data = await adoFetch(config, path);
  return Array.isArray(data?.value) ? data.value : [];
}

function isUserPullRequestComment(comment) {
  if (!comment || typeof comment !== "object" || comment.isDeleted === true) {
    return false;
  }

  const commentType = normalizePlainText(comment.commentType ?? comment.CommentType).toLowerCase();

  if (commentType === "system") {
    return false;
  }

  const author = comment.author ?? comment.Author;

  if (!author || typeof author !== "object" || author.isContainer === true) {
    return false;
  }

  return Boolean(normalizePlainText(author.id ?? author.Id));
}

/**
 * Последний комментарий участника reviewer-группы, который **открывает тред**
 * (`comments[0]`). Ответы внутри ветки не учитываются.
 *
 * @param {Array<any>} threads
 * @param {Set<string>} memberIds
 */
function resolveLatestGroupMemberComment(threads, memberIds) {
  if (memberIds.size === 0 || !Array.isArray(threads)) {
    return "";
  }

  let latest = "";
  let latestTimestamp = Number.NEGATIVE_INFINITY;

  for (const thread of threads) {
    if (thread?.isDeleted === true) {
      continue;
    }

    const comments = Array.isArray(thread?.comments) ? thread.comments : [];
    const comment = comments[0];

    if (!isUserPullRequestComment(comment)) {
      continue;
    }

    const authorId = normalizePlainText(
      comment.author?.id ?? comment.Author?.Id ?? "",
    );

    if (!memberIds.has(authorId)) {
      continue;
    }

    const publishedAt = normalizeIsoDate(
      comment.publishedDate ?? comment.PublishedDate ?? "",
    );

    if (!publishedAt) {
      continue;
    }

    const publishedTimestamp = Date.parse(publishedAt);

    if (!Number.isFinite(publishedTimestamp)) {
      continue;
    }

    if (publishedTimestamp > latestTimestamp) {
      latest = publishedAt;
      latestTimestamp = publishedTimestamp;
    }
  }

  return latest;
}

/**
 * Обогащает PR полным description, временем пуша source и (опционально) последним
 * комментарием участника reviewer-группы.
 *
 * @param {import("./ado-config.mjs").DEFAULT_ADO_CONFIG} config
 * @param {Array<any>} pullRequests
 * @param {Set<string>|Iterable<string>|null|undefined} [groupMemberIds]
 */
export async function attachPullRequestLastCommitTimes(config, pullRequests, groupMemberIds = null) {
  const memberIds = groupMemberIds instanceof Set
    ? groupMemberIds
    : new Set(groupMemberIds ?? []);

  return Promise.all(
    pullRequests.map(async (pullRequest) => {
      const prId = pullRequest?.pullRequestId;

      if (prId == null) {
        return pullRequest;
      }

      const project = pullRequest?.repository?.project?.id
        ?? pullRequest?.repository?.project?.name
        ?? config.project;
      const repo = pullRequest?.repository?.id
        ?? pullRequest?.repository?.name
        ?? config.repositoryId;

      try {
        const detail = await fetchGitPullRequestById(config, project, repo, prId);

        if (!detail || typeof detail !== "object") {
          return pullRequest;
        }

        const [lastCommitAtFromDetail, lastIterationAt] = await Promise.all([
          resolveLastCommitAtFromPrDetail(
            config,
            detail,
            pullRequest,
            project,
            repo,
          ),
          fetchLatestPullRequestIterationDate(config, project, repo, prId).catch((error) => {
            logAdoError(`fetchLatestPullRequestIterationDate ${prId}`, error);
            return "";
          }),
        ]);
        const lastCommitAt = pickLatestIsoDate(lastCommitAtFromDetail, lastIterationAt);

        const next = { ...pullRequest };

        if (typeof detail.description === "string") {
          next.description = detail.description;
        }

        if (lastCommitAt) {
          next.lastCommitAt = lastCommitAt;
        }

        if (memberIds.size > 0) {
          try {
            const threads = await fetchGitPullRequestThreads(config, project, repo, prId);
            const lastGroupCommentAt = resolveLatestGroupMemberComment(threads, memberIds);

            if (lastGroupCommentAt) {
              next.lastGroupCommentAt = lastGroupCommentAt;
            }
          } catch (error) {
            logAdoError(`fetchGitPullRequestThreads ${prId}`, error);
          }
        }

        return next;
      } catch (error) {
        logAdoError(`fetchGitPullRequestById ${prId}`, error);
        return pullRequest;
      }
    }),
  );
}

function normalizePlainText(value) {
  return String(value ?? "")
    .replace(/\s+/g, " ")
    .trim();
}

/**
 * Абсолютный URL для загрузки в popup расширения (img src).
 */
function resolveExtensionAssetUrl(apiRoot, raw) {
  const url = normalizePlainText(raw);

  if (!url) {
    return "";
  }

  if (/^https?:\/\//i.test(url)) {
    return url;
  }

  if (url.startsWith("//")) {
    return `https:${url}`;
  }

  if (url.startsWith("/")) {
    try {
      const origin = new URL(normalizeApiRoot(apiRoot)).origin;
      return `${origin}${url}`;
    } catch (_error) {
      return url;
    }
  }

  return url;
}

/**
 * У Azure DevOps часто два источника аватара: GraphProfile в _links.avatar и identityImage.
 * identityImage для части пользователей отдаёт плейсхолдер — приоритет у ссылки из _links.
 */
function pickPullRequestAuthorAvatarUrl(createdBy, apiRoot) {
  if (!createdBy || typeof createdBy !== "object") {
    return "";
  }

  const links = createdBy._links ?? {};
  const avatarHref = normalizePlainText(
    links.avatar?.href ?? links.Avatar?.href ?? "",
  );
  const imageUrl = normalizePlainText(
    createdBy.imageUrl ?? createdBy.ImageUrl ?? "",
  );
  const raw = avatarHref || imageUrl;

  return resolveExtensionAssetUrl(apiRoot, raw);
}

function normalizeDescription(value) {
  return String(value ?? "")
    .replace(/\\n/g, "\n")
    .replace(/[ \t]+/g, " ")
    .trim();
}

function normalizeIsoDate(value) {
  if (typeof value !== "string" || !value.trim()) {
    return null;
  }

  const dotNetMatch = value.match(/\/Date\(([-+]?\d+)/i);

  if (dotNetMatch) {
    const timestamp = Number(dotNetMatch[1]);

    if (Number.isFinite(timestamp)) {
      return new Date(timestamp).toISOString();
    }
  }

  const date = new Date(value);

  if (Number.isNaN(date.getTime())) {
    return null;
  }

  return date.toISOString();
}

function buildPullRequestWebUrl(apiRoot, project, repo, pullRequestId, pr) {
  const fromLinks = pr?._links?.web?.href || pr?._links?.html?.href;

  if (typeof fromLinks === "string" && fromLinks.startsWith("http")) {
    return fromLinks;
  }

  const encProject = encodeURIComponent(project);
  const encRepo = encodeURIComponent(repo);
  return `${apiRoot}/${encProject}/_git/${encRepo}/pullrequest/${pullRequestId}`;
}

function getIdentityProperty(identity, key) {
  return normalizePlainText(identity?.properties?.[key]?.$value ?? "");
}

function firstNonEmpty(...values) {
  for (const value of values) {
    const normalized = normalizePlainText(value);

    if (normalized) {
      return normalized;
    }
  }

  return "";
}

function getDescriptorCandidates(entry) {
  if (typeof entry === "string") {
    const normalized = normalizePlainText(entry);
    return normalized ? [normalized] : [];
  }

  const candidates = [
    entry?.subjectDescriptor,
    entry?.descriptor,
    entry?.identifier,
    entry?.value,
    entry?.id,
  ]
    .map((value) => normalizePlainText(value))
    .filter(Boolean);

  return [...new Set(candidates)];
}

function getIdentityLookupQueryField(value) {
  const normalized = normalizePlainText(value);

  if (!normalized) {
    return null;
  }

  if (/^vss/i.test(normalized)) {
    return "subjectDescriptors";
  }

  if (/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(normalized)) {
    return "identityIds";
  }

  return "descriptors";
}

function mapIdentityGroup(identity) {
  const id = normalizePlainText(identity?.id);

  if (!id) {
    return null;
  }

  const name = firstNonEmpty(
    identity?.providerDisplayName,
    identity?.customDisplayName,
    getIdentityProperty(identity, "Account"),
    identity?.uniqueName,
    id,
  );

  return {
    id,
    name: name || id,
    uniqueName: normalizePlainText(identity?.uniqueName),
    description: getIdentityProperty(identity, "Description"),
    scopeName: getIdentityProperty(identity, "ScopeName"),
    scopeType: getIdentityProperty(identity, "ScopeType"),
    schemaClassName: getIdentityProperty(identity, "SchemaClassName"),
    securityGroupKind: getIdentityProperty(identity, "SecurityGroup"),
    descriptor: normalizePlainText(identity?.descriptor),
    isContainer: Boolean(identity?.isContainer),
    isActive: identity?.isActive !== false,
  };
}

function isIdentityGroup(group) {
  return Boolean(group?.id)
    && group.isContainer
    && (
      group.schemaClassName === "Group"
      || group.securityGroupKind === "SecurityGroup"
    );
}

function isLikelyReviewerGroup(group) {
  if (!isIdentityGroup(group)) {
    return false;
  }

  const uniqueName = group.uniqueName;

  return group.scopeType === "TeamProject"
    || uniqueName.startsWith("vstfs:///Classification/TeamProject/")
    || /\\/.test(group.name);
}

function sortIdentityGroups(groups) {
  return [...groups].sort((left, right) => {
    const byName = left.name.localeCompare(right.name, "ru", { sensitivity: "base" });
    return byName || left.id.localeCompare(right.id);
  });
}

function dedupeIdentityGroups(groups) {
  const byId = new Map();

  for (const group of groups) {
    if (!group?.id) {
      continue;
    }

    const existing = byId.get(group.id);

    if (!existing || (!existing.description && group.description)) {
      byId.set(group.id, group);
    }
  }

  return [...byId.values()];
}

function membershipCacheKey(config, userId) {
  return [
    normalizeApiRoot(config.apiRoot),
    resolveApiVersion(config),
    String(userId),
  ].join("|");
}

function getCachedReviewerGroups(config, userId) {
  const entry = reviewerGroupsCache.get(membershipCacheKey(config, userId));

  if (!entry) {
    return null;
  }

  if (Date.now() - entry.updatedAt > REVIEWER_GROUPS_CACHE_TTL_MS) {
    reviewerGroupsCache.delete(membershipCacheKey(config, userId));
    return null;
  }

  return entry.groups;
}

function setCachedReviewerGroups(config, userId, groups) {
  reviewerGroupsCache.set(membershipCacheKey(config, userId), {
    groups,
    updatedAt: Date.now(),
  });
}

/**
 * На части Azure DevOps Server endpoint identities чувствителен к пути/регистру.
 * Сначала пробуем рабочий для пользователя вариант `/_apis/identities`.
 *
 * @param {import("./ado-config.mjs").DEFAULT_ADO_CONFIG} config
 * @param {URLSearchParams} query
 */
async function adoFetchIdentitiesApi(config, query) {
  const suffix = query.toString();
  const candidates = [
    `_apis/identities?${suffix}`,
    `_apis/Identities?${suffix}`,
  ];
  let lastError = null;

  for (const path of candidates) {
    try {
      return await adoFetch(config, path);
    } catch (error) {
      lastError = error;
      const message = error instanceof Error ? error.message : String(error);

      if (!/\(404\)|\b404\b/.test(message)) {
        throw error;
      }
    }
  }

  throw lastError ?? new Error("Identities API: не удалось выполнить запрос.");
}

/**
 * @param {import("./ado-config.mjs").DEFAULT_ADO_CONFIG} config
 * @param {string} userId
 */
async function fetchIdentityMembershipDescriptors(config, userId) {
  const query = new URLSearchParams({
    identityIds: String(userId),
    queryMembership: "Expanded",
    "api-version": resolveApiVersion(config),
  });

  let data;

  try {
    data = await adoFetchIdentitiesApi(config, query);
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    throw new Error(`Identities memberships: ${message}`);
  }

  const identities = Array.isArray(data?.value) ? data.value : [];
  const descriptors = new Set();

  for (const identity of identities) {
    const memberOf = Array.isArray(identity?.memberOf) ? identity.memberOf : [];

    for (const entry of memberOf) {
      const candidates = getDescriptorCandidates(entry);

      for (const descriptor of candidates) {
        descriptors.add(descriptor);
      }
    }
  }

  return [...descriptors];
}

/**
 * @param {import("./ado-config.mjs").DEFAULT_ADO_CONFIG} config
 * @param {string[]} descriptors
 */
async function resolveIdentityDescriptors(config, descriptors) {
  const resolved = [];
  const grouped = new Map();

  for (const descriptor of descriptors) {
    const field = getIdentityLookupQueryField(descriptor);

    if (!field) {
      continue;
    }

    if (!grouped.has(field)) {
      grouped.set(field, []);
    }

    grouped.get(field).push(descriptor);
  }

  for (const [field, values] of grouped.entries()) {
    for (let offset = 0; offset < values.length; offset += IDENTITY_BATCH_SIZE) {
      const batch = values.slice(offset, offset + IDENTITY_BATCH_SIZE);

      if (batch.length === 0) {
        continue;
      }

      let data = null;

      try {
        const query = new URLSearchParams({
          [field]: batch.join(","),
          queryMembership: "None",
          "api-version": resolveApiVersion(config),
        });

        data = await adoFetchIdentitiesApi(config, query);
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error);

        if (!/\(404\)|\b404\b/.test(message)) {
          throw error;
        }

        const byIdentityIdsQuery = new URLSearchParams({
          identityIds: batch.join(","),
          queryMembership: "None",
          "api-version": resolveApiVersion(config),
        });

        try {
          data = await adoFetchIdentitiesApi(config, byIdentityIdsQuery);
        } catch (fallbackError) {
          const fallbackMessage = fallbackError instanceof Error
            ? fallbackError.message
            : String(fallbackError);
          throw new Error(`Identities resolve ${field}: ${fallbackMessage}`);
        }
      }

      const batchValues = Array.isArray(data?.value) ? data.value : [];
      resolved.push(...batchValues);
    }
  }

  return resolved;
}

function normalizeConfiguredGroupIds(config) {
  return Array.isArray(config.selectedGroupIds)
    ? config.selectedGroupIds.map((id) => normalizePlainText(id)).filter(Boolean)
    : [];
}

/**
 * Возвращает группы пользователя, релевантные для reviewer picker:
 * это memberships из `Identities API`, отфильтрованные до TeamProject group identities.
 *
 * @param {import("./ado-config.mjs").DEFAULT_ADO_CONFIG} config
 * @param {string | null} [currentUserId]
 * @returns {Promise<{ groups: Array<{
 *   id: string,
 *   name: string,
 *   uniqueName: string,
 *   description: string,
 *   scopeName: string,
 *   scopeType: string,
 *   schemaClassName: string,
 *   securityGroupKind: string,
 *   descriptor: string,
 *   isContainer: boolean,
 *   isActive: boolean,
 * }>, mode: "membership" | "cache" | "empty" | "error", note?: string, error?: Error }>}
 */
export async function fetchMyReviewerGroupsWithDiagnostics(config, currentUserId = null) {
  let userId = currentUserId ? String(currentUserId) : "";

  if (!userId) {
    const identity = await fetchConnectionIdentity(config);
    userId = identity.id;
  }

  const cached = getCachedReviewerGroups(config, userId);

  if (cached) {
    return {
      groups: cached,
      mode: "cache",
    };
  }

  try {
    const descriptors = await fetchIdentityMembershipDescriptors(config, userId);

    if (descriptors.length === 0) {
      return {
        groups: [],
        mode: "empty",
        note: "Memberships пользователя не найдены в Identities API.",
      };
    }

    const identities = await resolveIdentityDescriptors(config, descriptors);
    const groups = sortIdentityGroups(
      dedupeIdentityGroups(
        identities
          .map((identity) => mapIdentityGroup(identity))
          .filter((group) => isLikelyReviewerGroup(group)),
      ),
    );

    if (groups.length === 0) {
      return {
        groups: [],
        mode: "empty",
        note: "Memberships получены, но reviewer-группы среди них не выделены.",
      };
    }

    setCachedReviewerGroups(config, userId, groups);

    return {
      groups,
      mode: "membership",
    };
  } catch (error) {
    const err = error instanceof Error ? error : new Error(String(error));
    logAdoError("fetchMyReviewerGroupsWithDiagnostics", err);

    return {
      groups: [],
      mode: "error",
      error: err,
    };
  }
}

function buildIdentitySearchTerms(filterValue) {
  const normalized = normalizePlainText(filterValue);

  if (!normalized) {
    return [];
  }

  const terms = [normalized];
  const tokens = normalized
    .split(/[^\p{L}\p{N}]+/u)
    .map((token) => token.trim())
    .filter((token) => token.length >= 2);

  if (tokens.length >= 2) {
    terms.push(tokens.slice(0, 2).join(" "));
    terms.push(tokens.slice(-2).join(" "));
  }

  for (const token of tokens) {
    terms.push(token);
  }

  return [...new Set(terms)];
}

function groupMatchesSearch(group, filterValue) {
  const normalized = normalizePlainText(filterValue).toLowerCase();

  if (!normalized) {
    return true;
  }

  const searchable = [
    group?.name,
    group?.description,
    group?.scopeName,
    group?.uniqueName,
    group?.id,
  ]
    .filter(Boolean)
    .join("\n")
    .toLowerCase();

  if (searchable.includes(normalized)) {
    return true;
  }

  const tokens = normalized
    .split(/[^\p{L}\p{N}]+/u)
    .map((token) => token.trim())
    .filter((token) => token.length >= 2);

  return tokens.length > 0 && tokens.every((token) => searchable.includes(token));
}

/**
 * Ручной поиск reviewer-групп через совместимый с on-prem `searchFilter=General`.
 *
 * @param {import("./ado-config.mjs").DEFAULT_ADO_CONFIG} config
 * @param {string} filterValue
 * @returns {Promise<{ groups: Array<{
 *   id: string,
 *   name: string,
 *   uniqueName: string,
 *   description: string,
 *   scopeName: string,
 *   scopeType: string,
 *   schemaClassName: string,
 *   securityGroupKind: string,
 *   descriptor: string,
 *   isContainer: boolean,
 *   isActive: boolean,
 * }>, mode: "search" | "empty" | "error", note?: string, error?: Error }>}
 */
export async function searchReviewerGroupsByName(config, filterValue) {
  const queryText = normalizePlainText(filterValue);

  if (!queryText) {
    return {
      groups: [],
      mode: "empty",
      note: "Введите часть имени группы и нажмите «Найти группы».",
    };
  }

  const searchTerms = buildIdentitySearchTerms(queryText);

  try {
    const batches = await Promise.all(
      searchTerms.map(async (term) => {
        const query = new URLSearchParams({
          searchFilter: "General",
          filterValue: term,
          queryMembership: "None",
          "api-version": resolveApiVersion(config),
        });

        const data = await adoFetchIdentitiesApi(config, query);
        return Array.isArray(data?.value) ? data.value : [];
      }),
    );

    const identities = batches.flat();
    const groups = sortIdentityGroups(
      dedupeIdentityGroups(
        identities
          .map((identity) => mapIdentityGroup(identity))
          .filter((group) => isLikelyReviewerGroup(group)),
      ),
    ).filter((group) => groupMatchesSearch(group, queryText));

    if (groups.length === 0) {
      return {
        groups: [],
        mode: "empty",
        note: "По этому запросу reviewer-группы не найдены.",
      };
    }

    return {
      groups,
      mode: "search"
    };
  } catch (error) {
    const err = error instanceof Error ? error : new Error(String(error));
    logAdoError("searchReviewerGroupsByName", err);
    return {
      groups: [],
      mode: "error",
      error: err,
    };
  }
}

/**
 * @param {import("./ado-config.mjs").DEFAULT_ADO_CONFIG} config
 * @param {string} myId
 * @returns {{ allowedReviewerIds: string[], matchedSectionTitle: string }}
 */
export function getExtensionReviewerContext(config, myId) {
  const currentUserId = String(myId);
  const selectedGroupIds = normalizeConfiguredGroupIds(config);

  if (selectedGroupIds.length > 0) {
    return {
      allowedReviewerIds: [currentUserId, ...selectedGroupIds],
      matchedSectionTitle: "Выбранные reviewer-группы",
    };
  }

  return {
    allowedReviewerIds: [currentUserId],
    matchedSectionTitle: "Назначено мне",
  };
}

function reviewerHasNoVote(reviewer) {
  return Number(reviewer?.vote ?? 0) === 0;
}

/** ADO: 10 — Approved, 5 — Approved with suggestions. */
function reviewerHasApprovedVote(reviewer) {
  return Number(reviewer?.vote ?? 0) >= 5;
}

/**
 * Выбранные reviewer-группы на PR, иначе пустой массив.
 *
 * @param {Array<any>} reviewers
 * @param {string[]} selectedGroupIds
 * @returns {Array<any>}
 */
function getSelectedGroupReviewers(reviewers, selectedGroupIds) {
  const groupIdSet = new Set(selectedGroupIds);

  if (groupIdSet.size === 0) {
    return [];
  }

  return reviewers.filter((reviewer) => groupIdSet.has(String(reviewer?.id ?? "")));
}

/**
 * Нужна ли реакция на Review.
 * Если на PR есть выбранные reviewer-группы — решают только их голоса:
 * апрув группы убирает карточку из ожидающих даже при личном `vote === 0`
 * (в ADO участника часто добавляют optional-ревьюером).
 * Иначе смотрим только голос текущего пользователя.
 *
 * @param {Array<any>} reviewers
 * @param {string} currentUserId
 * @param {string[]} selectedGroupIds
 */
function hasPendingAllowedReviewerVote(reviewers, currentUserId, selectedGroupIds) {
  const selectedGroupReviewers = getSelectedGroupReviewers(reviewers, selectedGroupIds);

  if (selectedGroupReviewers.length > 0) {
    return selectedGroupReviewers.some(reviewerHasNoVote);
  }

  return reviewers.some(
    (reviewer) => (
      String(reviewer?.id ?? "") === String(currentUserId)
      && reviewerHasNoVote(reviewer)
    ),
  );
}

/**
 * Уже одобрено «своими» ревьюерами: все выбранные группы на PR с vote >= 5,
 * иначе личный апрув текущего пользователя.
 *
 * @param {Array<any>} reviewers
 * @param {string} currentUserId
 * @param {string[]} selectedGroupIds
 */
function hasApprovedAllowedReviewerVote(reviewers, currentUserId, selectedGroupIds) {
  const selectedGroupReviewers = getSelectedGroupReviewers(reviewers, selectedGroupIds);

  if (selectedGroupReviewers.length > 0) {
    return selectedGroupReviewers.every(reviewerHasApprovedVote);
  }

  return reviewers.some(
    (reviewer) => (
      String(reviewer?.id ?? "") === String(currentUserId)
      && reviewerHasApprovedVote(reviewer)
    ),
  );
}

/**
 * @param {import("./ado-config.mjs").DEFAULT_ADO_CONFIG} config
 * @param {Array<any>} pullRequests
 * @param {string} myId
 */
export async function filterPullRequestsForExtension(config, pullRequests, myId) {
  const { matchedSectionTitle } = getExtensionReviewerContext(config, myId);
  const selectedGroupIds = normalizeConfiguredGroupIds(config);
  const filtered = [];
  const approved = [];

  for (const pullRequest of pullRequests) {
    if (!isVisiblePullRequestForExtension(pullRequest)) {
      continue;
    }

    const reviewers = Array.isArray(pullRequest?.reviewers) ? pullRequest.reviewers : [];

    if (hasPendingAllowedReviewerVote(reviewers, myId, selectedGroupIds)) {
      filtered.push(pullRequest);
      continue;
    }

    if (hasApprovedAllowedReviewerVote(reviewers, myId, selectedGroupIds)) {
      approved.push(pullRequest);
    }
  }

  return { filtered, approved, matchedSectionTitle };
}

function isVisiblePullRequestForExtension(pullRequest) {
  if (pullRequest?.isDraft === true) {
    return false;
  }

  return isActivePullRequestStatus(pullRequest?.status);
}

function isActivePullRequestStatus(status) {
  if (status === undefined || status === null) {
    return true;
  }

  if (typeof status === "number") {
    return status === 1;
  }

  return String(status).toLowerCase() === "active";
}

function normalizeTargetBranch(refName) {
  const raw = normalizePlainText(refName);

  if (!raw) {
    return "";
  }

  return raw.replace(/^refs\/heads\//i, "");
}

export function mapPullRequestToItem(pr, config) {
  const apiRoot = normalizeApiRoot(config.apiRoot);
  const project = config.project.trim();
  const repo = config.repositoryId.trim();
  const id = String(pr?.pullRequestId ?? "").trim();

  if (!id) {
    return null;
  }

  const title = normalizePlainText(pr?.title ?? "");
  const author = normalizePlainText(pr?.createdBy?.displayName ?? "");
  const avatarUrl = pickPullRequestAuthorAvatarUrl(pr?.createdBy, config.apiRoot);
  const createdAt = normalizeIsoDate(pr?.creationDate);
  const closedAt = normalizeIsoDate(pr?.closedDate);
  const status = normalizePullRequestStatus(pr?.status);
  const updatedAt = pickLatestIsoDate(
    normalizeIsoDate(pr?.lastGroupCommentAt),
    normalizeIsoDate(pr?.lastCommitAt),
    closedAt,
    createdAt,
  );
  const description = normalizeDescription(pr?.description ?? "");
  const url = buildPullRequestWebUrl(apiRoot, project, repo, id, pr);
  const targetBranch = normalizeTargetBranch(
    pr?.targetRefName ?? pr?.TargetRefName ?? "",
  );
  /** @type {string[] | null | undefined} */
  let blockingReasons;
  /** @type {string[] | null | undefined} */
  let optionalPolicyReasons;

  if (pr?.blockingReasons === null) {
    blockingReasons = null;
  } else if (Array.isArray(pr?.blockingReasons)) {
    blockingReasons = pr.blockingReasons
      .map((reason) => normalizePlainText(reason))
      .filter(Boolean);
  }

  if (pr?.optionalPolicyReasons === null) {
    optionalPolicyReasons = null;
  } else if (Array.isArray(pr?.optionalPolicyReasons)) {
    optionalPolicyReasons = pr.optionalPolicyReasons
      .map((reason) => normalizePlainText(reason))
      .filter(Boolean);
  }

  const rawConflictText = typeof pr?.conflictText === "string" ? pr.conflictText.trim() : "";
  const conflictText = rawConflictText || "";

  if (!title || !url) {
    return null;
  }

  return {
    id,
    title,
    author,
    avatarUrl,
    createdAt,
    updatedAt,
    status,
    lastCommitAt: normalizeIsoDate(pr?.lastCommitAt) || undefined,
    lastGroupCommentAt: normalizeIsoDate(pr?.lastGroupCommentAt) || undefined,
    description,
    url,
    ...(closedAt ? { closedAt } : {}),
    ...(targetBranch ? { targetBranch } : {}),
    ...(blockingReasons !== undefined ? { blockingReasons } : {}),
    ...(optionalPolicyReasons !== undefined ? { optionalPolicyReasons } : {}),
    ...(conflictText ? { conflictText } : {}),
  };
}

/**
 * @param {unknown} status
 * @returns {"active" | "abandoned" | "completed" | "unknown"}
 */
function normalizePullRequestStatus(status) {
  if (status === undefined || status === null || status === "") {
    return "active";
  }

  if (typeof status === "number") {
    if (status === 1) {
      return "active";
    }

    if (status === 2) {
      return "abandoned";
    }

    if (status === 3) {
      return "completed";
    }

    return "unknown";
  }

  const normalized = String(status).toLowerCase();

  if (normalized === "active" || normalized === "abandoned" || normalized === "completed") {
    return normalized;
  }

  return "unknown";
}

function pickLatestIsoDate(...values) {
  let latest = null;
  let latestTimestamp = Number.NEGATIVE_INFINITY;

  for (const value of values) {
    if (!value) {
      continue;
    }

    const timestamp = Date.parse(value);

    if (!Number.isFinite(timestamp)) {
      continue;
    }

    if (timestamp > latestTimestamp) {
      latest = value;
      latestTimestamp = timestamp;
    }
  }

  return latest;
}

export { sortPullRequestsOldestFirst } from "./working-time.mjs";

/**
 * @param {import("./ado-config.mjs").DEFAULT_ADO_CONFIG} config
 */
export async function setReviewerVoteApprove(config, pullRequestId, reviewerId) {
  const project = encodeURIComponent(config.project.trim());
  const repo = encodeURIComponent(config.repositoryId.trim());
  const prId = encodeURIComponent(String(pullRequestId).trim());
  const revId = encodeURIComponent(String(reviewerId).trim());
  const apiVersion = encodeURIComponent(resolveApiVersion(config));

  const path = `${project}/_apis/git/repositories/${repo}/pullRequests/${prId}/reviewers/${revId}?api-version=${apiVersion}`;
  const body = JSON.stringify({
    vote: 10,
    id: reviewerId,
  });

  await adoFetch(config, path, {
    method: "PUT",
    body,
  });
}

export function logAdoError(context, error) {
  const message = error instanceof Error ? error.message : String(error);
  console.error(`[PR Notify] ${context}:`, message);
}
