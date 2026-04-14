import {
  normalizeApiRoot,
  parseTeamReviewerIds,
  resolveApiVersion,
} from "./ado-config.mjs";

const PAGE_SIZE = 100;
const MAX_RETRIES = 2;
const RETRY_BASE_MS = 900;
const IDENTITY_BATCH_SIZE = 40;
const REVIEWER_GROUPS_CACHE_TTL_MS = 5 * 60 * 1000;
const reviewerGroupsCache = new Map();
const commitTimestampCache = new Map();

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
 * Список активных PR. При переданном reviewerId сервер отдаёт только PR, где эта identity в reviewers
 * (см. searchCriteria.reviewerId в Git REST API), без обхода всех активных PR репозитория.
 *
 * @param {import("./ado-config.mjs").DEFAULT_ADO_CONFIG} config
 * @param {string | null} reviewerId
 */
async function listActivePullRequests(config, reviewerId = null) {
  const project = encodeURIComponent(config.project.trim());
  const repo = encodeURIComponent(config.repositoryId.trim());
  const basePath = `${project}/_apis/git/repositories/${repo}/pullrequests`;

  const all = [];
  let skip = 0;

  for (;;) {
    const query = new URLSearchParams({
      "searchCriteria.status": "active",
      "api-version": resolveApiVersion(config),
      "$top": String(PAGE_SIZE),
      "$skip": String(skip),
    });

    if (reviewerId) {
      query.set("searchCriteria.reviewerId", String(reviewerId).trim());
    }

    const data = await adoFetch(config, `${basePath}?${query.toString()}`);
    const batch = Array.isArray(data?.value) ? data.value : [];

    all.push(...batch);

    if (batch.length < PAGE_SIZE) {
      break;
    }

    skip += PAGE_SIZE;
  }

  return all;
}

/**
 * Все активные PR репозитория (без фильтра по ревьюеру). Может быть очень медленно при большом числе активных PR.
 *
 * @param {import("./ado-config.mjs").DEFAULT_ADO_CONFIG} config
 */
export async function listAllActivePullRequests(config) {
  return listActivePullRequests(config, null);
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

  const batches = await Promise.all(ids.map((id) => listActivePullRequests(config, id)));
  return dedupePullRequestsById(batches.flat());
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
  const timestamp = normalizeIsoDate(
    commit?.author?.date ?? commit?.committer?.date ?? "",
  );

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

function timestampFromGitCommitRef(ref) {
  return normalizeIsoDate(ref?.author?.date ?? ref?.committer?.date ?? "");
}

/**
 * Дата последнего source-коммита: из детального PR / массива commits, иначе отдельный GET commit.
 */
async function resolveLastCommitAtFromPrDetail(config, detail, listPr, project, repo) {
  const sourceId = String(
    listPr?.lastMergeSourceCommit?.commitId
      ?? detail?.lastMergeSourceCommit?.commitId
      ?? "",
  ).trim();

  if (!sourceId) {
    return "";
  }

  let t = timestampFromGitCommitRef(detail?.lastMergeSourceCommit);

  if (t) {
    return t;
  }

  const commits = Array.isArray(detail?.commits) ? detail.commits : [];
  const hit = commits.find(
    (c) => String(c?.commitId ?? "").toLowerCase() === sourceId.toLowerCase(),
  );
  t = timestampFromGitCommitRef(hit);

  if (t) {
    return t;
  }

  try {
    return await fetchCommitTimestamp(config, project, repo, sourceId);
  } catch (_error) {
    return "";
  }
}

/**
 * Обогащает PR полным description и временем последнего source-коммита (один GET на PR с includeCommits).
 *
 * @param {import("./ado-config.mjs").DEFAULT_ADO_CONFIG} config
 * @param {Array<any>} pullRequests
 */
export async function attachPullRequestLastCommitTimes(config, pullRequests) {
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

        const lastCommitAt = await resolveLastCommitAtFromPrDetail(
          config,
          detail,
          pullRequest,
          project,
          repo,
        );

        const next = { ...pullRequest };

        if (typeof detail.description === "string") {
          next.description = detail.description;
        }

        if (lastCommitAt) {
          next.lastCommitAt = lastCommitAt;
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
  const selectedGroupIds = Array.isArray(config.selectedGroupIds)
    ? config.selectedGroupIds.map((id) => normalizePlainText(id)).filter(Boolean)
    : [];

  if (selectedGroupIds.length > 0) {
    return selectedGroupIds;
  }

  const legacySelectedTeamIds = Array.isArray(config.selectedTeamIds)
    ? config.selectedTeamIds.map((id) => normalizePlainText(id)).filter(Boolean)
    : [];

  if (legacySelectedTeamIds.length > 0) {
    return legacySelectedTeamIds;
  }

  const legacyManualIds = parseTeamReviewerIds(config.teamReviewerIds ?? "");

  if (legacyManualIds.length > 0) {
    return legacyManualIds.map((id) => normalizePlainText(id)).filter(Boolean);
  }

  return [];
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

/**
 * @param {import("./ado-config.mjs").DEFAULT_ADO_CONFIG} config
 * @param {string | null} [currentUserId]
 */
export async function fetchMyReviewerGroups(config, currentUserId = null) {
  const { groups } = await fetchMyReviewerGroupsWithDiagnostics(config, currentUserId);
  return groups;
}

/**
 * @param {import("./ado-config.mjs").DEFAULT_ADO_CONFIG} config
 * @param {string | null} [currentUserId]
 */
export async function fetchMyReviewerGroupIds(config, currentUserId = null) {
  const groups = await fetchMyReviewerGroups(config, currentUserId);
  return groups.map((group) => group.id);
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

  let allowedReviewerIds;
  let matchedSectionTitle;

  if (selectedGroupIds.length > 0) {
    allowedReviewerIds = [currentUserId, ...selectedGroupIds];
    matchedSectionTitle = "Выбранные reviewer-группы (API)";
  } else if (Array.isArray(config.selectedTeamIds) && config.selectedTeamIds.length > 0) {
    allowedReviewerIds = [currentUserId, ...config.selectedTeamIds.map(String)];
    matchedSectionTitle = "Выбранные reviewer-группы (legacy)";
  } else {
    allowedReviewerIds = [currentUserId];
    matchedSectionTitle = "Назначено мне (API)";
  }

  return { allowedReviewerIds, matchedSectionTitle };
}

/**
 * @param {import("./ado-config.mjs").DEFAULT_ADO_CONFIG} config
 * @param {Array<any>} pullRequests
 * @param {string} myId
 */
export async function filterPullRequestsForExtension(config, pullRequests, myId) {
  const { allowedReviewerIds, matchedSectionTitle } = getExtensionReviewerContext(config, myId);
  const allowedReviewerIdsSet = new Set(allowedReviewerIds);
  const filtered = pullRequests.filter((pullRequest) => {
    if (!isVisiblePullRequestForExtension(pullRequest)) {
      return false;
    }

    const reviewers = Array.isArray(pullRequest?.reviewers) ? pullRequest.reviewers : [];

    return reviewers.some(
      (reviewer) => (
        allowedReviewerIdsSet.has(String(reviewer?.id ?? ""))
        && Number(reviewer?.vote ?? 0) === 0
      ),
    );
  });

  return { filtered, matchedSectionTitle };
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
  const updatedAt = normalizeIsoDate(pr?.lastCommitAt) ?? createdAt;
  const description = normalizeDescription(pr?.description ?? "");
  const url = buildPullRequestWebUrl(apiRoot, project, repo, id, pr);

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
    description,
    url,
  };
}

/** Старые по дате обновления (`updatedAt ?? createdAt`) — выше в списке. */
export function sortPullRequestsOldestFirst(items) {
  return items
    .map((item, index) => {
      const sortKey = item.updatedAt ?? item.createdAt;
      return {
        item,
        index,
        timestamp: sortKey ? Date.parse(sortKey) : Number.NaN,
      };
    })
    .sort((left, right) => {
      const leftHasDate = Number.isFinite(left.timestamp);
      const rightHasDate = Number.isFinite(right.timestamp);

      if (leftHasDate && rightHasDate && left.timestamp !== right.timestamp) {
        return left.timestamp - right.timestamp;
      }

      if (leftHasDate !== rightHasDate) {
        return leftHasDate ? -1 : 1;
      }

      return left.index - right.index;
    })
    .map(({ item }) => item);
}

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
