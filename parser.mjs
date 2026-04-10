export const TARGET_PAGE_URL =
  "https://hqrndtfs.avp.ru/tfs/DefaultCollection/Monorepo/_git/Monorepo/pullrequests?_a=mine";

export const TARGET_SECTION_TITLES = ["Assigned to my teams", "Assigned to me"];

const DATA_PROVIDERS_KEY = "ms.vss-code-web.prs-list-data-provider";

export function parsePullRequests(html) {
  return parseFromDataProviders(html);
}

function parseFromDataProviders(html) {
  const dataProviders = extractDataProvidersJson(html);

  if (!dataProviders) {
    return null;
  }

  const prsDataProvider = dataProviders?.data?.[DATA_PROVIDERS_KEY];

  if (!prsDataProvider) {
    return null;
  }

  const queries = Array.isArray(prsDataProvider.queries) ? prsDataProvider.queries : [];
  const queryResults = prsDataProvider.queryResults ?? {};
  const pullRequests = prsDataProvider.pullRequests ?? {};

  const matchedQuery = findMatchedQuery(queries);

  if (!matchedQuery) {
    return {
      items: [],
      matchedSectionTitle: null,
      sectionFound: false,
    };
  }

  const ids = Array.isArray(queryResults?.[matchedQuery.id]?.ids)
    ? queryResults[matchedQuery.id].ids
    : [];

  const items = ids
    .map((id) => pullRequests[String(id)] ?? pullRequests[id])
    .filter(Boolean)
    .filter((pullRequest) => isActivePullRequestForQuery(pullRequest, matchedQuery))
    .map((pullRequest) => mapPullRequestFromDataProvider(pullRequest))
    .filter((item) => item?.title && item?.url);

  return {
    items: sortPullRequestsOldestFirst(items),
    matchedSectionTitle: matchedQuery.title ?? null,
    sectionFound: true,
  };
}

function extractDataProvidersJson(html) {
  const match = html.match(
    /<script id="dataProviders" type="application\/json">([\s\S]*?)<\/script>/i,
  );

  if (!match?.[1]) {
    return null;
  }

  try {
    return JSON.parse(match[1]);
  } catch (_error) {
    return null;
  }
}

function findMatchedQuery(queries) {
  for (const title of TARGET_SECTION_TITLES) {
    const match = queries.find((query) => normalizeText(query?.title ?? "") === title);

    if (match) {
      return match;
    }
  }

  return null;
}

function mapPullRequestFromDataProvider(pullRequest) {
  const pullRequestId = String(pullRequest?.pullRequestId ?? "").trim();
  const title = normalizeText(pullRequest?.title ?? "");
  const author = normalizeText(pullRequest?.createdBy?.displayName ?? "");
  const avatarUrl = normalizeText(
    pullRequest?.createdBy?._links?.avatar?.href ?? pullRequest?.createdBy?.imageUrl ?? "",
  );
  const createdAt = normalizeTimestamp(pullRequest?.creationDate);
  const description = normalizeDescription(pullRequest?.description ?? "");

  if (!pullRequestId || !title) {
    return null;
  }

  return {
    id: pullRequestId,
    title,
    author,
    avatarUrl,
    createdAt,
    description,
    url: buildPullRequestUrl(pullRequestId),
  };
}

function isActivePullRequestForQuery(pullRequest, query) {
  const reviewerIds = Array.isArray(query?.reviewerIds) ? query.reviewerIds : [];

  if (!query?.groupByVote || reviewerIds.length === 0) {
    return true;
  }

  const matchedReviewers = Array.isArray(pullRequest?.reviewers)
    ? pullRequest.reviewers.filter((reviewer) => reviewerIds.includes(reviewer?.id))
    : [];

  if (matchedReviewers.length === 0) {
    return true;
  }

  return matchedReviewers.some((reviewer) => Number(reviewer?.vote ?? 0) === 0);
}

function normalizeText(value) {
  return decodeHtmlEntities(value)
    .replace(/\s+/g, " ")
    .trim();
}

function normalizeDescription(value) {
  return decodeHtmlEntities(value)
    .replace(/\\n/g, "\n")
    .replace(/[ \t]+/g, " ")
    .trim();
}

function normalizeTimestamp(value) {
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

function sortPullRequestsOldestFirst(items) {
  return items
    .map((item, index) => ({
      item,
      index,
      timestamp: item.createdAt ? Date.parse(item.createdAt) : Number.NaN,
    }))
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

function buildPullRequestUrl(pullRequestId) {
  const pageUrl = new URL(TARGET_PAGE_URL);
  const pathname = pageUrl.pathname.replace(/\/pullrequests$/i, `/pullrequest/${pullRequestId}`);
  return new URL(`${pathname}${pageUrl.search ? "" : ""}`, pageUrl.origin).href;
}

function decodeHtmlEntities(value) {
  return value
    .replace(/&#(\d+);/g, (_, code) => String.fromCharCode(Number(code)))
    .replace(/&#x([0-9a-f]+);/gi, (_, code) =>
      String.fromCharCode(Number.parseInt(code, 16)),
    )
    .replace(/&nbsp;/gi, " ")
    .replace(/&amp;/gi, "&")
    .replace(/&lt;/gi, "<")
    .replace(/&gt;/gi, ">")
    .replace(/&quot;/gi, '"')
    .replace(/&#39;/gi, "'");
}
