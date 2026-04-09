export const TARGET_PAGE_URL =
  "https://hqrndtfs.avp.ru/tfs/DefaultCollection/Monorepo/_git/Monorepo/pullrequests?_a=mine";

export const TARGET_SECTION_TITLES = ["Assigned to my teams", "Assigned to me"];

const DATA_PROVIDERS_KEY = "ms.vss-code-web.prs-list-data-provider";
const SECTION_CARD_MARKER = '<div class="flex-noshrink repos-pr-section-card';
const ROW_CLASS =
  "bolt-list-row-marked bolt-table-row bolt-list-row single-click-activation v-align-middle selectable-text";

export function parsePullRequests(html) {
  const fromDataProviders = parseFromDataProviders(html);

  if (fromDataProviders) {
    return fromDataProviders;
  }

  return parseFromRenderedDom(html);
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

function parseFromRenderedDom(html) {
  const sections = splitIntoSections(html);
  const sectionSummaries = sections.map((sectionHtml) => ({
    html: sectionHtml,
    title: extractSectionTitle(sectionHtml),
  }));
  const matchedSection = sectionSummaries.find((section) =>
    TARGET_SECTION_TITLES.includes(section.title),
  );

  if (!matchedSection) {
    return {
      items: [],
      matchedSectionTitle: null,
      sectionFound: false,
    };
  }

  const rows = extractRows(matchedSection.html);
  const matchedSectionTitle = matchedSection.title;
  const items = rows
    .map((rowHtml, index) => parseRow(rowHtml, index))
    .filter(Boolean);

  return {
    items: sortPullRequestsOldestFirst(items),
    matchedSectionTitle,
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

function splitIntoSections(html) {
  return html
    .split(SECTION_CARD_MARKER)
    .slice(1)
    .map((sectionHtml) => `${SECTION_CARD_MARKER}${sectionHtml}`);
}

function extractSectionTitle(sectionHtml) {
  const match = sectionHtml.match(
    /<div class="repos-pr-section-header-title[^"]*">\s*<span[^>]*>([\s\S]*?)<\/span>/i,
  );

  return normalizeText(match?.[1] ?? "");
}

function extractRows(sectionHtml) {
  const rows = [];
  const pattern = new RegExp(
    `<a\\b[^>]*class="${escapeRegExp(ROW_CLASS)}"[^>]*>[\\s\\S]*?<\\/a>`,
    "gi",
  );

  for (const match of sectionHtml.matchAll(pattern)) {
    rows.push(match[0]);
  }

  return rows;
}

function parseRow(rowHtml, index) {
  const href = extractAttribute(rowHtml, "href");
  const title = extractTitle(rowHtml);
  const author = extractAuthor(rowHtml);
  const avatarUrl = extractAvatarUrl(rowHtml);
  const description = extractDescription(rowHtml);

  if (!href || !title) {
    return null;
  }

  return {
    id: extractPullRequestId(href) ?? `pr-${index + 1}`,
    title,
    author,
    avatarUrl,
    createdAt: null,
    description,
    url: new URL(href, TARGET_PAGE_URL).href,
  };
}

function extractAttribute(html, attributeName) {
  const pattern = new RegExp(`${attributeName}="([^"]+)"`, "i");
  return html.match(pattern)?.[1] ?? null;
}

function extractTitle(rowHtml) {
  const match = rowHtml.match(
    /<div class="body-l[^"]*font-weight-semibold[^"]*">([\s\S]*?)<\/div>/i,
  );

  return normalizeText(stripTags(match?.[1] ?? ""));
}

function extractAuthor(rowHtml) {
  const match = rowHtml.match(
    /<div class="secondary-text body-s text-ellipsis">([\s\S]*?)<\/div>/i,
  );

  const text = normalizeText(stripTags(match?.[1] ?? ""));
  const [author] = text.split(/\s+request\b/i);

  return author?.trim() ?? "";
}

function extractAvatarUrl(rowHtml) {
  const match = rowHtml.match(/<img[^>]*class="[^"]*bolt-coin-content[^"]*"[^>]*src="([^"]+)"/i);
  return match?.[1] ?? "";
}

function extractDescription(rowHtml) {
  const match = rowHtml.match(
    /<div class="body-s secondary-text text-ellipsis"[^>]*>([\s\S]*?)<\/div>/gi,
  );

  if (!match) {
    return "";
  }

  const descriptionMatch = match[1]?.match(/>([\s\S]*?)</i);

  return descriptionMatch ? normalizeDescription(stripTags(descriptionMatch[1] ?? "")) : "";
}

function extractPullRequestId(href) {
  const match = href.match(/\/pullrequest\/(\d+)/i);
  return match?.[1] ?? null;
}

function stripTags(value) {
  return value.replace(/<[^>]+>/g, " ");
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

function escapeRegExp(value) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}
