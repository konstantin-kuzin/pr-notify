export const ADO_CONFIG_KEY = "adoConfig";

/** Defaults match прежний сценарий (Monorepo на корпоративном TFS). */
export const DEFAULT_ADO_CONFIG = {
  apiRoot: "https://hqrndtfs.avp.ru/tfs/DefaultCollection",
  project: "Monorepo",
  repositoryId: "Monorepo",
  /**
   * Версия REST API (`api-version` в запросах). На странице настроек не показывается —
   * меняйте здесь (или вручную в `chrome.storage.local` → `adoConfig`), если серверу нужна
   * другая версия (например `7.1` для dev.azure.com).
   */
  apiVersion: "6.0-preview",
  /**
   * `session` — куки браузера; `pat` — токен из поля `pat` ниже (без UI PAT можно задать в storage).
   * На странице настроек не показывается.
   */
  authMode: "session",
  pat: "",
  /** Выбранные reviewer-группы пользователя (identity ids). */
  selectedGroupIds: [],
  /**
   * Отображаемые имена запомненных групп (id → название с последнего успешного поиска).
   */
  selectedGroupLabels: {},
};

export async function loadAdoConfig() {
  const stored = await chrome.storage.local.get(ADO_CONFIG_KEY);
  const partial = stored[ADO_CONFIG_KEY] ?? {};
  const raw = { ...DEFAULT_ADO_CONFIG, ...partial };

  let selectedGroupIds = normalizeSelectedIds(partial.selectedGroupIds);

  if (selectedGroupIds.length === 0) {
    selectedGroupIds = normalizeSelectedIds(partial.selectedTeamIds);
  }

  if (selectedGroupIds.length === 0 && partial.teamReviewerIds?.trim()) {
    selectedGroupIds = parseTeamReviewerIds(partial.teamReviewerIds);
  }

  const { savedGroupSearchPhrases: _legacySearchPhrases, ...rest } = raw;

  return {
    ...rest,
    selectedGroupIds,
    selectedGroupLabels: normalizeSelectedGroupLabels(rest.selectedGroupLabels),
  };
}

/**
 * @param {unknown} value
 * @returns {Record<string, string>}
 */
export function normalizeSelectedGroupLabels(value) {
  if (value === null || typeof value !== "object" || Array.isArray(value)) {
    return {};
  }

  /** @type {Record<string, string>} */
  const out = {};

  for (const [rawId, rawName] of Object.entries(value)) {
    const id = String(rawId ?? "").trim();

    if (!id) {
      continue;
    }

    const name = String(rawName ?? "").trim();
    out[id] = name || id;
  }

  return out;
}

function normalizeSelectedIds(value) {
  if (!Array.isArray(value)) {
    return [];
  }

  return value.map((id) => String(id ?? "").trim()).filter(Boolean);
}

export async function saveAdoConfig(partial) {
  const current = await loadAdoConfig();
  const next = { ...current, ...partial };
  await chrome.storage.local.set({ [ADO_CONFIG_KEY]: next });
  return next;
}

export function parseTeamReviewerIds(raw) {
  if (typeof raw !== "string" || !raw.trim()) {
    return [];
  }

  return raw
    .split(/[,;\s]+/g)
    .map((s) => s.trim())
    .filter(Boolean);
}

export function validateAdoConfig(config) {
  const errors = [];

  if (!config.apiRoot?.trim()) {
    errors.push("Укажите корень API (URL коллекции, например …/tfs/DefaultCollection).");
  }

  if (!config.project?.trim()) {
    errors.push("Укажите проект.");
  }

  if (!config.repositoryId?.trim()) {
    errors.push("Укажите репозиторий (имя или GUID).");
  }

  if (config.authMode === "pat" && !config.pat?.trim()) {
    errors.push("В режиме PAT укажите токен или переключитесь на «Сессия браузера».");
  }

  const ver = resolveApiVersion(config);

  if (!isPlausibleApiVersion(ver)) {
    errors.push(
      "Версия API: например 6.0-preview, 6.0 или 7.1 (как в документации к вашему серверу).",
    );
  }

  return errors;
}

export function resolveApiVersion(config) {
  const raw = config?.apiVersion ?? DEFAULT_ADO_CONFIG.apiVersion;
  return String(raw ?? "").trim() || DEFAULT_ADO_CONFIG.apiVersion;
}

function isPlausibleApiVersion(ver) {
  if (!ver || /[\s<>'"]/.test(ver)) {
    return false;
  }

  return /^\d+\.\d+([\w.-]+)?$/.test(ver);
}

export function normalizeApiRoot(apiRoot) {
  return String(apiRoot ?? "").replace(/\/+$/, "");
}
