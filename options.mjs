import {
  ADO_CONFIG_KEY,
  DEFAULT_ADO_CONFIG,
  loadAdoConfig,
  validateAdoConfig,
} from "./ado-config.mjs";
import { searchReviewerGroupsByName } from "./ado-api.mjs";

const form = document.querySelector("#options-form");
const apiRootInput = document.querySelector("#api-root");
const projectInput = document.querySelector("#project");
const repositoryIdInput = document.querySelector("#repository-id");
const groupsFilterInput = document.querySelector("#groups-filter");
const groupsStatus = document.querySelector("#groups-status");
const groupsList = document.querySelector("#groups-list");
const groupsReloadButton = document.querySelector("#groups-reload");
const saveStatus = document.querySelector("#save-status");

/** @type {string} */
let groupsStatusLead = "";
let groupsStatusIsError = false;

void init();

async function init() {
  const config = await loadAdoConfig();
  apiRootInput.value = config.apiRoot ?? DEFAULT_ADO_CONFIG.apiRoot;
  projectInput.value = config.project ?? DEFAULT_ADO_CONFIG.project;
  repositoryIdInput.value = config.repositoryId ?? DEFAULT_ADO_CONFIG.repositoryId;

  groupsReloadButton.addEventListener("click", () => {
    void searchAndRememberGroups();
  });
  groupsFilterInput.addEventListener("keydown", (event) => {
    if (event.key === "Enter") {
      event.preventDefault();
      void searchAndRememberGroups();
    }
  });

  form.addEventListener("submit", (event) => {
    event.preventDefault();
    void handleSubmit();
  });

  groupsStatusLead = "";
  groupsStatusIsError = false;
  renderRememberedGroups(config);
  paintGroupsStatus();
}

async function buildConfigForApi() {
  const stored = await loadAdoConfig();

  return {
    ...stored,
    apiRoot: apiRootInput.value.trim(),
    project: projectInput.value.trim(),
    repositoryId: repositoryIdInput.value.trim(),
  };
}

function paintGroupsStatus() {
  groupsStatus.textContent = groupsStatusLead.trim();
  groupsStatus.classList.toggle("options__groups-status--err", groupsStatusIsError);
}

/**
 * @param {{ id: string, name?: string }[]} foundGroups
 */
async function mergeFoundGroupsIntoStorage(foundGroups) {
  const stored = await loadAdoConfig();
  const ids = [...stored.selectedGroupIds];
  const labels = { ...stored.selectedGroupLabels };
  let addedCount = 0;

  for (const group of foundGroups) {
    const id = String(group?.id ?? "").trim();

    if (!id) {
      continue;
    }

    if (!ids.includes(id)) {
      ids.push(id);
      addedCount += 1;
    }

    const name = String(group?.name ?? "").trim();
    labels[id] = name || labels[id] || id;
  }

  await chrome.storage.local.set({
    [ADO_CONFIG_KEY]: { ...stored, selectedGroupIds: ids, selectedGroupLabels: labels },
  });

  return addedCount;
}

/**
 * @param {string} id
 */
async function removeRememberedGroup(id) {
  const sid = String(id);
  const stored = await loadAdoConfig();
  const ids = stored.selectedGroupIds.filter((x) => String(x) !== sid);
  const labels = { ...stored.selectedGroupLabels };
  delete labels[sid];

  await chrome.storage.local.set({
    [ADO_CONFIG_KEY]: { ...stored, selectedGroupIds: ids, selectedGroupLabels: labels },
  });

  const next = await loadAdoConfig();
  renderRememberedGroups(next);
  paintGroupsStatus();
}

/**
 * @param {Awaited<ReturnType<typeof loadAdoConfig>>} config
 */
function renderRememberedGroups(config) {
  groupsList.textContent = "";
  const ids = config.selectedGroupIds ?? [];
  const labels = config.selectedGroupLabels ?? {};

  for (const rawId of ids) {
    const id = String(rawId ?? "").trim();

    if (!id) {
      continue;
    }

    groupsList.append(createRememberedRow(id, labels));
  }
}

/**
 * @param {string} id
 * @param {Record<string, string>} labels
 */
function createRememberedRow(id, labels) {
  const row = document.createElement("div");
  row.className = "options__selected-row";
  row.setAttribute("role", "listitem");

  const storedLabel = labels[id];
  const hasPrettyName = Boolean(storedLabel && storedLabel !== id);

  if (!hasPrettyName) {
    row.classList.add("options__selected-row--orphan");
  }

  const nameWrap = document.createElement("div");
  nameWrap.className = "options__selected-name";

  const title = document.createElement("span");
  title.textContent = hasPrettyName ? storedLabel : "Сохранённый идентификатор (нет названия)";
  nameWrap.append(title);

  if (!hasPrettyName) {
    const idLine = document.createElement("span");
    idLine.className = "options__selected-id";
    idLine.textContent = id;
    nameWrap.append(idLine);
  }

  const removeButton = document.createElement("button");
  removeButton.type = "button";
  removeButton.className = "options__selected-remove";
  removeButton.setAttribute("aria-label", "Удалить группу из списка");
  removeButton.textContent = "×";
  removeButton.addEventListener("click", () => {
    void removeRememberedGroup(id);
  });

  row.append(nameWrap, removeButton);
  return row;
}

async function searchAndRememberGroups() {
  groupsStatusLead = "Поиск reviewer-групп…";
  groupsStatusIsError = false;
  paintGroupsStatus();

  const config = await buildConfigForApi();
  const errors = validateAdoConfig(config);

  if (errors.length > 0) {
    groupsStatusLead = errors.join(" ");
    groupsStatusIsError = true;
    paintGroupsStatus();
    return;
  }

  const result = await searchReviewerGroupsByName(config, groupsFilterInput.value);
  const configAfter = await loadAdoConfig();

  if (result.mode === "error") {
    groupsStatusLead = `Ошибка загрузки групп: ${result.error?.message ?? "неизвестно"}.`;
    groupsStatusIsError = true;
    paintGroupsStatus();
    return;
  }

  const added = await mergeFoundGroupsIntoStorage(result.groups);
  const next = await loadAdoConfig();

  groupsStatusLead = result.note ?? "";

  if (added > 0) {
    groupsStatusLead = groupsStatusLead
      ? `${groupsStatusLead} Добавлено новых групп: ${added}.`
      : `Добавлено новых групп: ${added}.`;
  }

  groupsStatusIsError = false;
  renderRememberedGroups(next);
  paintGroupsStatus();
}

async function handleSubmit() {
  saveStatus.textContent = "";
  saveStatus.classList.remove("options__status--ok", "options__status--err");

  const stored = await loadAdoConfig();
  const merged = {
    ...stored,
    apiRoot: apiRootInput.value.trim(),
    project: projectInput.value.trim(),
    repositoryId: repositoryIdInput.value.trim(),
    selectedTeamIds: [],
    teamReviewerIds: "",
  };

  const errors = validateAdoConfig(merged);

  if (errors.length > 0) {
    saveStatus.textContent = errors.join(" ");
    saveStatus.classList.add("options__status--err");
    return;
  }

  await chrome.storage.local.set({
    [ADO_CONFIG_KEY]: merged,
  });

  await requestDevAzureHostPermissionIfNeeded(merged.apiRoot);

  saveStatus.textContent = "Сохранено. Список PR обновится автоматически.";
  saveStatus.classList.add("options__status--ok");
}

async function requestDevAzureHostPermissionIfNeeded(apiRoot) {
  let hostname = "";

  try {
    hostname = new URL(apiRoot).hostname;
  } catch (_error) {
    return;
  }

  if (hostname !== "dev.azure.com") {
    return;
  }

  const origins = ["https://dev.azure.com/*"];
  const already = await chrome.permissions.contains({ origins });

  if (already) {
    return;
  }

  try {
    await chrome.permissions.request({ origins });
  } catch (_error) {
    // пользователь отказал — fetch к API в режиме сессии не сработает до выдачи прав
  }
}
