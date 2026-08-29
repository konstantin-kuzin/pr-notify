# PR Notify



## Назначение

Расширение **PR Notify** для **Google Chrome** показывает pull request’ы в **Git-репозитории Azure DevOps** в popup с двумя вкладками:

- вкладка **Review** — активные PR, назначенные текущему пользователю как ревьюеру (лично и/или через выбранные reviewer-группы);
- вкладка **My PRs** — PR, **созданные** текущим пользователем: сначала **активные**, ниже — завершённые (**Complete**), с догрузкой по 10.

В списках дополнительно:

- **Policies** — проблемы required/optional-политик, из‑за которых недоступен Complete или видны красные Optional (на Review — раскрываемый бейдж; на My PRs у активных — список под карточкой);
- **конфликты слияния** — бейдж «конфликт» и текст из Conflicts API (Review и активные My PRs).

Данные получаются через **REST API** (`_apis/...`).

Страница настроек: [`options.html`](../options.html), логика — [`options.mjs`](../options.mjs).

## Первоначальная настройка (пошагово в UI)

1. **Открыть настройки** — из popup кнопка «Настройки подключения» в футере или через страницу расширения в Chrome (вкладка «Параметры» / «Расширения»).

2. **Коллекция** (`#api-root`, поле «Коллекция») — URL корня API:
   - on-prem: `https://<хост>/tfs/<коллекция>`;
   - облако: `https://dev.azure.com/<организация>`.

3. **Проект** (`#project`) и **Репозиторий** (`#repository-id`) — текстовые поля формы `#options-form`, обязательны для `validateAdoConfig` и для вызова поиска групп.

4. **Reviewer-группы** (fieldset `#groups-block`):
   - поле **поиска** `#groups-filter` — подсказка в UI: полное название группы;
   - кнопка **«Найти группы»** `#groups-reload` или клавиша **Enter** в поле поиска запускает `searchAndRememberGroups()`;
   - при успехе найденные группы **мержатся** в `adoConfig.selectedGroupIds` и `selectedGroupLabels` (новые id добавляются, без дублей);
   - список `#groups-list` показывает сохранённые группы строками `options__selected-row`; кнопка **×** (`removeRememberedGroup`) удаляет id из `selectedGroupIds`;
   - статус операции — `#groups-status`.

5. **Сохранить** — submit формы: в storage пишется `adoConfig` с актуальными `apiRoot`, `project`, `repositoryId` и текущими `selectedGroupIds` / `selectedGroupLabels`; вызывается `requestDevAzureHostPermissionIfNeeded` для `dev.azure.com`. После смены конфига фон (`background.mjs`) получает `chrome.storage.onChanged` и перезапрашивает PR.

Поля **`apiVersion`**, **`authMode`**, **`pat`** на странице настроек **не выводятся**; задаются дефолтами в [`ado-config.mjs`](../ado-config.mjs) или правкой `adoConfig` в storage / в коде.

### Разрешения (`permissions`)

| Разрешение    | Назначение |
|---------------|------------|
| `alarms`      | Периодическое обновление списка PR |
| `storage`     | Состояние списка и конфигурация ADO |
| `notifications` | Системные уведомления о новых PR |

### Доступ к хостам

- **`host_permissions`:**
  - `https://hqrndtfs.avp.ru/*` — корпоративный инстанс ADO (по умолчанию в манифесте);
  - `https://raw.githubusercontent.com/konstantin-kuzin/pr-notify/*` — загрузка `manifest.json` с ветки `main` для проверки новой версии расширения.
- **`optional_host_permissions`:** `https://dev.azure.com/*` — может быть запрошен при сохранении настроек при работе с облаком.

Для другого сервера ADO/TFS путь в манифесте нужно менять вручную под ваш хост.

## Конфигурация подключения

Хранится в `chrome.storage.local`, ключ **`adoConfig`**. Значения по умолчанию и слияние с сохранённым — [`ado-config.mjs`](../ado-config.mjs).

Поля модели (смысл as is):

| Поле | Роль |
|------|------|
| `apiRoot` | Базовый URL коллекции/организации (на странице настроек — поле «Коллекция») |
| `project` | Имя проекта |
| `repositoryId` | Имя или GUID репозитория |
| `apiVersion` | Строка `api-version` для всех REST-вызовов; **в UI настроек не редактируется**, по умолчанию в коде задано под on-prem (например `6.0-preview`) |
| `authMode` | `session` (куки браузера) или `pat` |
| `pat` | PAT при `authMode === "pat"` |
| `selectedGroupIds` | Массив identity id выбранных reviewer-групп |
| `selectedGroupLabels` | Отображаемые имена для выбранных id (кэш с последнего поиска) |

Пошаговый сценарий первого запуска описан выше в разделе **«Первоначальная настройка (пошагово в UI)»**.

## Логика отбора PR

Реализация в [`ado-api.mjs`](../ado-api.mjs) и вызовы из [`background.mjs`](../background.mjs).

### Review (на ревью)

1. Загружается конфиг; при ошибках валидации обновление не идёт в API, в состояние пишется понятная ошибка.
2. Определяется текущий пользователь: **`/_apis/connectionData`** → `authenticatedUser.id`.
3. Строится контекст ревьюеров: **вы** + набор identity id групп (`getExtensionReviewerContext`): либо **только выбранные** в настройках `selectedGroupIds`, либо если **ничего не выбрано** — все найденные reviewer-группы пользователя (после загрузки memberships и фильтрации «групп ревьюеров»).
4. Для **каждого** id из этого набора выполняется запрос списка активных PR с фильтром «этот ревьюер в PR» (параллельно, затем дедупликация по `pullRequestId`).
5. Дополнительно в коде отбрасываются черновики (`isDraft`) и неактивные статусы.
6. Остаются PR, где среди `reviewers` есть участник с **`vote === 0`** и `id` из множества: **вы ∪ выбранные группы** (или вы ∪ все reviewer-группы, если выбор пуст).
7. Результат сортируется в [`sortPullRequestsOldestFirst`](../working-time.mjs): сначала PR, **ожидающие ревью** (есть новые пуши после последнего комментария группы или комментария ещё не было) — **старые выше** по **`updatedAt`**; PR **без новых обновлений** после комментария группы — **в конце** списка (см. раздел «Рабочее время и срочность»).

### My PRs PRs (мои PR)

Вкладка **My PRs** показывает PR текущего пользователя как **создателя** (`searchCriteria.creatorId` = id из `connectionData`; для QA может быть временно переопределён `MY_TAB_CREATOR_OVERRIDE_ID` в [`ado-api.mjs`](../ado-api.mjs)).

#### Активные PR

Параллельно с обновлением Review (`refreshPullRequests`):

1. Запрос активных PR с **`searchCriteria.creatorId`**.
2. Отбрасываются черновики и неактивные статусы (как в Review).
3. Для каждого активного PR подмешиваются **Policies** (`blockingReasons` / `optionalPolicyReasons`) и при необходимости **конфликты** (`conflictText`) — см. разделы ниже.
4. В элемент UI дополнительно: `targetBranch` (целевая ветка без `refs/heads/`), `status`.
5. Сортировка: **новые выше** по `createdAt` (`sortMyPullRequestsNewestFirst`).
6. Результат пишется в `prState.myItems` / `myCount` (счётчик вкладки и `#count-badge` на My PRs — только активные).

#### Завершённые PR (Complete)

Не входят в периодический `prState`: popup догружает их **по запросу** при открытии вкладки My PRs (и после обновления списка — кэш Complete сбрасывается).

1. Сообщение в background: **`load-my-completed-pull-requests`** с `$top` / `$skip` (по умолчанию **10**).
2. REST: `searchCriteria.status = completed`, тот же `creatorId`.
3. Без загрузки политик и конфликтов; маппинг в UI-элементы со `status: completed`, при наличии — `closedAt` из `closedDate`.
4. В списке: **после всех активных**; у карточки бейдж **Complete**, дата закрытия вместо даты создания; причин Policies нет.
5. Если порция полная (10) — кнопка **«Показать ещё»** подгружает следующую страницу; состояние Complete хранится в памяти popup до закрытия или до сброса при смене `prState`.

Пустое состояние My PRs («Нет ваших активных pull requests») — только если нет ни активных, ни загруженных Complete (и загрузка Complete уже завершилась).

## Policies (Required + Optional)

Реализация: [`attachPullRequestBlockingReasons`](../ado-api.mjs). Вызывается в фоне для **активных** PR **обеих** вкладок (Review и My PRs).

### API

1. **`GET {project}/_apis/policy/evaluations?artifactId=vstfs:///CodeReview/CodeReviewId/{projectGuid}/{pullRequestId}`**  
   (`api-version` для endpoint — preview, например `6.0-preview.1`).
2. Параллельно (best-effort): **`GET .../pullRequests/{id}/statuses`** — для текстов status-политик.
3. В списки попадают включённые политики (кроме merge strategy), у которых статус **не** `approved` / `notApplicable` — как красные пункты в overview Policies в ADO:
   - **`blockingReasons`** — Required (`isBlocking === true`);
   - **`optionalPolicyReasons`** — Optional (`isBlocking !== true`).
4. При ошибке запроса: оба поля = `null` (в UI — сообщение об ошибке загрузки).

### Тексты причин

Ближе к overview Policies в ADO:

| Тип политики | Пример текста |
|--------------|---------------|
| Minimum reviewers + downvote | `1 reviewer is blocking` / `N reviewers are blocking` |
| Minimum reviewers без downvote | `N of M reviewers approved` |
| Required reviewers | `Required reviewers have not approved` |
| Comment requirements | `Not all comments resolved` |
| Build с `context.isExpired` | `{displayName} expired` |
| Status policy | `description` из PR statuses / `defaultDisplayName` (например `Votes check`); при `isExpired` — суффикс `expired` |
| Merge strategy | **не** включается (в overview ADO тоже не показывается) |

### UI

| Вкладка | Показ |
|---------|--------|
| **Review** | Если есть проблемные Required и/или Optional — розовый бейдж с **числом** замечаний; клик раскрывает список (Required сверху, блок **Optional** ниже). Если Optional нет, а Required — ровно один из наборов «ожидание голосов» (`0 of 1 reviewers approved` + `Votes check` + `Required reviewers have not approved`, либо только `Votes check` + `Required reviewers have not approved`), бейдж показывает **Votes check** вместо числа. |
| **My PRs** (активные) | Под карточкой: список Required **или** **«Готов к Complete»** (если Required пуст), затем блок **Optional** при наличии; либо **«Не удалось загрузить политики Complete»** (`null`). |
| **My PRs** (Complete) | Policies не загружаются и не показываются. |

Поля в элементе списка: **`blockingReasons`**, **`optionalPolicyReasons`**: `string[]` | `null` | отсутствует.

## Конфликты слияния

Реализация: [`attachPullRequestConflictInfo`](../ado-api.mjs). Для активных PR Review и My PRs, у которых **`mergeStatus === conflicts`**.

### API

1. **`GET .../pullRequests/{id}/conflicts`** (preview `api-version`).
2. В объект пишется **`conflictText`**:
   - сводка вида `N conflict(s) prevent(s) automatic merging`;
   - строки `path — type` (например `Edited in both`), как в баннере Conflicts в ADO;
   - при ошибке запроса — только общая сводка без списка файлов.

### UI

- В строке метаданных — ярко-красный бейдж **«конфликт»** (белый текст); клик раскрывает/сворачивает панель с `conflictText`.
- На вкладке My PRs у карточек **Complete** конфликты не запрашиваются и бейдж не показывается.

Поле в элементе списка: **`conflictText`** (строка; отсутствует, если конфликтов нет).

## Обогащение PR перед UI

После фильтрации фон вызывает [`resolveConfiguredGroupMemberIds`](../ado-api.mjs) — user id участников **сохранённых** reviewer-групп (Identities API, `ExpandedDown`, кэш 5 мин).

Порядок в [`refreshPullRequests`](../background.mjs):

1. **Review only** — [`attachPullRequestLastCommitTimes`](../ado-api.mjs) для отфильтрованных PR на ревью:
   - **`GET .../pullrequests/{id}?includeCommits=true`**: полное `description`; **`lastCommitAt`** из push/commit (кэш);
   - если есть участники групп — **`GET .../threads`**, **`lastGroupCommentAt`** (открывающий тред комментарий участника группы, не system).

2. **Review и My PRs (активные)** — [`attachPullRequestBlockingReasons`](../ado-api.mjs) → `blockingReasons` / `optionalPolicyReasons` (см. раздел «Policies»).

3. **Review и My PRs (активные)** — [`attachPullRequestConflictInfo`](../ado-api.mjs) → `conflictText` при `mergeStatus === conflicts` (см. раздел «Конфликты слияния»).

Завершённые PR на My PRs (Complete) этим пайплайном **не** проходят — только list + `mapPullRequestToItem`.

В элементе для UI ([`mapPullRequestToItem`](../ado-api.mjs)):

- **`updatedAt`** = max(`lastGroupCommentAt`, `lastCommitAt`, `closedAt`, `createdAt`) — для **сортировки** и метаданных (на My PRs без enrichment commit-полей обычно = `createdAt` / `closedAt`);
- **`createdAt`** = дата создания PR;
- **`closedAt`** = дата закрытия (`closedDate`), если есть;
- **`status`** — `active` / `abandoned` / `completed` / `unknown`;
- **`lastCommitAt`**, **`lastGroupCommentAt`** — ISO-даты (в основном на Review; для рабочего времени);
- **`targetBranch`** — целевая ветка (на My PRs);
- **`blockingReasons`** — проблемы Required Policies;
- **`optionalPolicyReasons`** — проблемы Optional Policies;
- **`conflictText`** — при конфликтах слияния (иначе поле отсутствует).

## Состояние в `storage` (ключ `prState`)

[`background.mjs`](../background.mjs) сохраняет объект (см. `DEFAULT_STATE`):

- `items` — массив элементов вкладки **Review**: `id`, `title`, `author`, `avatarUrl`, `createdAt`, `updatedAt`, `status`, `lastCommitAt`, `lastGroupCommentAt`, `description`, `url`, при проблемах Policies — `blockingReasons` / `optionalPolicyReasons`, при конфликтах слияния — `conflictText`;
- `count` — длина `items` (именно он влияет на badge toolbar и уведомления о новых PR);
- `myItems` — массив **активных** элементов вкладки **My PRs**: те же базовые поля плюс `targetBranch`, `blockingReasons` / `optionalPolicyReasons` (`string[]` или `null` при ошибке загрузки политик), при конфликтах — `conflictText`;
- `myCount` — длина `myItems` (только активные; Complete в storage не кэшируются);
- `matchedSectionTitle` — служебное поле фильтра (заголовок секции совпадения, если используется);
- `lastCheckedAt` — ISO-время последней попытки обновления;
- `lastSuccessAt` — ISO-время последнего **успешного** обновления;
- `lastTrigger` — откуда вызвано обновление (`install`, `startup`, `alarm`, `manual`, `config-change`, `service-worker-load`, …);
- `lastError` — текст ошибки или `null`;
- `previousItemIds` — id из последнего успешного списка **Review** — для детекта **новых** PR и показа уведомлений.

При ошибке API: **иконка** `icon-*-error.png`, **badge** пустой, в `prState` пишется ошибка; **предыдущий успешный список в `items` / `myItems` не подменяется** на пустой в ветке catch (сохраняется прошлое состояние кроме полей ошибки/времени — см. `refreshPullRequests`).

## Рабочее время и срочность

Логика в [`working-time.mjs`](../working-time.mjs). Рабочие часы: **пн–пт**, **10:00–18:00** (конец не включается), **Europe/Moscow**.

### Ожидание ревью и точка отсчёта

Функция **`hasUpdatesAfterLastGroupComment(item)`** определяет, считается ли PR «ожидающим ревью»:

| Условие | Ожидание ревью |
|---------|----------------|
| `lastGroupCommentAt` отсутствует | да — открывающего тред комментария группы ещё не было |
| есть комментарий, но нет `lastCommitAt` | нет |
| `lastCommitAt` **строго позже** `lastGroupCommentAt` | да — автор допушил после комментария |
| иначе | нет — группа уже ответила, новых пушей нет |

**Точка отсчёта** рабочего времени — [`getItemWorkingTimeFrom(item)`](../working-time.mjs):

- если PR **не** ожидает ревью → `null` (счётчик не ведётся);
- если есть и комментарий группы, и пуш → от **`lastCommitAt`** (момент последнего пуша после комментария);
- иначе → **`lastCommitAt ?? createdAt`**.

От этой точки до **`lastCheckedAt`** считаются **рабочие минуты** ([`getWorkingElapsedMinutes`](../working-time.mjs)).

### Пороги срочности

| Порог (рабочие ч) | Уровень | Где применяется |
|-------------------|---------|-----------------|
| > 6               | yellow  | чип времени; **badge** toolbar |
| > 8               | orange  | чип времени; **badge** и **иконка** toolbar |
| > 16              | red     | чип времени; **badge** и **иконка** toolbar |

PR без ожидания ревью **не участвуют** в расчёте максимального возраста для иконки и badge.

### Сортировка списка

[`sortPullRequestsOldestFirst`](../working-time.mjs): PR с ожиданием ревью — выше, внутри группы **старые по `updatedAt` выше**; PR без новых обновлений после комментария группы — **в конце**. Сортировка выполняется в фоне при сохранении `prState` и повторно в popup при отрисовке.

### Иконка и badge на панели Chrome

Срочность по максимальному **рабочему** возрасту среди PR, **ожидающих ревью** ([`getBadgeUrgencyFromItems`](../working-time.mjs)):

| Порог (рабочие ч) | Иконка toolbar | Badge (число) |
|-------------------|----------------|---------------|
| ≤ 6               | default        | серый фон     |
| > 6               | default        | жёлтый фон    |
| > 8               | orange         | оранжевый фон |
| > 16              | red            | красный фон   |

Жёлтый уровень (>6 ч) **не меняет** иконку toolbar — только **чип времени** в карточке PR и цвет badge.

Счётчик в popup (`#count-badge`) всегда в **сером** стиле (`BADGE_STYLES.gray`), независимо от срочности.

## Состояние обновления расширения (ключ `prUpdateState`)

При **bootstrap** (установка, старт браузера, загрузка service worker) фон запрашивает  
`https://raw.githubusercontent.com/konstantin-kuzin/pr-notify/main/manifest.json` (таймаут 3 с, `cache: no-store`).

Поля в storage:

| Поле | Смысл |
|------|--------|
| `checkedAt` | ISO время последней проверки |
| `localVersion` | `chrome.runtime.getManifest().version` на момент проверки |
| `latestVersion` | `version` из удалённого manifest |
| `hasUpdate` | semver remote > local |
| `error` | текст ошибки или `null` |

Popup читает `hasUpdate` и `latestVersion`, показывает чип **«Новая версия — …»** в футере (ссылка на GitHub). Обновление установленной копии — вручную: `git pull` + перезагрузка на `chrome://extensions` (см. README).

## Периодичность и триггеры обновления

- Будильник Chrome: **каждые 10 минут** (`CHECK_INTERVAL_MINUTES`).
- При **установке**, **старте браузера**, **загрузке service worker**, **смене `adoConfig`**, **ручном обновлении из popup**, после **успешного Approve** (с задержкой/опросом) — обновление списка PR по той же цепочке.
- Проверка версии на GitHub — при bootstrap (не по alarm).

## Popup ([`popup.mjs`](../popup.mjs), [`popup.css`](../popup.css))

- Ширина документа: **600px**.
- Верх: заголовок; справа — время последней проверки (сегодня только время, иначе дата+время), кнопка обновления, **счётчик** PR на **активной вкладке** (**серый** badge; скрывается при ошибке). На **Review** — `count`, на **My PRs** — `myCount` (только активные). Отступ от заголовка до вкладок — **16px**.
- Вкладки **Review** / **My PRs** (выбор хранится в `chrome.storage.session` на сессию браузера).
- Середина: **прокручиваемый** список PR; при ошибке — текст сообщения; пустой Review — текст про reviewer-группы; пустой My PRs — «Нет ваших активных pull requests» (если нет ни активных, ни Complete).
- Низ: **футер** (сетка 3 колонки) — слева ссылка **GitHub**, по центру (если `prUpdateState.hasUpdate`) чип **«Новая версия — X.Y»**, справа **«Настройки подключения»** (`chrome.runtime.openOptionsPage()`).
- Высота колонки попапа ограничивается CSS-переменной **`--popup-max-height`**, выставляемой скриптом как **половина `screen.availHeight`** (fallback `innerHeight`), плюс слушатель `resize`.

Карточка PR (**Review**):

- клик по заголовку открывает PR в новой вкладке и закрывает popup;
- если новых пушей после последнего комментария группы нет — вся карточка с **opacity 0.8** (`popup__item--no-updates`), заголовок с обычным начертанием;
- под заголовком: автор и **относительное рабочее время** (`N мин` / `H ч M мин`) от точки [`getItemWorkingTimeFrom`](../working-time.mjs) до `lastCheckedAt`; если новых пушей после комментария группы нет — текст **«Нет обновлений»** (без цветного чипа);
- фрагмент времени — **чип** при **> 6 / > 8 / > 16** рабочих ч (жёлтый / оранжевый / красный); пороги — [`working-time.mjs`](../working-time.mjs);
- иконка **Storybook** (слева от иконки описания) открывает `https://storybook.s1.ksc-web.avp.ru/hexa-ui/<id PR>/` в новой вкладке и закрывает popup; **скрывается**, если есть конфликты слияния или в Policies есть `[OSMP] Storybook Hexa UI deploy for Review expired`;
- при наличии описания — иконка раскрывает **панель под карточкой** с **упрощённым markdown** (заголовки **h1–h6**, списки, ссылки, код, жирный/курсив, картинки по `http(s)`; разбор скобок в URL картинок с балансом `()`);
- **Policies** — розовый бейдж с числом замечаний или **Votes check** (см. раздел «Policies»);
- **конфликты** — бейдж **«конфликт»** (см. раздел «Конфликты слияния»);
- для PR с текстом описания, совпадающим с эвристикой «тех ПР» — бейдж **ТЕХ ПР** и кнопка **Approve** (отправка сообщения в background).

Карточка PR (**My PRs**, активные):

- клик по заголовку — как в Review;
- под заголовком: целевая ветка (`→ master`) и дата создания;
- **Policies** — список под карточкой (см. раздел «Policies»);
- **конфликты** — бейдж в строке метаданных (см. раздел «Конфликты слияния»);
- иконка Storybook и описание по иконке — как в Review; Approve / ТЕХ ПР на этой вкладке не показываются.

Карточка PR (**My PRs**, Complete):

- ниже всех активных; бейдж **Complete**; дата закрытия; без Policies и конфликтов;
- иконка Storybook — как у активных;
- догрузка по 10, кнопка **«Показать ещё»** (см. «My PRs → Завершённые PR»).

Сообщения popup → background:

| Тип | Назначение |
|-----|------------|
| `manual-refresh` | Ручное обновление активных списков Review/My PRs |
| `approve-pull-request` | Approve (vote 10) |
| `load-my-completed-pull-requests` | Страница Complete для вкладки My PRs (`skip`, `top`) |

## Approve

- Сообщение типа `approve-pull-request` с `pullRequestId`.
- Вызов **`PUT .../pullRequests/{id}/reviewers/{reviewerId}`** с телом `vote: 10` и `id` текущего пользователя (тот же GUID, что в `connectionData`).
- После успеха — фоновое обновление списка с таймаутами/интервалом опроса (см. константы в `background.mjs`).

## Уведомления

При успешном обновлении, если появились id, которых не было в `previousItemIds`, показывается **`chrome.notifications`** с кратким текстом по новым PR (ограничение числа строк в теле обработчика).

## Структура файлов (основное)

| Файл | Роль |
|------|------|
| `manifest.json` | MV3, права, иконки, entrypoints |
| `background.mjs` | Алармы, storage, badge, уведомления, оркестрация API |
| `ado-api.mjs` | Запросы ADO, фильтры, маппинг, approve, identities/groups |
| `ado-config.mjs` | Дефолты и нормализация `adoConfig` |
| `options.html` / `options.css` / `options.mjs` | UI настроек |
| `popup.html` / `popup.css` / `popup.mjs` | UI списка |
| `working-time.mjs` | Рабочие часы (МСК), точка отсчёта и ожидание ревью, пороги срочности, сортировка списка, стили badge |
| `icons/*` | default / orange / red / error (16 и 32 px для toolbar) |

## Ограничения (as is)

- Один проект и один репозиторий из конфига; нет мульти-репо в одном расширении без смены настроек.
- Отбор PR зависит от корректности membership/group id и полей `reviewers` в ответе API.
- `apiVersion` не на экране настроек — смена только в коде или вручную в storage.
- Упрощённый markdown в popup не равен полному движку ADO/Web; сложные конструкции могут отображаться иначе.

## Проверка синтаксиса модулей

```bash
node --check background.mjs
node --check ado-api.mjs
node --check ado-config.mjs
node --check options.mjs
node --check popup.mjs
node --check working-time.mjs
```

Установка: режим разработчика в `chrome://extensions`, «Загрузить распакованное». Актуальная версия в [`manifest.json`](../manifest.json) (на момент сборки документации — **2.11**).
