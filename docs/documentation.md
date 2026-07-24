# PR Notify



## Назначение

Расширение **PR Notify** для **Google Chrome** показывает активные pull request’ы в **Git-репозитории Azure DevOps**, назначенные текущему пользователю как ревьюеру (лично и/или через выбранные reviewer-группы). Данные получаются через **REST API** (`_apis/...`).

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

1. Загружается конфиг; при ошибках валидации обновление не идёт в API, в состояние пишется понятная ошибка.
2. Определяется текущий пользователь: **`/_apis/connectionData`** → `authenticatedUser.id`.
3. Строится контекст ревьюеров: **вы** + набор identity id групп (`getExtensionReviewerContext`): либо **только выбранные** в настройках `selectedGroupIds`, либо если **ничего не выбрано** — все найденные reviewer-группы пользователя (после загрузки memberships и фильтрации «групп ревьюеров»).
4. Для **каждого** id из этого набора выполняется запрос списка активных PR с фильтром «этот ревьюер в PR» (параллельно, затем дедупликация по `pullRequestId`).
5. Дополнительно в коде отбрасываются черновики (`isDraft`) и неактивные статусы.
6. Остаются PR, где среди `reviewers` есть участник с **`vote === 0`** и `id` из множества: **вы ∪ выбранные группы** (или вы ∪ все reviewer-группы, если выбор пуст).
7. Результат сортируется в [`sortPullRequestsOldestFirst`](../working-time.mjs): сначала PR, **ожидающие ревью** (есть новые пуши после последнего комментария группы или комментария ещё не было) — **старые выше** по **`updatedAt`**; PR **без новых обновлений** после комментария группы — **в конце** списка (см. раздел «Рабочее время и срочность»).

## Обогащение PR перед UI

После фильтрации фон вызывает [`resolveConfiguredGroupMemberIds`](../ado-api.mjs) — user id участников **сохранённых** reviewer-групп (Identities API, `ExpandedDown`, кэш 5 мин).

Для **каждого** PR:

1. **`GET .../pullrequests/{id}?includeCommits=true`**:
   - в объект подмешивается **полное** поле `description` (в списке API оно часто усечено);
   - вычисляется **`lastCommitAt`**: `push.date` или **`GET .../pushes/{pushId}`** (приоритет), затем **`GET .../commits/{commitId}`**, в конце — `author`/`committer` из PR (кэш по commit/push).

2. Если есть участники групп:
   - **`GET .../pullRequests/{id}/threads`**;
   - ищется **`lastGroupCommentAt`** — последний **не system** комментарий участника группы, который **открывает тред** (`comments[0]` в ветке); ответы внутри ветки не учитываются.

В элементе для UI ([`mapPullRequestToItem`](../ado-api.mjs)):

- **`updatedAt`** = max(`lastGroupCommentAt`, `lastCommitAt`, `createdAt`) — для **сортировки** и метаданных;
- **`createdAt`** = дата создания PR;
- **`lastCommitAt`**, **`lastGroupCommentAt`** — ISO-даты последнего пуша source и последнего комментария участника группы; попадают в элемент UI и используются для **рабочего времени** (см. ниже).

## Состояние в `storage` (ключ `prState`)

[`background.mjs`](../background.mjs) сохраняет объект (см. `DEFAULT_STATE`):

- `items` — массив элементов для popup: `id`, `title`, `author`, `avatarUrl`, `createdAt`, `updatedAt`, `lastCommitAt`, `lastGroupCommentAt`, `description`, `url`;
- `count` — длина `items`;
- `matchedSectionTitle` — служебное поле фильтра (заголовок секции совпадения, если используется);
- `lastCheckedAt` — ISO-время последней попытки обновления;
- `lastSuccessAt` — ISO-время последнего **успешного** обновления;
- `lastTrigger` — откуда вызвано обновление (`install`, `startup`, `alarm`, `manual`, `config-change`, `service-worker-load`, …);
- `lastError` — текст ошибки или `null`;
- `previousItemIds` — id из последнего успешного списка — для детекта **новых** PR и показа уведомлений.

При ошибке API: **иконка** `icon-*-error.png`, **badge** пустой, в `prState` пишется ошибка; **предыдущий успешный список в `items` не подменяется** на пустой в ветке catch (сохраняется прошлое состояние кроме полей ошибки/времени — см. `refreshPullRequests`).

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
- Верх: заголовок, строка «Последняя проверка», кнопка обновления, **счётчик** с числом PR (**серый** badge; скрывается при ошибке).
- Середина: **прокручиваемый** список PR; при ошибке — текст сообщения; пустой список — текст про reviewer-группы и личные назначения.
- Низ: **футер** (сетка 3 колонки) — слева ссылка **GitHub**, по центру (если `prUpdateState.hasUpdate`) чип **«Новая версия — X.Y»**, справа **«Настройки подключения»** (`chrome.runtime.openOptionsPage()`).
- Высота колонки попапа ограничивается CSS-переменной **`--popup-max-height`**, выставляемой скриптом как **половина `screen.availHeight`** (fallback `innerHeight`), плюс слушатель `resize`.

Карточка PR:

- клик по заголовку открывает PR в новой вкладке и закрывает popup;
- под заголовком: автор и **относительное рабочее время** (`N мин` / `H ч M мин`) от точки [`getItemWorkingTimeFrom`](../working-time.mjs) до `lastCheckedAt`; если новых пушей после последнего комментария группы нет — текст **«Нет обновлений»** (без цветного чипа);
- фрагмент времени — **чип** при **> 6 / > 8 / > 16** рабочих ч (жёлтый / оранжевый / красный); пороги — [`working-time.mjs`](../working-time.mjs);
- при наличии описания — иконка раскрывает **панель под карточкой** с **упрощённым markdown** (заголовки **h1–h6**, списки, ссылки, код, жирный/курсив, картинки по `http(s)`; разбор скобок в URL картинок с балансом `()`);
- для PR с текстом описания, совпадающим с эвристикой «тех ПР» — бейдж **ТЕХ ПР** и кнопка **Approve** (отправка сообщения в background).

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

Установка: режим разработчика в `chrome://extensions`, «Загрузить распакованное». Актуальная версия в [`manifest.json`](../manifest.json) (на момент сборки документации — **2.8**).
