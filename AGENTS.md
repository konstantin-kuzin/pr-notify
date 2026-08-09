# AGENTS.md

Инструкции для ИИ-агентов, работающих с репозиторием **PR Notify** — Chrome-расширением (Manifest V3), которое показывает pull request'ы в Azure DevOps: на вкладке **Review** — PR на ревью, на вкладке **My PRs** — собственные активные и Complete PR; в списках — Policies и конфликты слияния.

Подробное описание поведения и полей данных: [`docs/documentation.md`](docs/documentation.md). Краткий обзор для людей: [`README.md`](README.md).

## Обзор проекта

- **Стек:** vanilla JavaScript (ES modules, `.mjs`), HTML/CSS, Chrome Extension API (MV3).
- **Сборка:** отсутствует — расширение загружается как распакованная папка в `chrome://extensions`.
- **Зависимости:** нет `package.json`, npm/pnpm не используются.
- **Версия:** только в [`manifest.json`](manifest.json) (поле `version`).
- **Язык UI:** русский (тексты в popup, options, уведомлениях).
- **Целевой ADO:** по умолчанию on-prem `hqrndtfs.avp.ru`; облако `dev.azure.com` — через optional host permission.

## Команды разработки

### Установка и запуск

```bash
# Клонирование (если нужно)
git clone https://github.com/konstantin-kuzin/pr-notify.git
cd pr-notify
```

1. Открыть `chrome://extensions`, включить **Режим разработчика**.
2. **Загрузить распакованное расширение** — выбрать корень репозитория (где лежит `manifest.json`).
3. После правок кода — кнопка **Обновить** на карточке расширения.

### Проверка синтаксиса (единственный автоматический чек)

Перед завершением задачи прогоните все `.mjs`-модули:

```bash
node --check background.mjs
node --check ado-api.mjs
node --check ado-config.mjs
node --check options.mjs
node --check popup.mjs
node --check working-time.mjs
```

Или одной командой:

```bash
for f in background.mjs ado-api.mjs ado-config.mjs options.mjs popup.mjs working-time.mjs; do
  node --check "$f" || exit 1
done
```

### Тесты

Автотестов нет. Проверка — синтаксис (`node --check`) и ручное тестирование в Chrome с настроенным ADO-подключением.

## Архитектура и файлы

| Файл | Назначение |
|------|------------|
| [`manifest.json`](manifest.json) | MV3: permissions, service worker, popup, иконки |
| [`background.mjs`](background.mjs) | Service worker: alarms, storage, badge, уведомления, оркестрация API, approve |
| [`ado-api.mjs`](ado-api.mjs) | HTTP к Azure DevOps REST (`_apis/...`), фильтрация PR, маппинг в модель UI |
| [`ado-config.mjs`](ado-config.mjs) | Ключ `adoConfig`, дефолты, валидация, нормализация |
| [`working-time.mjs`](working-time.mjs) | Рабочие часы (пн–пт, 10:00–18:00 МСК), пороги срочности, стили badge |
| [`options.html`](options.html) / [`options.css`](options.css) / [`options.mjs`](options.mjs) | Страница настроек подключения и reviewer-групп |
| [`popup.html`](popup.html) / [`popup.css`](popup.css) / [`popup.mjs`](popup.mjs) | Popup со списком PR, markdown-описания, Approve |
| [`icons/`](icons/) | Иконки toolbar: default / orange / red / error (16 и 32 px) |
| [`docs/documentation.md`](docs/documentation.md) | Полная справка по поведению, storage, API |

### Ключи `chrome.storage.local`

- **`adoConfig`** — подключение к ADO (см. [`ado-config.mjs`](ado-config.mjs)).
- **`prState`** — списки PR (Review: `items` / My PRs: `myItems`), ошибки, метки времени обновления.
- **`prUpdateState`** — результат проверки новой версии на GitHub.

### Поток данных

```
background.mjs → ado-api.mjs (REST: review + my/policies; Complete на My PRs PRs — по запросу) → prState в storage → popup.mjs (вкладки Review/My PRs)
options.mjs → adoConfig в storage → background (onChanged) → refresh PR
```

## Стиль кода

- **Модули:** ES modules, расширение `.mjs`, относительные импорты (`./ado-api.mjs`).
- **Chrome API:** `chrome.storage`, `chrome.alarms`, `chrome.notifications`, `chrome.runtime.onMessage` и т.д.
- **JSDoc:** используется для типов параметров (`@param`, `@returns`) там, где это помогает.
- **Именование:** `camelCase` для функций и переменных, `UPPER_SNAKE_CASE` для констант.
- **Асинхронность:** `async/await`; fire-and-forget через `void fn()`.
- **Комментарии:** на русском, только для неочевидной бизнес-логики (не дублировать очевидный код).
- **Минимальный diff:** не рефакторить и не менять несвязанный код без запроса.
- **Соглашения проекта:** перед правками читайте соседний код и повторяйте его паттерны.

## Важная бизнес-логика (куда смотреть)

- **Отбор PR:** [`ado-api.mjs`](ado-api.mjs) — `filterPullRequestsForExtension`, `getExtensionReviewerContext`, `filterMyPullRequests`, `listCompletedPullRequestsByCreator`.
- **Обогащение PR** (commits, group comments, Policies, conflicts, `updatedAt`): `attachPullRequestLastCommitTimes`, `attachPullRequestBlockingReasons`, `attachPullRequestConflictInfo`, `mapPullRequestToItem`.
- **Рабочее время и срочность:** [`working-time.mjs`](working-time.mjs) — пороги 6 / 8 / 16 рабочих часов; точка отсчёта от `lastCommitAt` с учётом комментариев группы; сортировка «ожидающие ревью» выше.
- **Approve (vote: 10):** сообщение `approve-pull-request` в [`background.mjs`](background.mjs), REST в `ado-api.mjs`.
- **Complete на My PRs PRs:** сообщение `load-my-completed-pull-requests` в [`background.mjs`](background.mjs) / [`popup.mjs`](popup.mjs).
- **Markdown в popup:** упрощённый парсер в [`popup.mjs`](popup.mjs) — не полноценный CommonMark.
- **Периодическое обновление:** alarm каждые 10 мин (`CHECK_INTERVAL_MINUTES` в `background.mjs`).

При изменении поведения обновляйте [`docs/documentation.md`](docs/documentation.md) и при необходимости [`README.md`](README.md).

## Коммиты и версии

Формат сообщений коммитов в репозитории:

```
<version> <type>(<scope>): <краткое описание>
```

Примеры из истории: `2.7 docs: align README...`, `2.6 feat(update): notify when a newer extension version...`.

- **Версия** — префикс в subject (например `2.7`), синхронизируется с [`manifest.json`](manifest.json) → `version`.
- **Типы:** `feat`, `fix`, `docs`, `refactor`, `style`, `chore` и т.д.
- **Scope:** `ui`, `api`, `update` и др. по смыслу изменения.
- Коммиты создавать **только по явной просьбе** пользователя.

При bump версии меняйте `manifest.json` → `version`.

## Pull request

- Описание на русском или английском — как принято в задаче.
- Перед PR: все `node --check` проходят без ошибок.
- Если меняется поведение — обновить `docs/documentation.md`.
- Не коммитить секреты (PAT, credentials).

## Безопасность и ограничения

- **Авторизация:** по умолчанию `authMode: "session"` (куки браузера); PAT — через `adoConfig` в storage, без UI.
- **`host_permissions`** в манифесте захардкожены под `hqrndtfs.avp.ru`; другой on-prem хост требует правки [`manifest.json`](manifest.json).
- **`apiVersion`** не редактируется в UI — дефолт в [`ado-config.mjs`](ado-config.mjs) (`6.0-preview` для on-prem; для `dev.azure.com` часто нужен `7.1`).
- Один проект и один репозиторий из настроек — мульти-репо не поддерживается.
- При ошибке API последний успешный список PR **сохраняется** в storage (не затирать без необходимости).

## Типичные задачи агента

| Задача | Где править |
|--------|-------------|
| UI popup / вкладки Review·My PRs / markdown / approve | `popup.mjs`, `popup.css`, `popup.html` |
| Настройки ADO / группы | `options.mjs`, `options.css`, `ado-config.mjs` |
| Логика фильтрации / REST | `ado-api.mjs` |
| Badge, alarms, уведомления | `background.mjs` |
| Пороги срочности / рабочие часы | `working-time.mjs` |
| Права / версия расширения | `manifest.json` |
| Документация поведения | `docs/documentation.md`, `README.md` |

## Чего не делать

- Не добавлять `package.json`, bundler или TypeScript без явного запроса.
- Не создавать лишние абстракции и helper'ы на 1–2 строки.
- Не менять `host_permissions` «на всякий случай» — только если задача про новый хост.
- Не пушить в remote и не создавать коммиты без просьбы пользователя.
