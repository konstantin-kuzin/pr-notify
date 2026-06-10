/** Москва, пн–пт, 10:00–18:00 (конец интервала не включается). */
const MOSCOW_TZ = "Europe/Moscow";
const WORK_START_HOUR = 10;
const WORK_END_HOUR = 18;
const MOSCOW_OFFSET = "+03:00";

/**
 * Рабочие минуты между двумя моментами (ISO-строки или Date).
 * @param {string|Date} from
 * @param {string|Date} to
 * @returns {number|null}
 */
export function getWorkingElapsedMinutes(from, to) {
  if (from == null || to == null || from === "" || to === "") {
    return null;
  }

  const start = from instanceof Date ? from : new Date(from);
  const end = to instanceof Date ? to : new Date(to);

  if (Number.isNaN(start.getTime()) || Number.isNaN(end.getTime())) {
    return null;
  }

  if (end < start) {
    return null;
  }

  let totalMs = 0;
  let { year, month, day } = getMoscowYmd(start);
  const endYmd = getMoscowYmd(end);

  for (;;) {
    if (isMoscowWeekday(year, month, day)) {
      const windowStart = moscowLocalDate(year, month, day, WORK_START_HOUR, 0);
      const windowEnd = moscowLocalDate(year, month, day, WORK_END_HOUR, 0);
      const segmentStart = Math.max(start.getTime(), windowStart.getTime());
      const segmentEnd = Math.min(end.getTime(), windowEnd.getTime());

      if (segmentStart < segmentEnd) {
        totalMs += segmentEnd - segmentStart;
      }
    }

    if (year === endYmd.year && month === endYmd.month && day === endYmd.day) {
      break;
    }

    ({ year, month, day } = addMoscowCalendarDays(year, month, day, 1));
  }

  return Math.floor(totalMs / 60_000);
}

function pad2(n) {
  return String(n).padStart(2, "0");
}

function getMoscowYmd(date) {
  const parts = new Intl.DateTimeFormat("en-CA", {
    timeZone: MOSCOW_TZ,
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  }).formatToParts(date);

  const pick = (type) => parts.find((p) => p.type === type)?.value ?? "";

  return {
    year: Number(pick("year")),
    month: Number(pick("month")),
    day: Number(pick("day")),
  };
}

function moscowLocalDate(year, month, day, hour, minute) {
  return new Date(
    `${year}-${pad2(month)}-${pad2(day)}T${pad2(hour)}:${pad2(minute)}:00${MOSCOW_OFFSET}`,
  );
}

function isMoscowWeekday(year, month, day) {
  const noon = moscowLocalDate(year, month, day, 12, 0);
  const weekday = new Intl.DateTimeFormat("en-US", {
    timeZone: MOSCOW_TZ,
    weekday: "short",
  }).format(noon);

  return weekday !== "Sat" && weekday !== "Sun";
}

function addMoscowCalendarDays(year, month, day, days) {
  const anchor = moscowLocalDate(year, month, day, 12, 0);
  anchor.setUTCDate(anchor.getUTCDate() + days);
  return getMoscowYmd(anchor);
}

const BADGE_THRESHOLD_HOURS = {
  red: 16,
  orange: 8,
  yellow: 6,
};

/** @type {Record<string, { background: string, text: string }>} */
export const BADGE_STYLES = {
  gray: { background: "#9e9e9e", text: "#ffffff" },
  yellow: { background: "#f9a825", text: "#1a1a1a" },
  orange: { background: "#ef6c00", text: "#ffffff" },
  red: { background: "#ca2c2c", text: "#ffffff" },
};

/**
 * @param {number|null} minutes
 * @returns {"yellow"|"orange"|"red"|null}
 */
export function getWorkingTimeUrgency(minutes) {
  if (minutes === null || minutes <= BADGE_THRESHOLD_HOURS.yellow * 60) {
    return null;
  }

  if (minutes > BADGE_THRESHOLD_HOURS.red * 60) {
    return "red";
  }

  if (minutes > BADGE_THRESHOLD_HOURS.orange * 60) {
    return "orange";
  }

  return "yellow";
}

/**
 * Есть ли коммиты/пуши позже последнего комментария участника группы.
 * Без открывающего тред комментария группы считаем, что «ожидание ревью» — счётчик включается.
 *
 * @param {{ lastCommitAt?: string, lastGroupCommentAt?: string }} item
 * @returns {boolean}
 */
export function hasUpdatesAfterLastGroupComment(item) {
  const commentAt = item?.lastGroupCommentAt;

  if (!commentAt) {
    return true;
  }

  const commitAt = item?.lastCommitAt;

  if (!commitAt) {
    return false;
  }

  const commentTimestamp = Date.parse(commentAt);
  const commitTimestamp = Date.parse(commitAt);

  if (!Number.isFinite(commentTimestamp) || !Number.isFinite(commitTimestamp)) {
    return true;
  }

  return commitTimestamp > commentTimestamp;
}

/**
 * Точка отсчёта рабочего времени для PR или `null`, если счётчик не нужен.
 *
 * @param {{ lastCommitAt?: string, lastGroupCommentAt?: string, createdAt?: string }} item
 * @returns {string|null|undefined}
 */
export function getItemWorkingTimeFrom(item) {
  if (!hasUpdatesAfterLastGroupComment(item)) {
    return null;
  }

  if (item?.lastGroupCommentAt && item?.lastCommitAt) {
    return item.lastCommitAt;
  }

  return item?.lastCommitAt ?? item?.createdAt;
}

/**
 * @param {{ updatedAt?: string, createdAt?: string, lastCommitAt?: string, lastGroupCommentAt?: string }} item
 * @param {string|null|undefined} checkedAt
 * @returns {"yellow"|"orange"|"red"|null}
 */
export function getItemWorkingTimeUrgency(item, checkedAt) {
  const from = getItemWorkingTimeFrom(item);

  if (!from || !checkedAt) {
    return null;
  }

  return getWorkingTimeUrgency(getWorkingElapsedMinutes(from, checkedAt));
}

/**
 * Уровень срочности бейджа по максимальному рабочему возрасту PR в списке.
 * @param {Array<{ updatedAt?: string, createdAt?: string, lastCommitAt?: string, lastGroupCommentAt?: string }>} items
 * @param {string|null|undefined} checkedAt
 * @returns {"gray"|"yellow"|"orange"|"red"}
 */
export function getBadgeUrgencyFromItems(items, checkedAt) {
  if (!checkedAt || !Array.isArray(items) || items.length === 0) {
    return "gray";
  }

  let maxMinutes = 0;

  for (const item of items) {
    const from = getItemWorkingTimeFrom(item);

    if (!from) {
      continue;
    }

    const minutes = getWorkingElapsedMinutes(from, checkedAt);

    if (minutes !== null && minutes > maxMinutes) {
      maxMinutes = minutes;
    }
  }

  return getWorkingTimeUrgency(maxMinutes) ?? "gray";
}

/** PR без новых пушей после комментария — в конце списка; остальные — старые выше. */
export function sortPullRequestsOldestFirst(items) {
  return items
    .map((item, index) => {
      const sortKey = item.updatedAt ?? item.createdAt;
      return {
        item,
        index,
        waiting: hasUpdatesAfterLastGroupComment(item),
        timestamp: sortKey ? Date.parse(sortKey) : Number.NaN,
      };
    })
    .sort((left, right) => {
      if (left.waiting !== right.waiting) {
        return left.waiting ? -1 : 1;
      }

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
