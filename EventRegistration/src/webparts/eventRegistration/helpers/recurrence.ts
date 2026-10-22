export interface IRecurrenceSettings {
  type: string;
  interval: number;
  weekdays: number[];
  until: Date;
}

export interface IOccurrence {
  start: Date;
  end: Date;
  key: string;
}

function pad(value: number): string {
  return value < 10 ? '0' + String(value) : String(value);
}

export function occurrenceKey(eventId: number, start: Date): string {
  return String(eventId) + '-'
    + String(start.getUTCFullYear()) + pad(start.getUTCMonth() + 1) + pad(start.getUTCDate())
    + 'T' + pad(start.getUTCHours()) + pad(start.getUTCMinutes()) + pad(start.getUTCSeconds()) + 'Z';
}

export function normalizeWeekdays(values: number[]): number[] {
  var found: { [key: string]: boolean } = {};
  var result: number[] = [];
  (values || []).forEach(function(value: number): void {
    if (value >= 0 && value <= 6 && !found[String(value)]) {
      found[String(value)] = true;
      result.push(value);
    }
  });
  return result.sort(function(a: number, b: number): number { return a - b; });
}

function sameOrBefore(left: Date, right: Date): boolean {
  return left.getTime() <= right.getTime();
}

export function generateOccurrences(
  eventId: number,
  firstStart: Date,
  firstEnd: Date,
  settings: IRecurrenceSettings,
  maximum?: number
): IOccurrence[] {
  if (!firstStart || !firstEnd || isNaN(firstStart.getTime()) || isNaN(firstEnd.getTime())) {
    throw new Error('Start and end must be valid dates.');
  }
  if (firstEnd.getTime() <= firstStart.getTime()) {
    throw new Error('Session end must be after its start.');
  }
  var limit = maximum && maximum > 0 ? maximum : 366;
  var interval = settings.interval > 0 ? Math.floor(settings.interval) : 1;
  var until = settings.until && !isNaN(settings.until.getTime()) ? settings.until : firstStart;
  var duration = firstEnd.getTime() - firstStart.getTime();
  var type = String(settings.type || 'none').toLowerCase();
  var occurrences: IOccurrence[] = [];

  if (type === 'none') {
    return [{ start: firstStart, end: firstEnd, key: occurrenceKey(eventId, firstStart) }];
  }

  if (type === 'daily') {
    var daily = new Date(firstStart.getTime());
    while (sameOrBefore(daily, until) && occurrences.length < limit) {
      occurrences.push({
        start: new Date(daily.getTime()),
        end: new Date(daily.getTime() + duration),
        key: occurrenceKey(eventId, daily)
      });
      daily.setDate(daily.getDate() + interval);
    }
    return occurrences;
  }

  if (type !== 'weekly') {
    throw new Error('Recurrence type must be none, daily, or weekly.');
  }
  var weekdays = normalizeWeekdays(settings.weekdays || []);
  if (weekdays.length === 0) {
    weekdays = [firstStart.getDay()];
  }
  var cursor = new Date(firstStart.getTime());
  cursor.setHours(firstStart.getHours(), firstStart.getMinutes(), firstStart.getSeconds(), firstStart.getMilliseconds());
  while (sameOrBefore(cursor, until) && occurrences.length < limit) {
    var elapsedDays = Math.floor((cursor.getTime() - firstStart.getTime()) / 86400000);
    var weekIndex = Math.floor(elapsedDays / 7);
    if (weekIndex % interval === 0 && weekdays.indexOf(cursor.getDay()) >= 0) {
      occurrences.push({
        start: new Date(cursor.getTime()),
        end: new Date(cursor.getTime() + duration),
        key: occurrenceKey(eventId, cursor)
      });
    }
    cursor.setDate(cursor.getDate() + 1);
  }
  return occurrences;
}
