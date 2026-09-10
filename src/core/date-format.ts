// Rendering a `Date` through an Excel number-format code.
//
// Apart from `./date.ts`, which owns the serial arithmetic and the *classifier* (`isDateFormat`),
// because this is the only half either entry can decline to carry: the renderer is four kilobytes of
// month names and grammar, and `date.ts` is in the closure of every entry point there is while
// nothing but the CSV writer renders a date today. Splitting it costs one import at the two call
// sites and keeps four entries from paying for a table they never read.
//
// It is still the same vocabulary, and that is the point of ADR 0041: `Cell.numFmt` holds a format
// code, `isDateFormat` next door classifies one, and this renders one. The CSV writer used to carry a
// moment.js token table instead, in which the month is `MM` and `mm` is minutes, so a caller who
// passed a format code from this library got silent garbage back.

// The month and weekday names an Excel format code's `mmm`/`mmmm`/`ddd`/`dddd` render.
//
// English, and stated rather than taken from `Intl`. A format code carries its own locale directive
// (`[$-409]`), which is a fact about the *file*, while `Intl` would render whatever locale the
// process happens to run under: the same workbook would produce different CSV on a developer's
// machine and on a build server. Honouring the directive is the eventual answer; guessing from the
// environment is not a step toward it.
const MONTH_NAMES = [
  'January',
  'February',
  'March',
  'April',
  'May',
  'June',
  'July',
  'August',
  'September',
  'October',
  'November',
  'December',
] as const;

const WEEKDAY_NAMES = [
  'Sunday',
  'Monday',
  'Tuesday',
  'Wednesday',
  'Thursday',
  'Friday',
  'Saturday',
] as const;

// One run of a format code: a run of one placeholder letter, a literal, or a directive to drop. The
// scan produces these first and renders second, because two of the decisions -- whether `m` is a
// month or a minute, and whether `h` counts to 12 or to 24 -- cannot be made from the run alone.
interface FormatRun {
  readonly kind: 'placeholder' | 'literal';
  /** For a placeholder, the lower-cased letter; for a literal, the text to emit. */
  readonly text: string;
  /** For a placeholder, how many times the letter repeated. */
  readonly count: number;
  /** For a short meridiem (`a/p`), the token as the code spells it, whose letters keep their case. */
  readonly written?: string;
}

/**
 * Split a number-format code into the runs {@link formatSerialDate} renders.
 *
 * Only the code's first section is read. A format may carry up to four (`positive;negative;zero;
 * text`), and a date is never negative or zero in any sense the sections distinguish, so the first
 * is the one that applies. Bracketed directives -- colour, locale, condition -- are dropped, quoted
 * literals and backslash escapes emit their content, and everything else is literal text.
 */
function formatRuns(code: string): FormatRun[] {
  const runs: FormatRun[] = [];
  const literal = (text: string) => {
    if (text !== '') runs.push({kind: 'literal', text, count: text.length});
  };
  for (let i = 0; i < code.length; i++) {
    const ch = code[i] as string;
    if (ch === ';') break;
    if (ch === '[') {
      const end = code.indexOf(']', i);
      i = end === -1 ? code.length : end;
      continue;
    }
    if (ch === '"') {
      const end = code.indexOf('"', i + 1);
      literal(code.slice(i + 1, end === -1 ? code.length : end));
      i = end === -1 ? code.length : end;
      continue;
    }
    if (ch === '\\') {
      literal(code[i + 1] ?? '');
      i += 1;
      continue;
    }
    // A run of zeros behind a decimal point is a fractional second, not a numeric placeholder and
    // not literal text: `ss.00` renders hundredths. Emitting the `.00` verbatim, which is what a
    // renderer that did not know the form would do, is silently wrong rather than unsupported. Only a
    // seconds placeholder opens one: a quoted or escaped `s` is a literal whose text happens to match.
    const fraction = /^\.0+/.exec(code.slice(i));
    const previous = runs.at(-1);
    if (fraction !== null && previous?.kind === 'placeholder' && previous.text === 's') {
      runs.push({kind: 'placeholder', text: '.0', count: fraction[0].length - 1});
      i += fraction[0].length - 1;
      continue;
    }
    const lower = ch.toLowerCase();
    if (lower === 'y' || lower === 'm' || lower === 'd' || lower === 'h' || lower === 's') {
      let count = 1;
      while (code[i + count]?.toLowerCase() === lower) count += 1;
      runs.push({kind: 'placeholder', text: lower, count});
      i += count - 1;
      continue;
    }
    // `AM/PM` and its `A/P` short form are one token, not letters to be rendered individually. They
    // render differently, so the run keeps which one matched, and the short form keeps its spelling.
    const meridiem = /^(AM\/PM|A\/P)/i.exec(code.slice(i));
    if (meridiem !== null) {
      const token = meridiem[0];
      if (token.length === 3)
        runs.push({kind: 'placeholder', text: 'a/p', count: 3, written: token});
      else runs.push({kind: 'placeholder', text: 'am/pm', count: token.length});
      i += token.length - 1;
      continue;
    }
    literal(ch);
  }
  return runs;
}

/**
 * Whether the `m` run at `index` means minutes rather than months.
 *
 * Excel's own rule, and the whole reason this renderer cannot be a token table: `m` is a minute when
 * it sits immediately after an hour run or immediately before a seconds run, and a month otherwise.
 * Literals between the runs do not break the adjacency, which is what makes `hh:mm` and `h" hours "mm`
 * agree.
 */
function isMinuteRun(runs: readonly FormatRun[], index: number): boolean {
  for (let i = index - 1; i >= 0; i--) {
    const run = runs[i] as FormatRun;
    if (run.kind === 'literal') continue;
    if (run.text === 'h') return true;
    break;
  }
  for (let i = index + 1; i < runs.length; i++) {
    const run = runs[i] as FormatRun;
    if (run.kind === 'literal') continue;
    return run.text === 's';
  }
  return false;
}

/**
 * Render a `Date` through an Excel number-format code: `yyyy-mm-dd`, `d mmm yy`, `hh:mm:ss`.
 *
 * The one vocabulary this library speaks for dates. `Cell.numFmt` holds a code, `DEFAULT_DATE_NUMFMT`
 * is one, `isDateFormat` classifies one, and the CSV writer's `dateFormat` used to be a moment.js
 * token set instead -- case-sensitive, with `mm` meaning minutes and the month spelled `MM` -- so
 * `writeCsv(wb, {dateFormat: cell.numFmt})`, the obvious thing to write, rendered every date cell as
 * `2024-45-dd` with no throw and no warning. A library with two public date-format properties
 * speaking different languages has no defensible answer to "which one is this".
 *
 * Codes are case-insensitive, as Excel's are. What is *not* covered, deliberately: the elapsed-time
 * forms (`[h]`, `[mm]`), which measure a duration rather than name a moment, and the locale directive
 * (`[$-409]`), which is dropped rather than honoured -- month and weekday names render in English.
 * Both are recorded in ADR 0041.
 *
 * @param utc read the date's fields in UTC rather than in the runner's local time. A serial carries
 *   no timezone, so a `Date` that came from one is UTC wall-clock and this should be `true`; a `Date`
 *   a caller assigned may mean either.
 */
export function formatSerialDate(date: Date, code: string, utc: boolean): string {
  if (Number.isNaN(date.getTime())) return '';
  const runs = formatRuns(code);
  const year = utc ? date.getUTCFullYear() : date.getFullYear();
  const month = (utc ? date.getUTCMonth() : date.getMonth()) + 1;
  const day = utc ? date.getUTCDate() : date.getDate();
  const weekday = utc ? date.getUTCDay() : date.getDay();
  const hour24 = utc ? date.getUTCHours() : date.getHours();
  const minute = utc ? date.getUTCMinutes() : date.getMinutes();
  const second = utc ? date.getUTCSeconds() : date.getSeconds();
  // A `h` run counts to twelve only when the code asks for a meridiem, which is Excel's rule and the
  // reason the runs are collected before any of them is rendered.
  const twelveHour = runs.some(
    (run) => run.kind === 'placeholder' && (run.text === 'am/pm' || run.text === 'a/p'),
  );
  const hour = twelveHour ? hour24 % 12 || 12 : hour24;
  const pad = (value: number, width: number) => String(value).padStart(width, '0');

  let out = '';
  for (const [index, run] of runs.entries()) {
    if (run.kind === 'literal') {
      out += run.text;
      continue;
    }
    switch (run.text) {
      case 'y':
        // Excel takes one or two `y` as the two-digit year and three or more as the four-digit one.
        out += run.count <= 2 ? pad(year % 100, 2) : String(year);
        break;
      case 'm':
        if (isMinuteRun(runs, index)) {
          out += run.count >= 2 ? pad(minute, 2) : String(minute);
        } else if (run.count >= 5) {
          out += (MONTH_NAMES[month - 1] as string)[0];
        } else if (run.count === 4) {
          out += MONTH_NAMES[month - 1] as string;
        } else if (run.count === 3) {
          out += (MONTH_NAMES[month - 1] as string).slice(0, 3);
        } else {
          out += run.count === 2 ? pad(month, 2) : String(month);
        }
        break;
      case 'd':
        if (run.count >= 4) out += WEEKDAY_NAMES[weekday] as string;
        else if (run.count === 3) out += (WEEKDAY_NAMES[weekday] as string).slice(0, 3);
        else out += run.count === 2 ? pad(day, 2) : String(day);
        break;
      case 'h':
        out += run.count >= 2 ? pad(hour, 2) : String(hour);
        break;
      case 's':
        out += run.count >= 2 ? pad(second, 2) : String(second);
        break;
      case '.0': {
        // Milliseconds are all a `Date` carries, so more than three digits pad with zeros rather
        // than inventing precision.
        const digits = pad(utc ? date.getUTCMilliseconds() : date.getMilliseconds(), 3);
        out += `.${(digits + '000').slice(0, run.count)}`;
        break;
      }
      // Excel Desktop's rendering, read back over COM, not the Format Cells reference's: the long form
      // is capitals however the code spells it (`am/pm` and `Am/Pm` render `AM`), and the short form
      // keeps each letter's own case (`a/p` renders `a`, and `A/p` renders `A` before noon, `p` after).
      case 'am/pm':
        out += hour24 < 12 ? 'AM' : 'PM';
        break;
      case 'a/p': {
        const written = run.written ?? 'A/P';
        out += hour24 < 12 ? written.charAt(0) : written.charAt(2);
        break;
      }
      default:
        break;
    }
  }
  return out;
}
