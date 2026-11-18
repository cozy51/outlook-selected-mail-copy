// cozy51/OLM-Achiver の lib.js / markdown.js にある日時の読み取り・整形を移植。

// Outlook の日本語表示「2026-07-13 (月) 20:41」「2026年7月13日 20:41」などを
// 数値から組み立てて Date にする（ブラウザ依存の解釈を避ける）。
export function parseMailDate(value) {
  if (value instanceof Date) return Number.isNaN(value.valueOf()) ? null : value;
  const source = String(value || "").trim();
  if (!source) return null;
  const match = source.match(/(\d{4})\s*(?:年|[-/.])\s*(\d{1,2})\s*(?:月|[-/.])\s*(\d{1,2})\s*日?(?:\s*\([^)]*\))?\s*T?\s*(\d{1,2})\s*[:：]\s*(\d{2})/);
  if (match) {
    const [, year, month, day, hour, minute] = match.map(Number);
    const parsed = new Date(year, month - 1, day, hour, minute);
    if (parsed.getFullYear() === year && parsed.getMonth() === month - 1 && parsed.getDate() === day) return parsed;
    return null;
  }
  // 日付らしさのない文字列（添付ファイル名など）を Date.parse に通すと
  // 思わぬ値になるため、数字の並びが日付として読めるものだけを対象にする。
  if (!/\d{1,4}\s*[-/.]\s*\d{1,2}|\d{1,2}\s*[:：]\s*\d{2}/.test(source)) return null;
  const parsed = new Date(source);
  return Number.isNaN(parsed.valueOf()) ? null : parsed;
}

const WEEKDAYS = ["日", "月", "火", "水", "木", "金", "土"];

// 読めた日時は「2026-07-13 (月) 20:41」に揃える。読めなければ空文字（見出しごと出さない）。
export function formatMailDate(value) {
  const instant = parseMailDate(value);
  if (!instant) return "";
  const pad = number => String(number).padStart(2, "0");
  return `${instant.getFullYear()}-${pad(instant.getMonth() + 1)}-${pad(instant.getDate())} (${WEEKDAYS[instant.getDay()]}) ${pad(instant.getHours())}:${pad(instant.getMinutes())}`;
}
