const MAX_AMOUNTS = 5;

const MINUS = "[-\u2212]?";
const NUMBER = String.raw`(?:\d{1,3}(?:[,.]\d{3})+|\d+)(?:[.,]\d{1,2})?(?!\d)`;
const SCALE = "(?:[kKMB](?![A-Za-z]))?";
const SPACE = "[ \u00a0]?";
const CODE = "(?:USD|CAD|EUR|GBP)";
const PREFIX_SYMBOL = String.raw`(?:US\$|CA\$|AU\$|A\$|C\$|\$|€|£|¥)`;
const SUFFIX_SYMBOL = String.raw`(?:€|£|¥|\$|${CODE}\b)`;
// A trailing symbol followed by a number is that number's prefix instead, so "Qty 2 $24.00"
// gives "$24.00" rather than "2 $".
const AMOUNT_PATTERN = new RegExp(
  [
    `${MINUS}${PREFIX_SYMBOL}${NUMBER}${SCALE}`,
    String.raw`${MINUS}\b${CODE}${SPACE}${NUMBER}${SCALE}`,
    String.raw`(?<![\w.,])${MINUS}${NUMBER}${SPACE}${SUFFIX_SYMBOL}(?!${SPACE}\d)`,
  ].join("|"),
  "g",
);

const WEEKDAY = String.raw`(?:Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday|Tues|Thurs|Thur|Mon|Tue|Wed|Thu|Fri|Sat|Sun)\.?,?\s+`;
const MONTH = String.raw`(?:January|February|March|April|May|June|July|August|September|October|November|December|Sept|Jan|Feb|Mar|Apr|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\b\.?`;
const DAY = String.raw`\d{1,2}(?:st|nd|rd|th)?\b`;
const YEAR = String.raw`(?:,?\s+\d{4}\b)?`;
const DATE = `(?:${WEEKDAY})?(?:${[
  String.raw`${MONTH}\s+${DAY}${YEAR}`,
  String.raw`${DAY}\s+${MONTH}${YEAR}`,
  String.raw`\d{4}-\d{2}-\d{2}(?!\d)`,
  String.raw`\d{1,2}\/\d{1,2}(?:\/(?:\d{4}|\d{2}))?(?![\d/])`,
].join("|")})`;
// "due" alone must sit right before the date (or a connector), so "due to weather on November 3" and
// "due today" give nothing.
const DUE_TRIGGER = String.raw`\b(?:payment\s+due|due\s+date|due\s+on|due\s+by|due|pay\s+by|payment\s+date|will\s+be\s+charged\s+on)`;
const CONNECTOR = String.raw`[\s:,-]*(?:\b(?:is|on|by)\b[\s:,-]*)*`;
const DUE_DATE_PATTERN = new RegExp(`${DUE_TRIGGER}${CONNECTOR}(${DATE})`, "i");

export function extractMailFacts({
  subject,
  snippet,
  body,
}: {
  subject?: string | null;
  snippet?: string | null;
  body?: string | null;
}): { amounts?: string[]; dueDate?: string } {
  const texts = [subject, snippet, body].filter((text): text is string =>
    Boolean(text),
  );

  const amounts = findAmounts(texts);
  const dueDate = findDueDate(texts);

  return {
    ...(amounts.length ? { amounts } : {}),
    ...(dueDate ? { dueDate } : {}),
  };
}

function findAmounts(texts: string[]) {
  const amounts = new Set<string>();
  for (const text of texts) {
    for (const match of text.matchAll(AMOUNT_PATTERN)) {
      amounts.add(match[0]);
      if (amounts.size >= MAX_AMOUNTS) return [...amounts];
    }
  }
  return [...amounts];
}

function findDueDate(texts: string[]) {
  for (const text of texts) {
    const match = DUE_DATE_PATTERN.exec(text);
    if (match) return match[1].replace(/[\s.,;:]+$/, "");
  }
  return;
}
