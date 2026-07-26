export function workbookNameFromDocumentUrl(documentUrl?: string | null): string {
  const raw = documentUrl?.trim();
  if (!raw) return "未保存的工作簿";

  const withoutQuery = raw.split(/[?#]/, 1)[0].replace(/\\/g, "/");
  const encodedName = withoutQuery.slice(withoutQuery.lastIndexOf("/") + 1);
  if (!encodedName) return "未保存的工作簿";

  try {
    return decodeURIComponent(encodedName);
  } catch {
    return encodedName;
  }
}

function normalizedDate(year: string, month: string, day: string): string {
  return `${year}-${month.padStart(2, "0")}-${day.padStart(2, "0")}`;
}

export function extractWorkbookDataPeriod(
  workbookName: string
): string | null {
  const matches = [
    ...workbookName.matchAll(
      /(\d{4})\s*(?:-|\/|\.|年)\s*(\d{1,2})\s*(?:-|\/|\.|月)\s*(\d{1,2})\s*日?/g
    )
  ].map((match) => normalizedDate(match[1], match[2], match[3]));

  if (matches.length === 0) return null;
  if (matches.length === 1 || matches[0] === matches[1]) return matches[0];
  return `${matches[0]} 至 ${matches[1]}`;
}
