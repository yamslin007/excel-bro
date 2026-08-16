// A1 地址 / 选区解析的纯函数（从 App.tsx 抽出）。

// 从选区地址取首格：兼容 "Sheet1!E2:E50" / "E2:E50" / "E2"。
export function firstCellOfRange(address: string): string | null {
  const bare = address.includes("!") ? address.split("!").pop()! : address;
  const first = bare.split(":")[0]?.trim();
  if (!first || !/^\$?[A-Za-z]{1,3}\$?\d+$/.test(first)) return null;
  return first.replace(/\$/g, "");
}

// 列号 → 列字母（0→A, 25→Z, 26→AA）。
export function columnLetterFromIndex(index: number): string {
  let n = index;
  let letters = "";
  do {
    letters = String.fromCharCode(65 + (n % 26)) + letters;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return letters;
}

// 列字母 → 列号（A→0, Z→25, AA→26）。
export function columnIndexFromLetter(letter: string): number {
  let index = 0;
  for (const ch of letter.toUpperCase()) {
    index = index * 26 + (ch.charCodeAt(0) - 64);
  }
  return index - 1;
}

// 从 usedRange 地址取起始列字母（"A1:E20" → "A"，"Sheet1!C2:E9" → "C"）。
export function startColumnOfRange(address: string | null): string | null {
  if (!address) return null;
  const bare = address.includes("!") ? address.split("!").pop()! : address;
  const match = bare.match(/^\$?([A-Za-z]{1,3})/);
  return match ? match[1].toUpperCase() : null;
}

// 智能建议写入目标：数据区右侧第一空列 + 数据首行（表头下一行）。
// 解析 usedRange（如 "A1:D20"）取末列右移一列、起始行 +1。解析失败退回 "A2"。
export function suggestWriteTarget(usedRange: string | null): string {
  if (!usedRange) return "A2";
  const bare = usedRange.includes("!")
    ? usedRange.split("!").pop()!
    : usedRange;
  const match = bare.match(
    /^\$?([A-Za-z]{1,3})\$?(\d+)(?::\$?([A-Za-z]{1,3})\$?(\d+))?$/
  );
  if (!match) return "A2";
  const startRow = Number(match[2]);
  const endColLetter = (match[3] ?? match[1]).toUpperCase();
  const nextCol = columnLetterFromIndex(columnIndexFromLetter(endColLetter) + 1);
  return `${nextCol}${startRow + 1}`;
}
