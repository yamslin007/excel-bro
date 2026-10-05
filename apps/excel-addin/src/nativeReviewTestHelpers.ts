// Test-only stored ZIP reader (no production imports).
export function readSeedEntries(base64: string): Record<string, string> {
  const bytes = Uint8Array.from(atob(base64), c => c.charCodeAt(0));
  const view = new DataView(bytes.buffer), result: Record<string, string> = {};
  let at = 0;
  while (view.getUint32(at, true) === 0x04034b50) {
    const size = view.getUint32(at + 18, true), nameSize = view.getUint16(at + 26, true);
    const start = at + 30 + nameSize + view.getUint16(at + 28, true);
    if (view.getUint16(at + 8, true) !== 0) throw new Error("Expected stored ZIP");
    result[new TextDecoder().decode(bytes.subarray(at + 30, at + 30 + nameSize))] = new TextDecoder().decode(bytes.subarray(start, start + size));
    at = start + size;
  }
  if (view.getUint32(at, true) !== 0x02014b50 || view.getUint32(bytes.length - 22, true) !== 0x06054b50) throw new Error("Invalid ZIP directory");
  return result;
}
