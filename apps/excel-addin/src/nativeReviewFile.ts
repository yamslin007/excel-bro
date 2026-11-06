import type { CellValue } from "./contracts";

// Small, stored (uncompressed) OOXML package. No workbook export, network or persistence.
// Native Excel restores formatting from the authorized sheet after opening this seed.
const NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
export function nativeReviewXml(value: string): string {
  if (/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/.test(value)) throw new Error("单元格含不支持的控制字符，不能生成原生副本。");
  return value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}
const xml = nativeReviewXml;
export function nativeColumnName(column: number): string {
  let text = "";
  for (let n = column; n > 0; n = Math.floor((n - 1) / 26)) text = String.fromCharCode(65 + (n - 1) % 26) + text;
  return text;
}
function crc32(bytes: Uint8Array): number {
  let crc = 0xffffffff;
  for (const byte of bytes) {
    crc ^= byte;
    for (let bit = 0; bit < 8; bit++) crc = (crc >>> 1) ^ ((crc & 1) ? 0xedb88320 : 0);
  }
  return (crc ^ 0xffffffff) >>> 0;
}
export function storedZip(entries: Readonly<Record<string, string>>): Uint8Array {
  const encoder = new TextEncoder();
  const local: Uint8Array[] = [], central: Uint8Array[] = [];
  let offset = 0;
  for (const [path, content] of Object.entries(entries)) {
    const name = encoder.encode(path), data = encoder.encode(content), crc = crc32(data);
    const header = new Uint8Array(30 + name.length), h = new DataView(header.buffer);
    h.setUint32(0, 0x04034b50, true); h.setUint16(4, 20, true); h.setUint16(6, 0x800, true);
    h.setUint32(14, crc, true); h.setUint32(18, data.length, true); h.setUint32(22, data.length, true);
    h.setUint16(26, name.length, true); header.set(name, 30);
    const dir = new Uint8Array(46 + name.length), d = new DataView(dir.buffer);
    d.setUint32(0, 0x02014b50, true); d.setUint16(4, 20, true); d.setUint16(6, 20, true); d.setUint16(8, 0x800, true);
    d.setUint32(16, crc, true); d.setUint32(20, data.length, true); d.setUint32(24, data.length, true);
    d.setUint16(28, name.length, true); d.setUint32(42, offset, true); dir.set(name, 46);
    local.push(header, data); central.push(dir); offset += header.length + data.length;
  }
  const centralSize = central.reduce((sum, item) => sum + item.length, 0);
  const end = new Uint8Array(22), e = new DataView(end.buffer);
  e.setUint32(0, 0x06054b50, true); e.setUint16(8, central.length, true); e.setUint16(10, central.length, true);
  e.setUint32(12, centralSize, true); e.setUint32(16, offset, true);
  const output = new Uint8Array(offset + centralSize + end.length);
  let at = 0;
  for (const part of [...local, ...central, end]) { output.set(part, at); at += part.length; }
  return output;
}
export function nativeReviewSeed(sheetName: string, startRow: number, startColumn: number, values: readonly (readonly CellValue[])[], maxBytes: number): string {
  const rows = values.map((row, r) => `<row r="${startRow + r}">${row.map((value, c) => {
    if (value === null) return "";
    const address = `${nativeColumnName(startColumn + c)}${startRow + r}`;
    if (typeof value === "number") {
      if (!Number.isFinite(value)) throw new Error("单元格包含非有限数值，不能生成副本。");
      return `<c r="${address}"><v>${value}</v></c>`;
    }
    if (typeof value === "boolean") return `<c r="${address}" t="b"><v>${Number(value)}</v></c>`;
    return `<c r="${address}" t="inlineStr"><is><t xml:space="preserve">${xml(value)}</t></is></c>`;
  }).join("")}</row>`).join("");
  return nativeReviewPackage(sheetName, `<worksheet xmlns="${NS}"><sheetData>${rows}</sheetData></worksheet>`, maxBytes);
}
export function nativeReviewPackage(sheetName: string, sheetXml: string, maxBytes: number, stylesXml?: string, reviewCopyId?: string): string {
  return nativeReviewWorkbookPackage([{name: sheetName, xml: sheetXml}], maxBytes, stylesXml, reviewCopyId);
}
export function nativeReviewWorkbookPackage(sheets: Array<{name: string; xml: string}>, maxBytes: number, stylesXml?: string, reviewCopyId?: string, activeTab = 0): string {
  if (!sheets.length || new Set(sheets.map(sheet => sheet.name.toLowerCase())).size !== sheets.length ||
      sheets.some(sheet => !sheet.name || sheet.name.length > 31 || /[\\/?*\[\]:]/.test(sheet.name))) throw new Error('对比子表名称无效或重复。');
  if (reviewCopyId !== undefined && !/^EB_[0-9a-f]{20}$/.test(reviewCopyId)) throw new Error('对比副本标识无效。');
  const identity = reviewCopyId ? `<definedNames><definedName name="_EB_ReviewCopyId" hidden="1">&quot;${reviewCopyId}&quot;</definedName></definedNames>` : '';
  const bytes = storedZip({
    "[Content_Types].xml": `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>${sheets.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')}${stylesXml ? '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' : ''}</Types>`,
    "_rels/.rels": '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
    "xl/workbook.xml": `<workbook xmlns="${NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView activeTab="${activeTab}"/></bookViews><sheets>${sheets.map((sheet,i) => `<sheet name="${xml(sheet.name)}" sheetId="${i+1}" r:id="rId${i+1}"/>`).join('')}</sheets>${identity}${stylesXml ? '<calcPr calcId="0" fullCalcOnLoad="1" forceFullCalc="1"/>' : ''}</workbook>`,
    "xl/_rels/workbook.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${sheets.map((_,i) => `<Relationship Id="rId${i+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i+1}.xml"/>`).join('')}${stylesXml ? `<Relationship Id="rId${sheets.length+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>` : ''}</Relationships>`,
    ...Object.fromEntries(sheets.map((sheet,i) => [`xl/worksheets/sheet${i+1}.xml`, sheet.xml])),
    ...(stylesXml ? { "xl/styles.xml": stylesXml } : {})
  });
  if (bytes.length > maxBytes) throw new Error("原生审阅副本超过本机大小限额，请缩小工作表后重试。");
  let binary = "";
  for (let i = 0; i < bytes.length; i += 8192) binary += String.fromCharCode(...bytes.subarray(i, i + 8192));
  return btoa(binary);
}
