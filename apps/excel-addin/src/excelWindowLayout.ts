import { prepareExcelReviewLayout, arrangeExcelReviewLayout, measureExcelReviewPane } from './api';
import type { ExcelReviewLayoutTicket } from './contracts';

const sha256 = async (text: string) => [...new Uint8Array(await crypto.subtle.digest('SHA-256', new TextEncoder().encode(text)))].map(x => x.toString(16).padStart(2, '0')).join('');

// This normalization is for the transient layout bridge ONLY. It does not change
// officeIdentityFromUrl or any persisted execution/authorization identity.
export function localWindowsPath(identity: string): string {
  if (!identity.startsWith('url:')) throw new Error('需要已保存的本机工作簿才能自动排列。');
  let path = identity.slice(4).replace(/\\/g, '/');
  if (path.startsWith('file:')) {
    const url = new URL(path);
    path = decodeURIComponent(url.pathname);
    if (url.hostname && url.hostname !== 'localhost') path = `//${url.hostname}${path}`;
    else path = path.replace(/^\/([a-z]:\/)/i, '$1');
  }
  if (!/^(?:[a-z]:\/|\/\/[^/]+\/[^/]+)/i.test(path) || /[?#\u0000-\u001f]/.test(path)) throw new Error('该工作簿地址不支持本机自动窗口排列。');
  // Match ntpath.normpath on the native side; never expand relative paths.
  const unc = path.startsWith('//');
  const parts = path.slice(unc ? 2 : 3).split('/').filter(Boolean);
  const root = unc ? parts.splice(0, 2).join('/') : path.slice(0, 2);
  const normalized: string[] = [];
  for (const part of parts) {
    if (part === '.') continue;
    if (part === '..') normalized.pop(); else normalized.push(part);
  }
  return `${unc ? '//' : ''}${root}/${normalized.join('/')}`.replace(/\/$/, '').toLowerCase();
}
export async function prepareLocalExcelWindows(identity: string, sourceSheetName: string, copySheetName: string): Promise<ExcelReviewLayoutTicket> {
  return prepareExcelReviewLayout({workbookPathHash: await sha256(localWindowsPath(identity)),
    worksheetNameHash: await sha256(sourceSheetName), copySheetName});
}
export { arrangeExcelReviewLayout as arrangeLocalExcelWindows };
export async function measureLocalExcelWindow(identity: string, sourceSheetName: string, copySheetName: string) {
  return measureExcelReviewPane({workbookPathHash: await sha256(localWindowsPath(identity)),
    worksheetNameHash: await sha256(sourceSheetName), copySheetName});
}
