import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { createHash, webcrypto } from 'node:crypto';
import { localWindowsPath, prepareLocalExcelWindows, arrangeLocalExcelWindows, closeLocalExcelWindows } from './excelWindowLayout';
import { prepareExcelReviewLayout, measureExcelReviewPane, ExcelReviewLayoutVersionError } from './api';
import type { AnalysisPlan, ExcelReviewLayoutPrepareRequest, ExcelReviewLayoutTicket } from './contracts';

beforeEach(() => vi.stubGlobal('crypto', webcrypto));
afterEach(() => { vi.unstubAllGlobals(); vi.unstubAllEnvs(); });
const hash = (text: string) => createHash('sha256').update(text).digest('hex');
const token = 'T'.repeat(43), copySheetName = 'EB_' + 'a'.repeat(20);
const ticket: ExcelReviewLayoutTicket = {token, layoutVersion:3};
const arranged = {status:'arranged', layoutVersion:3, copySide:'left', originalSide:'right', synchronousScrolling:'enabled'};

describe('local-only window bridge', () => {
  it('measurement request sends metadata only, never the local plan or full path', async () => {
    const privatePlan = {actions:[{values:[['LOCAL_PRIVATE_SENTINEL']]}]};
    expect(JSON.stringify(privatePlan)).toContain('LOCAL_PRIVATE_SENTINEL');
    const request = {workbookPathHash:hash('e:/original.xlsx'),worksheetNameHash:hash('Ledger'),copySheetName,
      plan:privatePlan,fullPath:'E:/SECRET/Original.xlsx',hwnd:123};
    const fetch=vi.fn(async (_url:string,_init:RequestInit)=>({ok:true,json:async()=>({clientWidthCssPx:1200,workAreaWidthCssPx:2000})}));
    vi.stubGlobal('fetch',fetch);
    expect(await measureExcelReviewPane(request)).toEqual({clientWidthCssPx:1200,workAreaWidthCssPx:2000});
    expect(JSON.parse(String(fetch.mock.calls[0][1].body))).toEqual({workbookPathHash:request.workbookPathHash,worksheetNameHash:request.worksheetNameHash,copySheetName});
    expect(fetch.mock.calls[0][0]).toMatch(/review-layout\/measure$/);
    expect(fetch.mock.calls[0][1].redirect).toBe('error');
    expect(JSON.stringify(fetch.mock.calls)).not.toMatch(/LOCAL_PRIVATE_SENTINEL|SECRET|hwnd|plan|values/);
  });
  it.each([0,-1,NaN,Infinity,'1200',32769])('rejects invalid measured width %s', async width => {
    vi.stubGlobal('fetch',vi.fn(async()=>({ok:true,json:async()=>({clientWidthCssPx:1200,workAreaWidthCssPx:width})})));
    await expect(measureExcelReviewPane({workbookPathHash:hash('e:/original.xlsx'),worksheetNameHash:hash('Ledger'),copySheetName})).rejects.toThrow();
  });
  it.each([-1,NaN,Infinity,'2000',32769])('rejects invalid work-area width %s', async workAreaWidthCssPx => {
    vi.stubGlobal('fetch',vi.fn(async()=>({ok:true,json:async()=>({clientWidthCssPx:1200,workAreaWidthCssPx})})));
    await expect(measureExcelReviewPane({workbookPathHash:hash('e:/original.xlsx'),worksheetNameHash:hash('Ledger'),copySheetName})).rejects.toThrow();
  });
  it.each([
    ['legacy single-field measurement', {clientWidthCssPx:1200}],
    ['work area smuggled behind extra keys', {clientWidthCssPx:1200,workAreaWidthCssPx:2000,note:'x'}],
    ['client width missing', {workAreaWidthCssPx:2000}],
  ])('refuses to accept a %s as verified geometry', async (_label, payload) => {
    vi.stubGlobal('fetch',vi.fn(async()=>({ok:true,json:async()=>payload})));
    await expect(measureExcelReviewPane({workbookPathHash:hash('e:/original.xlsx'),worksheetNameHash:hash('Ledger'),copySheetName})).rejects.toThrow();
  });
  it.each([
    ['url:file:///E:/Case%20Folder/Original.xlsx', 'e:/case folder/original.xlsx'],
    ['url:E:\\Case Folder\\temp\\..\\Original.xlsx', 'e:/case folder/original.xlsx'],
    ['url:file://server/share/temp/../Original.xlsx', '//server/share/original.xlsx'],
    ['url:\\\\server\\share\\Original.xlsx', '//server/share/original.xlsx'],
  ])('canonical path hash agrees for %s', async (identity, canonical) => {
    expect(localWindowsPath(identity)).toBe(canonical);
    const fetch = vi.fn(async (_url: string, _init: RequestInit) => ({ok:true, json:async () => ticket})); vi.stubGlobal('fetch', fetch);
    await prepareLocalExcelWindows(identity, 'Ledger', copySheetName);
    const body = JSON.parse(String(fetch.mock.calls[0][1]?.body));
    expect(body).toEqual({workbookPathHash:hash(canonical),worksheetNameHash:hash('Ledger'),copySheetName});
  });
  it.each(['url:https://example.com/a.xlsx', 'unsaved', 'url:relative.xlsx', 'url:E:/file.xlsx?secret'])('does not send unsupported identity %s', async identity => {
    const fetch = vi.fn(); vi.stubGlobal('fetch', fetch);
    await expect(prepareLocalExcelWindows(identity, 'Ledger', copySheetName)).rejects.toThrow();
    expect(fetch).not.toHaveBeenCalled();
  });
  it('real fetch bodies project metadata only even if a local plan is accidentally attached', async () => {
    const plan: AnalysisPlan = {id:'private',title:'PRIVATE_TITLE',summary:'',assumptions:[],warnings:[],actions:[{type:'writeValues',sheet:'Ledger',range:'C2',values:[['LOCAL_PRIVATE_SENTINEL']]}]};
    expect(JSON.stringify(plan)).toContain('LOCAL_PRIVATE_SENTINEL'); // positive anchor, not empty privacy test
    const fetch = vi.fn(async (url: string, _init: RequestInit) => ({ok:true,json:async () => url.endsWith('/prepare') ? ticket : arranged}));
    vi.stubGlobal('fetch', fetch);
    const request = {workbookPathHash:hash('e:/original.xlsx'),worksheetNameHash:hash('Ledger'),copySheetName,plan,values:plan.actions,fullPath:'E:/SECRET/Original.xlsx'};
    const preparedTicket = await prepareExcelReviewLayout(request as ExcelReviewLayoutPrepareRequest);
    await arrangeLocalExcelWindows({...preparedTicket, plan} as typeof preparedTicket);
    expect(JSON.parse(String(fetch.mock.calls[0][1].body))).toEqual({workbookPathHash:request.workbookPathHash,worksheetNameHash:request.worksheetNameHash,copySheetName});
    expect(JSON.parse(String(fetch.mock.calls[1][1].body))).toEqual({token});
    expect(fetch.mock.calls.every(([,init]) => init.redirect === 'error')).toBe(true);
    expect(JSON.stringify(fetch.mock.calls)).not.toMatch(/LOCAL_PRIVATE_SENTINEL|PRIVATE_TITLE|SECRET|writeValues|values/);
    expect(fetch.mock.calls.every(([url]) => url.startsWith('http://127.0.0.1:8765/api/office/review-layout/'))).toBe(true);
  });
  it.each([{token:'short',layoutVersion:3}, {...ticket,values:[['secret']]}, {status:'arranged'}])('rejects malformed preparation evidence %#', async response => {
    vi.stubGlobal('fetch', vi.fn(async () => ({ok:true,json:async () => response})));
    await expect(prepareLocalExcelWindows('url:E:/Original.xlsx', 'Ledger', copySheetName)).rejects.toThrow();
  });
  it.each([{status:'requested'}, {...arranged,values:['secret']}, {}, {...arranged,copySide:'right',originalSide:'left'}, {...arranged,originalSide:'left'}, {...arranged,synchronousScrolling:'assumed'}, {status:'arranged',layoutVersion:3,copySide:'left',originalSide:'right'}])('does not claim unverified layout %#', async response => {
    vi.stubGlobal('fetch', vi.fn(async () => ({ok:true,json:async () => response})));
    await expect(arrangeLocalExcelWindows(ticket)).rejects.toThrow();
  });
  it.each(['enabled', 'different_instance', 'unavailable'])('preserves the independent scrolling receipt %s', async synchronousScrolling => {
    vi.stubGlobal('fetch', vi.fn(async () => ({ok:true,json:async () => ({...arranged,synchronousScrolling})})));
    expect((await arrangeLocalExcelWindows(ticket)).synchronousScrolling).toBe(synchronousScrolling);
  });
  it('rejects layout-only version 2 before attempting movement', async () => {
    const fetch = vi.fn(async () => ({ok:true,json:async () => ({token,layoutVersion:2})})); vi.stubGlobal('fetch',fetch);
    await expect(prepareLocalExcelWindows('url:E:/Original.xlsx','Ledger',copySheetName)).rejects.toBeInstanceOf(ExcelReviewLayoutVersionError);
    expect(fetch).toHaveBeenCalledOnce();
  });
  it('rejects the running legacy server before requesting its old window order', async () => {
    const fetch = vi.fn(async () => ({ok:true,json:async () => ({token})})); vi.stubGlobal('fetch',fetch);
    await expect(prepareLocalExcelWindows('url:E:/Original.xlsx','Ledger',copySheetName)).rejects.toBeInstanceOf(ExcelReviewLayoutVersionError);
    expect(fetch).toHaveBeenCalledOnce();
    await expect(arrangeLocalExcelWindows({token} as ExcelReviewLayoutTicket)).rejects.toBeInstanceOf(ExcelReviewLayoutVersionError);
    expect(fetch).toHaveBeenCalledOnce(); // no movement request using an old ticket
  });
  it('rejects a legacy generic success after a server switch', async () => {
    vi.stubGlobal('fetch',vi.fn(async () => ({ok:true,json:async () => ({status:'arranged'})})));
    await expect(arrangeLocalExcelWindows(ticket)).rejects.toBeInstanceOf(ExcelReviewLayoutVersionError);
  });
  it('does not send hashes or tickets to a configured remote server', async () => {
    vi.stubEnv('VITE_API_BASE_URL','https://example.com'); vi.resetModules();
    const api = await import('./api'), fetch = vi.fn(); vi.stubGlobal('fetch',fetch);
    await expect(api.arrangeExcelReviewLayout(ticket)).rejects.toThrow('仅能连接本机');
    expect(fetch).not.toHaveBeenCalled(); vi.resetModules();
  });
});

/*
 * 兼容副本的**关闭**走同一条本机桥，但语义完全不同：
 * 它回答的是"副本关没关"，绝不是"原表写回成没成"。任何自相矛盾、缺字段、
 * 版本过旧或越界的回执都不许被解释成"已关闭"。
 */
const CLOSE_STATUSES = ['closed','already_closed','close_unconfirmed','user_saved','identity_lost','ambiguous','unsupported','failed'] as const;
const receipt = (status: string, copyClosed = status === 'closed' || status === 'already_closed') => ({status, layoutVersion:4, copyClosed, recoveryCopyPath: 'E:/recovery/review-copy.xlsx'});
describe('review copy close bridge', () => {
  it('sends metadata only and projects the versioned close receipt and local archive path', async () => {
    const fetch = vi.fn(async (_url:string,_init:RequestInit) => ({ok:true, json: async () => receipt('closed')}));
    vi.stubGlobal('fetch', fetch);
    expect(await closeLocalExcelWindows('url:file:///E:/Original.xlsx','Ledger',copySheetName))
      .toEqual({status:'closed', layoutVersion:4, copyClosed:true, recoveryCopyPath:'E:/recovery/review-copy.xlsx'});
    expect(JSON.parse(String(fetch.mock.calls[0][1].body)))
      .toEqual({closeVersion:4,workbookPathHash:hash('e:/original.xlsx'), worksheetNameHash:hash('Ledger'), copySheetName});
    expect(fetch.mock.calls[0][0]).toMatch(/review-layout\/close$/);
    expect(fetch.mock.calls[0][1].redirect).toBe('error');
    expect(JSON.stringify(fetch.mock.calls)).not.toMatch(/LOCAL_PRIVATE_SENTINEL|SECRET|plan|values|hwnd/);
  });
  it('never lets an attached local plan or full path reach the close request', async () => {
    const plan: AnalysisPlan = {id:'private',title:'PRIVATE_TITLE',summary:'',assumptions:[],warnings:[],actions:[{type:'writeValues',sheet:'Ledger',range:'C2',values:[['LOCAL_PRIVATE_SENTINEL']]}]};
    expect(JSON.stringify(plan)).toContain('LOCAL_PRIVATE_SENTINEL'); // positive anchor
    const fetch = vi.fn(async (_url:string,_init:RequestInit) => ({ok:true, json: async () => receipt('closed')})); vi.stubGlobal('fetch', fetch);
    const request = {closeVersion:4,workbookPathHash:hash('e:/original.xlsx'),worksheetNameHash:hash('Ledger'),copySheetName,plan,values:plan.actions,fullPath:'E:/SECRET/Original.xlsx'} as never;
    const api = await import('./api');
    await api.closeExcelReviewCopy(request);
    expect(JSON.parse(String(fetch.mock.calls[0][1].body)))
      .toEqual({closeVersion:4,workbookPathHash:hash('e:/original.xlsx'),worksheetNameHash:hash('Ledger'),copySheetName});
    expect(JSON.stringify(fetch.mock.calls)).not.toMatch(/LOCAL_PRIVATE_SENTINEL|PRIVATE_TITLE|SECRET|writeValues/);
  });
  it.each(CLOSE_STATUSES)('preserves the real close outcome %s and its copyClosed flag', async status => {
    for (const copyClosed of [true, false]) {
      vi.stubGlobal('fetch', vi.fn(async () => ({ok:true, json: async () => receipt(status, copyClosed)})));
      expect(await closeLocalExcelWindows('url:E:/Original.xlsx','Ledger',copySheetName)).toEqual({status, layoutVersion:4, copyClosed, recoveryCopyPath: 'E:/recovery/review-copy.xlsx'});
    }
  });
  it.each([
    ['unknown status', {status:'done', layoutVersion:4, copyClosed:true}],
    ['missing copyClosed', {status:'closed', layoutVersion:4}],
    ['non-boolean copyClosed', {status:'closed', layoutVersion:4, copyClosed:'yes'}],
    ['extra smuggled keys', {status:'closed', layoutVersion:4, copyClosed:true, written:true}],
    ['writeback claim leaking into the receipt', {status:'closed', layoutVersion:4, copyClosed:true, originalWritten:true}],
    ['legacy version', {status:'closed', layoutVersion:3, copyClosed:true}],
    ['empty body', {}]
  ])('refuses a malformed close receipt: %s', async (_label, payload) => {
    vi.stubGlobal('fetch', vi.fn(async () => ({ok:true, json: async () => payload})));
    await expect(closeLocalExcelWindows('url:E:/Original.xlsx','Ledger',copySheetName)).rejects.toThrow();
  });
  it('reports the legacy server as a version error rather than as a closed copy', async () => {
    vi.stubGlobal('fetch', vi.fn(async () => ({ok:true, json: async () => ({status:'closed', layoutVersion:2, copyClosed:true})})));
    await expect(closeLocalExcelWindows('url:E:/Original.xlsx','Ledger',copySheetName)).rejects.toBeInstanceOf(ExcelReviewLayoutVersionError);
  });
  it('does not send close metadata to a configured remote server', async () => {
    vi.stubEnv('VITE_API_BASE_URL','https://example.com'); vi.resetModules();
    const api = await import('./api'), fetch = vi.fn(); vi.stubGlobal('fetch',fetch);
    await expect(api.closeExcelReviewCopy({closeVersion:4,workbookPathHash:hash('e:/original.xlsx'),worksheetNameHash:hash('Ledger'),copySheetName}))
      .rejects.toThrow('仅能连接本机');
    expect(fetch).not.toHaveBeenCalled(); vi.resetModules();
  });
});
