import { useRef, useState } from 'react';

/** Match ordered characters, ignoring case, full-width variants and spacing. */
export function matchesReviewSheetName(name: string, query: string): boolean {
  const normalize = (value: string) => Array.from(value.normalize('NFKC').toLocaleLowerCase().replace(/\s+/gu, ''));
  const target = normalize(name);
  const search = normalize(query);
  let index = 0;
  for (const character of target) {
    if (character === search[index]) index += 1;
    if (index === search.length) return true;
  }
  return search.length === 0;
}

/** Selection only requests an existing authorized comparison; it never accepts a review. */
export function ReviewSheetPicker({ sheets, selected, pending, onSelect }: {
  sheets: string[]; selected?: string; pending: boolean; onSelect(sheet: string): void;
}) {
  const [query, setQuery] = useState('');
  const [filter, setFilter] = useState('');
  const [page, setPage] = useState(() => Math.floor(Math.max(0, sheets.indexOf(selected ?? '')) / 3));
  const composing = useRef(false);
  const searching = filter.trim().length > 0;
  const normalized = (value: string) => value.normalize('NFKC').toLocaleLowerCase().replace(/\s+/gu, '');
  const filtered = sheets.filter(sheet => matchesReviewSheetName(sheet, filter));
  // Preserve workbook order within each group; contiguous matches are easier to scan first.
  if (searching) filtered.sort((a, b) => Number(normalized(b).includes(normalized(filter))) - Number(normalized(a).includes(normalized(filter))));
  const pages = Math.ceil(filtered.length / 3);
  const currentPage = Math.min(page, Math.max(0, pages - 1));
  const applyQuery = (value: string) => { setFilter(value); setPage(0); };
  const clear = () => { setQuery(''); setFilter(''); setPage(Math.floor(Math.max(0, sheets.indexOf(selected ?? '')) / 3)); };
  return <div className="review-sheet-picker">
    {sheets.length > 3 && <div className="review-sheet-search">
      <span className="review-sheet-search-icon" aria-hidden="true">⌕</span>
      <input aria-label="搜索对比工作表" placeholder="搜索工作表名称…" value={query} disabled={pending}
        onCompositionStart={() => { composing.current = true; }}
        onCompositionEnd={event => { composing.current = false; applyQuery(event.currentTarget.value); }}
        onChange={event => { setQuery(event.target.value); if (!composing.current) applyQuery(event.target.value); }} />
      {query && <button type="button" className="review-sheet-search-clear" aria-label="清空工作表搜索" disabled={pending} onClick={clear}>×</button>}
    </div>}
    <div className={`review-sheet-page${!searching && pages > 1 ? ' has-pages' : ''}`}>
      {!searching && pages > 1 && <button type="button" className="secondary-button" disabled={pending || currentPage === 0} onClick={() => setPage(currentPage - 1)}>‹ 上一页</button>}
      <div className={`review-sheet-list${searching ? ' is-searching' : ''}`} aria-label="对比工作表列表" tabIndex={searching ? 0 : undefined}>
        {(searching ? filtered : filtered.slice(currentPage * 3, (currentPage + 1) * 3)).map(sheet => <button type="button" className="secondary-button" key={sheet} aria-label={sheet} aria-pressed={sheet === selected} disabled={pending} onClick={() => { if (!pending) onSelect(sheet); }}><span className="review-sheet-check" aria-hidden="true">{sheet === selected ? '✓' : ''}</span><span>{sheet}</span>{sheet === selected && <span className="review-sheet-selected-label">正在对比</span>}</button>)}
        {!filtered.length && <span role="status">未找到匹配的工作表</span>}
      </div>
      {!searching && pages > 1 && <button type="button" className="secondary-button" disabled={pending || currentPage === pages - 1} onClick={() => setPage(currentPage + 1)}>下一页 ›</button>}
    </div>
    {(pages > 1 || searching) && <div className="review-sheet-page-count" aria-live="polite">{searching ? `找到 ${filtered.length} 张` : `第 ${currentPage + 1} / ${pages} 页`}</div>}
  </div>;
}
