import capabilities from '../../../config/capabilities.json';
import type { WpsApplication, WpsTaskPane, WpsWorkbook } from './wpsHost';
import type { WpsNativeComparison } from './wpsNativeComparison';
import { workbookFullName } from './wpsWorkbookReader';
import { TASK_PANE_STORAGE_KEY } from './wpsReviewPaneLease';

const positive = (n: unknown): n is number => typeof n === 'number' && Number.isFinite(n) && n > 0;
function metadata(book: WpsWorkbook) {
  return JSON.stringify([book.Name, workbookFullName(book), String(book.Path ?? '').trim()]);
}

/** Sizing reads only live identity/geometry. It must never activate or read ranges. */
export function createWpsReviewPaneAdapter(app: WpsApplication, review: WpsNativeComparison) {
  const storedId = app.PluginStorage.getItem(TASK_PANE_STORAGE_KEY);
  if (typeof storedId !== 'string' && typeof storedId !== 'number') throw new Error('No existing pane');
  const paneId = storedId;
  const hwnd = app.Hwnd;
  const pixelRatio = window.devicePixelRatio;
  if (!positive(hwnd)) throw new Error('No host identity');
  const originalIdentity = metadata(review.original);
  const copyIdentity = metadata(review.copy);
  if (!review.original.Path || originalIdentity === copyIdentity || review.copy.Path ||
      review.copy.Name !== review.copyIdentity.name || workbookFullName(review.copy) !== review.copyIdentity.fullName) {
    throw new Error('Review identities are unavailable');
  }
  function pane(): WpsTaskPane {
    if (app.Hwnd !== hwnd || String(app.PluginStorage.getItem(TASK_PANE_STORAGE_KEY)) !== String(paneId)) {
      throw new Error('Pane scope changed');
    }
    const live = app.GetTaskPane(paneId);
    if (!live || String(live.ID) !== String(paneId) || !live.Visible || live.DockPosition !== 2 || !positive(live.Width)) {
      throw new Error('Pane is not visible and docked right');
    }
    return live;
  }
  function measure() {
    const current = pane();
    if (window.devicePixelRatio !== pixelRatio) throw new Error('Display/WebView scale changed; recalibration required');
    const books = app.Workbooks;
    if (!books || !Number.isInteger(books.Count) || books.Count < 2 || books.Count > capabilities.wpsReviewPane.maxOpenWorkbooks) {
      throw new Error('Live workbooks unavailable');
    }
    let originals = 0, copies = 0;
    for (let index = 1; index <= books.Count; index++) {
      const identity = metadata(books.Item(index));
      originals += Number(identity === originalIdentity);
      copies += Number(identity === copyIdentity);
    }
    const active = app.ActiveWorkbook && metadata(app.ActiveWorkbook);
    if (originals !== 1 || copies !== 1 || (active !== originalIdentity && active !== copyIdentity)) {
      throw new Error('Live review changed or ambiguous');
    }
    if (!positive(app.Width)) throw new Error('Host width unavailable');
    // ET Application.Width is points (72/in), not iframe/screen pixels.
    const hostWidth = app.Width * 96 / 72;
    if (hostWidth > capabilities.reviewPaneSizing.maxHostWidthCssPx) throw new Error('Invalid host width');
    return { hostWidth, nativeWidth: current.Width! };
  }
  const initial = measure();
  const initialCssWidth = window.innerWidth;
  if (!positive(initialCssWidth)) throw new Error('No CSS width evidence');
  const nativeUnitsPerCssPx = initial.nativeWidth / initialCssWidth;
  return {
    paneId: String(paneId),
    checkPane: pane,
    measure,
    resize(cssWidth: number, actualCssWidth: number) {
      measure(); // recheck before mutating pane geometry
      if (!positive(actualCssWidth) || !positive(cssWidth)) throw new Error('No CSS width evidence');
      // TaskPane.Width's unit is not specified by WPS docs. Calibrate against
      // the initial native pane/viewport pair; never assume physical pixels.
      // Recalibrating while WebView resize is in flight creates a feedback drift.
      pane().Width = Math.round(cssWidth * nativeUnitsPerCssPx);
    },
    restore() { pane().Width = initial.nativeWidth; }
  };
}
