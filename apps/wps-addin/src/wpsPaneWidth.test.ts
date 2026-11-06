import { describe, expect, it, vi } from 'vitest';
import { widenWpsPane } from './wpsPaneWidth';
import type { WpsApplication, WpsTaskPane } from './wpsHost';

function fixture() {
  const pane: WpsTaskPane = { ID: 7, Visible: true, DockPosition: 2, Width: 600 };
  const items = new Map<string, unknown>([['excel_bro_taskpane_id', 7]]);
  const app = { Width: 1200, GetTaskPane: vi.fn(() => pane), PluginStorage: {
    getItem: (key: string) => items.get(key), setItem: (key: string, value: unknown) => items.set(key, value)
  }} as unknown as WpsApplication;
  return { app, pane, items };
}
describe('user pane width', () => {
  it('uses the existing right-docked pane and preserves scale', () => {
    const {app, pane} = fixture();
    expect(widenWpsPane(app, 400)).toContain('请求加宽');
    expect(pane.Width).toBe(840);
    expect(app.GetTaskPane).toHaveBeenCalledWith(7);
  });
  it('does not alter a pane under an active review lease', () => {
    const {app, pane, items} = fixture();
    items.set('excel_bro_review_pane_width_v1', JSON.stringify({paneId:'7',owner:'review',expiresAt:Date.now()+60000}));
    expect(widenWpsPane(app, 400)).toContain('先退出双视图');
    expect(pane.Width).toBe(600);
  });
  it('rejects missing or mismatched pane and clamps the width', () => {
    const {app, pane} = fixture(); pane.ID = 8;
    expect(widenWpsPane(app, 400)).toContain('无法核实');
    expect(pane.Width).toBe(600);
    pane.ID = 7;
    expect(widenWpsPane(app, 1200)).toContain('上限');
    expect(pane.Width).toBe(600);
  });
});
