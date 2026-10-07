import { expect, it } from 'vitest';
import { ReviewPaneSizing, reviewPaneTargetWidth } from './reviewPaneSizing';
it('uses a shared forty-percent target, readable minimum, and sheet-space bound', () => {
  expect([1200, 700, 500, 501].map(reviewPaneTargetWidth)).toEqual([480, 320, 250, 250]);
  for (const width of [0, -1, NaN, Infinity, 32769]) expect(() => reviewPaneTargetWidth(width)).toThrow();
});
it('requires actual width evidence, not just a successful setter', () => {
  const sizing = new ReviewPaneSizing(800);
  expect(sizing.observe(1200, 800)).toMatchObject({requestWidth:480, needsResize:true});
  expect(sizing.observe(1200, 800)).toMatchObject({requestWidth:undefined, needsResize:true});
  expect(sizing.observe(1200, 480)).toMatchObject({requestWidth:undefined, needsResize:false, manual:false});
});
it('follows a host resize but never fights a pane-only manual resize', () => {
  const sizing = new ReviewPaneSizing(800);
  sizing.observe(1200,800); sizing.observe(1200,480);
  expect(sizing.observe(1500,480).requestWidth).toBe(600);
  sizing.observe(1500,600);
  expect(sizing.observe(1500,430).manual).toBe(true);
  expect(sizing.observe(1800,430)).toMatchObject({manual:true,requestWidth:undefined,needsResize:false});
  expect(new ReviewPaneSizing(430).observe(1800,430).requestWidth).toBe(720);
});
