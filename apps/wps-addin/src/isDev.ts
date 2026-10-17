/**
 * 是否处于开发模式（DEV）。诊断路由只在 DEV 下可达。
 *
 * 独立成模块是为了让"非 DEV 门控"可测试：测试可 mock 本模块的 `isDev`，
 * 而 `AppRoutes` 从本模块 import，mock 能真正影响它内部的调用。
 */
export function isDev(): boolean {
  return import.meta.env.DEV === true;
}
