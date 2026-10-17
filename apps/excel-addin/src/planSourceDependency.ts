/**
 * 计划的**实际来源依赖**（Excel 侧入口）。
 *
 * 实现已上移到 `@excel-bro/core/planSourceDependency`：Excel 与 WPS 各有一份
 * 桥宿主适配器，但"读了谁"的记账口径必须**逐字同源** —— 同一份计划在两端
 * 一个可恢复、一个不可恢复是不可接受的。这里只做再导出，保留既有 import
 * 路径（App / 端到端用例）不被打断。
 *
 * 详细口径见 core 里的模块注释；断言纪律见 `planSourceDependency.test.ts`
 * 与 core 的同名用例。
 */

export {
  recordTrustedToolReads,
  trustedToolReads,
  resetTrustedSourceReads,
  attestSourceDependency,
  attestResultSourceDependency,
  type SourceDependencyOrigin,
  type SourceDependencyRecord
} from "@excel-bro/core/planSourceDependency";
