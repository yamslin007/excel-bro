/**
 * Office 通道**实际接通执行**的动作清单。
 *
 * 用途单一：给规划前的能力收窄（方案 §4.3）提供 `selfReportedActions`。
 * 它是"本端能执行哪些动作"的自述，**不是授权**——服务端按可信配置算上界后
 * 与它求交，多报会被上界砍掉，所以这份清单写错不会放开任何越权动作，
 * 只会让本端计划多出执行阶段才失败的步骤。因此它的正确性标准是
 * "与 `excel.ts` 的 `executeAction` 保持一致"，而不是"尽量多"。
 *
 * 维护约定：新增动作先落 `executeAction` 的 switch 分支，再登记到这里，
 * 并在 `officeActions.test.ts` 里对齐 —— 该测试从 `contracts.ts` 抓取
 * `ExcelAction` 全体成员，双向比对，任一侧漏登记立刻变红。
 */
export const OFFICE_EXECUTED_ACTION_TYPES = [
  "createWorksheet",
  "deleteWorksheet",
  "writeTable",
  "writeValues",
  "writeFormulas",
  "computeValues",
  "clearRange",
  "insertRange",
  "deleteRange",
  "copyRange",
  "sortRange",
  "removeDuplicates",
  "filterRange",
  "clearFilter",
  "splitGroupAggregate",
  "setFill",
  "setFont",
  "setNumberFormat",
  "setAlignment",
  "setBorders",
  "setDataValidation",
  "setConditionalFormat",
  "resizeRange",
  "mergeCells",
  "unmergeCells",
  "freezePanes",
  "setHyperlink",
  "addComment",
  "addNote",
  "createTable",
  "createChart",
  "createPivotTable",
  "addNamedRange",
  "addImage",
  "addShape",
  "autofit",
  "activateWorksheet"
] as const;

export type OfficeExecutedActionType =
  (typeof OFFICE_EXECUTED_ACTION_TYPES)[number];
