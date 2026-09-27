/**
 * L0 契约层：纯类型,零运行时依赖。
 *
 * 阶段 2 只迁入 L1 模块真正需要的类型种子。`apps/excel-addin/src/contracts.ts`
 * 的其余 1900 余行按方案 §5.2 拆分（纯 DTO 留 L0、运行时校验归 L2）后再迁入,
 * 本文件届时逐步扩充。
 */

/** 单元格标量值。两端读写与比对的最小公共类型。 */
export type CellValue = string | number | boolean | null;
