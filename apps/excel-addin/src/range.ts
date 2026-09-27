// 阶段 2 兼容 shim：实现已迁入 packages/core/src/range.ts。
// 保留此再导出，使现有导入点零改动；阶段 6 删除本文件并改导入为 @excel-bro/core。
export {
  firstCellOfRange,
  columnLetterFromIndex,
  columnIndexFromLetter,
  startColumnOfRange,
  suggestWriteTarget
} from "@excel-bro/core/range";
