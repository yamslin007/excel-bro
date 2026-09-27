// 宿主无关共享层。分层顺序：contracts → pure → semantics → ports → orchestration → ui，
// 依赖只能从上往下，不允许反向边。
//
// L0 contracts
export type { CellValue } from "./contracts";

// L1 pure
export type { A1Rect } from "./planCore";
export {
  parseExplicitA1Rect,
  assertRectangularValues,
  writeMatrixShapeIssue,
  blankCellMatrix
} from "./planCore";
export {
  firstCellOfRange,
  columnLetterFromIndex,
  columnIndexFromLetter,
  startColumnOfRange,
  suggestWriteTarget
} from "./range";
export { normalizeField, detectHeaderIndex, detectSheetFields } from "./tableSchema";
