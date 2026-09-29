// 宿主无关共享层。分层顺序：contracts → pure → semantics → ports → orchestration → ui，
// 依赖只能从上往下，不允许反向边。
//
// L0 contracts
export type { CellValue, WorkbookLocator, AuthorizedSheetRef } from "./contracts";
export type {
  WorksheetToolParameter,
  OutputWorksheetToolParameter,
  FieldToolParameter,
  RangeToolParameter,
  ToolParameter,
  ToolRiskApproval,
  ToolEligibilityIssue,
  ToolEligibility,
  SavedTool
} from "./tools";

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
// 写进任何用户可见 / 可持久化文本之前，先过一遍凭据脱敏。
export {
  REDACTED_PLACEHOLDER,
  redactCredentials,
  containsLikelyCredential
} from "./redact";

// L2 semantics：身份折叠是唯一允许拼接/解析 `url:` 前缀的地方。
export {
  officeIdentityFromUrl,
  officeIdentityFromLocator,
  locatorFromOfficeIdentity,
  sameLocator
} from "./locator";
// 与 server/app/column_mapping.py 镜像，一致性由 config/column_mapping_cases.json 锁住。
export type {
  ColumnMappingInput,
  ColumnMappingAction
} from "./columnMapping";
export {
  ColumnMappingError,
  UnmappedCategoryError,
  ShapeMismatchError,
  OverlappingRangeError,
  mapColumnValues,
  buildColumnMappingActions
} from "./columnMapping";

// L3 ports：只有接口，没有实现。
export type {
  HostId,
  KeyValueStore,
  ActionCapability,
  CapabilityResolver,
  HostPrimitiveOps,
  HostAdapter,
  ApiClient,
  Clock
} from "./ports";

// L2 semantics：宿主 × 版本 × 动作 的可信上界与交集（方案 §4.2/§4.3）。
// 与 server/app/action_capability.py 镜像，一致性由 config/action_capability_cases.json 锁住。
export type {
  HostActionTable,
  HostCapabilityMatrix,
  CapabilityMatrix
} from "./actionCapability";
export {
  matchesVersionPrefix,
  parseHostId,
  resolveHostActionBound,
  intersectActionCapabilities,
  resolveAllowedActionTypes,
  actionCapabilityKeys,
  isActionAllowedByKeys
} from "./actionCapability";
