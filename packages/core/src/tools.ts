/**
 * L0 工具 DTO。原在 `apps/excel-addin/src/storage.ts:17-123`,
 * 除 `SavedTool.planTemplate` 泛型化（见下）外定义逐字未变。
 *
 * 下沉原因：`format.ts`（L1）为取 `SavedTool` 而 import `storage.ts`（L2）,
 * 是 L1→L2 反向边。类型本身是纯 DTO,不该住在带 1000 行运行时逻辑的文件里。
 */

interface ParameterBase {
  id: string;
  label: string;
  defaultValue: string;
  required: true;
}

export interface WorksheetToolParameter extends ParameterBase {
  type: "worksheet";
  bindings: Array<{
    actionIndex: number;
    property: "sheet" | "sourceSheet";
  }>;
}

export interface OutputWorksheetToolParameter extends ParameterBase {
  type: "outputWorksheet";
  bindings: Array<{
    actionIndex: number;
    property: "sheet" | "sourceSheet";
  }>;
}

export interface FieldToolParameter extends ParameterBase {
  type: "field";
  sourceParameterId: string;
  bindings: Array<
    | {
        actionIndex: number;
        property: "splitBy";
      }
    | {
        actionIndex: number;
        property: "groupBy";
        itemIndex: number;
      }
    | {
        actionIndex: number;
        property: "metricField";
        itemIndex: number;
      }
    | {
        actionIndex: number;
        property: "pivotRowField" | "pivotColumnField";
        itemIndex: number;
      }
    | {
        actionIndex: number;
        property: "pivotValueField";
        itemIndex: number;
      }
  >;
}

export interface RangeToolParameter extends ParameterBase {
  type: "range";
  sourceParameterId?: string;
  bindings: Array<
    | {
        actionIndex: number;
        property: "range";
      }
    | {
        actionIndex: number;
        property: "sourceRange";
      }
  >;
}

export type ToolParameter =
  | WorksheetToolParameter
  | OutputWorksheetToolParameter
  | FieldToolParameter
  | RangeToolParameter;

export type ToolRiskApproval = "fixedContent" | "destructive";

export interface ToolEligibilityIssue {
  code:
    | "FIXED_CONTENT"
    | "DESTRUCTIVE_ACTION"
    | "EMBEDDED_IMAGE"
    | "SNAPSHOT_VERIFICATION";
  severity: "approval" | "blocked" | "warning";
  message: string;
  actionIndexes: number[];
  approval?: ToolRiskApproval;
}

export interface ToolEligibility {
  issues: ToolEligibilityIssue[];
  requiredApprovals: ToolRiskApproval[];
  blocked: boolean;
}

/**
 * `planTemplate` 泛型化,而不是让 L0 直接依赖 `AnalysisPlan`。
 *
 * `AnalysisPlan` 挂在 `ExcelAction` 这个 38 成员联合上,连带 1000 余行,
 * 本阶段不动（方案 §5.4 明确不从 contracts 开刀）。app 侧绑定具体类型后
 * 结构完全一致,使用点零改动。`AnalysisPlan` 迁入 L0 时可收紧此默认值。
 */
export interface SavedTool<TPlanTemplate = unknown> {
  id: string;
  version: 2;
  name: string;
  description: string;
  createdAt: string;
  updatedAt: string;
  verifiedAt: string;
  planTemplate: TPlanTemplate;
  parameters: ToolParameter[];
  approvals: ToolRiskApproval[];
}
