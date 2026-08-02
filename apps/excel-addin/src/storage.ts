import {
  assertAssistantResponse,
  type AnalysisPlan,
  type DataToolRequest,
  type DeterministicQueryTemplate,
  type ExcelAction,
  type WorkbookSnapshot
} from "./contracts";
import capabilities from "../../../config/capabilities.json";

const STORAGE_KEY = "excel-bro.tools.v2";
const LEGACY_TOOL_STORAGE_KEY = "excel-bro.tools.v1";
const LEGACY_AUTOMATION_STORAGE_KEY = "excel-bro.automations.v1";
const TOOL_LIMIT = capabilities.savedTools.maxItems;
const QUERY_TOOL_STORAGE_KEY = "excel-bro.query-tools.v1";

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

export interface SavedTool {
  id: string;
  version: 2;
  name: string;
  description: string;
  createdAt: string;
  updatedAt: string;
  verifiedAt: string;
  planTemplate: AnalysisPlan;
  parameters: ToolParameter[];
  approvals: ToolRiskApproval[];
}

export interface SavedQueryTool extends DeterministicQueryTemplate {
  version: 1;
  createdAt: string;
  updatedAt: string;
}

export interface QueryToolCompatibility {
  runnable: boolean;
  requiresModel: boolean;
  reasons: string[];
}

export function createQueryTool(
  name: string,
  description: string,
  request: DataToolRequest,
  sourceMode: "workbook" | "folder",
  sourceSheetNames: string[],
  sourceSheetIds: string[] = [],
  expectedHeaders: string[] = []
): SavedQueryTool {
  const now = new Date().toISOString();
  return {
    id: identifier(),
    version: 1,
    name: name.trim() || "本地查询",
    description: description.trim() || "重复运行确定性本地查询",
    sourceMode,
    request: structuredClone(request),
    sourceSheetNames: [...sourceSheetNames],
    sourceSheetIds: [...sourceSheetIds],
    expectedHeaders: [...expectedHeaders],
    createdAt: now,
    updatedAt: now
  };
}

export function loadQueryTools(): SavedQueryTool[] {
  return readArray(QUERY_TOOL_STORAGE_KEY).filter(
    (value): value is SavedQueryTool =>
      Boolean(
        value &&
          typeof value === "object" &&
          (value as Partial<SavedQueryTool>).version === 1 &&
          (value as Partial<SavedQueryTool>).request
      )
  );
}

export function saveQueryTool(tool: SavedQueryTool): SavedQueryTool[] {
  const next = [
    { ...tool, updatedAt: new Date().toISOString() },
    ...loadQueryTools().filter((item) => item.id !== tool.id)
  ].slice(0, TOOL_LIMIT);
  localStorage.setItem(QUERY_TOOL_STORAGE_KEY, JSON.stringify(next));
  return next;
}

export function deleteQueryTool(toolId: string): SavedQueryTool[] {
  const next = loadQueryTools().filter((tool) => tool.id !== toolId);
  localStorage.setItem(QUERY_TOOL_STORAGE_KEY, JSON.stringify(next));
  return next;
}

export function analyzeQueryToolCompatibility(
  tool: SavedQueryTool,
  workbook: WorkbookSnapshot
): QueryToolCompatibility {
  const reasons: string[] = [];
  const availableSheets = new Map(
    workbook.worksheets.map((sheet) => [normalize(sheet.name), sheet])
  );
  for (const name of tool.sourceSheetNames) {
    const sheet = availableSheets.get(normalize(name));
    if (!sheet) {
      reasons.push(`来源工作表「${name}」不存在`);
      continue;
    }
    const availableFields = new Set(
      sheet.headers.map((header) => normalize(String(header ?? "")))
    );
    const requestedFields = [
      ...(tool.request.arguments.fields ?? []),
      ...(tool.request.arguments.groupBy ?? []),
      ...(tool.request.arguments.filters ?? []).map((item) => item.field),
      ...(tool.request.arguments.metrics ?? [])
        .map((item) => item.field)
        .filter((field): field is string => Boolean(field)),
      ...(tool.request.arguments.profileField
        ? [tool.request.arguments.profileField]
        : [])
    ];
    const missing = requestedFields.filter(
      (field) => !availableFields.has(normalize(field))
    );
    if (missing.length > 0) {
      reasons.push(`「${name}」缺少字段：${[...new Set(missing)].join("、")}`);
    }
  }
  if (tool.sourceMode === "folder" && tool.sourceSheetIds.length === 0) {
    reasons.push("文件夹查询缺少稳定工作表 ID");
  } else if (tool.sourceMode === "folder") {
    const currentIds = tool.sourceSheetNames
      .map((name) => availableSheets.get(normalize(name))?.sourceSheetId)
      .filter((value): value is string => Boolean(value))
      .sort();
    const savedIds = [...tool.sourceSheetIds].sort();
    if (JSON.stringify(currentIds) !== JSON.stringify(savedIds)) {
      reasons.push("文件夹来源 ID 已变化，请重新确认数据来源");
    }
  }
  return {
    runnable: reasons.length === 0,
    requiresModel: reasons.length > 0,
    reasons
  };
}

export function instantiateQueryTool(tool: SavedQueryTool): DataToolRequest {
  return {
    ...structuredClone(tool.request),
    id: `saved-query-${tool.id}-${Date.now()}`
  };
}

interface LegacySavedTool {
  id: string;
  version: 1;
  name: string;
  description: string;
  createdAt: string;
  updatedAt: string;
  verifiedAt: string;
  plan: AnalysisPlan;
  parameters: Array<{
    id: string;
    label: string;
    type: "worksheet";
    defaultValue: string;
    required: true;
  }>;
}

function identifier(): string {
  return (
    globalThis.crypto?.randomUUID?.() ??
    `${Date.now()}-${Math.random().toString(16).slice(2)}`
  );
}

function normalize(value: string): string {
  return value.trim().toLocaleLowerCase();
}

function actionSheetReferences(
  action: ExcelAction
): Array<"sheet" | "sourceSheet"> {
  return action.type === "copyRange" || action.type === "createPivotTable"
    ? ["sheet", "sourceSheet"]
    : ["sheet"];
}

function actionSheetValue(
  action: ExcelAction,
  property: "sheet" | "sourceSheet"
): string {
  if (property === "sheet") return action.sheet;
  return action.type === "copyRange" || action.type === "createPivotTable"
    ? action.sourceSheet
    : "";
}

function referencedSourceSheets(
  plan: AnalysisPlan,
  sourceSheetNames: string[]
): string[] {
  const referenced = new Set(
    plan.actions
      .flatMap((action) =>
        actionSheetReferences(action).map((property) =>
          actionSheetValue(action, property)
        )
      )
      .map(normalize)
  );
  return sourceSheetNames.filter(
    (name, index) =>
      referenced.has(normalize(name)) &&
      sourceSheetNames.findIndex(
        (candidate) => normalize(candidate) === normalize(name)
      ) === index
  );
}

export function analyzeToolEligibility(
  plan: AnalysisPlan
): ToolEligibility {
  const fixedContent = plan.actions.flatMap((action, index) =>
    action.type === "writeValues" || action.type === "writeTable"
      ? [index]
      : []
  );
  const destructive = plan.actions.flatMap((action, index) => {
    const destructiveAction =
      action.type === "deleteWorksheet" ||
      action.type === "deleteRange" ||
      (action.type === "clearRange" &&
        (action.applyTo === "all" || action.applyTo === "contents")) ||
      (action.type === "splitGroupAggregate" &&
        action.existingSheetPolicy === "replace");
    return destructiveAction ? [index] : [];
  });
  const embeddedImages = plan.actions.flatMap((action, index) =>
    action.type === "addImage" ? [index] : []
  );
  const snapshotVerification = (plan.acceptanceCriteria ?? []).flatMap(
    (criterion, index) =>
      criterion.type === "rangeEquals" ? [index] : []
  );
  const issues: ToolEligibilityIssue[] = [];
  if (fixedContent.length > 0) {
    issues.push({
      code: "FIXED_CONTENT",
      severity: "approval",
      approval: "fixedContent",
      actionIndexes: fixedContent,
      message:
        "计划包含固定写入值。再次运行会写入保存时的内容，不会重新计算。"
    });
  }
  if (destructive.length > 0) {
    issues.push({
      code: "DESTRUCTIVE_ACTION",
      severity: "approval",
      approval: "destructive",
      actionIndexes: destructive,
      message:
        "计划包含删除、清空、覆盖或替换操作，每次运行前都必须再次预览。"
    });
  }
  if (embeddedImages.length > 0) {
    issues.push({
      code: "EMBEDDED_IMAGE",
      severity: "blocked",
      actionIndexes: embeddedImages,
      message:
        "计划包含内嵌图片数据，不能保存为工具；请改为运行时选择图片。"
    });
  }
  if (snapshotVerification.length > 0) {
    issues.push({
      code: "SNAPSHOT_VERIFICATION",
      severity: "warning",
      actionIndexes: snapshotVerification,
      message:
        "部分验收条件包含固定结果，运行时会按当前参数重新绑定工作表，但不会重新推导旧数值。"
    });
  }
  return {
    issues,
    requiredApprovals: issues.flatMap((issue) =>
      issue.approval ? [issue.approval] : []
    ),
    blocked: issues.some((issue) => issue.severity === "blocked")
  };
}

function buildWorksheetParameters(
  plan: AnalysisPlan,
  sourceSheetNames: string[]
): WorksheetToolParameter[] {
  return referencedSourceSheets(plan, sourceSheetNames).map(
    (sheet, index) => ({
      id: `worksheet_${index + 1}`,
      label:
        sourceSheetNames.length === 1
          ? "来源工作表"
          : `来源工作表：${sheet}`,
      type: "worksheet",
      defaultValue: sheet,
      required: true,
      bindings: plan.actions.flatMap((action, actionIndex) =>
        actionSheetReferences(action).flatMap((property) =>
          normalize(actionSheetValue(action, property)) === normalize(sheet)
            ? [{ actionIndex, property }]
            : []
        )
      )
    })
  );
}

function buildOutputWorksheetParameters(
  plan: AnalysisPlan,
  worksheets: WorksheetToolParameter[]
): OutputWorksheetToolParameter[] {
  const sourceNames = new Set(
    worksheets.map((parameter) => normalize(parameter.defaultValue))
  );
  const outputNames = plan.actions
    .flatMap((action) =>
      action.type === "createWorksheet" ? [action.sheet] : []
    )
    .filter(
      (name, index, all) =>
        !sourceNames.has(normalize(name)) &&
        all.findIndex(
          (candidate) => normalize(candidate) === normalize(name)
        ) === index
    );
  return outputNames.map((sheet, index) => ({
    id: `output_worksheet_${index + 1}`,
    label:
      outputNames.length === 1
        ? "输出工作表名称"
        : `输出工作表名称：${sheet}`,
    type: "outputWorksheet",
    defaultValue: sheet,
    required: true,
    bindings: plan.actions.flatMap((action, actionIndex) =>
      actionSheetReferences(action).flatMap((property) =>
        normalize(actionSheetValue(action, property)) === normalize(sheet)
          ? [{ actionIndex, property }]
          : []
      )
    )
  }));
}

function buildRangeParameters(
  plan: AnalysisPlan,
  worksheets: WorksheetToolParameter[]
): RangeToolParameter[] {
  const parameters = new Map<string, RangeToolParameter>();
  const worksheetForAction = (
    actionIndex: number,
    property: "sheet" | "sourceSheet"
  ) =>
    worksheets.find((parameter) =>
      parameter.bindings.some(
        (binding) =>
          binding.actionIndex === actionIndex &&
          binding.property === property
      )
    );
  const add = (
    value: string | null | undefined,
    actionIndex: number,
    property: "range" | "sourceRange",
    worksheetProperty: "sheet" | "sourceSheet"
  ) => {
    if (!value?.trim()) return;
    const source = worksheetForAction(actionIndex, worksheetProperty);
    if (!source) return;
    const key = `${source.id}:${property}:${normalize(value)}`;
    const current = parameters.get(key);
    if (current) {
      current.bindings.push({ actionIndex, property });
      return;
    }
    parameters.set(key, {
      id: `range_${parameters.size + 1}`,
      label: `数据范围：${value}`,
      type: "range",
      defaultValue: value,
      required: true,
      sourceParameterId: source.id,
      bindings: [{ actionIndex, property }]
    });
  };

  plan.actions.forEach((action, actionIndex) => {
    if (action.type === "sortRange" || action.type === "filterRange") {
      add(action.range, actionIndex, "range", "sheet");
    }
    if (action.type === "copyRange") {
      add(action.sourceRange, actionIndex, "sourceRange", "sourceSheet");
    }
    if (
      action.type === "createChart" ||
      action.type === "splitGroupAggregate"
    ) {
      add(action.sourceRange, actionIndex, "sourceRange", "sheet");
    }
    if (action.type === "createPivotTable") {
      add(action.sourceRange, actionIndex, "sourceRange", "sourceSheet");
    }
  });
  return [...parameters.values()];
}

function buildFieldParameters(
  plan: AnalysisPlan,
  worksheets: WorksheetToolParameter[]
): FieldToolParameter[] {
  const parameters = new Map<string, FieldToolParameter>();
  const worksheetForAction = (actionIndex: number, source = false) =>
    worksheets.find((parameter) =>
      parameter.bindings.some(
        (binding) =>
          binding.actionIndex === actionIndex &&
          binding.property === (source ? "sourceSheet" : "sheet")
      )
    );
  const add = (
    field: string,
    sourceParameterId: string,
    binding: FieldToolParameter["bindings"][number]
  ) => {
    const key = `${sourceParameterId}:${normalize(field)}`;
    const current = parameters.get(key);
    if (current) {
      current.bindings.push(binding);
      return;
    }
    parameters.set(key, {
      id: `field_${parameters.size + 1}`,
      label: `字段：${field}`,
      type: "field",
      defaultValue: field,
      required: true,
      sourceParameterId,
      bindings: [binding]
    });
  };

  plan.actions.forEach((action, actionIndex) => {
    if (action.type === "splitGroupAggregate") {
      const source = worksheetForAction(actionIndex);
      if (!source) return;
      add(action.splitBy, source.id, {
        actionIndex,
        property: "splitBy"
      });
      action.groupBy.forEach((field, itemIndex) =>
        add(field, source.id, {
          actionIndex,
          property: "groupBy",
          itemIndex
        })
      );
      action.metrics.forEach((metric, itemIndex) => {
        if (metric.field) {
          add(metric.field, source.id, {
            actionIndex,
            property: "metricField",
            itemIndex
          });
        }
      });
    }
    if (action.type === "createPivotTable") {
      const source = worksheetForAction(actionIndex, true);
      if (!source) return;
      action.rowFields.forEach((field, itemIndex) =>
        add(field, source.id, {
          actionIndex,
          property: "pivotRowField",
          itemIndex
        })
      );
      action.columnFields.forEach((field, itemIndex) =>
        add(field, source.id, {
          actionIndex,
          property: "pivotColumnField",
          itemIndex
        })
      );
      action.valueFields.forEach((value, itemIndex) =>
        add(value.field, source.id, {
          actionIndex,
          property: "pivotValueField",
          itemIndex
        })
      );
    }
  });
  return [...parameters.values()];
}

function validTool(value: unknown): value is SavedTool {
  if (!value || typeof value !== "object") return false;
  const tool = value as Partial<SavedTool>;
  const validShape =
    typeof tool.id === "string" &&
    typeof tool.name === "string" &&
    typeof tool.description === "string" &&
    typeof tool.verifiedAt === "string" &&
    tool.version === 2 &&
    Boolean(tool.planTemplate) &&
    Array.isArray(tool.parameters) &&
    tool.parameters.every(
      (parameter) =>
        parameter &&
        typeof parameter.id === "string" &&
        typeof parameter.label === "string" &&
        typeof parameter.defaultValue === "string" &&
        (parameter.type === "worksheet" ||
          parameter.type === "outputWorksheet" ||
          parameter.type === "field" ||
          parameter.type === "range") &&
        Array.isArray(parameter.bindings)
    ) &&
    Array.isArray(tool.approvals);
  if (!validShape) return false;
  try {
    assertAssistantResponse({
      kind: "plan",
      provider: "local",
      plan: tool.planTemplate
    });
    return true;
  } catch {
    return false;
  }
}

function migrateLegacyTool(tool: LegacySavedTool): SavedTool {
  const sourceSheets = tool.parameters.map(
    (parameter) => parameter.defaultValue
  );
  const worksheets = buildWorksheetParameters(tool.plan, sourceSheets);
  const outputs = buildOutputWorksheetParameters(tool.plan, worksheets);
  return {
    id: tool.id,
    version: 2,
    name: tool.name,
    description: tool.description,
    createdAt: tool.createdAt,
    updatedAt: new Date().toISOString(),
    verifiedAt: tool.verifiedAt,
    planTemplate: structuredClone(tool.plan),
    parameters: [
      ...worksheets,
      ...outputs,
      ...buildFieldParameters(tool.plan, worksheets),
      ...buildRangeParameters(tool.plan, worksheets)
    ],
    approvals: analyzeToolEligibility(tool.plan).requiredApprovals
  };
}

function refreshToolParameters(tool: SavedTool): SavedTool {
  const sourceSheets = tool.parameters
    .filter(
      (parameter): parameter is WorksheetToolParameter =>
        parameter.type === "worksheet"
    )
    .map((parameter) => parameter.defaultValue);
  const worksheets = buildWorksheetParameters(
    tool.planTemplate,
    sourceSheets
  );
  return {
    ...tool,
    parameters: [
      ...worksheets,
      ...buildOutputWorksheetParameters(tool.planTemplate, worksheets),
      ...buildFieldParameters(tool.planTemplate, worksheets),
      ...buildRangeParameters(tool.planTemplate, worksheets)
    ]
  };
}

function readArray(key: string): unknown[] {
  try {
    const value: unknown = JSON.parse(localStorage.getItem(key) ?? "[]");
    return Array.isArray(value) ? value : [];
  } catch {
    return [];
  }
}

export function loadTools(): SavedTool[] {
  if (localStorage.getItem(STORAGE_KEY) !== null) {
    const tools = readArray(STORAGE_KEY)
      .filter(validTool)
      .map(refreshToolParameters);
    persistTools(tools);
    return tools;
  }

  const legacyTools = readArray(LEGACY_TOOL_STORAGE_KEY)
    .filter(
      (value): value is LegacySavedTool =>
        Boolean(
          value &&
            typeof value === "object" &&
            (value as Partial<LegacySavedTool>).version === 1 &&
            (value as Partial<LegacySavedTool>).plan
        )
    )
    .map(migrateLegacyTool);
  const legacyPlans = readArray(LEGACY_AUTOMATION_STORAGE_KEY)
    .filter(
      (value): value is AnalysisPlan =>
        Boolean(value && typeof value === "object")
    )
    .flatMap((plan) => {
      try {
        return [
          createTool(plan, plan.title, plan.summary, [], {
            fixedContent: true,
            destructive: true
          })
        ];
      } catch {
        return [];
      }
    });
  const migrated = [...legacyTools, ...legacyPlans];
  persistTools(migrated);
  return migrated;
}

function persistTools(tools: SavedTool[]): SavedTool[] {
  const next = tools.slice(0, TOOL_LIMIT);
  localStorage.setItem(STORAGE_KEY, JSON.stringify(next));
  return next;
}

export function createTool(
  plan: AnalysisPlan,
  name: string,
  description: string,
  sourceSheetNames: string[],
  approvals: Partial<Record<ToolRiskApproval, boolean>> = {}
): SavedTool {
  const eligibility = analyzeToolEligibility(plan);
  if (eligibility.blocked) {
    throw new Error(
      eligibility.issues
        .filter((issue) => issue.severity === "blocked")
        .map((issue) => issue.message)
        .join("；")
    );
  }
  const missingApproval = eligibility.requiredApprovals.find(
    (approval) => !approvals[approval]
  );
  if (missingApproval) {
    throw new Error(
      missingApproval === "fixedContent"
        ? "请确认此工具会重复写入保存时的固定内容"
        : "请确认此工具包含破坏性或覆盖操作"
    );
  }
  const now = new Date().toISOString();
  const worksheets = buildWorksheetParameters(plan, sourceSheetNames);
  const outputs = buildOutputWorksheetParameters(plan, worksheets);
  return {
    id: identifier(),
    version: 2,
    name: name.trim() || plan.title,
    description: description.trim() || plan.summary,
    createdAt: now,
    updatedAt: now,
    verifiedAt: now,
    planTemplate: structuredClone(plan),
    parameters: [
      ...worksheets,
      ...outputs,
      ...buildFieldParameters(plan, worksheets),
      ...buildRangeParameters(plan, worksheets)
    ],
    approvals: eligibility.requiredApprovals
  };
}

export function saveTool(tool: SavedTool): SavedTool[] {
  const current = loadTools();
  return persistTools([
    { ...tool, updatedAt: new Date().toISOString() },
    ...current.filter((item) => item.id !== tool.id)
  ]);
}

export function deleteTool(toolId: string): SavedTool[] {
  return persistTools(loadTools().filter((tool) => tool.id !== toolId));
}

function assertCompileContext(
  parameter: ToolParameter,
  value: string,
  values: Record<string, string>,
  workbook?: WorkbookSnapshot | null
) {
  if (parameter.type === "outputWorksheet") {
    if (value.length > 31 || /[:\\/?*\[\]]/.test(value)) {
      throw new Error(
        `输出工作表名称「${value}」无效：不能超过 31 个字符，也不能包含 : \\ / ? * [ ]`
      );
    }
    if (value.startsWith("'") || value.endsWith("'")) {
      throw new Error("输出工作表名称不能以单引号开头或结尾");
    }
    if (
      workbook?.worksheets.some(
        (sheet) => normalize(sheet.name) === normalize(value)
      )
    ) {
      throw new Error(`工作表「${value}」已存在，请更换输出工作表名称`);
    }
    return;
  }
  if (parameter.type === "range") {
    if (/[\r\n\0]/.test(value)) {
      throw new Error(`${parameter.label}包含无效字符`);
    }
    if (parameter.sourceParameterId && workbook) {
      const source = values[parameter.sourceParameterId];
      if (
        source &&
        !workbook.worksheets.some(
          (sheet) => normalize(sheet.name) === normalize(source)
        )
      ) {
        throw new Error(`数据范围对应的来源工作表「${source}」不存在`);
      }
    }
    return;
  }
  if (!workbook) return;
  if (parameter.type === "worksheet") {
    if (
      !workbook.worksheets.some(
        (sheet) => normalize(sheet.name) === normalize(value)
      )
    ) {
      throw new Error(`工作簿中不存在工作表「${value}」`);
    }
    return;
  }
  const source = values[parameter.sourceParameterId];
  const sheet = workbook.worksheets.find((item) => item.name === source);
  if (!sheet) {
    throw new Error(`请先选择${parameter.label}对应的来源工作表`);
  }
  if (
    !sheet.headers.some(
      (header) => normalize(String(header ?? "")) === normalize(value)
    )
  ) {
    throw new Error(
      `工作表「${sheet.name}」中没有字段「${value}」，请重新选择`
    );
  }
}

export function instantiateTool(
  tool: SavedTool,
  values: Record<string, string>,
  workbook?: WorkbookSnapshot | null
): AnalysisPlan {
  const resolved: Record<string, string> = {};
  for (const parameter of tool.parameters) {
    const value = values[parameter.id]?.trim() || parameter.defaultValue;
    if (parameter.required && !value) {
      throw new Error(
        parameter.type === "outputWorksheet" || parameter.type === "range"
          ? `请输入${parameter.label}`
          : `请选择${parameter.label}`
      );
    }
    resolved[parameter.id] = value;
    assertCompileContext(parameter, value, resolved, workbook);
  }
  const plan = structuredClone(tool.planTemplate);
  plan.id = `tool-${tool.id}-${Date.now()}`;
  plan.title = tool.name;
  plan.summary = tool.description;

  for (const parameter of tool.parameters) {
    const value = resolved[parameter.id];
    if (
      parameter.type === "worksheet" ||
      parameter.type === "outputWorksheet"
    ) {
      for (const binding of parameter.bindings) {
        const action = plan.actions[binding.actionIndex];
        if (binding.property === "sheet") action.sheet = value;
        else if (
          action.type === "copyRange" ||
          action.type === "createPivotTable"
        ) {
          action.sourceSheet = value;
        }
      }
      plan.acceptanceCriteria = plan.acceptanceCriteria?.map((criterion) => ({
        ...criterion,
        sheet:
          normalize(criterion.sheet) === normalize(parameter.defaultValue)
            ? value
            : criterion.sheet
      }));
      continue;
    }
    if (parameter.type === "range") {
      for (const binding of parameter.bindings) {
        const action = plan.actions[binding.actionIndex];
        if (
          binding.property === "range" &&
          (action.type === "sortRange" || action.type === "filterRange")
        ) {
          action.range = value;
        }
        if (binding.property === "sourceRange") {
          if (
            action.type === "copyRange" ||
            action.type === "createChart" ||
            action.type === "createPivotTable" ||
            action.type === "splitGroupAggregate"
          ) {
            action.sourceRange = value;
          }
        }
      }
      continue;
    }
    for (const binding of parameter.bindings) {
      const action = plan.actions[binding.actionIndex];
      if (action.type === "splitGroupAggregate") {
        if (binding.property === "splitBy") action.splitBy = value;
        if (binding.property === "groupBy") {
          action.groupBy[binding.itemIndex] = value;
        }
        if (binding.property === "metricField") {
          action.metrics[binding.itemIndex].field = value;
        }
      }
      if (action.type === "createPivotTable") {
        if (binding.property === "pivotRowField") {
          action.rowFields[binding.itemIndex] = value;
        }
        if (binding.property === "pivotColumnField") {
          action.columnFields[binding.itemIndex] = value;
        }
        if (binding.property === "pivotValueField") {
          action.valueFields[binding.itemIndex].field = value;
        }
      }
    }
  }
  plan.assumptions = [
    ...plan.assumptions,
    ...tool.parameters.map(
      (parameter) => `${parameter.label}使用「${resolved[parameter.id]}」`
    )
  ];
  return plan;
}
