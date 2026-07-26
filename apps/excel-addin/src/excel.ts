import type {
  AnalysisPlan,
  ActionExecutionResult,
  CellValue,
  ExcelAction,
  PlanExecutionResult,
  VerificationCheck,
  VerificationCriterion,
  VerificationReport,
  WorkbookSnapshot,
  WorksheetSnapshot
} from "./contracts";
import capabilities from "../../../config/capabilities.json";
import { detectSheetFields } from "./tableSchema";
import {
  buildSplitAggregateOutputs,
  safeWorksheetBaseName
} from "./splitAggregate";
import { workbookNameFromDocumentUrl } from "./workbookIdentity";

const DATA_ROW_LIMIT = capabilities.snapshot.dataRows;
const DATA_COLUMN_LIMIT = capabilities.snapshot.dataColumns;
const STRUCTURE_ROW_LIMIT = capabilities.snapshot.structureRows;
const STRUCTURE_COLUMN_LIMIT = capabilities.snapshot.structureColumns;

function normalizeValue(value: unknown): CellValue {
  if (
    value === null ||
    typeof value === "string" ||
    typeof value === "number" ||
    typeof value === "boolean"
  ) {
    return value;
  }
  return String(value);
}

function excelErrorDetail(reason: unknown): string {
  if (!(reason instanceof Error)) return String(reason);
  const error = reason as Error & {
    code?: string;
    debugInfo?: {
      errorLocation?: string;
      statement?: string;
      surroundingStatements?: string[];
    };
  };
  const details = [
    error.message,
    error.code ? `代码 ${error.code}` : "",
    error.debugInfo?.errorLocation
      ? `位置 ${error.debugInfo.errorLocation}`
      : "",
    error.debugInfo?.statement
      ? `语句 ${error.debugInfo.statement}`
      : ""
  ].filter(Boolean);
  return details.join("；");
}

export function isRunningInExcel(): boolean {
  return typeof Office !== "undefined" && Office.context?.host === Office.HostType.Excel;
}

export async function captureSelectionContext(): Promise<{
  activeWorksheet: string;
  selectedRange: string;
}> {
  return Excel.run(async (context) => {
    const activeWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const selectedRange = context.workbook.getSelectedRange();
    activeWorksheet.load("name");
    selectedRange.load("address");
    await context.sync();
    return {
      activeWorksheet: activeWorksheet.name,
      selectedRange: selectedRange.address
    };
  });
}

export async function captureWorkbook(
  dataSheetNames?: string[]
): Promise<WorkbookSnapshot> {
  return Excel.run(async (context) => {
    const workbook = context.workbook;
    const worksheets = workbook.worksheets;
    const activeWorksheet = worksheets.getActiveWorksheet();
    const selectedRange = workbook.getSelectedRange();

    worksheets.load("items/name");
    activeWorksheet.load("name");
    selectedRange.load("address");
    await context.sync();

    const pending = worksheets.items.map((worksheet) => {
      const usedRange = worksheet.getUsedRangeOrNullObject(true);
      usedRange.load("address,rowCount,columnCount,rowIndex,columnIndex,isNullObject");
      return { worksheet, usedRange };
    });
    await context.sync();

    const sheetsToRead = new Set(
      dataSheetNames?.length ? dataSheetNames : [activeWorksheet.name]
    );
    const dataRanges = pending.map(({ worksheet, usedRange }) => {
      if (usedRange.isNullObject || !sheetsToRead.has(worksheet.name)) {
        return null;
      }
      const rowCount = Math.min(usedRange.rowCount, DATA_ROW_LIMIT + 1);
      const columnCount = Math.min(usedRange.columnCount, DATA_COLUMN_LIMIT);
      const range = worksheet.getRangeByIndexes(
        usedRange.rowIndex,
        usedRange.columnIndex,
        rowCount,
        columnCount
      );
      range.load("values");
      return range;
    });
    await context.sync();

    const snapshots: WorksheetSnapshot[] = pending.map(
      ({ worksheet, usedRange }, index) => {
        if (usedRange.isNullObject) {
          return {
            name: worksheet.name,
            usedRange: null,
            rowCount: 0,
            columnCount: 0,
            headers: [],
            dataRows: [],
            truncated: false
          };
        }

        const values = dataRanges[index]?.values ?? [];
        return {
          name: worksheet.name,
          usedRange: usedRange.address,
          rowCount: usedRange.rowCount,
          columnCount: usedRange.columnCount,
          headers: (values[0] ?? []).map(normalizeValue),
          dataRows: values.slice(1).map((row) => row.map(normalizeValue)),
          truncated:
            usedRange.rowCount > DATA_ROW_LIMIT + 1 ||
            usedRange.columnCount > DATA_COLUMN_LIMIT
        };
      }
    );

    return {
      name: workbookNameFromDocumentUrl(Office.context.document.url),
      capturedAt: new Date().toISOString(),
      activeWorksheet: activeWorksheet.name,
      selectedRange: selectedRange.address,
      worksheets: snapshots
    };
  });
}

export async function captureWorkbookStructure(
  dataSheetNames?: string[]
): Promise<WorkbookSnapshot> {
  return Excel.run(async (context) => {
    const workbook = context.workbook;
    const worksheets = workbook.worksheets;
    const activeWorksheet = worksheets.getActiveWorksheet();
    const selectedRange = workbook.getSelectedRange();

    worksheets.load("items/name");
    activeWorksheet.load("name");
    selectedRange.load("address");
    await context.sync();

    const pending = worksheets.items.map((worksheet) => {
      const usedRange = worksheet.getUsedRangeOrNullObject(true);
      usedRange.load(
        "address,rowCount,columnCount,rowIndex,columnIndex,isNullObject"
      );
      return { worksheet, usedRange };
    });
    await context.sync();

    const sheetsToRead = new Set(
      dataSheetNames?.length ? dataSheetNames : [activeWorksheet.name]
    );
    const previews = pending.map(({ worksheet, usedRange }) => {
      if (usedRange.isNullObject || !sheetsToRead.has(worksheet.name)) {
        return null;
      }
      const preview = worksheet.getRangeByIndexes(
        usedRange.rowIndex,
        usedRange.columnIndex,
        Math.min(usedRange.rowCount, STRUCTURE_ROW_LIMIT),
        Math.min(usedRange.columnCount, STRUCTURE_COLUMN_LIMIT)
      );
      preview.load("values");
      return preview;
    });
    await context.sync();

    return {
      name: workbookNameFromDocumentUrl(Office.context.document.url),
      capturedAt: new Date().toISOString(),
      activeWorksheet: activeWorksheet.name,
      selectedRange: selectedRange.address,
      worksheets: pending.map(({ worksheet, usedRange }, index) => {
        if (usedRange.isNullObject) {
          return {
            name: worksheet.name,
            usedRange: null,
            rowCount: 0,
            columnCount: 0,
            headers: [],
            dataRows: [],
            truncated: false
          };
        }
        const values = previews[index]?.values ?? [];
        return {
          name: worksheet.name,
          usedRange: usedRange.address,
          rowCount: usedRange.rowCount,
          columnCount: usedRange.columnCount,
          headers: detectSheetFields(
            values.map((row) => row.map(normalizeValue))
          ),
          dataRows: [],
          truncated:
            usedRange.rowCount > STRUCTURE_ROW_LIMIT ||
            usedRange.columnCount > STRUCTURE_COLUMN_LIMIT
        };
      })
    };
  });
}

async function getOrCreateWorksheet(
  context: Excel.RequestContext,
  sheetName: string
): Promise<Excel.Worksheet> {
  const collection = context.workbook.worksheets;
  const existing = collection.getItemOrNullObject(sheetName);
  existing.load("isNullObject");
  await context.sync();
  return existing.isNullObject ? collection.add(sheetName) : existing;
}

function assertRectangular(values: CellValue[][]): void {
  if (values.length === 0) return;
  const width = values[0].length;
  if (width === 0 || values.some((row) => row.length !== width)) {
    throw new Error("写入数据必须是非空的规则二维数组");
  }
}

function columnNumber(name: string): number {
  return [...name.toUpperCase()].reduce(
    (result, character) => result * 26 + character.charCodeAt(0) - 64,
    0
  );
}

function columnName(number: number): string {
  let result = "";
  while (number > 0) {
    const remainder = (number - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    number = Math.floor((number - 1) / 26);
  }
  return result;
}

function matrixRange(
  startCell: string,
  rowCount: number,
  columnCount: number
): string | null {
  const match = /^\$?([A-Za-z]+)\$?(\d+)$/.exec(startCell.trim());
  if (!match) return null;
  const startColumn = columnNumber(match[1]);
  const startRow = Number(match[2]);
  return `${match[1].toUpperCase()}${startRow}:${columnName(
    startColumn + columnCount - 1
  )}${startRow + rowCount - 1}`;
}

function inferredCriteria(
  plan: AnalysisPlan,
  dynamicCriteria: VerificationCriterion[] = []
): VerificationCriterion[] {
  const criteria: VerificationCriterion[] = [
    ...(plan.acceptanceCriteria ?? []),
    ...dynamicCriteria
  ];
  const criterionKeys = new Set(criteria.map((criterion) => JSON.stringify(criterion)));
  const addCriterion = (criterion: VerificationCriterion) => {
    const key = JSON.stringify(criterion);
    if (!criterionKeys.has(key)) {
      criteria.push(criterion);
      criterionKeys.add(key);
    }
  };
  const seenSheets = new Set<string>();
  for (const action of plan.actions) {
    if (action.type === "deleteWorksheet") {
      addCriterion({ type: "worksheetMissing", sheet: action.sheet });
      continue;
    }
    if (!seenSheets.has(action.sheet)) {
      addCriterion({ type: "worksheetExists", sheet: action.sheet });
      seenSheets.add(action.sheet);
    }
    if (action.type === "writeValues") {
      addCriterion({
        type: "rangeEquals",
        sheet: action.sheet,
        range: action.range,
        expected: action.values
      });
    } else if (action.type === "writeTable") {
      const expected = [action.headers, ...action.rows];
      const range = matrixRange(
        action.startCell,
        expected.length,
        action.headers.length
      );
      if (range) {
        addCriterion({
          type: "rangeEquals",
          sheet: action.sheet,
          range,
          expected
        });
      }
    } else if (action.type === "writeFormulas") {
      addCriterion({
        type: "formulasEqual",
        sheet: action.sheet,
        range: action.range,
        expected: action.formulas
      });
    } else if (
      action.type === "clearRange" &&
      (action.applyTo === "all" || action.applyTo === "contents")
    ) {
      addCriterion({
        type: "rangeEmpty",
        sheet: action.sheet,
        range: action.range
      });
    }
  }
  return criteria;
}

export function valuesEqual(actual: CellValue, expected: CellValue): boolean {
  if ((actual === null || actual === "") && (expected === null || expected === "")) {
    return true;
  }
  if (typeof actual === "number" && typeof expected === "number") {
    return (
      actual === expected ||
      Math.abs(actual - expected) <=
        Number.EPSILON * 16 * Math.max(1, Math.abs(actual), Math.abs(expected))
    );
  }
  if (
    (typeof actual === "number" && typeof expected === "string") ||
    (typeof actual === "string" && typeof expected === "number")
  ) {
    const numberValue = typeof actual === "number" ? actual : expected;
    const textValue = String(
      typeof actual === "string" ? actual : expected
    ).trim();
    if (!textValue || !Number.isFinite(Number(textValue))) return false;
    const losesLeadingZero =
      /^[-+]?0\d/.test(textValue) && !/^[-+]?0(?:\.\d+)?$/.test(textValue);
    return !losesLeadingZero && Number(textValue) === numberValue;
  }
  return actual === expected;
}

function matricesEqual(actual: CellValue[][], expected: CellValue[][]): boolean {
  return (
    actual.length === expected.length &&
    actual.every(
      (row, rowIndex) =>
        row.length === expected[rowIndex].length &&
        row.every((value, columnIndex) =>
          valuesEqual(value, expected[rowIndex][columnIndex])
        )
    )
  );
}

function firstMatrixDifference(
  actual: CellValue[][],
  expected: CellValue[][]
): string | null {
  if (actual.length !== expected.length) {
    return `实际 ${actual.length} 行，预期 ${expected.length} 行`;
  }
  for (let rowIndex = 0; rowIndex < expected.length; rowIndex += 1) {
    if (actual[rowIndex].length !== expected[rowIndex].length) {
      return `第 ${rowIndex + 1} 行实际 ${actual[rowIndex].length} 列，预期 ${
        expected[rowIndex].length
      } 列`;
    }
    for (
      let columnIndex = 0;
      columnIndex < expected[rowIndex].length;
      columnIndex += 1
    ) {
      if (
        !valuesEqual(
          actual[rowIndex][columnIndex],
          expected[rowIndex][columnIndex]
        )
      ) {
        return `${columnName(columnIndex + 1)}${rowIndex + 1} 实际 ${JSON.stringify(
          actual[rowIndex][columnIndex]
        )}，预期 ${JSON.stringify(expected[rowIndex][columnIndex])}`;
      }
    }
  }
  return null;
}

function isBlankMatrix(values: CellValue[][]): boolean {
  return values.every((row) =>
    row.every((value) => value === null || value === "")
  );
}

async function verifyPlan(
  context: Excel.RequestContext,
  plan: AnalysisPlan,
  dynamicCriteria: VerificationCriterion[] = []
): Promise<VerificationReport> {
  const checks: VerificationCheck[] = [];
  for (const criterion of inferredCriteria(plan, dynamicCriteria)) {
    const sheet = context.workbook.worksheets.getItemOrNullObject(criterion.sheet);
    sheet.load("isNullObject");
    await context.sync();

    if (criterion.type === "worksheetMissing") {
      checks.push({
        criterion,
        passed: sheet.isNullObject,
        message: sheet.isNullObject
          ? `工作表「${criterion.sheet}」已删除`
          : `工作表「${criterion.sheet}」仍然存在`
      });
      continue;
    }
    if (sheet.isNullObject) {
      checks.push({
        criterion,
        passed: false,
        message: `未找到工作表「${criterion.sheet}」`
      });
      continue;
    }
    if (criterion.type === "worksheetExists") {
      checks.push({
        criterion,
        passed: true,
        message: `工作表「${criterion.sheet}」存在`
      });
      continue;
    }

    const range = sheet.getRange(criterion.range);
    range.load(criterion.type === "formulasEqual" ? "formulas" : "values");
    await context.sync();
    const actual =
      criterion.type === "formulasEqual"
        ? range.formulas.map((row) => row.map(normalizeValue))
        : range.values.map((row) => row.map(normalizeValue));
    const passed =
      criterion.type === "rangeEmpty"
        ? isBlankMatrix(actual)
        : matricesEqual(actual, criterion.expected);
    const difference =
      !passed && criterion.type !== "rangeEmpty"
        ? firstMatrixDifference(actual, criterion.expected)
        : null;
    checks.push({
      criterion,
      passed,
      message: passed
        ? criterion.type === "rangeEmpty"
          ? `「${criterion.sheet}」${criterion.range} 已清空`
          : criterion.type === "formulasEqual"
            ? `「${criterion.sheet}」${criterion.range} 公式一致`
            : `「${criterion.sheet}」${criterion.range} 写入值一致`
        : `「${criterion.sheet}」${criterion.range} 与预期不一致${
            difference ? `（${difference}）` : ""
          }`,
      actual
    });
  }
  return {
    passed: checks.length > 0 && checks.every((check) => check.passed),
    checks
  };
}

function enumText(value: string): string {
  return value
    .trim()
    .replace(/[-_\s]+(.)/g, (_, character: string) => character.toUpperCase())
    .replace(/^./, (character) => character.toUpperCase());
}

function dataValidationRule(
  action: Extract<ExcelAction, { type: "setDataValidation" }>
): Excel.DataValidationRule {
  if (action.validationType === "list") {
    return {
      list: {
        inCellDropDown: true,
        source: action.values.map(String).join(",")
      }
    };
  }
  if (action.validationType === "custom") {
    return { custom: { formula: String(action.formula1 ?? "") } };
  }
  const basic = {
    formula1: action.formula1 ?? 0,
    ...(action.formula2 !== undefined && action.formula2 !== null
      ? { formula2: action.formula2 }
      : {}),
    operator: enumText(action.operator) as Excel.DataValidationOperator
  };
  if (action.validationType === "date") {
    return {
      date: {
        formula1: String(basic.formula1),
        operator: basic.operator,
        ...(basic.formula2 !== undefined
          ? { formula2: String(basic.formula2) }
          : {})
      }
    };
  }
  return { [action.validationType]: basic } as Excel.DataValidationRule;
}

function positionShape(
  shape: Excel.Shape,
  target: Excel.Range
): void {
  shape.left = target.left;
  shape.top = target.top;
  shape.width = target.width;
  shape.height = target.height;
  shape.placement = "TwoCell";
}

function minimumExcelApiVersion(action: ExcelAction): string {
  const versions: Partial<Record<ExcelAction["type"], string>> = {
    sortRange: "1.2",
    setConditionalFormat: "1.6",
    freezePanes: "1.7",
    setHyperlink: "1.7",
    setDataValidation: "1.8",
    createPivotTable: "1.8",
    copyRange: "1.9",
    filterRange: "1.9",
    clearFilter: "1.9",
    addImage: "1.9",
    addShape: "1.9",
    addComment: "1.10",
    addNote: "1.18"
  };
  return versions[action.type] ?? "1.1";
}

async function executeAction(
  context: Excel.RequestContext,
  action: ExcelAction
): Promise<VerificationCriterion[] | void> {
  if (action.type === "deleteWorksheet") {
    const existing = context.workbook.worksheets.getItemOrNullObject(action.sheet);
    existing.load("isNullObject");
    await context.sync();
    if (existing.isNullObject) {
      throw new Error(`未找到工作表「${action.sheet}」`);
    }
    existing.delete();
    return;
  }

  if (action.type === "splitGroupAggregate") {
    const worksheets = context.workbook.worksheets;
    const sourceSheet = worksheets.getItemOrNullObject(action.sheet);
    sourceSheet.load("isNullObject,name");
    try {
      await context.sync();
    } catch (reason) {
      throw new Error(`定位源工作表失败：${excelErrorDetail(reason)}`);
    }
    if (sourceSheet.isNullObject) {
      throw new Error(`未找到源工作表「${action.sheet}」`);
    }

    const sourceRange = action.sourceRange
      ? sourceSheet.getRange(action.sourceRange)
      : sourceSheet.getUsedRangeOrNullObject(true);
    sourceRange.load("isNullObject,values");
    try {
      await context.sync();
    } catch (reason) {
      throw new Error(`读取源数据失败：${excelErrorDetail(reason)}`);
    }
    if (sourceRange.isNullObject) {
      throw new Error(`源工作表「${action.sheet}」没有可读取的数据`);
    }
    const outputs = buildSplitAggregateOutputs(
      sourceRange.values.map((row) => row.map(normalizeValue)),
      {
        ...action,
        maxOutputSheets: action.maxOutputSheets ?? 200
      }
    );

    worksheets.load("items/name");
    await context.sync();
    const existingNames = new Map(
      worksheets.items.map((worksheet) => [
        worksheet.name.toLocaleLowerCase(),
        worksheet.name
      ])
    );
    const sourceName = sourceSheet.name.toLocaleLowerCase();
    const generatedNames = new Set<string>();
    const renamed = (base: string): string => {
      if (!existingNames.has(base.toLocaleLowerCase())) return base;
      let suffix = 2;
      while (true) {
        const marker = ` (${suffix})`;
        const candidate = `${base.slice(0, 31 - marker.length)}${marker}`;
        if (!existingNames.has(candidate.toLocaleLowerCase())) return candidate;
        suffix += 1;
      }
    };
    const criteria: VerificationCriterion[] = [];
    const createdInThisRun: string[] = [];
    const rollbackCreatedSheets = async (): Promise<void> => {
      if (createdInThisRun.length === 0) return;
      const candidates = createdInThisRun.map((name) => {
        const worksheet = worksheets.getItemOrNullObject(name);
        worksheet.load("isNullObject");
        return worksheet;
      });
      await context.sync();
      for (const worksheet of candidates) {
        if (!worksheet.isNullObject) worksheet.delete();
      }
      await context.sync();
    };

    for (const output of outputs) {
      const baseName = safeWorksheetBaseName(output.splitValue);
      const normalizedBase = baseName.toLocaleLowerCase();
      let targetName = baseName;
      let targetSheet: Excel.Worksheet;
      let createdNewSheet = false;
      const existingName = existingNames.get(normalizedBase);
      const collidesWithCurrentRun = generatedNames.has(normalizedBase);

      if (
        existingName &&
        action.existingSheetPolicy === "skip" &&
        !collidesWithCurrentRun
      ) {
        criteria.push({
          type: "worksheetExists",
          sheet: existingName
        });
        continue;
      }
      if (
        existingName &&
        action.existingSheetPolicy === "replace" &&
        !collidesWithCurrentRun
      ) {
        if (normalizedBase === sourceName) {
          throw new Error(`不能用拆分结果覆盖源工作表「${sourceSheet.name}」`);
        }
        targetSheet = worksheets.getItem(existingName);
        const used = targetSheet.getUsedRangeOrNullObject();
        used.load("isNullObject");
        await context.sync();
        if (!used.isNullObject) used.clear("All");
      } else {
        if (existingName) targetName = renamed(baseName);
        targetSheet = worksheets.add(targetName);
        createdNewSheet = true;
        existingNames.set(targetName.toLocaleLowerCase(), targetName);
      }
      generatedNames.add(targetName.toLocaleLowerCase());

      const matrix = [output.headers, ...output.rows];
      const target = targetSheet
        .getRange("A1")
        .getResizedRange(matrix.length - 1, matrix[0].length - 1);
      for (
        let columnIndex = 0;
        columnIndex < action.groupBy.length + 1;
        columnIndex += 1
      ) {
        if (
          matrix
            .slice(1)
            .some((row) => typeof row[columnIndex] === "string")
        ) {
          const dimensionRange = targetSheet.getRange(
            `${columnName(columnIndex + 1)}1:${columnName(columnIndex + 1)}${
              matrix.length
            }`
          );
          dimensionRange.numberFormat = Array.from(
            { length: matrix.length },
            () => ["@"]
          );
        }
      }
      target.values = matrix;
      const header = target.getRow(0);
      header.format.font.bold = true;
      header.format.fill.color = "#DFF3E4";
      for (const columnIndex of output.ratioColumnIndexes) {
        const ratioRange = targetSheet.getRange(
          `${columnName(columnIndex + 1)}2:${columnName(columnIndex + 1)}${
            matrix.length
          }`
        );
        ratioRange.numberFormat = Array.from(
          { length: matrix.length - 1 },
          () => ["0.00%"]
        );
      }
      target.format.autofitColumns();
      target.format.autofitRows();
      try {
        await context.sync();
        if (createdNewSheet) createdInThisRun.push(targetName);
      } catch (reason) {
        try {
          await rollbackCreatedSheets();
        } catch {
          // Keep the original Excel error; cleanup is best-effort.
        }
        throw new Error(
          `生成工作表「${targetName}」失败：${excelErrorDetail(reason)}`
        );
      }

      const range = matrixRange("A1", matrix.length, matrix[0].length);
      criteria.push(
        range && matrix.length <= 500
          ? {
              type: "rangeEquals",
              sheet: targetName,
              range,
              expected: matrix
            }
          : { type: "worksheetExists", sheet: targetName }
      );
    }
    return criteria;
  }

  const sheet = await getOrCreateWorksheet(context, action.sheet);

  switch (action.type) {
    case "createWorksheet":
      return;
    case "writeTable": {
      const values = [action.headers, ...action.rows];
      assertRectangular(values);
      const start = sheet.getRange(action.startCell);
      const target = start.getResizedRange(values.length - 1, values[0].length - 1);
      target.values = values;
      if (action.headers.length > 0) {
        const header = start.getResizedRange(0, action.headers.length - 1);
        header.format.font.bold = true;
        header.format.fill.color = "#DFF3E4";
      }
      return;
    }
    case "writeValues": {
      assertRectangular(action.values);
      const target = sheet.getRange(action.range);
      target.values = action.values;
      return;
    }
    case "writeFormulas": {
      const target = sheet.getRange(action.range);
      target.formulas = action.formulas;
      return;
    }
    case "clearRange":
      sheet
        .getRange(action.range)
        .clear(
          ({ all: "All", contents: "Contents", formats: "Formats", hyperlinks: "Hyperlinks" } as const)[
            action.applyTo
          ]
        );
      return;
    case "insertRange":
      sheet.getRange(action.range).insert(action.shift === "down" ? "Down" : "Right");
      return;
    case "deleteRange":
      sheet.getRange(action.range).delete(action.shift === "up" ? "Up" : "Left");
      return;
    case "copyRange": {
      const sourceSheet = context.workbook.worksheets.getItem(action.sourceSheet);
      sheet.getRange(action.targetRange).copyFrom(
        sourceSheet.getRange(action.sourceRange),
        enumText(action.copyType) as Excel.RangeCopyType,
        action.skipBlanks,
        action.transpose
      );
      return;
    }
    case "sortRange":
      sheet.getRange(action.range).sort.apply(
        action.keys.map((key) => ({ key: key.column, ascending: key.ascending })),
        false,
        action.hasHeaders
      );
      return;
    case "filterRange":
      sheet.autoFilter.apply(sheet.getRange(action.range), action.column, {
        filterOn: "Values",
        values: action.values.map((value) => String(value))
      });
      return;
    case "clearFilter":
      sheet.autoFilter.clearCriteria();
      return;
    case "setDataValidation": {
      const validation = sheet.getRange(action.range).dataValidation;
      validation.rule = dataValidationRule(action);
      validation.ignoreBlanks = action.allowBlank;
      if (action.prompt) {
        validation.prompt = {
          showPrompt: true,
          title: "输入提示",
          message: action.prompt
        };
      }
      if (action.errorMessage) {
        validation.errorAlert = {
          showAlert: true,
          style: "Stop",
          title: "输入无效",
          message: action.errorMessage
        };
      }
      return;
    }
    case "setConditionalFormat": {
      const range = sheet.getRange(action.range);
      if (action.ruleType === "colorScale") {
        const format = range.conditionalFormats.add("ColorScale").colorScale;
        format.criteria = {
          minimum: {
            type: "LowestValue",
            color: action.minColor ?? "#F8696B"
          },
          ...(action.midColor
            ? {
                midpoint: {
                  type: "Percent",
                  formula: "50",
                  color: action.midColor
                }
              }
            : {}),
          maximum: {
            type: "HighestValue",
            color: action.maxColor ?? "#63BE7B"
          }
        };
      } else if (action.ruleType === "custom") {
        const format = range.conditionalFormats.add("Custom").custom;
        format.rule.formula = String(action.formula1 ?? "");
        format.format.fill.color = action.color ?? "#FFF2CC";
      } else {
        const format = range.conditionalFormats.add("CellValue").cellValue;
        const conditionalOperator = {
          greaterThanOrEqualTo: "GreaterThanOrEqual",
          lessThanOrEqualTo: "LessThanOrEqual"
        }[action.operator ?? ""] ?? enumText(action.operator ?? "equalTo");
        format.rule = {
          operator: conditionalOperator as Excel.ConditionalCellValueOperator,
          formula1: String(action.formula1 ?? ""),
          ...(action.formula2 !== undefined && action.formula2 !== null
            ? { formula2: String(action.formula2) }
            : {})
        };
        format.format.fill.color = action.color ?? "#FFF2CC";
      }
      return;
    }
    case "setNumberFormat": {
      const range = sheet.getRange(action.range);
      range.load("rowCount,columnCount");
      await context.sync();
      range.numberFormat = Array.from({ length: range.rowCount }, () =>
        Array.from({ length: range.columnCount }, () => action.formatCode)
      );
      return;
    }
    case "setBorders": {
      const borderIndexes = {
        top: "EdgeTop",
        bottom: "EdgeBottom",
        left: "EdgeLeft",
        right: "EdgeRight",
        insideHorizontal: "InsideHorizontal",
        insideVertical: "InsideVertical"
      } as const;
      const borders = sheet.getRange(action.range).format.borders;
      for (const side of action.sides) {
        const border = borders.getItem(borderIndexes[side]);
        border.style = enumText(action.style) as Excel.BorderLineStyle;
        border.color = action.color;
        border.weight = enumText(action.weight) as Excel.BorderWeight;
      }
      return;
    }
    case "setAlignment": {
      const format = sheet.getRange(action.range).format;
      if (action.horizontal) {
        format.horizontalAlignment = enumText(
          action.horizontal
        ) as Excel.HorizontalAlignment;
      }
      if (action.vertical) {
        format.verticalAlignment = enumText(
          action.vertical
        ) as Excel.VerticalAlignment;
      }
      if (action.wrapText !== undefined && action.wrapText !== null) {
        format.wrapText = action.wrapText;
      }
      return;
    }
    case "mergeCells":
      sheet.getRange(action.range).merge(action.across);
      return;
    case "unmergeCells":
      sheet.getRange(action.range).unmerge();
      return;
    case "resizeRange": {
      const format = sheet.getRange(action.range).format;
      if (action.rowHeight !== undefined && action.rowHeight !== null) {
        format.rowHeight = action.rowHeight;
      }
      if (action.columnWidth !== undefined && action.columnWidth !== null) {
        format.columnWidth = action.columnWidth;
      }
      return;
    }
    case "freezePanes":
      sheet.freezePanes.unfreeze();
      if (action.rows > 0) sheet.freezePanes.freezeRows(action.rows);
      if (action.columns > 0) sheet.freezePanes.freezeColumns(action.columns);
      return;
    case "setHyperlink":
      sheet.getRange(action.range).hyperlink = {
        address: action.address,
        ...(action.text ? { textToDisplay: action.text } : {}),
        ...(action.screenTip ? { screenTip: action.screenTip } : {})
      };
      return;
    case "addComment":
      context.workbook.comments.add(sheet.getRange(action.cell), action.text);
      return;
    case "addNote":
      sheet.notes.add(sheet.getRange(action.cell), action.text);
      return;
    case "createTable": {
      const table = sheet.tables.add(action.range, action.hasHeaders);
      if (action.name) table.name = action.name;
      if (action.style) table.style = action.style;
      return;
    }
    case "createChart": {
      const chart = sheet.charts.add(
        enumText(action.chartType) as Excel.ChartType,
        sheet.getRange(action.sourceRange)
      );
      if (action.title) {
        chart.title.text = action.title;
        chart.title.visible = true;
      }
      if (action.targetRange) {
        chart.setPosition(sheet.getRange(action.targetRange));
      }
      return;
    }
    case "createPivotTable": {
      const sourceSheet = context.workbook.worksheets.getItem(action.sourceSheet);
      const pivot = sheet.pivotTables.add(
        action.name,
        sourceSheet.getRange(action.sourceRange),
        sheet.getRange(action.destinationCell)
      );
      for (const field of action.rowFields) {
        pivot.rowHierarchies.add(pivot.hierarchies.getItem(field));
      }
      for (const field of action.columnFields) {
        pivot.columnHierarchies.add(pivot.hierarchies.getItem(field));
      }
      for (const valueField of action.valueFields) {
        const hierarchy = pivot.dataHierarchies.add(
          pivot.hierarchies.getItem(valueField.field)
        );
        hierarchy.summarizeBy = enumText(
          valueField.aggregation
        ) as Excel.AggregationFunction;
      }
      return;
    }
    case "addNamedRange":
      context.workbook.names.add(
        action.name,
        sheet.getRange(action.range),
        action.comment ?? undefined
      );
      return;
    case "addImage": {
      const target = sheet.getRange(action.targetRange);
      target.load("left,top,width,height");
      await context.sync();
      const shape = sheet.shapes.addImage(action.base64);
      if (action.name) shape.name = action.name;
      positionShape(shape, target);
      return;
    }
    case "addShape": {
      const target = sheet.getRange(action.targetRange);
      target.load("left,top,width,height");
      await context.sync();
      const shape =
        action.shapeType === "line"
          ? sheet.shapes.addLine(
              target.left,
              target.top,
              target.left + target.width,
              target.top + target.height,
              "Straight"
            )
          : sheet.shapes.addGeometricShape(
              ({
                rectangle: "Rectangle",
                roundedRectangle: "RoundRectangle",
                ellipse: "Ellipse",
                triangle: "Triangle",
                diamond: "Diamond"
              } as Record<string, Excel.GeometricShapeType>)[action.shapeType] ??
                "Rectangle"
            );
      if (action.text) shape.textFrame.textRange.text = action.text;
      if (action.fillColor) shape.fill.setSolidColor(action.fillColor);
      positionShape(shape, target);
      return;
    }
    case "setFill":
      sheet.getRange(action.range).format.fill.color = action.color;
      return;
    case "setFont": {
      const font = sheet.getRange(action.range).format.font;
      if (action.bold !== undefined) font.bold = action.bold;
      if (action.color) font.color = action.color;
      return;
    }
    case "autofit": {
      const format = sheet.getRange(action.range).format;
      format.autofitColumns();
      format.autofitRows();
      return;
    }
    case "activateWorksheet":
      sheet.activate();
      return;
  }
}

export async function executePlan(plan: AnalysisPlan): Promise<PlanExecutionResult> {
  return Excel.run(async (context) => {
    const actionResults: ActionExecutionResult[] = [];
    const dynamicCriteria: VerificationCriterion[] = [];
    for (const [index, action] of plan.actions.entries()) {
      try {
        const requiredVersion = minimumExcelApiVersion(action);
        if (
          !Office.context.requirements.isSetSupported(
            "ExcelApi",
            requiredVersion
          )
        ) {
          throw new Error(
            `当前 Excel 不支持此操作（需要 ExcelApi ${requiredVersion} 或更高版本）`
          );
        }
        const actionCriteria = await executeAction(context, action);
        if (actionCriteria) dynamicCriteria.push(...actionCriteria);
        await context.sync();
      } catch (reason) {
        const detail = excelErrorDetail(reason);
        throw new Error(`第 ${index + 1} 步执行失败：${detail}`);
      }
      actionResults.push({
        index,
        type: action.type,
        sheet: action.sheet,
        status: "succeeded"
      });
    }
    const verification = await verifyPlan(context, plan, dynamicCriteria);
    return { actionResults, verification };
  });
}
