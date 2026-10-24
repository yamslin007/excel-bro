import type { CellValue, ExcelAction } from "../../excel-addin/src/contracts";
import { executeQueryTableData } from "../../excel-addin/src/dataTools";
import { normalizeCellValue } from "../../excel-addin/src/cellNormalization";
import { parseExplicitA1Rect } from "@excel-bro/core/planCore";
import { columnLetterFromIndex } from "@excel-bro/core/range";
import { WRITE_TABLE_HEADER_FILL } from "@excel-bro/core/writeTableLayout";
import type { WpsRange } from "./wpsHost";

type ComputeAction = Extract<ExcelAction, { type: "computeValues" }>;

export function isCompleteSourceMatrix(raw: unknown, rows: number, columns: number): boolean {
  const scalar = (value: unknown) => value === null || typeof value === "string" || typeof value === "boolean" || (typeof value === "number" && Number.isFinite(value));
  if (!Array.isArray(raw)) return rows === 1 && columns === 1 && scalar(raw);
  if (raw.some(Array.isArray)) return raw.length === rows && raw.every(row => Array.isArray(row) && row.length === columns && row.every(scalar));
  return (rows === 1 || columns === 1) && raw.length === rows * columns && raw.every(scalar);
}

/** Local presentation evidence is never projected into the server recovery descriptor. */
export function computationPresentation(range: WpsRange, rows: number, columns: number): { formats: string[][]; texts: string[][] } {
  if (rows * columns > 1 && typeof range.Item !== "function") throw new Error("计算来源无法逐格读取数字格式，未写入。");
  const formats: string[][] = [], texts: string[][] = [];
  for (let r = 0; r < rows; r++) {
    formats.push([]); texts.push([]);
    for (let c = 0; c < columns; c++) {
      const cell = rows * columns === 1 ? range : range.Item!(r + 1, c + 1);
      if (typeof cell.NumberFormat !== "string") throw new Error("计算来源的数字格式未完整取得，未写入。");
      const text = cell.Text;
      if ((cell.NumberFormat === "@" || /^0+$/.test(cell.NumberFormat)) && typeof text !== "string") {
        throw new Error("计算来源包含文本编号格式，但显示文本未取得，未写入。");
      }
      formats[r].push(cell.NumberFormat);
      texts[r].push(typeof text === "string" ? text : "");
    }
  }
  return { formats, texts };
}

/** Reuses the same deterministic query engine as Excel; no host calls, eval or model calls. */
export function computedValuesPlan(action: ComputeAction, rawValues: CellValue[][],
  presentation: ReturnType<typeof computationPresentation>, maxCells: number): ExcelAction[] {
  const destination = parseExplicitA1Rect(action.destinationCell);
  if (!destination || destination.rows !== 1 || destination.columns !== 1) throw new Error("计算结果目标必须是单个 A1 起点。");
  const input = rawValues.map((row, r) => row.map((value, c) =>
    normalizeCellValue(value, presentation.texts[r][c], presentation.formats[r][c])));
  const values = action.hasHeaders ? input : [input[0].map((_, c) => `列${c + 1}`), ...input];
  const result = executeQueryTableData({ id: "computeValues", tool: "query_table", arguments: action.query },
    [{ name: action.sourceSheet, values }]);
  if (!result.complete) throw new Error(`计算来源不能完整处理，未写入：${result.warnings.join("；")}`);
  const output = action.includeHeaders ? [result.headers, ...result.rows] : result.rows;
  if (!output.length || !output[0].length) throw new Error("本次计算没有可写入结果，原表未修改。");
  const rows = output.length, columns = output[0].length;
  if (rows * columns > maxCells || output.some(row => row.length !== columns)) throw new Error("计算结果尺寸不符或超过 WPS 单次单元格限额。");
  if (destination.startRow + rows - 1 > 1048576 || destination.startColumn + columns - 1 > 16384) throw new Error("计算结果超出工作表边界。");
  const write: Extract<ExcelAction, {type:"writeValues"}> = { type: "writeValues", sheet: action.sheet,
    range: `${columnLetterFromIndex(destination.startColumn - 1)}${destination.startRow}:${columnLetterFromIndex(destination.startColumn + columns - 2)}${destination.startRow + rows - 1}`,
    values: output };
  if (!action.includeHeaders) return [write];
  const header = `${columnLetterFromIndex(destination.startColumn - 1)}${destination.startRow}:${columnLetterFromIndex(destination.startColumn + columns - 2)}${destination.startRow}`;
  return [write, {type:"setFont",sheet:action.sheet,range:header,bold:true},
    {type:"setFill",sheet:action.sheet,range:header,color:WRITE_TABLE_HEADER_FILL}];
}
