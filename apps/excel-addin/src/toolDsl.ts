import type { AnalysisPlan, ExcelAction } from "./contracts";

function escapeHtml(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function quote(value: string): string {
  return `"${value
    .replaceAll("\\", "\\\\")
    .replaceAll('"', '\\"')
    .replaceAll("\r", "\\r")
    .replaceAll("\n", "\\n")}"`;
}

function stableValue(value: unknown): string {
  if (value === null) return "NULL";
  if (typeof value === "string") return quote(value);
  if (typeof value === "number" || typeof value === "boolean") {
    return String(value).toUpperCase();
  }
  if (Array.isArray(value)) {
    return `[${value.map(stableValue).join(", ")}]`;
  }
  if (typeof value === "object") {
    const entries = Object.entries(value as Record<string, unknown>).sort(
      ([left], [right]) => (left < right ? -1 : left > right ? 1 : 0)
    );
    return `{ ${entries
      .map(([key, nested]) => `${toDslName(key)}: ${stableValue(nested)}`)
      .join(", ")} }`;
  }
  return quote(String(value));
}

function toDslName(value: string): string {
  return value
    .replace(/([a-z0-9])([A-Z])/g, "$1_$2")
    .replace(/[^a-zA-Z0-9]+/g, "_")
    .replace(/^_+|_+$/g, "")
    .toUpperCase();
}

function matrixSummary(value: unknown): string | null {
  if (!Array.isArray(value)) return null;
  const rows = value.length;
  const columns = value.reduce(
    (maximum, row) =>
      Math.max(maximum, Array.isArray(row) ? row.length : 0),
    0
  );
  return `<MATRIX ROWS=${rows} COLUMNS=${columns}>`;
}

function formatArgument(key: string, value: unknown): string {
  if (key === "base64" && typeof value === "string") {
    return `<EMBEDDED_IMAGE CHARACTERS=${value.length}>`;
  }
  if (["rows", "values", "formulas"].includes(key)) {
    return matrixSummary(value) ?? stableValue(value);
  }
  return stableValue(value);
}

function renderAction(action: ExcelAction, index: number): string[] {
  const entries = Object.entries(action)
    .filter(([key]) => key !== "type")
    .sort(([left], [right]) => (left < right ? -1 : left > right ? 1 : 0));
  return [
    `STEP ${index + 1} ${toDslName(action.type)}`,
    ...entries.map(
      ([key, value]) => `  ${toDslName(key)} = ${formatArgument(key, value)}`
    )
  ];
}

/**
 * Turns a saved white-list AnalysisPlan into a deterministic, read-only DSL.
 * The output is presentation only; execution continues to use AnalysisPlan.
 */
export function renderToolDsl(plan: AnalysisPlan): string {
  const lines = [
    `TOOL ${quote(plan.title)}`,
    `SUMMARY ${quote(plan.summary)}`,
    "FORMAT EXCEL_BRO_CONTROLLED_PLAN_V1",
    "PREVIEW REQUIRED",
    `SOURCE_SCOPE ${
      plan.sourceFingerprintSheets?.length
        ? stableValue(plan.sourceFingerprintSheets)
        : "BOUND_AT_RUN_TIME"
    }`
  ];

  if (plan.assumptions.length > 0) {
    lines.push("", "ASSUMPTIONS", ...plan.assumptions.map((item) => `  - ${quote(item)}`));
  }
  if (plan.warnings.length > 0) {
    lines.push("", "WARNINGS", ...plan.warnings.map((item) => `  - ${quote(item)}`));
  }

  lines.push("", "ACTIONS");
  plan.actions.forEach((action, index) => {
    if (index > 0) lines.push("");
    lines.push(...renderAction(action, index));
  });

  lines.push(
    "",
    "GUARDRAILS",
    "  EXECUTION = WHITE_LIST_ANALYSIS_PLAN",
    "  WRITE_POLICY = PREVIEW_THEN_CONFIRM",
    "  MODEL_REINTERPRETATION = DISABLED",
    "  ARBITRARY_CODE = DISABLED",
    "  VBA = DISABLED",
    "  EXTERNAL_PROGRAM = DISABLED"
  );

  return lines.join("\n");
}

const DSL_TOKEN_PATTERN =
  /\b(TOOL|SUMMARY|FORMAT|PREVIEW|SOURCE_SCOPE|ASSUMPTIONS|WARNINGS|ACTIONS|GUARDRAILS|STEP|BOUND_AT_RUN_TIME|REQUIRED|EXCEL_BRO_CONTROLLED_PLAN_V1)\b|([A-Z][A-Z0-9_]*)(\s*=\s*)|("(?:[^"\\]|\\.)*")|(<(?:[A-Z_ ]+\s*=\s*\d+\s*)+>)|(\bTRUE\b|\bFALSE\b)|\bNULL\b|(-?\d+(?:\.\d+)?)/g;

/**
 * 只读 DSL 展示用的轻量高亮：关键字、参数名、字符串、矩阵标签、
 * 布尔/空值和数字分色，其余保持原样。仅做展示，不改动 DSL 文本。
 */
export function highlightToolDsl(dsl: string): string {
  let out = "";
  let last = 0;
  DSL_TOKEN_PATTERN.lastIndex = 0;
  let match: RegExpExecArray | null;
  while ((match = DSL_TOKEN_PATTERN.exec(dsl)) !== null) {
    out += escapeHtml(dsl.slice(last, match.index));
    if (match[1]) {
      out += `<span class="tk-kw">${escapeHtml(match[1])}</span>`;
    } else if (match[2]) {
      out += `<span class="tk-key">${escapeHtml(match[2])}</span>${escapeHtml(
        match[3]
      )}`;
    } else if (match[4]) {
      out += `<span class="tk-str">${escapeHtml(match[4])}</span>`;
    } else if (match[5]) {
      out += `<span class="tk-tag">${escapeHtml(match[5])}</span>`;
    } else if (match[6]) {
      out += `<span class="tk-bool">${escapeHtml(match[6])}</span>`;
    } else if (match[0] === "NULL") {
      out += `<span class="tk-null">NULL</span>`;
    } else {
      out += `<span class="tk-num">${escapeHtml(match[0])}</span>`;
    }
    last = match.index + match[0].length;
  }
  out += escapeHtml(dsl.slice(last));
  return out;
}
