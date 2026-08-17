// 意图 / 需求确认相关的纯函数（从 App.tsx 抽出）。
import type { ChatMessage } from "./types/chat";
import type {
  FormulaDictionarySheet,
  IntentCheckResponse,
  IntentScopeContext,
  QueryTableArguments,
  ResultContext
} from "./contracts";
import { INTENT_MAX_PRIOR_RESULT_ROWS } from "./conversation";

// 把本地工具查询参数翻译成一句人能读懂的话，供「查看详情」里展示。
export function describeQueryArguments(args: QueryTableArguments): string {
  const parts: string[] = [];
  const modeLabel =
    args.mode === "aggregate"
      ? "分组汇总"
      : args.mode === "profile"
        ? "字段画像"
        : "取明细行";
  parts.push(`方式：${modeLabel}`);
  if (args.fields?.length) parts.push(`字段：${args.fields.join("、")}`);
  if (args.groupBy?.length) parts.push(`分组：${args.groupBy.join("、")}`);
  if (args.metrics?.length) {
    parts.push(
      `指标：${args.metrics
        .map((metric) => `${metric.operation}(${metric.field ?? "*"})`)
        .join("、")}`
    );
  }
  if (args.filters?.length) parts.push(`筛选：${args.filters.length} 个条件`);
  if (args.combine?.mode) parts.push(`合并：${args.combine.mode}`);
  if (typeof args.limit === "number") parts.push(`上限：${args.limit} 行`);
  return parts.join("\n");
}

// describeIntentDecision：把需求确认的结构化结果翻译成「用户看得懂的判断」。
// 返回 label（步骤标题）、note（模型的真实理解，第一层默认可见）、
// detail（原始明细，第二层展开「查看详情」才显示），让用户判断模型是否偏离了轨迹。
export function describeIntentDecision(intent: IntentCheckResponse): {
  label: string;
  note?: string;
  detail?: string;
} {
  const source = intent.provider === "model" ? "模型" : "本地规则";
  if (intent.kind === "clarification") {
    const clarification = intent.clarification;
    const optionText = clarification.options
      .map((option, index) => `${index + 1}. ${option.label}`)
      .join("\n");
    return {
      label: `${source}判断需要先澄清需求`,
      note: clarification.summary || clarification.question,
      detail: [
        `理解：${clarification.summary}`,
        `追问：${clarification.question}`,
        clarification.reason ? `原因：${clarification.reason}` : "",
        clarification.options.length ? `可选项：\n${optionText}` : ""
      ]
        .filter(Boolean)
        .join("\n\n")
    };
  }
  if (intent.kind === "tool_request") {
    return {
      label: `${source}决定先用本地工具取数`,
      note: intent.summary,
      detail: [
        `理解：${intent.summary}`,
        `锁定执行：${intent.confirmedPrompt}`,
        `本地查询：\n${describeQueryArguments(intent.request.arguments)}`
      ]
        .filter(Boolean)
        .join("\n\n")
    };
  }
  return {
    label: `${source}确认可以直接分析`,
    note: intent.summary,
    detail: [
      `理解：${intent.summary}`,
      `锁定执行：${intent.confirmedPrompt}`
    ]
      .filter(Boolean)
      .join("\n\n")
  };
}

export function intentScopeFingerprint(scope: IntentScopeContext): string {
  // selectedRange (光标位置) 不纳入指纹：查询不依赖它，写入落点由预览确认兜底。
  // 把它算进来只会在用户移动光标时造成误判（"数据范围已经变化"）。
  // activeWorksheet 在 auto 模式下决定扫描哪张表，切表意味着数据真的变了，仍需保留。
  return JSON.stringify({
    workbookName: scope.workbookName,
    sourceMode: scope.sourceMode,
    selectionMode: scope.selectionMode,
    sheets: scope.sheets.map((sheet) => sheet.name),
    ...(scope.selectionMode === "auto"
      ? {
          activeWorksheet: scope.activeWorksheet
        }
      : {})
  });
}

export function latestResultContext(items: ChatMessage[]): ResultContext | null {
  const result =
    [...items].reverse().find((message) => message.resultContext)
      ?.resultContext ?? null;
  return result
    ? {
        ...result,
        rows: result.rows.slice(0, INTENT_MAX_PRIOR_RESULT_ROWS)
      }
    : null;
}

// /function 上下文：扫描名字含"字典"/"映射"的表，读其内容注入生成提示。
export async function loadDictionaryForFormula(
  activeSheetName: string
): Promise<FormulaDictionarySheet | null> {
  try {
    return await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      sheets.load("items/name");
      await context.sync();

      const match = sheets.items.find(
        (sheet) =>
          sheet.name !== activeSheetName &&
          (sheet.name.includes("字典") || sheet.name.includes("映射"))
      );
      if (!match) return null;

      const usedRange = match.getUsedRangeOrNullObject(true);
      usedRange.load("values,rowCount,isNullObject");
      await context.sync();
      if (usedRange.isNullObject || usedRange.rowCount === 0) return null;

      const rows = (usedRange.values as unknown[][])
        .slice(0, 200)
        .map((row) => row.map((cell) => String(cell ?? "")));
      return { name: match.name, rows };
    });
  } catch {
    return null;
  }
}
