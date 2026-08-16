// 展示用的格式化纯函数（从 App.tsx 抽出）。
import type { AnalysisPlan, VerificationReport } from "./contracts";
import type { SavedTool } from "./storage";

// 生成耗时格式化：<1s 显示毫秒，否则显示秒（保留一位小数）。
export function formatGenerateMs(ms: number): string {
  if (ms < 1000) return `${ms} 毫秒`;
  return `${(ms / 1000).toFixed(1)} 秒`;
}

export function formatStepElapsed(ms: number): string {
  if (ms < 1000) return `${Math.max(0, Math.round(ms))}ms`;
  return `${(ms / 1000).toFixed(1)}s`;
}

// 工作流工具的展示归类：按动作类型给出标签 / 图标字 / 色板，用于工具箱卡片与说明页。
export function workflowToolPresentation(tool: SavedTool): {
  label: string;
  glyph: string;
  tone: string;
} {
  const actionTypes = new Set(
    tool.planTemplate.actions.map((action) => action.type)
  );
  if (actionTypes.has("splitGroupAggregate")) {
    return { label: "拆分统计", glyph: "分", tone: "sage" };
  }
  if (
    actionTypes.has("createPivotTable") ||
    actionTypes.has("createChart")
  ) {
    return { label: "分析呈现", glyph: "析", tone: "blue" };
  }
  if (
    actionTypes.has("filterRange") ||
    actionTypes.has("sortRange")
  ) {
    return { label: "整理数据", glyph: "整", tone: "amber" };
  }
  return { label: "工作流程", glyph: "工", tone: "slate" };
}

export function verificationSummary(report: VerificationReport): string {
  if (report.status === "verified") return "并通过独立验证";
  if (report.status === "executed_unverified") {
    return "；写入已完成，但部分操作暂时无法独立验证";
  }
  return "，但结果验证未通过";
}

export function actionLabel(
  action: AnalysisPlan["actions"][number]
): string {
  switch (action.type) {
    case "createWorksheet":
      return `新建或复用工作表「${action.sheet}」`;
    case "writeTable":
      return `向「${action.sheet}」${action.startCell} 写入 ${
        action.rows.length + 1
      } 行表格`;
    case "writeValues":
      return `写入「${action.sheet}」${action.range}`;
    case "setFill":
      return `设置「${action.sheet}」${action.range} 的填充色`;
    case "setFont":
      return `设置「${action.sheet}」${action.range} 的字体`;
    case "autofit":
      return `自动调整「${action.sheet}」${action.range}`;
    case "activateWorksheet":
      return `切换到工作表「${action.sheet}」`;
    case "deleteWorksheet":
      return `删除工作表「${action.sheet}」`;
    case "clearRange":
      return `清除「${action.sheet}」${action.range}（${action.applyTo}）${
        action.filters?.length ? "，仅限命中行" : ""
      }`;
    case "insertRange":
      return `在「${action.sheet}」${action.range} 插入单元格`;
    case "deleteRange":
      return `删除「${action.sheet}」${action.range} 的单元格${
        action.filters?.length ? "（仅限命中行）" : ""
      }`;
    case "copyRange":
      return `复制「${action.sourceSheet}」${action.sourceRange} 到「${action.sheet}」${action.targetRange}${
        action.filters?.length ? "（仅限命中行）" : ""
      }`;
    case "writeFormulas":
      return `向「${action.sheet}」${action.range} 写入公式`;
    case "sortRange":
      return `排序「${action.sheet}」${action.range}`;
    case "removeDuplicates":
      return `去重「${action.sheet}」${action.range}${
        action.filters?.length ? "（仅限命中行）" : ""
      }`;
    case "filterRange":
      return `筛选「${action.sheet}」${action.range}`;
    case "clearFilter":
      return `清除「${action.sheet}」的筛选条件`;
    case "setDataValidation":
      return `设置「${action.sheet}」${action.range} 的数据验证`;
    case "setConditionalFormat":
      return `设置「${action.sheet}」${action.range} 的条件格式`;
    case "setNumberFormat":
      return `设置「${action.sheet}」${action.range} 的数字格式`;
    case "setBorders":
      return `设置「${action.sheet}」${action.range} 的边框`;
    case "setAlignment":
      return `设置「${action.sheet}」${action.range} 的对齐方式`;
    case "mergeCells":
      return `合并「${action.sheet}」${action.range}`;
    case "unmergeCells":
      return `取消合并「${action.sheet}」${action.range}`;
    case "resizeRange":
      return `调整「${action.sheet}」${action.range} 的行高列宽`;
    case "freezePanes":
      return `冻结「${action.sheet}」的窗格`;
    case "setHyperlink":
      return `为「${action.sheet}」${action.range} 添加超链接`;
    case "addComment":
      return `为「${action.sheet}」${action.cell} 添加批注`;
    case "addNote":
      return `为「${action.sheet}」${action.cell} 添加备注`;
    case "createTable":
      return `将「${action.sheet}」${action.range} 转换为表格`;
    case "createChart":
      return `基于「${action.sheet}」${action.sourceRange} 创建图表`;
    case "createPivotTable":
      return `在「${action.sheet}」创建数据透视表「${action.name}」`;
    case "splitGroupAggregate":
      return `按「${action.splitBy}」拆分「${action.sheet}」，并按 ${action.groupBy.join(
        "、"
      )} 汇总`;
    case "addNamedRange":
      return `为「${action.sheet}」${action.range} 创建名称「${action.name}」`;
    case "addImage":
      return `在「${action.sheet}」${action.targetRange} 添加图片`;
    case "addShape":
      return `在「${action.sheet}」${action.targetRange} 添加形状`;
  }
}
