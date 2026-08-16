// 工具抽屉：工具箱 / 工具说明 / 运行工具 三种视图。
// 纯展示组件，JSX 与内部渲染辅助函数（renderToolParameter、workflowToolPresentation、
// formatToolDate）从 App.tsx 逐字搬移；状态与回调由 props 传入。
import type { Dispatch, SetStateAction } from "react";
import type { WorkbookSnapshot } from "./contracts";
import type {
  SavedQueryTool,
  SavedTool,
  ToolParameter
} from "./storage";
import { renderToolDsl } from "./toolDsl";
import { actionLabel, workflowToolPresentation } from "./format";
import type {
  ToolDetailMode,
  ToolDrawerView
} from "./hooks/useToolManagement";

interface ToolDrawerProps {
  toolsOpen: boolean;
  tools: SavedTool[];
  queryTools: SavedQueryTool[];
  toolDrawerView: ToolDrawerView;
  setToolDrawerView: Dispatch<SetStateAction<ToolDrawerView>>;
  selectedToolId: string | null;
  selectedQueryToolId: string | null;
  toolDetailMode: ToolDetailMode;
  setToolDetailMode: Dispatch<SetStateAction<ToolDetailMode>>;
  copiedToolDslId: string | null;
  toolParameterValues: Record<string, string>;
  workbook: WorkbookSnapshot | null;
  busy: boolean;
  setToolsOpen: Dispatch<SetStateAction<boolean>>;
  openWorkflowToolDetail: (tool: SavedTool) => void;
  openQueryToolDetail: (tool: SavedQueryTool) => void;
  requestToolDeletion: (
    kind: "workflow" | "query",
    tool: Pick<SavedTool | SavedQueryTool, "id" | "name">
  ) => void;
  copyToolDsl: (tool: SavedTool) => void;
  prepareToolRun: (tool: SavedTool) => void;
  runQueryTool: (tool: SavedQueryTool) => void;
  previewTool: (tool: SavedTool) => void;
  fieldOptions: (
    tool: SavedTool,
    parameter: Extract<ToolParameter, { type: "field" }>
  ) => string[];
  updateToolParameter: (
    tool: SavedTool,
    parameter: ToolParameter,
    value: string
  ) => void;
}

export default function ToolDrawer({
  toolsOpen,
  tools,
  queryTools,
  toolDrawerView,
  setToolDrawerView,
  selectedToolId,
  selectedQueryToolId,
  toolDetailMode,
  setToolDetailMode,
  copiedToolDslId,
  toolParameterValues,
  workbook,
  busy,
  setToolsOpen,
  openWorkflowToolDetail,
  openQueryToolDetail,
  requestToolDeletion,
  copyToolDsl,
  prepareToolRun,
  runQueryTool,
  previewTool,
  fieldOptions,
  updateToolParameter
}: ToolDrawerProps) {
  function renderToolParameter(tool: SavedTool, parameter: ToolParameter) {
    return (
      <label key={parameter.id}>
        <span>{parameter.label}</span>
        {parameter.type === "outputWorksheet" ||
        parameter.type === "range" ? (
          <input
            value={
              toolParameterValues[parameter.id] ?? parameter.defaultValue
            }
            placeholder={
              parameter.type === "outputWorksheet"
                ? "输入新的工作表名称"
                : "例如 A1:E812"
            }
            onChange={(event) =>
              updateToolParameter(tool, parameter, event.target.value)
            }
          />
        ) : (
          <select
            value={
              toolParameterValues[parameter.id] ?? parameter.defaultValue
            }
            onChange={(event) =>
              updateToolParameter(tool, parameter, event.target.value)
            }
          >
            {parameter.type === "worksheet"
              ? (workbook?.worksheets ?? []).map((sheet) => (
                  <option key={sheet.name} value={sheet.name}>
                    {sheet.name}
                  </option>
                ))
              : [
                  <option key="field-placeholder" value="">
                    请选择对应字段
                  </option>,
                  ...fieldOptions(tool, parameter).map((field) => (
                    <option key={field} value={field}>
                      {field}
                    </option>
                  ))
                ]}
          </select>
        )}
      </label>
    );
  }

  function formatToolDate(value: string): string {
    const parsed = new Date(value);
    if (Number.isNaN(parsed.getTime())) return "";
    return new Intl.DateTimeFormat("zh-CN", {
      month: "numeric",
      day: "numeric"
    }).format(parsed);
  }

  if (!toolsOpen) return null;
  return (
    <aside className="tool-drawer" aria-label="我的工具">
      <div className="tool-drawer-header">
        <div>
          <strong>
            {toolDrawerView === "library"
              ? "我的工具"
              : toolDrawerView === "detail"
                ? "工具说明"
                : "运行工具"}
          </strong>
          <span>
            {toolDrawerView === "library"
              ? "先选择工具，了解清楚后再运行"
              : toolDrawerView === "detail"
                ? "确认用途、输入和结果"
                : "逐步确认本次运行所需信息"}
          </span>
        </div>
        <button onClick={() => setToolsOpen(false)} aria-label="关闭">
          ×
        </button>
      </div>

      {toolDrawerView === "library" && (
        <div className="tool-library">
          {tools.length === 0 && queryTools.length === 0 ? (
            <div className="tool-empty">
              <i>▦</i>
              <strong>工具箱还是空的</strong>
              <span>执行并验证一个计划后，可以把它保存到这里。</span>
            </div>
          ) : (
            <>
              <div className="tool-library-intro">
                <span>工具箱</span>
                <strong>今天想用哪一个？</strong>
                <p>
                  每个工具都保留自己的处理规则。选择后先看说明，不会立即执行。
                </p>
              </div>
              <div className="tool-card-grid">
                {tools.map((tool) => {
                  const presentation = workflowToolPresentation(tool);
                  return (
                    <article className="tool-card-wrap" key={tool.id}>
                      <button
                        className="tool-card"
                        onClick={() => openWorkflowToolDetail(tool)}
                      >
                        <span
                          className={`tool-card-icon tone-${presentation.tone}`}
                        >
                          {presentation.glyph}
                        </span>
                        <span className="tool-card-copy">
                          <small>{presentation.label}</small>
                          <strong>{tool.name}</strong>
                          <span>{tool.description}</span>
                        </span>
                      </button>
                      <div className="tool-card-footer">
                        <span className="tool-card-meta">
                          {tool.planTemplate.actions.length} 个步骤
                          {formatToolDate(tool.updatedAt)
                            ? ` · ${formatToolDate(tool.updatedAt)} 更新`
                            : ""}
                        </span>
                        <button
                          className="tool-card-delete"
                          aria-label={`删除工具「${tool.name}」`}
                          title={`删除「${tool.name}」`}
                          onClick={() =>
                            requestToolDeletion("workflow", tool)
                          }
                        >
                          删除
                        </button>
                      </div>
                    </article>
                  );
                })}
                {queryTools.map((tool) => (
                  <article className="tool-card-wrap" key={tool.id}>
                    <button
                      className="tool-card"
                      onClick={() => openQueryToolDetail(tool)}
                    >
                      <span className="tool-card-icon tone-blue">查</span>
                      <span className="tool-card-copy">
                        <small>本地查询</small>
                        <strong>{tool.name}</strong>
                        <span>{tool.description}</span>
                      </span>
                    </button>
                    <div className="tool-card-footer">
                      <span className="tool-card-meta">
                        模型调用 0 次
                        {formatToolDate(tool.updatedAt)
                          ? ` · ${formatToolDate(tool.updatedAt)} 更新`
                          : ""}
                      </span>
                      <button
                        className="tool-card-delete"
                        aria-label={`删除工具「${tool.name}」`}
                        title={`删除「${tool.name}」`}
                        onClick={() => requestToolDeletion("query", tool)}
                      >
                        删除
                      </button>
                    </div>
                  </article>
                ))}
              </div>
            </>
          )}
        </div>
      )}

      {toolDrawerView === "detail" &&
        tools
          .filter((tool) => tool.id === selectedToolId)
          .map((tool) => {
            const presentation = workflowToolPresentation(tool);
            return (
              <section className="tool-page tool-detail-page" key={tool.id}>
                <button
                  className="tool-back-button"
                  onClick={() => setToolDrawerView("library")}
                >
                  ← 返回工具箱
                </button>
                <div className="tool-detail-hero">
                  <span
                    className={`tool-card-icon large tone-${presentation.tone}`}
                  >
                    {presentation.glyph}
                  </span>
                  <div>
                    <small>{presentation.label}</small>
                    <strong>{tool.name}</strong>
                    <span>
                      版本 {tool.version} ·{" "}
                      {tool.planTemplate.actions.length} 个处理步骤
                    </span>
                  </div>
                </div>
                <div
                  className="tool-view-switch"
                  role="tablist"
                  aria-label="工具说明视图"
                >
                  <button
                    role="tab"
                    aria-selected={toolDetailMode === "standard"}
                    className={toolDetailMode === "standard" ? "active" : ""}
                    onClick={() => setToolDetailMode("standard")}
                  >
                    普通视图
                  </button>
                  <button
                    role="tab"
                    aria-selected={toolDetailMode === "expert"}
                    className={toolDetailMode === "expert" ? "active" : ""}
                    onClick={() => setToolDetailMode("expert")}
                  >
                    专家视图
                  </button>
                </div>
                {toolDetailMode === "standard" ? (
                  <>
                    <div className="tool-purpose">
                      <strong>它会为你完成什么</strong>
                      <p>{tool.description}</p>
                    </div>
                    <div className="tool-facts">
                      <div>
                        <span>处理方式</span>
                        <strong>{presentation.label}</strong>
                      </div>
                      <div>
                        <span>执行步骤</span>
                        <strong>{tool.planTemplate.actions.length} 步</strong>
                      </div>
                      <div>
                        <span>安全机制</span>
                        <strong>运行前预览</strong>
                      </div>
                    </div>
                    <div className="tool-conversation-note">
                      <span>运行时怎么沟通</span>
                      <p>
                        我会先识别当前工作簿和字段。能够自动匹配的内容不会打扰你，
                        只有发现字段变化时才会请你确认。
                      </p>
                    </div>
                    <details className="tool-steps">
                      <summary>
                        查看完整处理步骤（{tool.planTemplate.actions.length}）
                      </summary>
                      <ol>
                        {tool.planTemplate.actions.map((action, index) => (
                          <li key={`${tool.id}-${index}`}>
                            {actionLabel(action)}
                          </li>
                        ))}
                      </ol>
                    </details>
                  </>
                ) : (
                  <div className="tool-expert-view">
                    <div className="tool-expert-heading">
                      <div>
                        <small>CONTROLLED PLAN</small>
                        <strong>专家脚本</strong>
                        <span>由保存的白名单计划确定性生成</span>
                      </div>
                      <button
                        className="tool-copy-dsl"
                        onClick={() => void copyToolDsl(tool)}
                      >
                        {copiedToolDslId === tool.id ? "已复制" : "复制脚本"}
                      </button>
                    </div>
                    <div className="tool-expert-badges">
                      <span>只读展示</span>
                      <span>运行前预览</span>
                      <span>禁止任意代码</span>
                    </div>
                    <pre className="tool-dsl" tabIndex={0}>
                      <code>{renderToolDsl(tool.planTemplate)}</code>
                    </pre>
                    <p className="tool-expert-note">
                      这里和普通视图是同一份工具。复制脚本不会执行；实际运行仍使用受控
                      AnalysisPlan，并在写入前让你确认预览。
                    </p>
                  </div>
                )}
                <div className="tool-run-actions">
                  <button
                    className="tool-preview-button"
                    disabled={busy || !workbook}
                    onClick={() => void prepareToolRun(tool)}
                  >
                    使用这个工具
                  </button>
                  <span>下一步先确认本次数据来源，不会立即修改工作簿</span>
                </div>
              </section>
            );
          })}

      {toolDrawerView === "detail" &&
        queryTools
          .filter((tool) => tool.id === selectedQueryToolId)
          .map((tool) => (
            <section className="tool-page tool-detail-page" key={tool.id}>
              <button
                className="tool-back-button"
                onClick={() => setToolDrawerView("library")}
              >
                ← 返回工具箱
              </button>
              <div className="tool-detail-hero">
                <span className="tool-card-icon large tone-blue">查</span>
                <div>
                  <small>本地查询</small>
                  <strong>{tool.name}</strong>
                  <span>确定性执行 · 不调用模型</span>
                </div>
              </div>
              <div className="tool-purpose">
                <strong>它会为你查什么</strong>
                <p>{tool.description}</p>
              </div>
              <div className="tool-facts">
                <div>
                  <span>运行位置</span>
                  <strong>本地</strong>
                </div>
                <div>
                  <span>模型调用</span>
                  <strong>0 次</strong>
                </div>
                <div>
                  <span>字段变化</span>
                  <strong>停止并提醒</strong>
                </div>
              </div>
              <div className="tool-conversation-note">
                <span>运行前检查</span>
                <p>
                  如果来源或字段发生变化，工具会停止，不会用错误字段继续计算。
                </p>
              </div>
              <div className="tool-run-actions">
                <button
                  className="tool-preview-button"
                  disabled={busy || !workbook}
                  onClick={() => void runQueryTool(tool)}
                >
                  运行这个查询
                </button>
                <span>查询结果会回到当前对话，不会写入工作簿</span>
              </div>
            </section>
          ))}

      {toolDrawerView === "run" &&
        tools
          .filter((tool) => tool.id === selectedToolId)
          .map((tool) => {
            const primaryParameters = tool.parameters.filter(
              (parameter) =>
                parameter.type === "worksheet" ||
                parameter.type === "outputWorksheet"
            );
            const advancedParameters = tool.parameters.filter(
              (parameter) =>
                parameter.type === "field" ||
                parameter.type === "range"
            );
            const fieldParameters = advancedParameters.filter(
              (parameter) => parameter.type === "field"
            );
            const missingFieldCount = fieldParameters.filter(
              (parameter) =>
                !(toolParameterValues[parameter.id] ?? "").trim()
            ).length;
            return (
              <section className="tool-page tool-run-page" key={tool.id}>
                <button
                  className="tool-back-button"
                  onClick={() => setToolDrawerView("detail")}
                >
                  ← 返回工具说明
                </button>
                <div className="tool-run-heading">
                  <small>正在准备</small>
                  <strong>{tool.name}</strong>
                  <span>
                    我会逐项确认本次运行环境，已匹配的内容保持收起。
                  </span>
                </div>
                <div className="tool-run-guide">
                  <section className="tool-guide-step">
                    <span className="tool-guide-number">1</span>
                    <div className="tool-guide-content">
                      <div className="tool-guide-title">
                        <div>
                          <strong>选择数据来源</strong>
                          <span>这次要处理哪张工作表？</span>
                        </div>
                        <small>需要确认</small>
                      </div>
                      <div className="tool-guide-fields">
                        {primaryParameters.length > 0 ? (
                          primaryParameters.map((parameter) =>
                            renderToolParameter(tool, parameter)
                          )
                        ) : (
                          <span>使用保存时的固定数据来源。</span>
                        )}
                      </div>
                    </div>
                  </section>
                  <section
                    className={`tool-guide-step${
                      missingFieldCount > 0 ? " needs-attention" : ""
                    }`}
                  >
                    <span className="tool-guide-number">2</span>
                    <div className="tool-guide-content">
                      <div className="tool-guide-title">
                        <div>
                          <strong>核对字段</strong>
                          <span>
                            {fieldParameters.length === 0
                              ? "这个工具不需要字段匹配"
                              : missingFieldCount > 0
                              ? `有 ${missingFieldCount} 个字段需要你指定`
                              : `${fieldParameters.length} 个字段已按表头自动匹配`}
                          </span>
                        </div>
                        <small>
                          {missingFieldCount > 0 ? "待处理" : "已完成"}
                        </small>
                      </div>
                      {advancedParameters.length > 0 && (
                        <details
                          className="tool-field-mapping"
                          open={missingFieldCount > 0}
                        >
                          <summary>
                            <span>
                              {missingFieldCount > 0
                                ? "完成字段匹配"
                                : "查看字段映射"}
                            </span>
                            <small>
                              {fieldParameters.length > 0
                                ? `${
                                    fieldParameters.length -
                                    missingFieldCount
                                  }/${fieldParameters.length} 已匹配`
                                : `${advancedParameters.length} 项设置`}
                            </small>
                          </summary>
                          <div>
                            <p>
                              只有字段名称发生变化时才需要调整；数据范围也可以在这里检查。
                            </p>
                            {advancedParameters.map((parameter) =>
                              renderToolParameter(tool, parameter)
                            )}
                          </div>
                        </details>
                      )}
                    </div>
                  </section>
                  <section className="tool-guide-step">
                    <span className="tool-guide-number">3</span>
                    <div className="tool-guide-content">
                      <div className="tool-guide-title">
                        <div>
                          <strong>生成安全预览</strong>
                          <span>
                            下一步只展示将要执行的内容，不会立即修改工作簿。
                          </span>
                        </div>
                        <small>最后确认</small>
                      </div>
                    </div>
                  </section>
                </div>
                <div className="tool-run-actions">
                  <button
                    className="tool-preview-button"
                    disabled={
                      busy || !workbook || missingFieldCount > 0
                    }
                    onClick={() => void previewTool(tool)}
                  >
                    查看执行预览
                  </button>
                  <span>
                    预览确认后才会执行；当前页面不会直接写入 Excel
                  </span>
                </div>
              </section>
            );
          })}
    </aside>
  );
}
