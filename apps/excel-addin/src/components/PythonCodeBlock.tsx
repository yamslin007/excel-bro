import { useEffect, useId, useState } from "react";
import {
  pythonCodeCopyText,
  pythonCodeView,
  DEFAULT_PYTHON_CODE_LIMITS,
  PYTHON_CODE_BODY_LOADING_NOTE,
  PYTHON_CODE_BODY_UNAVAILABLE_NOTE,
  type PythonCodeArtifact,
  type PythonCodeLimits
} from "@excel-bro/core/pythonCodeArtifact";
import { fetchPythonCodeBody } from "../api";

interface PythonCodeBlockProps {
  artifact: PythonCodeArtifact;
  /**
   * 这段代码所属的任务身份（服务端 `turnId`）。
   *
   * 取回正文必须带任务身份：只有 `codeId` 就能取的话，任何能读到一个 id 的
   * 进程都能读到别的任务的代码。拿不到任务身份就不发请求，直说正文没保留。
   */
  taskId?: string | null;
  limits?: PythonCodeLimits;
  /** 窄窗格里默认折叠（默认按行数判断）。 */
  defaultExpanded?: boolean;
  /** 测试与宿主可注入（默认走本机回环端点）。 */
  loadBody?: (request: { taskId: string; codeId: string }) => Promise<
    { status: "ready"; code: string; chars: number } | { status: "unavailable"; reason: string }
  >;
}

type BodyState =
  | { kind: "absent" }
  | { kind: "loading" }
  | { kind: "ready"; code: string }
  | { kind: "unavailable" };

/**
 * Python 代码内容块（Excel / WPS 共用）。
 *
 * 四条硬性约束，改动前先读：
 * 1. **只读**。这里**不提供**任何"运行"入口：本仓库只有"生成白名单计划 → 预览 →
 *    用户确认 → 执行"这一条写入路径，代码块不是第二条。
 * 2. **安全转义**。代码一律作为文本节点渲染，不用 `dangerouslySetInnerHTML`。
 * 3. **不假装**。"已运行"只在工件记录的 `executedVersion` 与当前版本一致时出现；
 *    正文取不回来就直说没保留；未完成的草稿不提供复制全文。
 * 4. **正文按需取回**。代码不进持久账本，事件里只有状态与规模，因此本组件在
 *    挂载时按 `codeId` 向本机回环端点取一次正文（失败即定为"没保留"）。
 */
export function PythonCodeBlock({
  artifact,
  taskId = null,
  limits = DEFAULT_PYTHON_CODE_LIMITS,
  defaultExpanded = false,
  loadBody = fetchPythonCodeBody
}: PythonCodeBlockProps) {
  const [expanded, setExpanded] = useState(defaultExpanded);
  const [copied, setCopied] = useState(false);
  const [loaded, setLoaded] = useState<BodyState>(
    artifact.code.length > 0 ? { kind: "ready", code: artifact.code } : { kind: "absent" }
  );
  const bodyId = useId();

  // 正文不在手上时按 codeId 取回一次。换了工件（codeId 变了）就重新判断：
  // 不重置会把上一段代码显示成这一段（"已运行"就会指错版本）。
  const needsBody = artifact.code.length === 0;
  useEffect(() => {
    if (!needsBody) {
      setLoaded({ kind: "ready", code: artifact.code });
      return;
    }
    if (!taskId) {
      // 没有任务身份 ⇒ 根本不允许发这个请求。照实说没保留。
      setLoaded({ kind: "unavailable" });
      return;
    }
    let alive = true;
    setLoaded({ kind: "loading" });
    void loadBody({ taskId, codeId: artifact.codeId })
      .then((result) => {
        if (!alive) return;
        setLoaded(
          result.status === "ready"
            ? { kind: "ready", code: result.code }
            : { kind: "unavailable" }
        );
      })
      .catch(() => {
        if (alive) setLoaded({ kind: "unavailable" });
      });
    return () => {
      alive = false;
    };
  }, [artifact.codeId, artifact.code, needsBody, taskId, loadBody]);

  // 取回的正文只在本组件里合并成"可展示的工件"：状态机与事件归并那条链不受影响，
  // 也不存在"界面有一份、状态里有一份"的双真相。
  const effective: PythonCodeArtifact =
    artifact.code.length > 0 || loaded.kind !== "ready"
      ? artifact
      : { ...artifact, code: loaded.code, codeRetained: true };
  const view = pythonCodeView(effective, limits);
  if (!view) return null;

  const copyAll = async () => {
    const text = pythonCodeCopyText(effective, limits);
    if (!text) return;
    try {
      await navigator.clipboard?.writeText(text);
      setCopied(true);
      window.setTimeout(() => setCopied(false), 1500);
    } catch {
      // 复制失败不改状态：宁可让用户再点一次，也不谎报"已复制"。
      setCopied(false);
    }
  };

  const showBody = view.hasCode && (!view.collapsible || expanded);
  const note =
    view.note ??
    (loaded.kind === "loading"
      ? PYTHON_CODE_BODY_LOADING_NOTE
      : loaded.kind === "unavailable"
        ? PYTHON_CODE_BODY_UNAVAILABLE_NOTE
        : undefined);

  return (
    <section
      className="python-code-block"
      data-stage={artifact.stage}
      data-draft={view.draft ? "true" : "false"}
      data-body={loaded.kind}
      aria-label={`Python 代码 · ${view.label}`}
    >
      <div className="python-code-header">
        <span className="python-code-lang">Python</span>
        <span className="python-code-state" data-stage={artifact.stage}>
          {view.label}
        </span>
        {view.executedCodeMatches && (
          <span className="python-code-badge" title="界面上的这一版就是实际执行的那一版">
            已运行
          </span>
        )}
        <div className="python-code-actions">
          {view.hasCode && view.collapsible && (
            <button
              type="button"
              className="python-code-toggle"
              aria-expanded={expanded}
              aria-controls={bodyId}
              onClick={() => setExpanded((current) => !current)}
            >
              {expanded ? "收起代码" : "查看代码"}
            </button>
          )}
          {view.canCopyAll && (
            <button type="button" className="python-code-copy" onClick={() => void copyAll()}>
              {copied ? "已复制" : "复制"}
            </button>
          )}
        </div>
      </div>
      {note && <p className="python-code-note">{note}</p>}
      {view.hasCode && (
        <pre
          id={bodyId}
          className="python-code-body"
          data-truncated={view.truncated ? "true" : "false"}
        >
          <code>{showBody ? view.body : view.preview}</code>
        </pre>
      )}
      {!showBody && view.hasCode && view.collapsible && (
        <p className="python-code-more">
          还有 {Math.max(0, view.body.split("\n").length - view.preview.split("\n").length)} 行未展开
        </p>
      )}
    </section>
  );
}

export default PythonCodeBlock;
