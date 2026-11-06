/**
 * 本地查询结果卡。
 *
 * 服务端只下发**引用与结构**（没有值）；真实值由本机回环端点按 ``resultRef``
 * 取回后在这里渲染。这样值既不进模型提示，也不进聊天历史与持久账本。
 *
 * 取不到值时必须**如实**呈现：结果已过保留期、或本端已不再持有该任务的授权
 * （换簿/重授权）都会走到「不可用」分支，而不是渲染一张空表冒充成功。
 */

import { useEffect, useRef, useState } from "react";

import {
  describeIncompleteReasons,
  describeTruncation
} from "@excel-bro/core";
import type { LocalResultCard as LocalResultPointer } from "../contracts";
import type { SdkBridgeResultResponse } from "../contracts";

export type LocalResultState =
  | { status: "loading" }
  | { status: "ready"; payload: SdkBridgeResultResponse }
  | { status: "unavailable" };

export function LocalResultCard({
  card,
  load
}: {
  card: LocalResultPointer;
  /** 按引用取回值；失败（过期/授权不符/断线）返回 null。 */
  load: (resultRef: string) => Promise<SdkBridgeResultResponse | null>;
}) {
  const [state, setState] = useState<LocalResultState>({ status: "loading" });
  // ``load`` 通常是内联闭包，每次渲染都会换身份；把它放进 ref，
  // 让 effect 只依赖 resultRef，否则会陷入"取一次 → 重渲染 → 再取"的循环。
  const loadRef = useRef(load);
  loadRef.current = load;

  useEffect(() => {
    let active = true;
    setState({ status: "loading" });
    void loadRef
      .current(card.resultRef)
      .then((payload) => {
        if (!active) return;
        setState(payload ? { status: "ready", payload } : { status: "unavailable" });
      })
      .catch(() => {
        if (active) setState({ status: "unavailable" });
      });
    return () => {
      active = false;
    };
  }, [card.resultRef]);

  /*
   * 两句提示都由**事实**推出来：
   * - 不完整：`incomplete` 里有什么就说什么（哪张表缺哪个字段 / 哪张表没有数据行）；
   * - 截断：`truncation` 的计数，只在 limit 真的切掉了行时出现。
   * 没有任何原因却 `complete===false` 属于数据不一致 —— 那就如实说"原因未记录"，
   * 不替它编一个（尤其不编"扫描已达上限"）。
   */
  const incompleteText = (() => {
    const described = describeIncompleteReasons(card.incomplete ?? []);
    if (described) return `结果不完整：${described}`;
    if (card.complete === false) {
      return "结果不完整，但这次没有记录到具体原因。";
    }
    return null;
  })();
  const truncationText = card.truncation
    ? `列表已截断：${describeTruncation(card.truncation)}`
    : null;

  return (
    <section className="local-result-card" aria-label="本地查询结果">
      <div className="local-result-card-head">
        <strong>{card.title}</strong>
        <span className="local-result-card-shape">
          {card.rowCount} 行 × {card.columnCount} 列
        </span>
      </div>
      {card.sourceSheets.length > 0 && (
        <p className="local-result-card-source">
          来源：{card.sourceSheets.join("、")}
        </p>
      )}
      {/**
       * 提示语**按真实原因**写成，不按一个布尔猜。
       *
       * 原来这里只有一句「结果不完整（扫描已达上限）」：任何 `complete=false`
       * 都这么写。2026-11-01 真机里三张表缺「列2」也被说成扫描达上限，用户自然
       * 以为是读表读崩了。现在两条事实分开：内容不完整（有原因）与展示截断
       * （只有 limit 生效），各自只说自己。
       */}
      {incompleteText && (
        <p className="local-result-card-partial">{incompleteText}</p>
      )}
      {truncationText && (
        <p className="local-result-card-truncated">{truncationText}</p>
      )}
      {state.status === "loading" && (
        <p className="local-result-card-status" role="status">
          正在取回本机结果…
        </p>
      )}
      {state.status === "unavailable" && (
        <p className="local-result-card-status local-result-card-unavailable" role="status">
          结果已不可用（可能已过保留期，或本端已不再持有该任务的授权）。请重新发起查询。
        </p>
      )}
      {state.status === "ready" && <LocalResultTable payload={state.payload} />}
    </section>
  );
}

function LocalResultTable({ payload }: { payload: SdkBridgeResultResponse }) {
  /*
   * 告警必须显式渲染。
   *
   * "有 N 条未匹配""检测到重复键，已按 first 取值"这类信息决定了用户能不能
   * 信任这份结果 —— 它们只存在于 `warnings` 里。界面不渲染，用户就只看到一张
   * 看起来正常的表，误以为数据完整。空数组时不渲染空区块，避免制造噪声。
   */
  const warnings = payload.warnings.filter(
    (warning): warning is string => typeof warning === "string" && warning.trim().length > 0
  );
  return (
    <div className="local-result-card-table-wrap">
      <table className="local-result-card-table">
        <thead>
          <tr>
            {payload.headers.map((header, index) => (
              <th key={`h-${index}`}>{header === null ? "" : String(header)}</th>
            ))}
          </tr>
        </thead>
        <tbody>
          {payload.rows.map((row, rowIndex) => (
            <tr key={`r-${rowIndex}`}>
              {row.map((cell, cellIndex) => (
                <td key={`c-${rowIndex}-${cellIndex}`}>
                  {cell === null ? "" : String(cell)}
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
      <p className="local-result-card-calculation">{payload.calculation}</p>
      {warnings.length > 0 && (
        <ul className="local-result-card-warnings">
          {warnings.map((warning, index) => (
            <li key={`w-${index}`}>{warning}</li>
          ))}
        </ul>
      )}
    </div>
  );
}

export default LocalResultCard;
