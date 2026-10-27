import { resultRefFaultTarget, type ResultRefFaultTarget } from "@excel-bro/core/resultRefFault";
import type { WpsApplication } from "./wpsHost";
import { listWpsWorksheets, wpsWorkbookBridgeIdentity } from "./wpsWorkbookReader";

/**
 * A1-E1「结果引用过期」验收入口的**启用目标**（WPS 端）。
 *
 * 这个函数的唯一职责是把"当前这本工作簿"翻译成服务端认得的两个字段，而它必须
 * 与 `createWpsBridgeLiveContext`（桥建绑定时真正发出去的那一组值）**逐字相同**：
 * 服务端 `task_bridge.open()` 里 `dev_result_ref_fault.bind_task(...)` 是按
 * `host + 工作簿名 + 工作簿身份` 三者全等匹配的。任何一项不同，这次启用就永远
 * 绑不上任务 —— 表现是"点了按钮什么都没发生"，比报错更难查。
 *
 * 所以两处读的是同一对同源取值，不在这里另行拼装：
 *
 * - 名字 = `listWpsWorksheets().name`（宿主的 `Workbook.Name`）。**不要**用磁盘路径的
 *   basename 顶替：桥发的是 `Workbook.Name`，两者可能差一个扩展名；
 * - 身份 = `wpsWorkbookBridgeIdentity()`，即桥自己用的那串（已保存 = `wps-path:`
 *   加小写完整路径，未保存 = 会话内对象令牌）。
 *
 * 读不到现场（没有打开的工作簿、宿主抛错）时返回 `null`：宁可这一次入口不可用，
 * 也不退化成"一个对所有工作簿都生效的全局开关"。
 */
export function wpsResultRefFaultTarget(
  application: WpsApplication
): ResultRefFaultTarget | null {
  try {
    const { name } = listWpsWorksheets(application);
    return resultRefFaultTarget(
      "wps-et",
      name,
      wpsWorkbookBridgeIdentity(application)
    );
  } catch {
    return null;
  }
}
