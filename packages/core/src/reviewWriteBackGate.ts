/**
 * 「这次写回请求还对着同一份现场吗」—— 写回入口的**异步**复核判据。
 *
 * 写回入口先做同步闸门（勾选了没有 / 计划与副本在不在场 / 忙不忙），再发起
 * **异步**授权检查。等它回来时，用户完全可能已经点了「放弃变更」，甚至放弃之后
 * 又生成了一份**新副本**并勾了新的确认框。
 *
 * 「当前有审阅」只说明"有"，不说明"还是那一次"。只凭它放行，旧请求就会把
 * **新副本**的计划、拿着**新副本**的确认状态写进原表 —— 用户核对过的那一份
 * 从来没被写，没核对过的那一份反而被写了。
 *
 * 所以请求要在**发起点**记下一份身份三元组，授权检查回来之后逐项比对：
 * 同一次审阅、同一计划版本、同一原工作簿，三项全同才继续。
 *
 * 三项各管一件事，**不能互相顶替**：
 *  - `review` 防的是"换了副本"（放弃后又生成新副本）；
 *  - `plan` 防的是"换了计划"（重新预览、编辑后重新校验）；
 *  - `workbookIdentity` 防的是"换了原工作簿"（切文件、另存为另一个路径）。
 */
export type ReviewWriteBackTarget = {
  /** 同一次审阅：每次生成 / 放弃副本都会换新值。 */
  review: unknown;
  /** 同一计划版本：计划被替换 / 重新预览 / 编辑后重新校验都会换新值。 */
  plan: unknown;
  /** 同一原工作簿：`url:` 或本机路径形式的身份；读不到时为空串。 */
  workbookIdentity: string;
};

/**
 * 请求发起时记下的现场，与异步检查回来时的现场**是否已经不是同一份**。
 *
 * 任一侧缺席（`null`）一律算"已变化"：拿不到身份就不该继续写回 —— 这是
 * fail-closed，不是"没变化"。
 *
 * 空串表示"宿主没给出工作簿身份"。两侧都读不到时这一项不构成判别力，所以写回侧
 * 另有独立校验兜着（WPS 是 `wpsPlanExecutor.workbookIdentity`：身份缺失或不是本地
 * `.xlsx` 路径就直接拒绝写入），能走到这里时它一定非空。这里不额外加"空即拦"，
 * 免得把一个读得到、写不了的身份问题伪装成写回失败。
 */
export function reviewWriteBackTargetChanged(
  issued: ReviewWriteBackTarget | null,
  current: ReviewWriteBackTarget | null
): boolean {
  if (!issued || !current) return true;
  return issued.review !== current.review ||
    issued.plan !== current.plan ||
    issued.workbookIdentity !== current.workbookIdentity;
}
