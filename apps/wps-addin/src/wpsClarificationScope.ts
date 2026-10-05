/**
 * 澄清卡的**授权身份指纹**。
 *
 * 为什么不能用工作表索引：索引只是"第几张表"，重排、删除或替换之后，
 * 同一个索引会指向**另一张表**。用户对着"订单"回答的问题，可能被拿去算
 * "客户"的数据 —— 授权边界被绕过，而界面上看不出来。
 *
 * 宿主没有跨会话稳定的工作表 ID 可用（`WpsWorksheet.Index` 只有实机才能验证），
 * 所以这里采取**保守**策略：只要名称、顺序或工作表集合**任何一项**变化，指纹
 * 就变化，待答的澄清随之失效并要求重新授权、读取与预览。代价是"表改名"
 * 会让上一张卡失效 —— 这正是要付的代价，不得为了"改名不断对话"牺牲授权边界。
 *
 * 指纹里还带**工作簿名**：换一本工作簿同样不该沿用上一本的澄清。
 * 之所以不用完整路径：路径含用户名等个人信息，而这个指纹会被持久化进对话历史。
 */

export type WpsAuthorizedSheetIdentity = {
  index: number;
  name: string;
};

/**
 * 授权身份指纹。
 *
 * 刻意**不排序**：顺序本身是身份的一部分 —— 重排会让索引指向另一张表。
 */
export function wpsSheetIdentity(input: {
  workbookName: string;
  sheets: readonly WpsAuthorizedSheetIdentity[];
}): string {
  return JSON.stringify({
    host: "wps",
    workbookName: input.workbookName,
    sheets: input.sheets.map((sheet) => [sheet.index, sheet.name])
  });
}

/**
 * 表头指纹：结构扫描真正读到的**表头行**内容。
 *
 * 它回答的是另一个问题 —— "我这张卡上的问题，还算不算是在问同一份结构"。
 * 只由表头行构成（数据行不进指纹，否则指纹会随用户改数据而抖动，
 * 并且会把单元格值写进对话历史）。
 */
export function wpsHeaderFingerprint(headerRows: readonly (readonly unknown[])[]): string {
  return JSON.stringify(
    headerRows.map((row) =>
      row.map((cell) => (cell == null || cell === "" ? "" : String(cell)))
    )
  );
}

/** 两个指纹是否相同；`undefined` 表示"这一轮没有读取结构"，不参与比较。 */
export function wpsScopeChanged(
  recorded: string | undefined,
  current: string
): boolean {
  if (recorded == null) return false;
  return recorded !== current;
}
