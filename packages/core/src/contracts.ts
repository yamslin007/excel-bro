/**
 * L0 契约层：纯类型,零运行时依赖。
 *
 * 阶段 2 迁入 `CellValue`;阶段 3 追加 `WorkbookLocator` 与工具 DTO。
 * `apps/excel-addin/src/contracts.ts` 的其余部分按方案 §5.2 拆分
 * （纯 DTO 留 L0、运行时校验归 L2）后再迁入,本文件届时逐步扩充。
 */

/** 单元格标量值。两端读写与比对的最小公共类型。 */
export type CellValue = string | number | boolean | null;

/**
 * 工作簿身份的共享层内部表示（方案 §2.3 第 1 条）。
 *
 * 三条通道的身份语义本就不同：Office 是内存态文档 URL、WPS 是已存盘全路径、
 * 文件夹通道是服务端分配的文件标识。用裸字符串会让共享层不得不去猜"这是哪种身份",
 * 判别式把这件事变成类型问题。
 *
 * **只存在于 packages/core 内部。** 出网协议字段 `workbookIdentity: string`
 * 保持不变（`models.py:2017` 的校验强制 `url:` 前缀、`contracts.ts` 同为 string）,
 * 由 `locator.ts` 在唯一一处折叠。理由见方案 §2.4。
 */
export type WorkbookLocator =
  | { kind: "office-url"; value: string }
  | { kind: "wps-path"; value: string }
  | { kind: "folder-relative"; value: string };

/** 已授权工作表在共享层的引用方式。WPS 用 1 基序号定位,故序号与表名都需带上。 */
export interface AuthorizedSheetRef {
  /** WPS 的 `Sheets(index)` 是 1 基;Office 通道该字段为 undefined。 */
  index?: number;
  name: string;
}
