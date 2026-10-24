/**
 * 会话内**宿主对象身份**（工作簿 + 工作表）。
 *
 * 为什么指纹不够：`wpsClarificationScope.ts` 的指纹只由**可序列化的名字**构成
 * （工作簿名 + `[index, name]` + 表头），因此有两类它**结构性看不见**的替换：
 *
 *  1. 两个**不同目录的同名工作簿** —— 名字完全一样，换本子指纹不变；
 *  2. **同索引、同名、同表头**的重建工作表 —— 删掉「订单」再新建一张叫「订单」
 *     且表头一致的，「名字」层面无从区分，但它**不是**用户当初授权的那张对象。
 *
 * 第 2 类尤其要紧：它不需要用户做任何"看起来危险"的操作，复制粘贴一张表就够。
 *
 * 指纹（可序列化，进对话历史）与对象身份（不可序列化，只活在内存）回答的是
 * 两个不同问题，**不能互相替代**，所以两个都留：
 *
 *  - 指纹负责**跨会话**能证明的部分（也负责"面板重开后还能看见历史"）；
 *  - 对象身份负责**同一会话内**"到底还是不是那个对象"。
 *
 * 隐私边界：本模块产出的值**一律不进 localStorage / 请求 / 日志 / 错误文案**。
 * `workbookPath` 与 `sheetIds` 只放在内存 `ref` 里，且用于比较的返回值是
 * **布尔**，不是身份本身。
 */
import type { WpsApplication, WpsWorkbook, WpsWorksheet } from "./wpsHost";
import { isWpsSystemSheet, workbookFullName } from "./wpsWorkbookReader";
import { wpsReadWorkbook } from "./wpsResultReviewTarget";

/**
 * 一个已授权工作表的对象身份。
 *
 * `id` 是宿主提供的稳定标识，**只有探测成功才有值** —— 宿主没给就保持 undefined，
 * 由调用方回退到"对象引用 + 结构"判据，而不是凭空造一个假 ID。
 */
export type WpsSheetObjectIdentity = {
  index: number;
  name: string;
  /** 宿主稳定 ID（`CodeName` / `SId`）。**不可序列化进历史**，只在内存比较。 */
  id?: string;
  /** 结构信号：`[起始行, 起始列, 行数, 列数]`。用来感知"同索引重建"后尺寸变化。 */
  shape?: readonly [number, number, number, number];
  /** 宿主对象引用；宿主可以返回新包装对象，引用不同不等于文件发生变化。 */
  ref: WpsWorksheet;
};

export type WpsObjectIdentity = {
  workbookRef: WpsWorkbook;
  workbookName: string;
  /**
   * 完整路径 / 完整文件名。**只在内存里比对，绝不落盘、绝不出站、绝不进文案**。
   * 未保存工作簿为 `null`。
   */
  workbookFullName: string | null;
  workbookPath: string | null;
  sheets: WpsSheetObjectIdentity[];
};

/** 结果三态：能证明相同 / 能证明不同 / **不能证明**（后者一律当作不安全）。 */
export type WpsIdentityCheck = "same" | "different" | "unknown";

function workbookPath(workbook: WpsWorkbook): string | null {
  const path = String(workbook.Path ?? "").trim();
  return path || null;
}

/**
 * 读宿主稳定 ID —— **探测式**，读不到就返回 `undefined`。
 *
 * 候选按"文档里的稳定语义"排序：
 *  - `CodeName`：WPS 基础接口文档写明"**只读**、**改名不影响**、**运行期不可改**"，
 *    是桌面 ET 上唯一被文档确认与表名解耦的标识。
 *  - `SId`：WPS 内核分配的**静态** ID（与动态 `Id` 不同）。
 *
 * 两者都是"文档写了但本项目未实机验证"的能力，所以这里**只做存在性探测**：
 * 拿到了就用（作为第二判据），拿不到不报错、不降级造假。
 * 注意读属性本身可能抛 —— 必须包起来，读一个属性不该把整轮任务打掉。
 */
function readSheetStableId(sheet: WpsWorksheet): string | undefined {
  const candidates: unknown[] = [];
  try {
    candidates.push((sheet as { CodeName?: unknown }).CodeName);
  } catch {
    // 属性不存在或宿主拒绝访问：继续试下一个候选。
  }
  try {
    candidates.push((sheet as { SId?: unknown }).SId);
  } catch {
    // 同上。
  }
  for (const candidate of candidates) {
    if (candidate == null) continue;
    const text = String(candidate).trim();
    if (text) return text;
  }
  return undefined;
}

/**
 * 结构信号：已用区域的起点与尺寸。
 *
 * 它**不是**身份判据（尺寸相同的两张表会被看成一模一样，这正是本模块要解决的
 * 问题），但它是"对象身份不可用"时唯一能感知**同索引重建**的可靠信号之一，
 * 所以一并保存、参与保守失效。
 */
function readSheetShape(
  sheet: WpsWorksheet
): readonly [number, number, number, number] | undefined {
  try {
    const used = sheet.UsedRange;
    if (!used) return undefined;
    const row = Number(used.Row ?? 1);
    const column = Number(used.Column ?? 1);
    const rows = Number(used.Rows?.Count ?? 0);
    const columns = Number(used.Columns?.Count ?? 0);
    if (!Number.isFinite(row) || !Number.isFinite(column)) return undefined;
    if (!Number.isFinite(rows) || !Number.isFinite(columns)) return undefined;
    return [row, column, rows, columns];
  } catch {
    return undefined;
  }
}

/**
 * 捕获"建卡这一刻"的对象身份。
 *
 * 只在**能读到工作簿对象**时成功。授权表序号越界或表不存在会抛
 * `WpsAuthorizationError` —— 这与 `authorizeWpsSheets` 的口径一致：拿不到
 * 就说明授权前提已经不成立，不该静默降级。
 */
export function captureWpsObjectIdentity(
  application: WpsApplication,
  selectedSheetIndexes: readonly number[]
): WpsObjectIdentity {
  // 本次副本不是新目标：身份按原表取，避免"副本一获得焦点就被判成换了文件"。
  const workbook = wpsReadWorkbook(application);
  if (!workbook) {
    throw new Error("WPS 表格中当前没有打开的工作簿。");
  }
  const sheets: WpsSheetObjectIdentity[] = [];
  for (const rawIndex of [...new Set(selectedSheetIndexes.map(Number))]) {
    const sheet = workbook.Worksheets.Item(rawIndex);
    if (!sheet?.Name) {
      throw new Error(`无法识别第 ${rawIndex} 张工作表。`);
    }
    const name = String(sheet.Name);
    if (isWpsSystemSheet(name)) continue;
    sheets.push({
      index: rawIndex,
      name,
      id: readSheetStableId(sheet),
      shape: readSheetShape(sheet),
      ref: sheet
    });
  }
  return {
    workbookRef: workbook,
    workbookName: String(workbook.Name || "未命名工作簿"),
    workbookFullName: workbookFullName(workbook),
    workbookPath: workbookPath(workbook),
    sheets
  };
}

/**
 * 复核"现在还是不是建卡时那个对象"。
 *
 * **放行需要正面证明，失败一律保守**。返回 `"same"` 的唯一途径是对象引用
 * 逐项对上；读不到、对不上、宿主能力不足，全部返回 `"different"` 或 `"unknown"`，
 * 调用方对这两个值同等处理（作废待答澄清）。
 *
 * 判据分三层，**任一层能证明不同就判不同**：
 *
 *  1. **工作簿对象引用**：引用相等提供正面证据；引用不同可能只是新包装对象，
 *     不据此单独断言换了工作簿。
 *  2. **完整路径 / 完整名**：即使引用恰好相同（宿主复用对象），路径不同
 *     也是换了本子。未保存工作簿两边都为空，退回第 1 层，
 *     **不因为"两边都为空"就判相同**。
 *  3. **工作表对象引用**：删掉再重建的同名同表头工作表，即使名称、表头、
 *     尺寸全都一致，**对象引用必然不同** —— 这是唯一能识破它的信号。
 *     权限内再核对稳定 ID，防宿主对象池复用。引用不同且 ID 没有正面变化则判未知，
 *     不拿"名称/表头/ID 相同"反推相同。
 */
export function verifyWpsObjectIdentity(
  application: WpsApplication,
  identity: WpsObjectIdentity
): WpsIdentityCheck {
  let workbook: WpsWorkbook | undefined;
  try {
    workbook = application.ActiveWorkbook;
  } catch {
    return "unknown";
  }
  if (!workbook) {
    // 当前压根没有打开的工作簿 —— 授权对象确实不在了，这是正面证据。
    return "different";
  }

  /*
   * 先取**正面的变化证据**，再看能不能证明相同。
   *
   * 顺序很关键：对象引用与路径是两条独立证据。引用不等**不能**直接判"变了"
   * （宿主可能只是换了包装对象），但完整路径/目录不同就是真的换了文件。
   */
  let currentFullName: string | null;
  let currentPath: string | null;
  try {
    currentFullName = workbookFullName(workbook);
    currentPath = workbookPath(workbook);
  } catch {
    return "unknown";
  }
  if (identity.workbookFullName != null && currentFullName !== identity.workbookFullName) {
    return "different";
  }
  if (identity.workbookPath != null && currentPath !== identity.workbookPath) {
    return "different";
  }

  // 第 3 层：逐张授权表比对。
  let sheetCount: number;
  try {
    sheetCount = Number(workbook.Worksheets?.Count ?? 0);
  } catch {
    return "unknown";
  }
  for (const recorded of identity.sheets) {
    if (!Number.isInteger(recorded.index) || recorded.index < 1 || recorded.index > sheetCount) {
      // 索引越界 = 那张表已经被删/被挪走。
      return "different";
    }
    let sheet: WpsWorksheet | undefined;
    try {
      sheet = workbook.Worksheets.Item(recorded.index);
    } catch {
      return "unknown";
    }
    // 读不到名称 = **证明不了**，不是"变了"。读取失败不许冒充变化证据。
    if (!sheet?.Name) return "unknown";
    if (String(sheet.Name) !== recorded.name) return "different";
    if (isWpsSystemSheet(String(sheet.Name))) return "different";
    const currentId = readSheetStableId(sheet);
    if (recorded.id != null && currentId != null && currentId !== recorded.id) {
      return "different";
    }

    /**
     * **放行的唯一充分条件：对象引用未变。**
     *
     * 这是本模块的核心纪律 —— "能证明相同"与"没发现不同"是两回事。
     * 宿主不保证包装对象稳定，所以引用不等仅为 unknown，必须显式重新授权。
     * 引用相等则可以正面证明。引用不等时其它信号不能反推相同：
     *
     *  - 名称相同、表头相同、尺寸相同 —— 一张重建的复制表全都满足，正是要堵的缺口；
     *  - 稳定 ID 相同 —— 只能说"编号一样"，而编号在"删除后按模板重建"
     *    或"复制整表"的场景里并不保证换新；拿它放行等于把 fail-closed 又拧回 fail-open。
     *
     * 所以这里只在引用相等时放行；引用不等且没有正面变化则判未知（fail-closed），
     * 由调用方要求用户重新选范围、重新预览。稳定 ID 与结构信号则作为
     * **加固**：它们与引用一致地指向"变了"时，同样判不同（多一条证据而已，
     * 不会把"不同"翻成"相同"）。
     */
    if (sheet === recorded.ref) {
      // 正面证明：还是建卡时那个对象。再核对一遍稳定 ID，防宿主对象池复用。
      continue;
    }

    /**
     * 引用不同 ⇒ 判**证明不了**（`"unknown"`），**不是** `"different"`。
     *
     * 宿主只是换了包装、还是真的换了表 —— 在引用不等的前提下**分辨不了**：
     * 名称、表头、尺寸、甚至稳定 ID 都可以在"删除后重建"与"整表复制"里保持一致。
     * 真正换了表的那几种情形（同名异路径、索引越界、表名不符、系统表）**前面已经
     * 用正面证据判掉了**；能走到这里的，只剩"可能没变、只是换了包装"。
     *
     * 把它判成 `"different"` 的后果已经见过：用户只改了一格又撤回，宿主换了包装
     * 对象，界面却告诉他"数据范围已经变化" —— 那是假话，而且他无从纠正。
     * 判 `"unknown"` 由调用方要求重新确认授权，代价同样是"多问一次"，
     * 但**不会撒谎**。具体是 `"changed"` 还是 `"unknown"` 由调用方统一的判定模块
     * 决定（见 `wpsClarificationVerdict`）。
     */
    return "unknown";
  }
  return workbook === identity.workbookRef ? "same" : "unknown";
}

/**
 * 会话身份是否达到"可用于放行"的强度。
 *
 * 未保存工作簿拿不到路径，只能靠对象引用 —— 同会话内够用，**跨重开必然失效**，
 * 这正是期望行为。这个函数存在的意义是把"能不能放行"显式化，而不是靠
 * 调用方各自解释。
 */
export function wpsIdentityIsConclusive(identity: WpsObjectIdentity | null): boolean {
  return identity != null && identity.sheets.length > 0;
}
