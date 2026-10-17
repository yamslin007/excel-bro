/**
 * 工作目录（folder）数据范围的**纯判据**：Excel 与 WPS 共用同一份声明。
 *
 * 这一层不碰宿主、不碰网络、不碰 localStorage。它只回答三件事：
 *   ① 扫描结果里哪些工作表**可勾选**（= 排除 EB 系统遗留表）；
 *   ② 勾选键集合与「文件 → 工作表」选择之间怎么互相翻译；
 *   ③ 一次刷新之后**还算不算同一件授权**（刷新不是同意）。
 *
 * 为什么要单独一层：上一版这两端各写各的。WPS 因为不敢确认这份口径，
 * 干脆把「工作目录」那张卡禁用掉（`WpsProbePanel.tsx` 里的 disabled 按钮）。
 * 判据只有一份、两端都读它，才不会一端放开、另一端继续猜。
 *
 * 分层：`folderSheetKey` / `isEBSystemSheet` 是 L1 纯函数；
 * `canPreserveFolderAuthorization` 是 L2 语义（读 `folderRootIdentity`）。
 * 与 `scopeAdmission` / `scopeFingerprint` 同属 subpath 导出家族，不进 index.ts。
 */
import { folderRootIdentity } from "./scopeFingerprint";

/**
 * 文件夹扫描结果里本层真正用到的字段。
 *
 * **故意只声明这几个**：app 层的 `FolderCatalog` 结构上满足它，多出来的字段
 * （`totalFiles` / `truncated` / `expiresAt` / `rowCount` …）本层不读，也就不该在场。
 * 这样 WPS 端不必为了用这份判据去 import Excel 侧的 contracts。
 */
export interface FolderScopeFile {
  id: string;
  worksheets: ReadonlyArray<{ name: string }>;
}

export interface FolderScopeCatalog {
  sessionId: string;
  folderPath: string;
  files: ReadonlyArray<FolderScopeFile>;
}

/** 一次「文件 → 工作表名」的勾选投影。与 app 层 `FolderSelection` 结构一致。 */
export interface FolderScopeSelection {
  fileId: string;
  sheets: string[];
}

/**
 * EB 规则系统早期把规则存进这两张隐藏工作表（见 README「EB 函数」）。
 * 现已改为纯内置规则，代码不再读写它们，但用户的旧文件里可能残留。
 * 用精确名单过滤，**不是** `#` 前缀模糊匹配，免得误伤用户自己以 `#` 开头命名的正常表。
 * 服务端 `server/app/models.py` 用同一口径（「不含 `#EB_*` 系统表」）。
 */
const EB_SYSTEM_SHEETS = new Set(["#EB_RULES", "#EB_RULES_BACKUP"]);

export function isEBSystemSheet(name: string): boolean {
  return EB_SYSTEM_SHEETS.has(name);
}

/**
 * 文件夹模式下的工作表标识键：文件 id + 表名合成，供勾选集合记录。
 *
 * 分隔符必须是 NUL：文件 id 是 uuid，表名里不可能出现 NUL，所以
 * 「键 = 哪一个文件的哪张表」这件事没有歧义。
 * 与 `server/app/folder_admission.py:21` 的 `f"{file_id}\x00{name}"` **逐字一致**，
 * 承认汇总才能把服务端返回的每一项对回到本地的每一个勾选上。
 */
export function folderSheetKey(fileId: string, sheetName: string): string {
  return `${fileId}\u0000${sheetName}`;
}

/**
 * 扫描结果里**全部**工作表的键（**不过滤** EB 系统表）。
 *
 * 用于「这个键在新的一次扫描里还有没有对应的工作表」这类**存在性**判断 ——
 * 过滤掉 EB 表会让一个真存在的键被判成"不存在"，从而误判成"目标已消失"。
 */
export function folderSheetKeysOf(catalog: FolderScopeCatalog): string[] {
  return catalog.files.flatMap((file) =>
    file.worksheets.map((sheet) => folderSheetKey(file.id, sheet.name))
  );
}

/** 扫描结果里**可勾选**的工作表键（过滤 EB 系统遗留表）。勾选/全选都只认这一份。 */
export function selectableFolderSheetKeys(catalog: FolderScopeCatalog): string[] {
  return catalog.files.flatMap((file) =>
    file.worksheets
      .filter((sheet) => !isEBSystemSheet(sheet.name))
      .map((sheet) => folderSheetKey(file.id, sheet.name))
  );
}

/**
 * 比较两次文件夹扫描，返回**新出现**且允许用户选择的工作表键。
 *
 * 只报"新增"，不作废原有勾选 —— 新增的表不会自动被勾上（自动勾 = 替用户扩权）。
 */
export function addedFolderSheetKeys(
  previous: FolderScopeCatalog | null,
  current: FolderScopeCatalog
): string[] {
  if (!previous) return [];
  const previousKeys = new Set(selectableFolderSheetKeys(previous));
  return selectableFolderSheetKeys(current).filter((key) => !previousKeys.has(key));
}

/**
 * 按勾选键集合生成「文件 → 工作表」投影。
 *
 * 没勾到任何表的文件被丢掉（服务端 `FolderSelection` 不接受空 sheets）；
 * 顺序跟扫描结果一致，保证「已选 N 个文件 · M 张表」的计数与展示稳定。
 */
export function folderSelectionsFromKeys(
  catalog: FolderScopeCatalog,
  keys: readonly string[]
): FolderScopeSelection[] {
  const selected = new Set(keys);
  return catalog.files
    .map((file) => ({
      fileId: file.id,
      sheets: file.worksheets
        .filter((sheet) => selected.has(folderSheetKey(file.id, sheet.name)))
        .map((sheet) => sheet.name)
    }))
    .filter((selection) => selection.sheets.length > 0);
}

export interface FolderViewAfterFileClear {
  remainingKeys: string[];
  showSelected: boolean;
  activeFileId: string;
}

/**
 * 清空当前文件后修正文件夹选择视图，避免「只看已选」落入无法退出的空结果：
 * 剩下的勾选里没有别的文件时，必须顺手把「只看已选」关掉，否则界面一片空白
 * 而用户找不到关掉它的开关。
 */
export function resolveFolderViewAfterFileClear(
  selectedKeys: readonly string[],
  clearedFileId: string,
  orderedFileIds: readonly string[],
  showSelected: boolean
): FolderViewAfterFileClear {
  const clearedPrefix = folderSheetKey(clearedFileId, "");
  const remainingKeys = selectedKeys.filter(
    (key) => !key.startsWith(clearedPrefix)
  );
  if (!showSelected) {
    return {
      remainingKeys,
      showSelected: false,
      activeFileId: clearedFileId
    };
  }
  const nextSelectedFileId = orderedFileIds.find((fileId) => {
    const prefix = folderSheetKey(fileId, "");
    return remainingKeys.some((key) => key.startsWith(prefix));
  });
  return {
    remainingKeys,
    showSelected: Boolean(nextSelectedFileId),
    activeFileId: nextSelectedFileId ?? clearedFileId
  };
}

/**
 * 一次刷新之后，**还算不算同一件授权**。
 *
 * 刷新不是同意。只有下面四条同时成立才允许保留原来的确认：
 *   ① 两次扫描的根身份一致（`folderRootIdentity` 把 `C:\A\` 与 `c:/a` 折叠成同一个，
 *      POSIX 大小写保留；无绝对根 ⇒ `null` ⇒ 失败关闭）；
 *   ② 会话 id 一致（换会话 = 换一次扫描，不能拿旧会话的确认接着用）；
 *   ③ 至少勾了一个表（勾选归零 = 用户把授权撤了，不是"顺便保留"）；
 *   ④ 每个勾选在新扫描里都能找到对应的工作表（被删/改名的表不算还在）。
 *
 * 任一条不成立都退回未确认。localStorage 里恢复出来的勾选与界面上的草稿
 * **永远**推不出授权 —— 它们只是草稿。
 */
export function canPreserveFolderAuthorization(
  previous: FolderScopeCatalog | null,
  next: FolderScopeCatalog,
  keys: readonly string[]
): boolean {
  const previousRoot = folderRootIdentity(previous?.folderPath);
  const nextRoot = folderRootIdentity(next.folderPath);
  if (previousRoot === null || previousRoot !== nextRoot) return false;
  if (previous?.sessionId !== next.sessionId) return false;
  if (keys.length === 0) return false;
  const present = new Set(folderSheetKeysOf(next));
  return keys.every((key) => present.has(key));
}

/**
 * 两个勾选集合是不是**同一件事**。
 *
 * 先比长度再比排序后连接：重复项会让长度先不等，避免 `["a","a"]` 与 `["a"]`
 * 被判成一致。这是 Excel 侧 `scopeDraftDirty` 一直在用的判据，原样搬到共享层，
 * 口径不变（不改成 Set 比较 —— 那会把重复项判成"没变"）。
 */
export function sameScopeSelection(
  a: readonly string[],
  b: readonly string[]
): boolean {
  return (
    a.length === b.length &&
    [...a].sort().join("\u0000") === [...b].sort().join("\u0000")
  );
}

/** 草稿与已提交授权是否已经不一致（界面据此提示「还有未确认的范围修改」）。 */
export function folderScopeDraftDirty(
  draftKeys: readonly string[],
  committedKeys: readonly string[]
): boolean {
  return !sameScopeSelection(draftKeys, committedKeys);
}

/**
 * 工作目录范围下的**发送判定**。
 *
 * `null` = 放行；返回文案 = 拦住，且这段文案就是用户该看到的原因。
 *
 * 为什么必须有这道闸：工作目录的数据在**服务端**（openpyxl），当前宿主自己的
 * 工作簿与它毫无关系。要是放行，下游就只能拿"当前打开的工作簿"当数据源 ——
 * 那不是用户选的东西，是用一个他没授权的对象冒充他要的数据。
 *
 * 文案里**不许**出现当前工作簿的名字，也不许说"已改用当前工作簿"：
 * 一提工作簿名，用户就会以为"那这次读的是它"。两件事都要说死：
 * 什么都没读、当前工作簿也没被动过。措辞是**结构性**的（只看来源），
 * 不是从宿主文本里读出来的。
 */
export function workDirectorySendRefusal(scope: {
  sourceMode: "workbook" | "folder";
  folderName?: string | null;
  sheetCount?: number | null;
}): string | null {
  if (scope.sourceMode !== "folder") return null;
  const label = scope.folderName
    ? `工作目录「${scope.folderName}」（已选 ${scope.sheetCount ?? 0} 张表）`
    : "工作目录";
  return (
    `这次对话的数据范围是${label}。工作目录的读取通道本轮还没接到这个面板 —— ` +
    "这条消息没有发出，也没有读取任何文件；你当前打开的工作簿没有被读取或修改。" +
    "请先用上方的范围选择器切回「跟随当前」或「选择多个」，再发送。"
  );
}
