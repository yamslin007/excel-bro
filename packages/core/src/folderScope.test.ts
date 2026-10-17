import { describe, expect, it } from "vitest";
import {
  addedFolderSheetKeys,
  canPreserveFolderAuthorization,
  folderScopeDraftDirty,
  folderSelectionsFromKeys,
  folderSheetKey,
  folderSheetKeysOf,
  isEBSystemSheet,
  resolveFolderViewAfterFileClear,
  sameScopeSelection,
  selectableFolderSheetKeys,
  workDirectorySendRefusal,
  type FolderScopeCatalog
} from "./folderScope";

function catalog(overrides: Partial<FolderScopeCatalog> = {}): FolderScopeCatalog {
  return {
    sessionId: "session-1",
    folderPath: "C:/Data",
    files: [
      {
        id: "file-1",
        worksheets: [{ name: "Data" }, { name: "Summary" }]
      },
      {
        id: "file-2",
        worksheets: [{ name: "明细" }]
      }
    ],
    ...overrides
  };
}

describe("folderScope 标识键", () => {
  it("键用 NUL 连接，且与服务端 folder_admission.py 的拼法逐字一致", () => {
    expect(folderSheetKey("file-1", "Data")).toBe("file-1\u0000Data");
    expect(folderSheetKey("file-1", "Data")).toBe("file-1\x00Data");
    // 正向锚点：不同文件 / 不同表名都不能撞成同一个键。
    expect(folderSheetKey("file-1", "Data")).not.toBe(folderSheetKey("file-2", "Data"));
    expect(folderSheetKey("file-1", "Data")).not.toBe(folderSheetKey("file-1", "Summary"));
  });
});

describe("folderScope 可勾选与存在性是两个判据", () => {
  it("可勾选名单只排除 EB 系统遗留表，不误伤用户自己的 # 前缀表", () => {
    const scanned = catalog({
      files: [
        {
          id: "file-1",
          worksheets: [
            { name: "Data" },
            { name: "#EB_RULES" },
            { name: "#EB_RULES_BACKUP" },
            { name: "#报表" }
          ]
        }
      ]
    });
    expect(selectableFolderSheetKeys(scanned)).toEqual([
      "file-1\u0000Data",
      "file-1\u0000#报表"
    ]);
    expect(isEBSystemSheet("#EB_RULES")).toBe(true);
    expect(isEBSystemSheet("#EB_RULES_BACKUP")).toBe(true);
    // 负向断言的锚点：`#` 前缀本身不是判据。
    expect(isEBSystemSheet("#报表")).toBe(false);
    expect(isEBSystemSheet("#EB_RULES_V2")).toBe(false);
    expect(isEBSystemSheet("#eb_rules")).toBe(false);
  });

  it("存在性名单不过滤 EB 表 —— 判据不同，所以读的不是同一个集合", () => {
    const scanned = catalog({
      files: [{ id: "file-1", worksheets: [{ name: "Data" }, { name: "#EB_RULES" }] }]
    });
    expect(folderSheetKeysOf(scanned)).toEqual([
      "file-1\u0000Data",
      "file-1\u0000#EB_RULES"
    ]);
    expect(selectableFolderSheetKeys(scanned)).toEqual(["file-1\u0000Data"]);
    // 唯一判别点：一个指向 EB 表的键，在新扫描里**确实还在**，
    // 因此不得判成「目标已消失」。若存在性名单也过滤 EB 表，这里会掉到 false。
    expect(
      canPreserveFolderAuthorization(scanned, scanned, ["file-1\u0000#EB_RULES"])
    ).toBe(true);
  });
});

describe("folderScope 刷新只报新增、不作废", () => {
  it("只报新增的可勾选工作表，且新增 EB 表不算新增", () => {
    const before = catalog();
    const after = catalog({
      files: [
        { id: "file-1", worksheets: [{ name: "Data" }, { name: "Summary" }, { name: "新增" }, { name: "#EB_RULES" }] },
        { id: "file-2", worksheets: [{ name: "明细" }] }
      ]
    });
    expect(addedFolderSheetKeys(before, after)).toEqual(["file-1\u0000新增"]);
    // 没有上一次扫描就没有"新增"这一说。
    expect(addedFolderSheetKeys(null, after)).toEqual([]);
    // 正向锚点：同一份扫描自己跟自己比，结果必须是空。
    expect(addedFolderSheetKeys(after, after)).toEqual([]);
  });
});

describe("folderScope 勾选键 → 文件/表 投影", () => {
  it("只投影勾到的表，丢掉一张表都没勾的文件，顺序跟扫描一致", () => {
    const scanned = catalog();
    expect(
      folderSelectionsFromKeys(scanned, ["file-2\u0000明细", "file-1\u0000Summary"])
    ).toEqual([
      { fileId: "file-1", sheets: ["Summary"] },
      { fileId: "file-2", sheets: ["明细"] }
    ]);
    // 一个都没勾 ⇒ 空投影，而不是「每个文件都空 sheets」。
    expect(folderSelectionsFromKeys(scanned, [])).toEqual([]);
    // 勾了不存在的键 ⇒ 也推不出任何一项。
    expect(folderSelectionsFromKeys(scanned, ["file-9\u0000Ghost"])).toEqual([]);
  });

  it("清空当前文件后视图不会卡死在空的「只看已选」", () => {
    expect(
      resolveFolderViewAfterFileClear(
        ["file-1\u0000Data", "file-2\u0000明细"],
        "file-1",
        ["file-1", "file-2"],
        true
      )
    ).toEqual({ remainingKeys: ["file-2\u0000明细"], showSelected: true, activeFileId: "file-2" });
    // 清掉唯一一个已选文件 ⇒ 必须退出「只看已选」。
    expect(
      resolveFolderViewAfterFileClear(["file-1\u0000Data"], "file-1", ["file-1", "file-2"], true)
    ).toEqual({ remainingKeys: [], showSelected: false, activeFileId: "file-1" });
    // 本来就不是「只看已选」⇒ 视图模式不动，只留当前文件。
    expect(
      resolveFolderViewAfterFileClear(["file-1\u0000Data"], "file-1", ["file-1", "file-2"], false)
    ).toEqual({ remainingKeys: [], showSelected: false, activeFileId: "file-1" });
  });
});

describe("folderScope 刷新不是同意", () => {
  it("同根（含路径拼写变化）+ 同会话 + 目标都在 ⇒ 可以保留原确认", () => {
    const before = catalog({ folderPath: "C:\\Data\\" });
    expect(
      canPreserveFolderAuthorization(before, catalog(), ["file-1\u0000Data"])
    ).toBe(true);
  });

  it("换目录 / 换会话 / 勾选归零 / 目标消失，逐条都不保留确认", () => {
    const before = catalog();
    const keys = ["file-1\u0000Data"];
    expect(
      canPreserveFolderAuthorization(before, catalog({ folderPath: "D:/Data" }), keys)
    ).toBe(false);
    expect(
      canPreserveFolderAuthorization(before, catalog({ sessionId: "session-2" }), keys)
    ).toBe(false);
    expect(canPreserveFolderAuthorization(before, catalog(), [])).toBe(false);
    const missing = catalog({ files: [{ id: "file-1", worksheets: [{ name: "Summary" }] }] });
    expect(canPreserveFolderAuthorization(before, missing, keys)).toBe(false);
  });

  it("没有上一次扫描 / 没有绝对根 ⇒ 失败关闭", () => {
    expect(canPreserveFolderAuthorization(null, catalog(), ["file-1\u0000Data"])).toBe(false);
    expect(
      canPreserveFolderAuthorization(catalog({ folderPath: "Data" }), catalog({ folderPath: "Data" }), ["file-1\u0000Data"])
    ).toBe(false);
    expect(
      canPreserveFolderAuthorization(
        catalog({ folderPath: "" }),
        catalog({ folderPath: "" }),
        ["file-1\u0000Data"]
      )
    ).toBe(false);
  });

  it("POSIX 根保留大小写 ⇒ 两个只差大小写的目录不是同一个根", () => {
    expect(
      canPreserveFolderAuthorization(
        catalog({ folderPath: "/home/Data" }),
        catalog({ folderPath: "/home/data" }),
        ["file-1\u0000Data"]
      )
    ).toBe(false);
  });
});

describe("folderScope 草稿判据", () => {
  it("顺序不算变化，重复项与增删算变化", () => {
    expect(sameScopeSelection(["a", "b"], ["b", "a"])).toBe(true);
    expect(sameScopeSelection(["a", "a"], ["a"])).toBe(false);
    expect(sameScopeSelection(["a"], ["a", "b"])).toBe(false);
    expect(folderScopeDraftDirty(["a", "b"], ["b", "a"])).toBe(false);
    expect(folderScopeDraftDirty(["a"], ["a", "b"])).toBe(true);
    expect(folderScopeDraftDirty([], [])).toBe(false);
  });
});

describe("folderScope 发送闸：工作目录下不偷用当前工作簿", () => {
  it("工作簿来源一律放行", () => {
    expect(workDirectorySendRefusal({ sourceMode: "workbook" })).toBeNull();
    // 正向锚点：带上无关的工作目录字段也照样放行 —— 判据只看来源。
    expect(
      workDirectorySendRefusal({ sourceMode: "workbook", folderName: "数据", sheetCount: 9 })
    ).toBeNull();
  });

  it("工作目录来源一律拦住，并说清「什么都没读、当前工作簿也没被动过」", () => {
    const refusal = workDirectorySendRefusal({
      sourceMode: "folder",
      folderName: "数据",
      sheetCount: 3
    });
    expect(refusal).not.toBeNull();
    expect(refusal).toContain("工作目录「数据」（已选 3 张表）");
    expect(refusal).toContain("没有发出");
    expect(refusal).toContain("没有读取任何文件");
    expect(refusal).toContain("当前打开的工作簿没有被读取或修改");
    // 负向：不许出现"已改用当前工作簿"这类冒充说法。
    expect(refusal).not.toContain("已改用");
    expect(refusal).not.toContain("代替");
  });

  it("还没确认目录名时也拦住，且不编造目录名", () => {
    const refusal = workDirectorySendRefusal({ sourceMode: "folder" });
    expect(refusal).not.toBeNull();
    expect(refusal).toContain("这次对话的数据范围是工作目录。");
    // 不编造目录名：目录名位置的「」不该出现（后面「跟随当前」那种引号是文案本身）。
    expect(refusal).not.toContain("工作目录「");
  });
});
