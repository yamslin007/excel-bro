// @vitest-environment jsdom
/**
 * A1-E1：WPS 端「启用目标」必须与桥建绑定时发出去的那一组值**逐字相同**。
 *
 * 服务端 `task_bridge.open()` 拿 `host + 工作簿名 + 工作簿身份` 三者全等去匹配这次
 * 启用。任何一项不同，按钮按下去就"什么都不发生"——既没报错也没生效，比报错更难查。
 * 所以这里不复述面板里那几行代码，而是**把桥的实际输出当判据**：
 * `createWpsBridgeLiveContext()` 给什么，`wpsResultRefFaultTarget()` 就得给什么。
 */
import { describe, expect, it } from "vitest";
import type { WpsApplication, WpsRange, WpsWorksheet } from "./wpsHost";
import { createWpsBridgeLiveContext } from "./wpsSdkBridgeHost";
import { wpsResultRefFaultTarget } from "./wpsResultRefFaultTarget";

function makeSheet(name: string, values: unknown[][]): WpsWorksheet {
  return {
    Name: name,
    UsedRange: {
      Row: 1,
      Column: 1,
      Rows: { Count: values.length },
      Columns: { Count: Math.max(0, ...values.map((row) => row.length)) },
      Item: (rowIndex: number, columnIndex = 1) => ({ Row: rowIndex, Column: columnIndex })
    },
    Range: (first: WpsRange, last?: WpsRange) => {
      const startRow = Number(first.Row ?? 1);
      const startColumn = Number(first.Column ?? 1);
      const endRow = Number(last?.Row ?? startRow);
      const endColumn = Number(last?.Column ?? startColumn);
      return {
        Value2: values
          .slice(startRow - 1, endRow)
          .map((row) => row.slice(startColumn - 1, endColumn))
      };
    }
  };
}

const SHEETS = [
  makeSheet("Sheet1", [
    ["类别", "数量"],
    ["甲", 2]
  ]),
  makeSheet("Sheet2", [
    ["类别", "数量"],
    ["乙", 1]
  ])
];

function makeApplication(
  overrides: { name?: string; fullName?: string | null; path?: string | null } = {}
): WpsApplication {
  const workbook = {
    Name: overrides.name ?? "影院.xlsx",
    // 用 `in`：`{fullName: null}` 必须真的是 null（未保存），否则测不出回退分支。
    FullName: "fullName" in overrides ? overrides.fullName : "C:\\Data\\影院.xlsx",
    Path: "path" in overrides ? overrides.path : "C:\\Data",
    Worksheets: {
      Count: SHEETS.length,
      Item: (index: number | string) => SHEETS[Number(index) - 1]
    },
    ActiveSheet: SHEETS[0]
  };
  return {
    PluginStorage: { getItem: () => null, setItem: () => undefined },
    CreateTaskPane: () => ({ ID: 1, Visible: true }),
    GetTaskPane: () => undefined,
    ActiveWorkbook: workbook,
    ActiveSheet: SHEETS[0],
    Worksheets: workbook.Worksheets
  } as unknown as WpsApplication;
}

describe("启用目标 = 桥发出去的那一组值", () => {
  it("已保存工作簿：名字与身份都与桥逐字相同", () => {
    const application = makeApplication();
    const live = createWpsBridgeLiveContext(application, [1, 2]);
    const target = wpsResultRefFaultTarget(application);

    expect(target).not.toBeNull();
    expect(target!.hostId).toBe("wps-et");
    expect(target!.workbookName).toBe(live.workbookName);
    expect(target!.workbookIdentity).toBe(live.workbookIdentity);
    expect(target!.workbookIdentity).toBe("wps-path:c:\\data\\影院.xlsx");
  });

  it("名字取 Workbook.Name，不是磁盘路径的 basename", () => {
    // 桥发 `Workbook.Name`。若这里改用路径 basename（"影院.xlsx" 带扩展名、
    // 或大小写不同的写法），绑定永远匹配不上 —— 这条用例就是那个判别点。
    const application = makeApplication({ name: "影院", fullName: "C:\\Data\\影院.xlsx" });
    const target = wpsResultRefFaultTarget(application);
    expect(target!.workbookName).toBe("影院");
    expect(target!.workbookName).toBe(createWpsBridgeLiveContext(application, [1]).workbookName);
    expect(target!.workbookName).not.toBe("影院.xlsx");
  });

  it("只有工作簿名、没有磁盘路径时：身份走桥的会话令牌，仍与桥一致", () => {
    // `FullName` 为空但 `Path` 还在时，读取层会自己拼出完整路径（上一个用例覆盖）。
    // 真正的"未保存"是两者都没有：此时桥发的是会话内对象令牌，面板必须发同一串。
    const application = makeApplication({ fullName: null, path: null });
    const live = createWpsBridgeLiveContext(application, [1]);
    const target = wpsResultRefFaultTarget(application);
    expect(target!.workbookName).toBe(live.workbookName);
    expect(target!.workbookIdentity).toBe(live.workbookIdentity);
    expect(target!.workbookIdentity.startsWith("wps-session:")).toBe(true);
  });

  it("FullName 缺失但 Path 还在：读取层拼出的路径就是身份，与桥一致", () => {
    const application = makeApplication({ fullName: null });
    const live = createWpsBridgeLiveContext(application, [1]);
    const target = wpsResultRefFaultTarget(application);
    expect(target!.workbookIdentity).toBe(live.workbookIdentity);
    expect(target!.workbookIdentity).toBe("wps-path:c:\\data\\影院.xlsx");
  });

  it("读不到现场时不给目标，而不是给一个空目标", () => {
    // 宿主抛错（没有打开的工作簿）→ 入口整体不可用；绝不退化成"全局生效"。
    const broken = {
      ...makeApplication(),
      ActiveWorkbook: null
    } as unknown as WpsApplication;
    expect(wpsResultRefFaultTarget(broken)).toBeNull();
  });

  it("同一本簿连问两次得到同一个身份（不是每次发新令牌）", () => {
    const application = makeApplication({ fullName: null });
    expect(wpsResultRefFaultTarget(application)!.workbookIdentity).toBe(
      wpsResultRefFaultTarget(application)!.workbookIdentity
    );
  });
});
