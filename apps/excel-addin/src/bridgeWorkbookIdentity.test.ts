// @vitest-environment jsdom
/**
 * 未保存同名工作簿：**读路径**的桥身份。
 *
 * 缺陷形状（A1-E3 矩阵第 2 行）：Office.js 在文件未保存时 `document.url` 为空，
 * 身份退回 `name:<名字>`；两个都叫「工作簿1」的未保存簿于是在桥的匹配
 * （`hostId + workbookName + workbookIdentity`）上**三项全等**。规划期要读表时，
 * 复核按"同一本簿"放行，任务就可能读到另一个簿的数据 —— 授权是给 A 的，读的是 B。
 *
 * 这条用例只问一件事：两个不同的未保存簿（两个宿主文档对象，同名）能不能共用
 * 一条桥。要求是**不能** —— 复核必须给 `source_changed`，且不派给宿主。
 *
 * 判据必须读生产函数（`currentOfficeBridgeWorkbookIdentity`），不许在用例里自己
 * 造一个"会话内身份"：替身自己写了被测结构，这条变异永远抓不到。
 */
import { afterEach, describe, expect, it } from "vitest";

import {
  bridgeWorkbookIdentity,
  verifyBridgeAuthorization,
  type BridgeLiveAuthorization,
  type PendingBridgeCall
} from "@excel-bro/core";

import { currentOfficeBridgeWorkbookIdentity } from "./excel";
import { createSdkBridgeBinding } from "./sdkBridgeHost";

const OP = "op_" + "2".repeat(32);
const FP = "b".repeat(64);

interface StubDocument {
  url: string;
}

function setDocument(documentRef: StubDocument | undefined): void {
  (globalThis as unknown as { Office?: unknown }).Office = {
    context: documentRef ? { document: documentRef } : {},
    HostType: { Excel: "Excel" }
  };
}

function call(): PendingBridgeCall {
  return {
    operationId: OP,
    tool: "query_table",
    arguments: { mode: "aggregate", groupBy: ["类别"] },
    requestFingerprint: FP,
    expiresAt: 0
  };
}

function live(
  name: string,
  identity: string | undefined,
  sourceFingerprint: string | null = "fp-1"
): BridgeLiveAuthorization {
  return {
    // 与生产同一个折叠点（App.tsx 的 `bridgeLiveContext`）：有身份用身份，
    // 没有才退回名字。用例不在这里另写一套"更严格"的规则。
    workbookIdentity: bridgeWorkbookIdentity(name, identity ?? null),
    sourceFingerprint,
    authorizedSheetNames: ["Sheet1"],
    activeWorksheet: "Sheet1"
  };
}

afterEach(() => {
  setDocument(undefined);
});

describe("未保存同名工作簿的桥身份", () => {
  it("两个未保存的同名簿不能共用同一条桥", () => {
    const first: StubDocument = { url: "" };
    const second: StubDocument = { url: "" };

    setDocument(first);
    const firstIdentity = currentOfficeBridgeWorkbookIdentity();
    const binding = createSdkBridgeBinding({
      taskId: "turn-unsaved-a",
      hostId: "excel-office",
      workbookName: "工作簿1",
      workbookIdentity: bridgeWorkbookIdentity("工作簿1", firstIdentity ?? null),
      sourceFingerprint: "fp-1",
      authorizedSheetNames: ["Sheet1"]
    });

    // 换到另一个「也还没保存、也叫工作簿1」的簿上执行前复核。
    setDocument(second);
    const verdict = verifyBridgeAuthorization(
      binding,
      call(),
      live("工作簿1", currentOfficeBridgeWorkbookIdentity())
    );

    expect(verdict).toEqual({ ok: false, errorCode: "source_changed" });
  });

  it("未保存的工作簿不再退化成 name: 身份", () => {
    setDocument({ url: "" });
    const identity = currentOfficeBridgeWorkbookIdentity();
    expect(bridgeWorkbookIdentity("工作簿1", identity ?? null)).not.toBe(
      "name:工作簿1"
    );
    expect(identity).toMatch(/^unsaved:/);
  });

  it("同一次加载里身份稳定（窗格内多次读取拿到同一个令牌）", () => {
    const documentRef: StubDocument = { url: "" };
    setDocument(documentRef);
    const first = currentOfficeBridgeWorkbookIdentity();
    const second = currentOfficeBridgeWorkbookIdentity();
    expect(first).toBeTruthy();
    expect(second).toBe(first);
  });

  it("已保存的工作簿仍然用 url: 身份，与账本同一口径", () => {
    setDocument({ url: "https://contoso.example/报表.xlsx" });
    expect(currentOfficeBridgeWorkbookIdentity()).toBe(
      "url:https://contoso.example/报表.xlsx"
    );
    // 两个不同路径的已保存簿本来就是两个身份。
    setDocument({ url: "https://contoso.example/另一个/报表.xlsx" });
    expect(currentOfficeBridgeWorkbookIdentity()).toBe(
      "url:https://contoso.example/另一个/报表.xlsx"
    );
  });

  it("没有宿主文档对象时不伪造身份（非 Excel 宿主照旧 fail-closed）", () => {
    setDocument(undefined);
    expect(currentOfficeBridgeWorkbookIdentity()).toBeUndefined();
  });
});
