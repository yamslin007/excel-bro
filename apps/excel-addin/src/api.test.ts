import { afterEach, describe, expect, it, vi } from "vitest";
import { checkHealth, checkIntent, listModels } from "./api";

afterEach(() => {
  vi.unstubAllGlobals();
});

describe("service discovery", () => {
  it("returns detailed health without exposing configuration secrets", async () => {
    vi.stubGlobal(
      "fetch",
      vi.fn().mockResolvedValue(
        new Response(
          JSON.stringify({
            status: "ok",
            configured: true,
            mode: "model",
            model: "configured-model"
          }),
          { status: 200 }
        )
      )
    );

    await expect(checkHealth()).resolves.toEqual({
      status: "ok",
      configured: true,
      mode: "model",
      model: "configured-model"
    });
  });

  it("treats an invalid health payload as offline", async () => {
    vi.stubGlobal(
      "fetch",
      vi.fn().mockResolvedValue(
        new Response(JSON.stringify({ status: "ok" }), { status: 200 })
      )
    );

    await expect(checkHealth()).resolves.toBeNull();
  });

  it("loads the server-side model allowlist", async () => {
    vi.stubGlobal(
      "fetch",
      vi.fn().mockResolvedValue(
        new Response(
          JSON.stringify({
            defaultModelId: "configured-model",
            models: [
              {
                id: "local",
                label: "基础模式",
                provider: "local",
                available: true,
                supportsVision: false
              },
              {
                id: "configured-model",
                label: "configured-model",
                provider: "model",
                available: true,
                supportsVision: true
              }
            ]
          }),
          { status: 200 }
        )
      )
    );

    const catalog = await listModels();

    expect(catalog.defaultModelId).toBe("configured-model");
    expect(catalog.models).toHaveLength(2);
  });

  it("requests a lightweight intent check before planning", async () => {
    const fetchMock = vi.fn().mockResolvedValue(
      new Response(
        JSON.stringify({
          kind: "proceed",
          provider: "local",
          summary: "只分析当前表",
          confirmedPrompt: "只分析当前工作表"
        }),
        { status: 200 }
      )
    );
    vi.stubGlobal("fetch", fetchMock);

    await expect(
      checkIntent({
        prompt: "分析当前表",
        scope: {
          workbookName: "测试.xlsx",
          sourceMode: "workbook",
          selectionMode: "auto",
          activeWorksheet: "Sheet1",
          selectedRange: "Sheet1!A1:C4",
          totalWorksheetCount: 1,
          sheets: [
            {
              name: "Sheet1",
              usedRange: "Sheet1!A1:C4",
              rowCount: 4,
              columnCount: 3,
              headers: ["姓名", "学科", "分数"]
            }
          ]
        },
        modelId: "local"
      })
    ).resolves.toMatchObject({ kind: "proceed" });

    expect(fetchMock).toHaveBeenCalledWith(
      expect.stringContaining("/api/turn"),
      expect.objectContaining({ method: "POST" })
    );
  });

  it("parses structured service errors without exposing the JSON envelope", async () => {
    vi.stubGlobal(
      "fetch",
      vi.fn().mockResolvedValue(
        new Response(
          JSON.stringify({
            detail: {
              code: "TURN_ALREADY_COMPLETED",
              message: "该轮次已经完成，请开始新的请求。",
              retryable: false
            }
          }),
          { status: 409 }
        )
      )
    );

    const request = checkIntent({
      turnId: "turn-completed-test",
      prompt: "继续处理",
      scope: {
        workbookName: "测试.xlsx",
        sourceMode: "workbook",
        selectionMode: "auto",
        activeWorksheet: "Sheet1",
        totalWorksheetCount: 1,
        sheets: [
          {
            name: "Sheet1",
            usedRange: null,
            rowCount: 1,
            columnCount: 1,
            headers: ["字段"]
          }
        ]
      },
      modelId: "local"
    });

    await expect(request).rejects.toMatchObject(
      {
        status: 409,
        code: "TURN_ALREADY_COMPLETED",
        message: "该轮次已经完成，请开始新的请求。",
        retryable: false
      }
    );
  });
});
