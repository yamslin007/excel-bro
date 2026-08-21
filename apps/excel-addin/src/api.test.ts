import { afterEach, describe, expect, it, vi } from "vitest";
import {
  ApiRequestError,
  apiErrorCategory,
  checkHealth,
  checkIntent,
  deleteModelConnection,
  executeFolderQuery,
  getModelSettings,
  listModels,
  refreshFolder,
  saveModelConnection,
  streamAssistantResponse,
  testModelConnection,
  updateModelSettings
} from "./api";
import type { PlanRequest, TurnStepEvent } from "./contracts";

afterEach(() => {
  vi.unstubAllGlobals();
});

const streamPlanRequest: PlanRequest = {
  turnId: "turn-stream-test",
  prompt: "分析当前表",
  workbook: {
    name: "测试.xlsx",
    capturedAt: "2026-07-30T00:00:00.000Z",
    activeWorksheet: "Sheet1",
    selectedRange: null,
    worksheets: [
      {
        name: "Sheet1",
        sourceFile: null,
        sourceSheet: null,
        usedRange: "Sheet1!A1:C4",
        rowCount: 4,
        columnCount: 3,
        headers: ["姓名", "学科", "分数"],
        dataRows: [],
        truncated: false
      }
    ]
  },
  lastResult: null,
  images: [],
  dataResults: [],
  modelId: "local"
};

function sseStreamResponse(chunks: string[]): Response {
  const encoder = new TextEncoder();
  const stream = new ReadableStream<Uint8Array>({
    start(controller) {
      for (const chunk of chunks) {
        controller.enqueue(encoder.encode(chunk));
      }
      controller.close();
    }
  });
  return new Response(stream, {
    status: 200,
    headers: { "Content-Type": "text/event-stream" }
  });
}

describe("service discovery", () => {
  it("distinguishes local connectivity failures from model responses", () => {
    expect(apiErrorCategory(new TypeError("Failed to fetch"))).toBe(
      "local_service"
    );
    expect(
      apiErrorCategory(
        new ApiRequestError(502, "MODEL_TRANSPORT_ERROR", "模型不可用", true)
      )
    ).toBe("model");
    expect(
      apiErrorCategory(
        new ApiRequestError(422, "FOLDER_DATA_ERROR", "数据无效", true)
      )
    ).toBe("service");
  });
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

  it("reads masked model settings without receiving the complete key", async () => {
    vi.stubGlobal(
      "fetch",
      vi.fn().mockResolvedValue(
        new Response(
          JSON.stringify({
            baseUrl: "https://api.example.test/v1",
            defaultModel: "configured-model",
            apiKeyConfigured: true,
            apiKeyHint: "••••cret",
            connections: []
          }),
          { status: 200 }
        )
      )
    );

    await expect(getModelSettings()).resolves.toEqual({
      baseUrl: "https://api.example.test/v1",
      defaultModel: "configured-model",
      apiKeyConfigured: true,
      apiKeyHint: "••••cret",
      connections: []
    });
  });

  it("updates the API key through the local settings endpoint", async () => {
    const fetchMock = vi.fn().mockResolvedValue(
      new Response(
        JSON.stringify({
          baseUrl: "https://api.example.test/v1",
          defaultModel: "configured-model",
          apiKeyConfigured: true,
          apiKeyHint: "••••-key",
          connections: []
        }),
        { status: 200 }
      )
    );
    vi.stubGlobal("fetch", fetchMock);

    await updateModelSettings({ apiKey: "replacement-key" });

    expect(fetchMock).toHaveBeenCalledWith(
      expect.stringContaining("/api/settings/model"),
      expect.objectContaining({
        method: "PUT",
        body: JSON.stringify({ apiKey: "replacement-key" })
      })
    );
  });

  it("creates and deletes independent model connections", async () => {
    const responsePayload = {
      baseUrl: null,
      defaultModel: null,
      apiKeyConfigured: false,
      apiKeyHint: null,
      connections: [
        {
          id: "model-example",
          catalogModelId: "connection:model-example",
          label: "DeepSeek",
          baseUrl: "https://api.deepseek.com/v1",
          modelId: "deepseek-chat",
          supportsVision: false,
          apiKeyConfigured: true,
          apiKeyHint: "••••cret"
        }
      ]
    };
    const fetchMock = vi.fn().mockImplementation(() =>
      Promise.resolve(new Response(JSON.stringify(responsePayload)))
    );
    vi.stubGlobal("fetch", fetchMock);

    await saveModelConnection({
      label: "DeepSeek",
      baseUrl: "https://api.deepseek.com/v1",
      modelId: "deepseek-chat",
      apiKey: "secret",
      supportsVision: false
    });
    await deleteModelConnection("model-example");

    expect(fetchMock).toHaveBeenNthCalledWith(
      1,
      expect.stringContaining("/api/settings/model/connections"),
      expect.objectContaining({ method: "POST" })
    );
    expect(fetchMock).toHaveBeenNthCalledWith(
      2,
      expect.stringContaining(
        "/api/settings/model/connections/model-example"
      ),
      { method: "DELETE" }
    );
  });

  it("tests a model connection without saving it", async () => {
    const fetchMock = vi.fn().mockResolvedValue(
      new Response(
        JSON.stringify({
          ok: true,
          message: "连接成功，服务地址、模型 ID 和 API Key 可用。"
        })
      )
    );
    vi.stubGlobal("fetch", fetchMock);
    const request = {
      label: "DeepSeek",
      baseUrl: "https://api.deepseek.com/v1",
      modelId: "deepseek-chat",
      apiKey: "secret",
      clearApiKey: false,
      supportsVision: false
    };

    await expect(testModelConnection(request)).resolves.toMatchObject({
      ok: true
    });
    expect(fetchMock).toHaveBeenCalledWith(
      expect.stringContaining("/api/settings/model/connections/test"),
      expect.objectContaining({
        method: "POST",
        body: JSON.stringify(request)
      })
    );
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
          worksheetNames: ["Sheet1"],
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

  it("refreshes a folder catalog while keeping the session id", async () => {
    const fetchMock = vi.fn().mockResolvedValue(
      new Response(
        JSON.stringify({
          sessionId: "session-1",
          folderName: "reports",
          folderPath: "C:\\reports",
          files: [
            {
              id: "file-1",
              name: "renamed.xlsx",
              relativePath: "renamed.xlsx",
              worksheets: [{ name: "得分", rowCount: 3, columnCount: 2 }],
              error: null
            }
          ],
          totalFiles: 1,
          truncated: false,
          expiresAt: "2026-08-20T00:00:00.000Z"
        }),
        { status: 200 }
      )
    );
    vi.stubGlobal("fetch", fetchMock);

    const catalog = await refreshFolder("session-1");

    expect(catalog.sessionId).toBe("session-1");
    expect(catalog.files[0].name).toBe("renamed.xlsx");
    expect(fetchMock).toHaveBeenCalledWith(
      expect.stringContaining("/api/folders/refresh"),
      expect.objectContaining({
        method: "POST",
        body: JSON.stringify({ sessionId: "session-1" })
      })
    );
  });

  it("surfaces folder refresh session errors", async () => {
    vi.stubGlobal(
      "fetch",
      vi.fn().mockResolvedValue(
        new Response(
          JSON.stringify({
            detail: {
              code: "FOLDER_DATA_ERROR",
              message: "文件夹会话已失效，请重新选择文件夹",
              retryable: true
            }
          }),
          { status: 422 }
        )
      )
    );

    await expect(refreshFolder("missing-session")).rejects.toMatchObject({
      status: 422,
      code: "FOLDER_DATA_ERROR",
      message: "文件夹会话已失效，请重新选择文件夹"
    });
  });

  it("rejects an invalid folder query response at runtime", async () => {
    vi.stubGlobal(
      "fetch",
      vi.fn().mockResolvedValue(
        new Response(JSON.stringify({ tool: "query_table", rows: "invalid" }))
      )
    );

    await expect(
      executeFolderQuery("session", {
        id: "query",
        tool: "query_table",
        arguments: { mode: "rows" }
      })
    ).rejects.toThrow("无效的数据工具结果");
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
        worksheetNames: ["Sheet1"],
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

describe("streaming assistant response", () => {
  it("dispatches step events and returns the final result", async () => {
    const resultFrame =
      "event: result\n" +
      `data: ${JSON.stringify({
        kind: "answer",
        message: "分析完成",
        provider: "model",
        turnId: "turn-stream-test"
      })}\n\n`;
    const fetchMock = vi.fn().mockResolvedValue(
      sseStreamResponse([
        'event: step\ndata: {"phase":"planning","title":"正在理解需求"}\n\n',
        'event: step\ndata: {"phase":"planning","title":"正在查找字段","detail":"分数","completedStep":"正在理解需求"}\n\n',
        resultFrame
      ])
    );
    vi.stubGlobal("fetch", fetchMock);

    const steps: TurnStepEvent[] = [];
    const response = await streamAssistantResponse(streamPlanRequest, {
      onStep: (step) => steps.push(step)
    });

    expect(fetchMock).toHaveBeenCalledWith(
      expect.stringContaining("/api/turn/stream"),
      expect.objectContaining({ method: "POST" })
    );
    expect(steps).toEqual([
      { phase: "planning", title: "正在理解需求", detail: null, completedStep: null },
      {
        phase: "planning",
        title: "正在查找字段",
        detail: "分数",
        completedStep: "正在理解需求"
      }
    ]);
    expect(response).toMatchObject({ kind: "answer", message: "分析完成" });
  });

  it("reassembles events split across stream chunks", async () => {
    const fetchMock = vi.fn().mockResolvedValue(
      sseStreamResponse([
        'event: step\ndata: {"phase":"planning","ti',
        'tle":"正在规划操作"}\n\n' +
          'event: result\ndata: {"kind":"answer","message":"完成","provi',
        'der":"model","turnId":"turn-stream-test"}\n\n'
      ])
    );
    vi.stubGlobal("fetch", fetchMock);

    const steps: TurnStepEvent[] = [];
    const response = await streamAssistantResponse(streamPlanRequest, {
      onStep: (step) => steps.push(step)
    });

    expect(steps).toEqual([
      { phase: "planning", title: "正在规划操作", detail: null, completedStep: null }
    ]);
    expect(response).toMatchObject({ kind: "answer", message: "完成" });
  });

  it("throws ApiRequestError when the stream emits an error event", async () => {
    const fetchMock = vi.fn().mockResolvedValue(
      sseStreamResponse([
        'event: error\n' +
          `data: ${JSON.stringify({
            status: 504,
            code: "MODEL_TIMEOUT",
            message: "模型响应超时，可以重试当前对话轮次。",
            retryable: true
          })}\n\n`
      ])
    );
    vi.stubGlobal("fetch", fetchMock);

    await expect(
      streamAssistantResponse(streamPlanRequest)
    ).rejects.toMatchObject({
      status: 504,
      code: "MODEL_TIMEOUT",
      retryable: true
    });
  });

  it("falls back to the non-streaming endpoint when the stream endpoint is missing", async () => {
    const fetchMock = vi
      .fn()
      .mockResolvedValueOnce(new Response("Not Found", { status: 404 }))
      .mockResolvedValueOnce(
        new Response(
          JSON.stringify({
            kind: "answer",
            message: "回退结果",
            provider: "model",
            turnId: "turn-stream-test"
          }),
          { status: 200 }
        )
      );
    vi.stubGlobal("fetch", fetchMock);

    const response = await streamAssistantResponse(streamPlanRequest);

    expect(fetchMock).toHaveBeenNthCalledWith(
      1,
      expect.stringContaining("/api/turn/stream"),
      expect.objectContaining({ method: "POST" })
    );
    expect(fetchMock).toHaveBeenNthCalledWith(
      2,
      expect.stringContaining("/api/turn"),
      expect.objectContaining({ method: "POST" })
    );
    expect(response).toMatchObject({ kind: "answer", message: "回退结果" });
  });
});
