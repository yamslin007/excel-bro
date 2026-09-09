import { afterEach, describe, expect, it, vi } from "vitest";
import {
  ApiRequestError,
  apiErrorCategory,
  checkHealth,
  checkIntent,
  deleteFolderBackup,
  deleteModelConnection,
  cleanupFolderBackups,
  collectLegacyFolderBackups,
  continuePlanningTask,
  executeFolderPlan,
  executeFolderQuery,
  exportFolderBackup,
  getFolderExecutionStatus,
  getPlanningTaskStatus,
  getModelSettings,
  listModels,
  listFolderBackups,
  listManagedFolderRecoveryDirectories,
  openFolderBackupDirectory,
  removeManagedFolderRecoveryDirectory,
  resumeFolderExecution,
  refreshFolder,
  restoreManagedFolderBackup,
  saveModelConnection,
  scanFolderByPath,
  streamAssistantResponse,
  setFolderBackupPinned,
  testModelConnection,
  undoFolderExecution,
  updateModelSettings
} from "./api";
import type {
  PlanRequest,
  TaskProgressEvent,
  TurnStepEvent
} from "./contracts";

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

  it("scans a folder by path to rebuild a lost session", async () => {
    const fetchMock = vi.fn().mockResolvedValue(
      new Response(
        JSON.stringify({
          sessionId: "session-rebuilt",
          folderName: "reports",
          folderPath: "C:\\reports",
          files: [
            {
              id: "file-1",
              name: "scores.xlsx",
              relativePath: "scores.xlsx",
              worksheets: [{ name: "得分", rowCount: 3, columnCount: 2 }],
              error: null
            }
          ],
          totalFiles: 1,
          truncated: false,
          expiresAt: "2026-09-30T00:00:00.000Z"
        }),
        { status: 200 }
      )
    );
    vi.stubGlobal("fetch", fetchMock);

    const catalog = await scanFolderByPath("C:\\reports");

    expect(catalog.sessionId).toBe("session-rebuilt");
    expect(catalog.files[0].id).toBe("file-1");
    expect(fetchMock).toHaveBeenCalledWith(
      expect.stringContaining("/api/folders/scan"),
      expect.objectContaining({
        method: "POST",
        body: JSON.stringify({ folderPath: "C:\\reports" })
      })
    );
  });

  it("surfaces a missing folder path when rebuilding", async () => {
    vi.stubGlobal(
      "fetch",
      vi.fn().mockResolvedValue(
        new Response(
          JSON.stringify({
            detail: {
              code: "FOLDER_DATA_ERROR",
              message: "文件夹不存在：C:\\gone",
              retryable: true
            }
          }),
          { status: 422 }
        )
      )
    );

    await expect(scanFolderByPath("C:\\gone")).rejects.toMatchObject({
      status: 422,
      code: "FOLDER_DATA_ERROR",
      message: "文件夹不存在：C:\\gone"
    });
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

  it("requests a folder undo with the server-issued recovery id", async () => {
    const fetchMock = vi.fn().mockResolvedValue(
      new Response(
        JSON.stringify({
          filesRestored: ["scores.xlsx"],
          executionsUndone: 1,
          remainingUndoCount: 0
        }),
        { status: 200 }
      )
    );
    vi.stubGlobal("fetch", fetchMock);

    await expect(
      undoFolderExecution("session-1", "undo-1")
    ).resolves.toEqual({
      filesRestored: ["scores.xlsx"],
      executionsUndone: 1,
      remainingUndoCount: 0
    });
    expect(fetchMock).toHaveBeenCalledWith(
      expect.stringContaining("/api/folders/undo"),
      expect.objectContaining({
        method: "POST",
        body: JSON.stringify({
          sessionId: "session-1",
          undoId: "undo-1"
        })
      })
    );
  });

  it("lists backups for the selected folder session", async () => {
    const payload = {
      folderPath: "C:\\work",
      backupDirectory: "C:\\work\\.excel-bro-backups",
      openDirectoryToken: "open-token",
      backups: [],
      legacyCandidates: [],
      warnings: []
    };
    const fetchMock = vi.fn().mockResolvedValue(
      new Response(JSON.stringify(payload), { status: 200 })
    );
    vi.stubGlobal("fetch", fetchMock);

    await expect(
      listFolderBackups({ sessionId: "session 1" })
    ).resolves.toEqual(payload);
    expect(fetchMock).toHaveBeenCalledWith(
      expect.stringContaining("/api/folders/backups?sessionId=session+1")
    );
  });

  it("opens only a backend-registered backup directory token", async () => {
    const payload = {
      opened: true,
      directory: "C:\\work\\.excel-bro-backups"
    };
    const fetchMock = vi.fn().mockResolvedValue(
      new Response(JSON.stringify(payload), { status: 200 })
    );
    vi.stubGlobal("fetch", fetchMock);

    await expect(openFolderBackupDirectory("open-token")).resolves.toEqual(payload);
    expect(fetchMock).toHaveBeenCalledWith(
      expect.stringContaining("/api/folders/backups/open"),
      expect.objectContaining({
        method: "POST",
        body: JSON.stringify({ directoryToken: "open-token" })
      })
    );
  });

  it("exports, deletes, cleans, collects and pins folder backups with scoped requests", async () => {
    const fetchMock = vi
      .fn()
      .mockResolvedValueOnce(
        new Response(
          JSON.stringify({
            deletedUndoId: "undo-1",
            discardedUndo: true,
            filesDeleted: 1,
            bytesFreed: 20,
            remainingUndoCount: 0
          })
        )
      )
      .mockResolvedValueOnce(
        new Response(
          JSON.stringify({ exportDirectory: "C:\\work\\exports", files: ["a.xlsx"] })
        )
      )
      .mockResolvedValueOnce(
        new Response(
          JSON.stringify({
            deletedExecutions: 1,
            deletedFiles: 1,
            bytesFreed: 20,
            skippedActive: 0,
            failures: [],
            legacyCandidateCount: 0
          })
        )
      )
      .mockResolvedValueOnce(
        new Response(JSON.stringify({ collectedFiles: ["legacy/a.xlsx"], failures: [] }))
      )
      .mockResolvedValueOnce(
        new Response(JSON.stringify({ undoId: "undo-1", pinned: true }))
      );
    vi.stubGlobal("fetch", fetchMock);

    await exportFolderBackup(
      { sessionId: "session-1" },
      "undo-1",
      "export-operation-1"
    );
    await deleteFolderBackup("manage-token", "undo-1");
    await cleanupFolderBackups({ sessionId: "session-1" });
    await collectLegacyFolderBackups({ sessionId: "session-1" });
    await setFolderBackupPinned("manage-token", "undo-1", true);

    expect(fetchMock).toHaveBeenNthCalledWith(
      1,
      expect.stringContaining("/api/folders/backups/export"),
      expect.objectContaining({
        body: JSON.stringify({
          sessionId: "session-1",
          undoId: "undo-1",
          operationId: "export-operation-1"
        })
      })
    );
    expect(fetchMock).toHaveBeenNthCalledWith(
      2,
      expect.stringContaining("/api/folders/backups/delete"),
      expect.objectContaining({
        body: JSON.stringify({
          backupManagementToken: "manage-token",
          undoId: "undo-1",
          confirmed: true
        })
      })
    );
    expect(fetchMock).toHaveBeenNthCalledWith(
      3,
      expect.stringContaining("/api/folders/backups/cleanup"),
      expect.objectContaining({ body: JSON.stringify({ sessionId: "session-1" }) })
    );
    expect(fetchMock).toHaveBeenNthCalledWith(
      4,
      expect.stringContaining("/api/folders/backups/legacy/collect"),
      expect.objectContaining({
        body: JSON.stringify({ sessionId: "session-1", confirmed: true })
      })
    );
    expect(fetchMock).toHaveBeenNthCalledWith(
      5,
      expect.stringContaining("/api/folders/backups/pin"),
      expect.objectContaining({
        body: JSON.stringify({
          backupManagementToken: "manage-token",
          undoId: "undo-1",
          pinned: true
        })
      })
    );
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

  it("returns a durable planning pause instead of treating it as an error", async () => {
    const pause = {
      kind: "planning_pause",
      turnId: "turn-stream-test",
      status: "waiting_user",
      revision: 4,
      lastSequence: 2,
      reason: "本轮处理时间已用完，当前进度已安全保存。",
      activeElapsedSeconds: 60,
      approvedWallTimeSeconds: 60,
      extensionSeconds: 60,
      extensionCount: 0,
      maximumExtensions: 4,
      canContinue: true,
      requiresExtension: true,
      sourceFingerprint: "sha256:source",
      sourceFingerprintSheets: ["Sheet1"]
    };
    vi.stubGlobal(
      "fetch",
      vi.fn().mockResolvedValue(
        sseStreamResponse([`event: paused\ndata: ${JSON.stringify(pause)}\n\n`])
      )
    );

    await expect(streamAssistantResponse(streamPlanRequest)).resolves.toEqual(pause);
  });

  it("checks persisted status and explicitly continues with a stable request id", async () => {
    const statusPayload = {
      kind: "planning_pause",
      turnId: "turn-stream-test",
      status: "waiting_user",
      revision: 7,
      lastSequence: 3,
      reason: "进度已保存",
      activeElapsedSeconds: 60,
      approvedWallTimeSeconds: 60,
      extensionSeconds: 60,
      extensionCount: 0,
      maximumExtensions: 4,
      canContinue: true,
      requiresExtension: true,
      sourceFingerprintSheets: ["Sheet1"]
    };
    const fetchMock = vi
      .fn()
      .mockResolvedValueOnce(
        new Response(JSON.stringify(statusPayload), {
          status: 200,
          headers: { "Content-Type": "application/json" }
        })
      )
      .mockResolvedValueOnce(
        sseStreamResponse([
          'event: result\ndata: {"kind":"answer","message":"续跑完成","provider":"model","turnId":"turn-stream-test"}\n\n'
        ])
      );
    vi.stubGlobal("fetch", fetchMock);

    const status = await getPlanningTaskStatus("turn-stream-test");
    const response = await continuePlanningTask({
      turnId: "turn-stream-test",
      expectedRevision: status.revision,
      requestId: "continue-request-0001",
      afterSequence: status.lastSequence
    });

    expect(status.status).toBe("waiting_user");
    expect(response).toMatchObject({ kind: "answer", message: "续跑完成" });
    expect(fetchMock).toHaveBeenNthCalledWith(
      2,
      expect.stringContaining("/api/turn/continue/stream"),
      expect.objectContaining({
        method: "POST",
        body: expect.stringContaining('"confirmed":true')
      })
    );
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

describe("executeFolderPlan streaming", () => {
  it("reconnects with the same operationId and forwards progress", async () => {
    const resultFrame = `event: result\ndata: ${JSON.stringify({
      filesModified: ["数据.xlsx"],
      backups: [],
      backupDirectory: null,
      backupWarnings: [],
      actionResults: [
        {
          index: 0,
          type: "createWorksheet",
          sheet: "结果",
          status: "succeeded"
        }
      ],
      verification: {
        status: "verified",
        passed: true,
        checks: [],
        unverifiedActions: []
      },
      undo: null,
      executionMs: 1,
      verificationMs: 1
    })}\n\n`;
    const progressFrame = `event: progress\ndata: ${JSON.stringify({
      schemaVersion: 1,
      operationId: "folder-operation-client-test",
      executionId: "0123456789abcdef0123456789abcdef",
      sequence: 1,
      phase: "executing",
      stepId: "folder.action.1",
      status: "completed",
      title: "第 1 步已完成",
      current: 1,
      total: 1,
      unit: "actions",
      timestamp: "2026-09-02T00:00:00Z"
    })}\n\n`;
    const fetchMock = vi
      .fn()
      .mockRejectedValueOnce(new TypeError("connection interrupted"))
      .mockResolvedValueOnce(sseStreamResponse([progressFrame, resultFrame]));
    vi.stubGlobal("fetch", fetchMock);
    const progress: TaskProgressEvent[] = [];
    const connectionStates: string[] = [];

    const result = await executeFolderPlan(
      "folder-session-test",
      {
        id: "folder-plan-test",
        title: "创建结果表",
        summary: "创建一张结果表",
        assumptions: [],
        warnings: [],
        actions: [{ type: "createWorksheet", sheet: "结果" }]
      },
      {
        operationId: "folder-operation-client-test",
        onProgress: (event) => progress.push(event),
        onConnectionState: (state) => connectionStates.push(state)
      }
    );

    expect(fetchMock).toHaveBeenCalledTimes(2);
    const firstBody = JSON.parse(fetchMock.mock.calls[0][1].body);
    const secondBody = JSON.parse(fetchMock.mock.calls[1][1].body);
    expect(firstBody.operationId).toBe("folder-operation-client-test");
    expect(secondBody.operationId).toBe(firstBody.operationId);
    expect(firstBody.afterSequence).toBe(0);
    expect(secondBody.afterSequence).toBe(0);
    expect(progress).toHaveLength(1);
    expect(connectionStates).toEqual(["reconnecting", "connected"]);
    expect(result.filesModified).toEqual(["数据.xlsx"]);
  });

  it("does not fall back to the legacy write endpoint when a successful stream has no body", async () => {
    const fetchMock = vi.fn().mockResolvedValue(new Response(null, { status: 200 }));
    vi.stubGlobal("fetch", fetchMock);

    await expect(
      executeFolderPlan("folder-session-test", {
        id: "folder-plan-test",
        title: "创建结果表",
        summary: "创建一张结果表",
        assumptions: [],
        warnings: [],
        actions: [{ type: "createWorksheet", sheet: "结果" }]
      })
    ).rejects.toMatchObject({
      status: 502,
      code: "EMPTY_EXECUTION_STREAM",
      retryable: false
    });
    expect(fetchMock).toHaveBeenCalledTimes(1);
  });
});

describe("folder recovery API", () => {
  it("resumes only the identified checkpoint revision and forwards progress", async () => {
    const progressFrame = `event: progress\ndata: ${JSON.stringify({
      schemaVersion: 1,
      operationId: "execution-resume",
      sequence: 8,
      phase: "saving",
      stepId: "folder.resume.publish",
      status: "progress",
      title: "正在继续未完成的文件发布",
      current: 1,
      total: 1,
      unit: "files",
      timestamp: "2026-09-09T00:00:00Z"
    })}\n\n`;
    const resultFrame = `event: result\ndata: ${JSON.stringify({
      executionId: "a".repeat(32),
      filesModified: ["cinema.xlsx"],
      backups: [],
      backupWarnings: [],
      actionResults: [],
      verification: {
        status: "verified",
        passed: true,
        checks: [],
        unverifiedActions: []
      },
      undo: null,
      executionMs: 0,
      verificationMs: 1
    })}\n\n`;
    const fetchMock = vi
      .fn()
      .mockResolvedValue(sseStreamResponse([progressFrame, resultFrame]));
    vi.stubGlobal("fetch", fetchMock);
    const progress: TaskProgressEvent[] = [];

    const result = await resumeFolderExecution("execution-resume", 12, {
      onProgress: (event) => progress.push(event)
    });

    expect(result.filesModified).toEqual(["cinema.xlsx"]);
    expect(progress).toHaveLength(1);
    expect(fetchMock).toHaveBeenCalledWith(
      expect.stringContaining("/api/folders/execute/resume/stream"),
      expect.objectContaining({
        method: "POST",
        body: JSON.stringify({
          operationId: "execution-resume",
          expectedRevision: 12,
          confirmed: true,
          afterSequence: 0
        })
      })
    );
  });

  it("queries the existing operation without creating a new execution", async () => {
    const fetchMock = vi.fn().mockResolvedValue(
      new Response(
        JSON.stringify({
          operationId: "execution-existing",
          status: "needs_reconciliation",
          lastSequence: 7,
          result: null,
          error: { message: "需要核对" }
        })
      )
    );
    vi.stubGlobal("fetch", fetchMock);

    const status = await getFolderExecutionStatus("execution-existing");

    expect(status.status).toBe("needs_reconciliation");
    expect(fetchMock).toHaveBeenCalledWith(
      expect.stringContaining(
        "/api/folders/execute/status?operationId=execution-existing"
      )
    );
  });

  it("lists, restores and removes only an explicitly identified managed directory", async () => {
    const directoryId = "d".repeat(64);
    const undoId = "a".repeat(32);
    const fetchMock = vi
      .fn()
      .mockResolvedValueOnce(new Response(JSON.stringify({ directories: [] })))
      .mockResolvedValueOnce(
        new Response(
          JSON.stringify({
            operationId: "recovery-operation-1",
            undoId,
            status: "completed",
            filesRestored: ["scores.xlsx"],
            alreadyRestored: false
          })
        )
      )
      .mockResolvedValueOnce(
        new Response(JSON.stringify({ directoryId, removed: true }))
      );
    vi.stubGlobal("fetch", fetchMock);

    await listManagedFolderRecoveryDirectories();
    await restoreManagedFolderBackup({
      directoryId,
      undoId,
      expectedRevision: 3,
      operationId: "recovery-operation-1",
      confirmed: true
    });
    await removeManagedFolderRecoveryDirectory(directoryId);

    expect(fetchMock).toHaveBeenNthCalledWith(
      2,
      expect.stringContaining("/api/folders/recovery/restore"),
      expect.objectContaining({
        method: "POST",
        body: JSON.stringify({
          directoryId,
          undoId,
          expectedRevision: 3,
          operationId: "recovery-operation-1",
          confirmed: true
        })
      })
    );
    expect(fetchMock).toHaveBeenNthCalledWith(
      3,
      expect.stringContaining(`/api/folders/recovery/directories/${directoryId}`),
      { method: "DELETE" }
    );
  });
});
