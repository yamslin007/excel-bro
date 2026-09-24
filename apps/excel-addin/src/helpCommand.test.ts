import { describe, it, expect } from "vitest";
import {
  BASE_MODE_HELP_TEXT,
  MODEL_REQUIRED_HELP_TEXT,
} from "./helpCommand";

describe("MODEL_REQUIRED_HELP_TEXT", () => {
  it("leads with model configuration instead of a base mode", () => {
    expect(MODEL_REQUIRED_HELP_TEXT).toContain("📘 Excel Bro 使用说明");
    expect(MODEL_REQUIRED_HELP_TEXT).toContain("添加模型");
    expect(MODEL_REQUIRED_HELP_TEXT).toContain("服务地址");
    expect(MODEL_REQUIRED_HELP_TEXT).toContain("API Key");
  });

  it("still documents the deterministic address commands", () => {
    expect(MODEL_REQUIRED_HELP_TEXT).toContain("1️⃣ 清空区域");
    expect(MODEL_REQUIRED_HELP_TEXT).toContain("2️⃣ 写入值");
    expect(MODEL_REQUIRED_HELP_TEXT).toContain("3️⃣ 写入公式");
    expect(MODEL_REQUIRED_HELP_TEXT).toContain("📝 语法说明");
    expect(MODEL_REQUIRED_HELP_TEXT).toContain("执行预览");
  });

  it("does not present a base mode or a fake zero-latency offline mode", () => {
    expect(MODEL_REQUIRED_HELP_TEXT).not.toContain("基础模式");
    expect(MODEL_REQUIRED_HELP_TEXT).not.toContain("离线");
    expect(MODEL_REQUIRED_HELP_TEXT).not.toContain("零延迟");
    expect(MODEL_REQUIRED_HELP_TEXT).not.toContain("本地 AI");
  });

  it("keeps the deprecated alias pointing at the current text", () => {
    expect(BASE_MODE_HELP_TEXT).toBe(MODEL_REQUIRED_HELP_TEXT);
  });
});
