import { describe, it, expect } from "vitest";
import { BASE_MODE_HELP_TEXT } from "./helpCommand";

describe("BASE_MODE_HELP_TEXT", () => {
  it("covers the base-mode command documentation", () => {
    expect(BASE_MODE_HELP_TEXT).toContain("📘 基础模式功能说明");
    expect(BASE_MODE_HELP_TEXT).toContain("1️⃣ 清空区域");
    expect(BASE_MODE_HELP_TEXT).toContain("2️⃣ 写入值");
    expect(BASE_MODE_HELP_TEXT).toContain("3️⃣ 写入公式");
    expect(BASE_MODE_HELP_TEXT).toContain("📝 语法说明");
    expect(BASE_MODE_HELP_TEXT).toContain("❌ 不支持的功能");
    expect(BASE_MODE_HELP_TEXT).toContain("/model");
  });
});
