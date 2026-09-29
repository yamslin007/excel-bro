import { describe, expect, it } from "vitest";
import {
  REDACTED_PLACEHOLDER,
  containsLikelyCredential,
  redactCredentials
} from "./redact";

describe("redactCredentials", () => {
  it("抹掉 Bearer token，保留前缀与占位符", () => {
    const out = redactCredentials("上游拒绝 Authorization: Bearer sk-abcdefghijklmnop");
    expect(out).not.toContain("sk-abcdefghijklmnop");
    expect(out).toContain("Bearer");
    expect(out).toContain(REDACTED_PLACEHOLDER);
  });

  it("抹掉 URL 内嵌的用户名密码", () => {
    const out = redactCredentials("连接失败：https://alice:hunter2@example.com/v1");
    expect(out).not.toContain("hunter2");
    expect(out).not.toContain("alice:");
    expect(out).toContain("example.com");
  });

  it("抹掉 sk- 前缀的 API Key", () => {
    const out = redactCredentials("key=sk-live-leaked-from-upstream-9f3a");
    expect(out).not.toContain("sk-live");
    expect(out).not.toContain("9f3a");
  });

  it("抹掉带标签的赋值形态（大小写与分隔符变体）", () => {
    for (const text of [
      "apiKey=secret-value-1234",
      "api_key: secret-value-1234",
      "API-KEY=secret-value-1234",
      "token=secret-value-1234",
      "password: secret-value-1234",
      "authorization=secret-value-1234"
    ]) {
      const out = redactCredentials(text);
      expect(out, text).not.toContain("secret-value-1234");
      expect(out, text).toContain(REDACTED_PLACEHOLDER);
    }
  });

  it("短得像密钥但不是的字符串不动，避免把普通文本误吞", () => {
    // 普通错误文案不能被改坏，否则用户读不懂为什么失败。
    const text = "无法连接 Excel Bro 后端。请确认本地服务已启动，然后重试。";
    expect(redactCredentials(text)).toBe(text);
  });

  it("抹掉 JSON 形态的键值凭据，保留键名以便诊断", () => {
    const cases: Array<[string, string]> = [
      ['{"apiKey":"testsecret123456"}', "testsecret123456"],
      ['{"api_key":"testsecret123456"}', "testsecret123456"],
      ['{"api-key": "testsecret123456"}', "testsecret123456"],
      ['{"api key": "testsecret123456"}', "testsecret123456"],
      ['{"access_token":"testsecret123456"}', "testsecret123456"],
      ['{"token": "testsecret123456"}', "testsecret123456"],
      ['{"secret":"testsecret123456"}', "testsecret123456"],
      ['{"password":"testsecret123456"}', "testsecret123456"],
      ['{"Authorization":"testsecret123456"}', "testsecret123456"]
    ];
    for (const [text, secret] of cases) {
      const out = redactCredentials(text);
      expect(out, text).not.toContain(secret);
      expect(out, text).toContain(REDACTED_PLACEHOLDER);
      // JSON 结构必须还在 —— 只换值，不破坏引号与冒号。
      expect(out, text).toMatch(/^\{"[^"]+": ?"【已隐藏】"\}$/);
    }
  });

  it("抹掉单引号形式的类 JSON 键值（不是合法 JSON，但同样会进日志）", () => {
    for (const text of [
      "{'api_key': 'testsecret123456'}",
      "'apiKey': 'testsecret123456'",
      "apiKey: 'testsecret123456'"
    ]) {
      const out = redactCredentials(text);
      expect(out, text).not.toContain("testsecret123456");
      expect(out, text).toContain(REDACTED_PLACEHOLDER);
      // 键名保留：出现 problem 时能看出是哪一项被隐藏。
      expect(out, text).toMatch(/api[_ -]?key/i);
    }
  });

  it("JSON 里的 Authorization: Bearer 仍保留方案名，只换 token", () => {
    const out = redactCredentials(
      '{"Authorization":"Bearer sk-abcdefghijklmnop"}'
    );
    expect(out).not.toContain("sk-abcdefghijklmnop");
    expect(out).toContain("Bearer");
    expect(out).toContain(REDACTED_PLACEHOLDER);
  });

  it("保留可诊断的部分：状态码、错误码与中文说明都还在", () => {
    const out = redactCredentials(
      "模型连接校验失败（MODEL_CONNECTION_DUPLICATE）：请换一个标签。Bearer sk-abcdefghijklmnop"
    );
    expect(out).toContain("MODEL_CONNECTION_DUPLICATE");
    expect(out).toContain("请换一个标签");
    expect(out).not.toContain("sk-abcdefghijklmnop");
  });

  it("不要把 Bearer 这个方案名本身误判成密钥", () => {
    // 「Bearer 认证失败」这类说明没有 token，不该被当成凭据改写。
    const text = "上游返回 401：Bearer 认证失败，请检查密钥是否正确。";
    expect(redactCredentials(text)).toBe(text);
  });

  it("containsLikelyCredential 与脱敏结果一致", () => {
    expect(containsLikelyCredential("Bearer sk-abcdefghijklmnop")).toBe(true);
    expect(containsLikelyCredential('{"apiKey":"testsecret123456"}')).toBe(true);
    expect(containsLikelyCredential("普通错误文案")).toBe(false);
  });

  it("占位符是带分隔符的结构化写法，不会被当成自然语言", () => {
    // 读取历史时要能把"脱敏留下的标记"和"用户自己写的'已隐藏'"区分开。
    // 裸的"已隐藏"三个字做不到这一点，所以占位符带方括号。
    expect(REDACTED_PLACEHOLDER).toContain("已隐藏");
    expect(REDACTED_PLACEHOLDER).not.toBe("已隐藏");
    const plain = "这一列我暂时已隐藏，先不用管。";
    expect(redactCredentials(plain)).toBe(plain);
    expect(containsLikelyCredential(plain)).toBe(false);
  });
});
