/**
 * 桥调用前的**授权与来源复核**：直接跑生产共享判据，不在这里另写一份。
 *
 * 为什么需要这个入口：`scripts/drive_real_sdk_result_expiry.py` 扮演的是**宿主适配器**
 * （真实面板里 `useSdkToolBridge` / `wpsSdkBridgeHost` 那一侧）。面板每次执行受控读取
 * 之前调用的判据是 `packages/core/src/sdkToolBridge.ts` 的 `verifyBridgeAuthorization`：
 * 换簿、来源变化、授权收窄三种情况一律拒交值。
 *
 * 早先驱动在 Python 里又写了一遍"指纹不等就回 source_changed"，那是**第二份行为相似
 * 的判据** —— 生产改一次，两份就会漂移，而它产出的工件还会被当成真实证据。这里改成
 * 直接 import 那份 TS（Node 22.18+ 默认支持 type stripping，模块本身没有任何依赖），
 * 于是"拒绝码是谁判的"只有一处答案。
 *
 * 输入（stdin，一个 JSON 对象）：
 *   { binding: SdkBridgeBinding, call: PendingBridgeCall, live: BridgeLiveAuthorization }
 * 输出（stdout，一个 JSON 对象）：
 *   { ok: true } | { ok: false, errorCode: "source_changed" | "unauthorized_scope" }
 *
 * 输入不合法时以退出码 2 结束：那说明**驱动写错了**，不是"拒绝了一次读取"，不能被
 * 当成 source_changed 混进证据里。
 */
import { readFileSync } from "node:fs";
import { verifyBridgeAuthorization } from "../packages/core/src/sdkToolBridge.ts";

function fail(message) {
  process.stdout.write(JSON.stringify({ error: message }) + "\n");
  process.exit(2);
}

let payload;
try {
  payload = JSON.parse(readFileSync(0, "utf8"));
} catch {
  fail("stdin 不是合法 JSON");
}

const { binding, call, live } = payload ?? {};
if (!binding || typeof binding !== "object") fail("缺少 binding");
if (!call || typeof call !== "object") fail("缺少 call");
if (!live || typeof live !== "object") fail("缺少 live");
if (!Array.isArray(binding.authorizedSheetNames)) fail("binding.authorizedSheetNames 必须是数组");
if (!Array.isArray(live.authorizedSheetNames)) fail("live.authorizedSheetNames 必须是数组");
if (typeof live.workbookIdentity !== "string") fail("live.workbookIdentity 必须是字符串");
if (typeof live.activeWorksheet !== "string") fail("live.activeWorksheet 必须是字符串");

let verdict;
try {
  verdict = verifyBridgeAuthorization(binding, call, live);
} catch (error) {
  fail(`verifyBridgeAuthorization 抛出异常：${error instanceof Error ? error.message : String(error)}`);
}
process.stdout.write(JSON.stringify(verdict) + "\n");
