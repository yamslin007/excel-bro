// 反向变异脚本：WPS「完整退出宿主后仍可核对并恢复原执行」的材料链（2026-10-27）
//
// 背景：真实 WPS 验收发现，完整退出 WPS 再打开后提示「找不到原执行的本机恢复记录」。
// 修复把不可变恢复材料挪到**本机后端应用数据目录**，`PluginStorage` 退化为缓存/索引。
// 这两条机制各自必须有一条能把它打红的用例 —— 否则"已修好"只是自述。
//
// M1 = `ensureWpsRecoveryMaterial` 退回"只信宿主缓存"（没有缓存就直接 `none`）
//      ⇒ 「full process restart loses PluginStorage but resumes only third from durable material」
//      与面板「完整退出 WPS 清空宿主缓存后重开…」必须红。判别点：**持久材料的读取路径**。
// M2 = 执行器拆掉「首次宿主写入前必须落盘」
//      ⇒ 同一个用例里"中断后后端确实有材料"的锚点必须红。
//      判别点：**写前落盘**，与 M1 是两条不同机制，互不代替。
//
// —— 2026-10-27 第二轮 ——
// M3 = 把「有效期内一律按持久包 expiresAt 判」改回无条件的 30 分钟检查
//      ⇒ 「a durable package older than the 30-minute cache window still resumes only the third
//      action」必须红。判别点：**有效期判据的来源**（不是身份/来源/授权那几条）。
// M4 = 服务端 `save()` 退回"同正文但已过期就重新算一个期限覆盖"
//      ⇒ 「test_expired_material_cannot_be_re_saved_into_a_new_window」与
//      「test_pruned_material_cannot_be_re_saved_past_the_first_expiry」必须红。
//      判别点：**重复保存不能续期**，与 M3 分属两端，互不代替。
//
// 四判据（不许只卡"退出码非 0"）：
//   ① 退出码非 0；② 红的那条用例**名字对得上**；③ 失败原文里带它自己的证据；
//   ④ 还原后文件逐字节 sha256 与基线一致。
//
// 用法：
//   node scripts/mutate_wps_recovery_material.mjs                 # 打印基线 sha256
//   node scripts/mutate_wps_recovery_material.mjs M1               # 应用变异
//   node scripts/mutate_wps_recovery_material.mjs M1 restore        # 还原
import { readFileSync, writeFileSync } from "node:fs";
import { createHash } from "node:crypto";
import { fileURLToPath } from "node:url";
import { dirname, resolve } from "node:path";

const here = dirname(fileURLToPath(import.meta.url));
const REPO = resolve(here, "..");
const LF = "\n";

const PACKAGE = resolve(REPO, "apps", "wps-addin", "src", "wpsRecoveryPackage.ts");
const EXECUTOR = resolve(REPO, "apps", "wps-addin", "src", "wpsPlanExecutor.ts");
const STORE = resolve(REPO, "server", "app", "execution_recovery.py");

const TARGETS = {
  M1: {
    file: PACKAGE,
    seed: [
      "  const existing = wpsUndoRecordFor(application, state.operationId);",
      "  const load = options.load ?? loadExecutionRecoveryPackage;"
    ].join(LF),
    mutant: [
      "  const existing = wpsUndoRecordFor(application, state.operationId);",
      "  if (!existing) return { kind: \"none\" };",
      "  const load = options.load ?? loadExecutionRecoveryPackage;"
    ].join(LF)
  },
  M2: {
    file: EXECUTOR,
    seed: [
      "    if (record.resume && !recovery?.resume) {",
      "      // 只在新执行时落盘：续跑时包已经在盘上（正是刚才读出来构建这份记录的），",
      "      // 重写会把 `expiresAt` 顺带刷新一次 —— 那等于\"每次继续都偷偷延长保留期\"。",
      "      if (!recovery?.persistMaterial)",
      "        throw new Error(\"这次执行需要先保存本机恢复包，但未提供持久化通道；未执行写入。\");",
      "      await recovery.persistMaterial(record);",
      "    }"
    ].join(LF),
    mutant: "    // M2: 写前落盘被拆掉（变异体，勿提交）"
  },
  M3: {
    file: EXECUTOR,
    seed: [
      "  const window = wpsResumeWindow(record, now);",
      "  if (!window.ok) throw new Error(window.reason);"
    ].join(LF),
    mutant: [
      "  const window = wpsResumeWindow(record, now);",
      "  if (!window.ok) throw new Error(window.reason);",
      "  if (now - Date.parse(record.createdAt) > capabilities.taskRecovery.wpsResumeMaxAgeSeconds * 1000)",
      "    throw new Error(\"本机恢复记录已过期或创建时间无效。\");"
    ].join(LF)
  },
  M4: {
    file: STORE,
    seed: [
      "            if deadline is None:",
      "                raise RecoveryPackageCorrupt(\"恢复包期限字段损坏，拒绝覆盖。\")",
      "            if deadline <= moment:",
      "                raise RecoveryPackageExpired(\"恢复材料已超过保留期，不能重新保存或续期。\")",
      "            # 幂等：不重写文件，首次 `savedAt`/`expiresAt` 原样返回。",
      "            return _material(header, package_json)"
    ].join(LF),
    mutant: [
      "            if deadline is None:",
      "                raise RecoveryPackageCorrupt(\"恢复包期限字段损坏，拒绝覆盖。\")",
      "            if deadline > moment:",
      "                return _material(header, package_json)"
    ].join(LF)
  },
  /*
   * —— 2026-10-27 第三轮 ——
   * M5 = 去掉 `_instant()` 的时区校验（恢复旧解析行为：naive 时间照收）
   *      ⇒ 「test_naive_expiry_reads_as_corrupt_not_a_type_error」等本轮新增用例必须红。
   *      判别点：**"没有时区"算不算一个可比较的时刻**。与 M4 分属"解析"和"续期"两条机制，
   *      互不代替：M4 打的是保存路径的期限语义，M5 打的是三条入口共用的解析器。
   */
  M5: {
    file: STORE,
    seed: [
      "    if parsed.tzinfo is None or parsed.utcoffset() is None:",
      "        return None",
      "    return parsed"
    ].join(LF),
    mutant: [
      "    # M5: 时区校验被拆掉（变异体，勿提交）",
      "    return parsed"
    ].join(LF)
  }
};

function normalize(text) {
  return text.split(/\r\n/).join(LF);
}
function denormalize(text, original) {
  const native = original.includes("\r\n") ? "\r\n" : LF;
  return text.split(LF).join(native);
}
function sha256(path) {
  return createHash("sha256").update(readFileSync(path)).digest("hex");
}

const which = process.argv[2];
if (!which) {
  for (const [id, target] of Object.entries(TARGETS)) {
    console.log(`BASELINE ${id} ${target.file} sha256=${sha256(target.file)}`);
  }
  process.exit(0);
}
if (!TARGETS[which]) {
  console.error("用法: node scripts/mutate_wps_recovery_material.mjs [M1|M2|M3|M4|M5] [restore]");
  process.exit(2);
}

const target = TARGETS[which];
const original = readFileSync(target.file, "utf8");

/*
 * 写回必须读回校验并重试：Windows 上偶发"写进去了、读回来还是旧的"。
 * 变异体留在工作区比变异失败更糟 —— 后面的回归会以"红"的面目出现，而人以为那是产品缺陷。
 */
function writeVerified(path, text) {
  const expected = normalize(text);
  let last = "";
  for (let attempt = 1; attempt <= 5; attempt += 1) {
    writeFileSync(path, text, "utf8");
    last = normalize(readFileSync(path, "utf8"));
    if (last === expected) return true;
  }
  console.error(`写回校验失败（5 次）：${path}`);
  console.error(`期望片段：${expected.slice(0, 80)}`);
  console.error(`实际片段：${last.slice(0, 80)}`);
  return false;
}

/*
 * 进程被打断也要还原：变异体泄漏进工作区，会让下一次回归以"红"的面目出现，
 * 而那与产品无关。SIGINT/TERM/HUP 与未捕获异常都装钩子。
 */
function restoreQuietly() {
  try {
    const current = normalize(readFileSync(target.file, "utf8"));
    if (!current.includes(target.mutant)) return;
    writeVerified(target.file, denormalize(current.split(target.mutant).join(target.seed), original));
    console.error(`[hook] 已还原 ${which} → ${target.file}`);
  } catch (error) {
    console.error(`[hook] 还原失败：${error instanceof Error ? error.message : String(error)}`);
  }
}
for (const signal of ["SIGINT", "SIGTERM", "SIGHUP"]) {
  process.on(signal, () => {
    restoreQuietly();
    process.exit(130);
  });
}
process.on("uncaughtException", (error) => {
  restoreQuietly();
  console.error(error);
  process.exit(1);
});

if (process.argv[3] === "restore") {
  const hits = normalize(original).split(target.mutant).length - 1;
  if (hits === 0) {
    console.log("ALREADY-CLEAN");
    process.exit(0);
  }
  const restored = denormalize(normalize(original).split(target.mutant).join(target.seed), original);
  if (!writeVerified(target.file, restored)) process.exit(4);
  console.log(`restore: OK ${target.file} sha256=${sha256(target.file)}`);
  process.exit(0);
}

const occurrences = normalize(original).split(target.seed).length - 1;
if (occurrences !== 1) {
  // SEED 未命中 = 假装有证据。宁可退出 3，也不要"跑了一条不该红的用例然后说没事"。
  console.log(`SKIP: SEED 出现 ${occurrences} 次（期望 1）于 ${target.file}`);
  process.exit(3);
}

const mutated = denormalize(normalize(original).replace(target.seed, target.mutant), original);
if (!writeVerified(target.file, mutated)) process.exit(4);
console.log(`APPLIED ${which} → ${target.file} sha256=${sha256(target.file)}`);
