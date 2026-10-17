/**
 * 反向变异驱动：证明「WPS 的工作目录（第三种数据来源）」这一轮的判据**真有判别力**。
 *
 * 本轮改的是三处**互相独立**的机制，所以一个变异可能只动 WPS、只动 core，也可能两端都动
 * （W5 改的是共享层的 EB 名单，两端都读它）。因此这里**每个变异都要跑两个套件**：
 * 声明红的那个套件必须恰好红在声明的用例上，另一个套件必须一条都不红 ——
 * 「不该红的也红了」和「该红的没红」一样是失败。
 *
 * 四条判据缺一不算过：
 * 1. 未变异时**两个套件**都基线 GREEN（退出码 0 + 有效测试摘要 + 用例数达标）；
 * 2. 每个变异 apply 后必须**按声明的样子**变红：`failing`（恰好这几条）+ `errorContains`
 *    （失败原文里出现**它自己那一处**的证据）。为什么不能只看"退出码非 0"：
 *    这一轮改的是 JSX 与判据函数，写崩了 jsdom 一样会让 vitest 非 0 退出 ——
 *    那是运行期错误冒充判红。所以还要求失败原文里出现该机制**自己**的判据文本。
 * 3. restore 后三个目标文件**逐字节一致**（sha256 比对）；
 * 4. 还原后**两个套件**复跑全绿。
 *
 * 变异自身不许靠"引用一个已删除的标识符"变红（W1/W5/W6/W7 都是就地改值或就地内联，
 * 不依赖已被删掉的 `extractWorkbookDataPeriod` 之类的符号）。
 *
 * 用法：node scripts/run_wps_work_directory_mutations.mjs
 */
import { spawnSync } from "node:child_process";
import { openSync, closeSync, readFileSync, mkdirSync, rmSync } from "node:fs";
import { createHash } from "node:crypto";
import { fileURLToPath } from "node:url";
import { dirname, resolve } from "node:path";

const here = dirname(fileURLToPath(import.meta.url));
const repoRoot = resolve(here, "..");
const scratchDir = resolve(repoRoot, "outputs");
const scratch = resolve(scratchDir, "wps-work-directory-mutation.log");
const vitest = resolve(repoRoot, "node_modules", "vitest", "vitest.mjs");
const mutateScript = resolve(here, "mutate_wps_work_directory.mjs");

const WPS_DIR = resolve(repoRoot, "apps", "wps-addin");
const CORE_DIR = resolve(repoRoot, "packages", "core");

/** 三个被变异的文件都要核 sha256（一个变异可能只动其中一个，还原后三个都不许变）。 */
const TARGET_FILES = [
  resolve(WPS_DIR, "src", "WpsProbePanel.tsx"),
  resolve(WPS_DIR, "src", "WpsAgentPanel.tsx"),
  resolve(CORE_DIR, "src", "folderScope.ts")
];

const WPS_SUITE = {
  name: "wps",
  cwd: WPS_DIR,
  // `wpsScopeDraft.test.tsx` 也在里面：这一轮在滚动容器与动作条之间插进了工作目录页脚与
  // 承认汇总，那条"动作条住在弹层上、不在滚动容器里"的判据必须**换个来源也成立**，
  // 所以它的工作目录分支（W8 钉的正是它）也属于本轮的判据面。
  files: [
    "src/wpsWorkDirectoryScope.test.tsx",
    "src/wpsWorkDirectorySendGate.test.tsx",
    "src/wpsScopeDraft.test.tsx"
  ],
  /** 13 + 3 + 17：少一条就说明判据被删。 */
  minTests: 33
};
const CORE_SUITE = {
  name: "core",
  cwd: CORE_DIR,
  files: ["src/folderScope.test.ts"],
  minTests: 14
};

/* 用例短名 —— 与两个测试文件里的 it(...) 标题逐字一致（取够区分的片段）。 */
const CARD_OPEN = "「工作目录」不再禁用";
const TWO_LEVEL_PICK = "文件/表两级勾选；勾选本身不是授权";
const REFRESH_TARGET_GONE = "已勾的表在刷新后消失";
const NEW_ROOT = "换成另一个根目录";
const DRAFT_CANCEL = "取消把草稿退回已提交的勾选";
const CLEAR_DRAFT_ONLY = "清空」只清草稿";
const TRIGGER_TEXT = "触发器在工作目录模式下只说工作目录";
const SEND_ZERO_OUTBOUND = "零出站、零工作簿读取";
const SELECTABLE_ONLY = "可勾选名单只排除 EB 系统遗留表";
const EXISTENCE_UNFILTERED = "存在性名单不过滤 EB 表";
const ADDED_ONLY = "只报新增的可勾选工作表";
const LAYER_SIBLING = "动作条是滚动容器的兄弟";
const FOLDER_FOOTER_LAYER = "页脚住在弹层上、不在滚动容器里";

/**
 * 变异 → 两个套件各自应当（且仅应当）红在哪几条，以及失败原文里必须出现的证据。
 * `failing: []` 表示"这个套件必须一条都不红"。
 */
const MUTATIONS = {
  W1: {
    note: "把「WPS 当前仅支持打开的工作簿」请回卡片副标题（能力说明复活）",
    wps: { failing: [CARD_OPEN], evidence: "WPS 当前仅支持打开的工作簿" },
    core: { failing: [], evidence: null }
  },
  W2: {
    note: "勾选即提交授权（草稿 / 已提交的分离被拆掉）",
    wps: { failing: [DRAFT_CANCEL, CLEAR_DRAFT_ONLY], evidence: "'数据|2'" },
    core: { failing: [], evidence: null }
  },
  W3: {
    note: "刷新被当成同意（目标消失也不再作废授权）",
    wps: { failing: [REFRESH_TARGET_GONE], evidence: "expected '数据|1' to be 'none'" },
    core: { failing: [], evidence: null }
  },
  W4: {
    note: "换目录不作废旧授权",
    // 与 W3 红在**不同的用例**上（W3 走刷新路径、W4 走换目录路径），两条用例名互不为子串
    // ⇒ 条数相等 + 每条在场就已经把它们分开了。
    // 证据串是实测来的，不是推的：换根之后旧根那份授权还活着，于是面板把**旧根上已确认的
    // 表数**贴到了**新目录的名字**上 —— 收到的是 `'别的数据|1'` 而不是 `'数据|1'`。
    // 这恰好就是这条变异要暴露的缺陷（1 张表来自旧根，名字却是新目录）。
    wps: { failing: [NEW_ROOT], evidence: "expected '别的数据|1' to be 'none'" },
    core: { failing: [], evidence: null }
  },
  W5: {
    note: "EB 系统遗留表也算可勾选（共享层的名单被清空）",
    wps: { failing: [TWO_LEVEL_PICK], evidence: "'0/2'" },
    core: {
      failing: [SELECTABLE_ONLY, EXISTENCE_UNFILTERED, ADDED_ONLY],
      evidence: "#EB_RULES"
    }
  },
  W6: {
    note: "发送闸读不到来源（等于永远放行 → 偷用当前工作簿）",
    wps: { failing: [SEND_ZERO_OUTBOUND], evidence: "to deeply equal []" },
    core: { failing: [], evidence: null }
  },
  W7: {
    note: "触发器在工作目录模式下顺带显示当前工作簿名",
    // 这条同时打破两个调用点：触发器正文里多出工作簿名，以及 `<small>` 重新出现。
    wps: { failing: [REFRESH_TARGET_GONE, TRIGGER_TEXT], evidence: "not to contain '测试.xlsx'" },
    core: { failing: [], evidence: null }
  },
  C1: {
    note: "存在性名单也过滤 EB 表（把「两个判据」合成一个）",
    wps: { failing: [], evidence: null },
    core: { failing: [EXISTENCE_UNFILTERED], evidence: "#EB_RULES" }
  },
  W8: {
    note: "工作目录页脚与承认汇总被卷进滚动容器里（工作簿动作条仍在容器外）",
    /*
      这条专门证明**工作目录那半条**层次判据有判别力。
      E2（`mutate_scope_action_bar_layer.mjs`）把动作条本身卷进滚动容器，于是
      `actions.parentElement === popover` 那一句先红，工作目录那半句永远轮不到 ——
      等于它从没被证明过。W8 只动工作目录侧（页脚 + 承认汇总），动作条层面保持正确，
      所以红的只能是"工作目录那两段也在滚动容器外面"这一组 —— 两个调用点各红一条：
      `wpsScopeDraft.test.tsx` 的工作簿来源那半句不受影响、工作目录那半句红；
      `wpsWorkDirectoryScope.test.tsx` 新加的「固定页脚的层与格式」整条红。
      证据串 `folder-selection-footer` 来自失败原文里 vitest 完整打印的那个元素
      （`class="folder-selection-footer"`）—— 光看 `<div …(2)>` 是认不出谁被卷进去的。
    */
    wps: {
      failing: [LAYER_SIBLING, FOLDER_FOOTER_LAYER],
      evidence: "folder-selection-footer"
    },
    core: { failing: [], evidence: null }
  }
};

const SUITES = [WPS_SUITE, CORE_SUITE];

mkdirSync(scratchDir, { recursive: true });

/**
 * 用**文件描述符**而不是 pipe 接子进程输出：流水线里创建管道会被拦，
 * 而判据又必须读到真实的摘要行与失败原文。vitest 输出带 ANSI 色码，先剥掉。
 */
function runSuite(suite) {
  const fd = openSync(scratch, "w");
  const result = spawnSync(process.execPath, [vitest, "run", ...suite.files], {
    cwd: suite.cwd,
    stdio: ["ignore", fd, fd]
  });
  closeSync(fd);
  const raw = readFileSync(scratch, "utf8").replace(/\x1b\[[0-9;]*m/g, "");
  const summary =
    raw.split(/\r?\n/).find((line) => /Tests\s+\d+ (passed|failed)/.test(line)) ?? "";
  const passed = /Tests\s+(?:(\d+) failed \| )?(\d+) passed/.exec(summary);
  const failed = /Tests\s+(\d+) failed/.exec(summary);
  // 文件级清单里失败用例的标记是 `×`（ANSI 已剥离），拿它的短名做比对。
  const failingTitles = raw
    .split(/\r?\n/)
    .filter((line) => /^\s*×\s+/.test(line))
    .map((line) => line.trim());
  return {
    status: result.status,
    summary: summary.trim(),
    passed: passed ? Number(passed[2]) : 0,
    failed: failed ? Number(failed[1]) : 0,
    failingTitles,
    raw
  };
}

/** 某条短名是不是**出现在失败清单里**（而不是只出现在别的什么提示文本里）。 */
function isFailing(run, shortName) {
  return run.failingTitles.some((line) => line.includes(shortName));
}

function sha256(path) {
  return createHash("sha256").update(readFileSync(path)).digest("hex");
}

function snapshot() {
  return Object.fromEntries(TARGET_FILES.map((file) => [file, sha256(file)]));
}

function mutate(which, restore) {
  const args = [mutateScript, which, ...(restore ? ["restore"] : [])];
  const fd = openSync(scratch, "w");
  const result = spawnSync(process.execPath, args, {
    cwd: repoRoot,
    stdio: ["ignore", fd, fd]
  });
  closeSync(fd);
  return { status: result.status, out: readFileSync(scratch, "utf8").trim() };
}

const shortName = (file) => file.split(/[\\/]/).pop();

const baselineShas = snapshot();
for (const file of TARGET_FILES) console.log(`[mutate] 基线 ${shortName(file)} sha256=${baselineShas[file]}`);

/* 前置状态闸：九个变异都必须完整在位或完整缺失。 */
console.log("[mutate] 先跑变异串定点闸…");
const preflight = mutate("check", false);
if (preflight.status !== 0) {
  console.error(preflight.out);
  console.error("[mutate] 变异串已失配，拒绝继续（先修 scripts/mutate_wps_work_directory.mjs 的种子）。");
  process.exit(1);
}
console.log(preflight.out.split(/\r?\n/).filter(Boolean).join("  |  "));

console.log("[mutate] 基线（未变异）两个套件都应为 GREEN…");
const baseline = {};
for (const suite of SUITES) {
  const run = runSuite(suite);
  baseline[suite.name] = run;
  if (run.status !== 0 || !run.summary || run.passed < suite.minTests || run.failed !== 0) {
    console.error(
      `基线就不绿：${suite.name} status=${run.status}, 摘要=${run.summary || "无"}`
    );
    process.exit(1);
  }
  console.log(`[mutate] 基线 GREEN ${suite.name}：${run.summary}`);
}

let failures = 0;
for (const [which, spec] of Object.entries(MUTATIONS)) {
  const applied = mutate(which, false);
  if (applied.status !== 0) {
    console.error(`[mutate] ${which} apply 失败：status=${applied.status} out=${applied.out}`);
    failures += 1;
    continue;
  }
  // 套件之间的判定是**独立**的：一个套件不合声明，不影响另一个套件照样取证。
  for (const suite of SUITES) {
    const declared = spec[suite.name];
    const run = runSuite(suite);
    // 「红了几条」+「红的正是声明的那几条」：条数相等 + 每条都在场 ⇒ 失败集合恰好等于声明集合，
    // 不会出现"该红的没红、不该红的替它红了"。（短名互不为子串，比对上不会串。）
    const titlesOk =
      declared.failing.length === 0
        ? run.status === 0 && run.failed === 0
        : run.status !== 0 &&
          run.failed === declared.failing.length &&
          declared.failing.every((name) => isFailing(run, name));
    // 只有当声明里真有失败用例时才要求证据；`failing: []` 的套件本来就没有失败原文可读。
    const evidenceOk = declared.evidence === null || run.raw.includes(declared.evidence);
    const verdict = titlesOk && evidenceOk ? "OK" : "NG!";
    console.log(
      `[mutate] ${which}(${suite.name}) → RED(${verdict}) ${run.summary || "(无摘要)"}｜` +
        `失败清单=${run.failingTitles.length}${verdict === "OK" ? "" : "  ← " + spec.note}`
    );
    if (!titlesOk || !evidenceOk) {
      console.error(
        `  ${which}/${suite.name} 期望恰好红在 ${declared.failing.length} 条 ` +
          `[${declared.failing.join(" / ")}]` +
          (declared.evidence ? `、失败原文含「${declared.evidence}」` : "") +
          `；实际 status=${run.status} failed=${run.failed} 清单=[${run.failingTitles.join(" | ")}]` +
          ` evidence=${evidenceOk}；见 ${scratch}`
      );
      failures += 1;
    }
  }
  const restored = mutate(which, true);
  if (restored.status !== 0 || restored.out === "ALREADY-CLEAN") {
    console.error(`[mutate] ${which} restore 失败：${restored.out}`);
    failures += 1;
    continue;
  }
  const after = snapshot();
  for (const file of TARGET_FILES) {
    if (after[file] !== baselineShas[file]) {
      console.error(`[mutate] ${which} 还原后 ${shortName(file)} 与基线不一致（sha256 变了）`);
      failures += 1;
    }
  }
}

const finalShas = snapshot();
for (const file of TARGET_FILES) {
  if (finalShas[file] !== baselineShas[file]) {
    console.error(
      `[mutate] 最终 ${shortName(file)} 与基线不一致：${finalShas[file]} != ${baselineShas[file]}`
    );
    failures += 1;
  }
}

console.log("[mutate] 还原后两个套件复跑应为 GREEN…");
for (const suite of SUITES) {
  const after = runSuite(suite);
  if (after.status !== 0 || after.passed < suite.minTests || after.failed !== 0) {
    console.error(`[mutate] 还原后 ${suite.name} 未全绿：${after.summary}`);
    failures += 1;
  } else {
    console.log(`[mutate] 还原后 GREEN ${suite.name}：${after.summary}`);
  }
}
rmSync(scratch, { force: true });

if (failures > 0) {
  console.error(`反向变异未通过：${failures} 处问题。`);
  process.exit(1);
}
console.log(
  "反向变异通过：W1→卡片文案、W2→勾选即提交、W3/W4→刷新与换目录各自作废授权、" +
    "W5→EB 名单（两端同时红）、W6→发送闸、W7→触发器、C1→存在性名单、" +
    "W8→工作目录那半条层次判据；每个变异只红它声明的那几条，还原逐字节一致，两个套件复跑全绿。"
);
