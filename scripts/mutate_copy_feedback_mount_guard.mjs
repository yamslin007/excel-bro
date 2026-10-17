// 反向变异脚本：复制反馈的「挂载期开关」与可见文案（2026-10-16）
//
// 背景：Excel 任务窗格入口用 `React.StrictMode` 包 `<App/>`，开发构建下 React 会
// **有意**重放 effect（setup → cleanup → setup）。`useCopyFeedback` 早期的写法只在
// cleanup 里把 `mountedRef` 置 false、不在 setup 里置回 true，于是重放之后这个开关
// **永久**停在 false —— 点复制时"写『已复制』状态"那一步被它拦掉，界面表现是
// "点复制毫无反应"。WPS 面板不走这个 hook，所以只有 Excel 端出问题。
//
// 用法：node scripts/mutate_copy_feedback_mount_guard.mjs [M1|M2|M3] [restore]
// 不带参数：只打印各目标文件的基线 sha256。
// M1 = 去掉 setup 里的「重新武装」（退回"只写不读"的挂载期开关）
// M2 = 把消息复制的自动清除窗口从 1600ms 拉长到 3600ms
// M3 = 按钮不再渲染「已复制」这三个字（只剩图标）
import { readFileSync, writeFileSync } from "node:fs";
import { createHash } from "node:crypto";
import { fileURLToPath } from "node:url";
import { dirname, resolve } from "node:path";

const here = dirname(fileURLToPath(import.meta.url));
const REPO = resolve(here, "..");
const LF = "\n";

const HOOK = resolve(REPO, "apps", "excel-addin", "src", "hooks", "useCopyFeedback.ts");
const BUTTON = resolve(
  REPO,
  "apps",
  "excel-addin",
  "src",
  "components",
  "MessageCopyButton.tsx"
);

const TARGETS = {
  M1: {
    file: HOOK,
    seed: ["      mountedRef.current = true;", "      return () => {"].join(LF),
    mutant: [
      "      // [变异M1] 只在 cleanup 里置 false，不在 setup 里重新武装",
      "      return () => {"
    ].join(LF)
  },
  M2: {
    file: HOOK,
    seed: "      }, 1600);",
    mutant: "      }, 3600); // [变异M2] 清除窗口被拉长"
  },
  M3: {
    file: BUTTON,
    seed: ["          </svg>", "          <span>已复制</span>", "        </>"].join(LF),
    mutant: ["          </svg>", "        </>"].join(LF)
  }
};

/** 目标文件自带换行（CRLF/LF）都要先归一化，否则 includes/split 会误报 ALREADY-CLEAN。 */
function normalize(text) {
  return text.split(/\r\n/).join(LF);
}
/** 写回必须还原**原文件自己的**换行，否则"逐字节一致"这条判据会自己失败。 */
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
  console.error(
    "用法: node scripts/mutate_copy_feedback_mount_guard.mjs [M1|M2|M3] [restore]"
  );
  process.exit(2);
}

const target = TARGETS[which];
const original = readFileSync(target.file, "utf8");

if (process.argv[3] === "restore") {
  const hits = normalize(original).split(target.mutant).length - 1;
  if (hits === 0) {
    console.log("ALREADY-CLEAN");
    process.exit(0);
  }
  writeFileSync(
    target.file,
    denormalize(normalize(original).split(target.mutant).join(target.seed), original),
    "utf8"
  );
  console.log(`restore: OK ${target.file}`);
  process.exit(0);
}

const occurrences = normalize(original).split(target.seed).length - 1;
if (occurrences !== 1) {
  console.log(`SKIP: SEED 出现 ${occurrences} 次（期望 1）于 ${target.file}`);
  process.exit(3);
}

writeFileSync(
  target.file,
  denormalize(normalize(original).replace(target.seed, target.mutant), original),
  "utf8"
);
console.log(`APPLIED ${which} → ${target.file}`);
