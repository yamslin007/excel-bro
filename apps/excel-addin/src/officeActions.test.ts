/**
 * 锁定 Office 自报动作清单与真实执行入口的一致性。
 *
 * 两条红线：
 * 1. 不能漏报 —— 清单少一个动作，该动作永远进不了计划，功能静默消失。
 * 2. 不能虚报 —— 清单多一个动作，规划前交集会放它过去，但 `executeAction`
 *    里没有分支，用户会在确认后才看到"未处理的动作"报错。评审最怕后者。
 *
 * 因此这里双向比对：`ExcelAction` 全体成员 ↔ `executeAction` 的 switch 分支
 * ↔ `OFFICE_EXECUTED_ACTION_TYPES`。三者必须完全一致。
 */

import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { dirname, resolve } from "node:path";
import { OFFICE_EXECUTED_ACTION_TYPES } from "./officeActions";

const here = dirname(fileURLToPath(import.meta.url));
const contractsSource = readFileSync(resolve(here, "contracts.ts"), "utf8");
const excelSource = readFileSync(resolve(here, "excel.ts"), "utf8");

/** 从 `export type ExcelAction = ...` 的联合体里抓动作名。 */
function excelActionUnionMembers(): string[] {
  const start = contractsSource.indexOf("export type ExcelAction =");
  if (start < 0) throw new Error("未能在 contracts.ts 中定位 ExcelAction 联合体");
  // 联合体在下一次顶层 `export ` 前结束（不分隔空行，所以不能用 \n\n 定位）。
  const nextExport = contractsSource.indexOf("\nexport ", start + 30);
  const body = contractsSource.slice(
    start,
    nextExport < 0 ? undefined : nextExport
  );
  const names = [...body.matchAll(/type:\s*"(\w+)"/g)].map((entry) => entry[1]);
  const unique = [...new Set(names)];
  if (unique.length === 0) {
    throw new Error("在 ExcelAction 联合体中未抓到任何动作名");
  }
  return unique;
}

/**
 * 抓 `executeAction` 覆盖的动作名。
 *
 * 两种分支形态都要收：主 switch 的 `case "xxx":`，以及函数开头针对
 * 需要提前 sync 的动作写的 `if (action.type === "xxx") { ...; return; }`
 * ——`deleteWorksheet` / `computeValues` / `splitGroupAggregate` 走的是后者。
 * 只看 switch 会误报这三个为"虚报"，反过来只看 if 会漏掉其余。
 */
function executedActionTypes(): string[] {
  const start = excelSource.indexOf("export async function executeAction(");
  if (start < 0) throw new Error("未能在 excel.ts 中定位 executeAction");
  const rest = excelSource.slice(start);
  const endMatch = /\n}\n/.exec(rest);
  const body = endMatch ? rest.slice(0, endMatch.index) : rest;
  const cases = [...body.matchAll(/case\s+"(\w+)":/g)].map((entry) => entry[1]);
  const early = [...body.matchAll(/action\.type\s*===\s*"(\w+)"/g)].map(
    (entry) => entry[1]
  );
  return [...new Set([...cases, ...early])];
}

describe("Office 自报动作清单", () => {
  it("与 ExcelAction 联合体成员完全一致", () => {
    expect([...OFFICE_EXECUTED_ACTION_TYPES].sort()).toEqual(
      excelActionUnionMembers().sort()
    );
  });

  it("不虚报：清单里每个动作都在 executeAction 中有执行分支", () => {
    const cases = new Set(executedActionTypes());
    const missing = OFFICE_EXECUTED_ACTION_TYPES.filter(
      (action) => !cases.has(action)
    );
    expect(missing).toEqual([]);
  });

  it("清单无重复项", () => {
    expect(new Set(OFFICE_EXECUTED_ACTION_TYPES).size).toBe(
      OFFICE_EXECUTED_ACTION_TYPES.length
    );
  });

  it("快照：37 个动作（契约不变时这个数字不该动）", () => {
    expect(OFFICE_EXECUTED_ACTION_TYPES).toHaveLength(
      excelActionUnionMembers().length
    );
    expect(OFFICE_EXECUTED_ACTION_TYPES).toHaveLength(37);
  });
});
