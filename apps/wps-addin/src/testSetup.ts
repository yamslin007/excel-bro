/**
 * WPS 端测试的 `setupFiles`。
 *
 * 只有一件事：把**未替身的网络出口**变成失败，而不是悄悄打到用户正在用的
 * 127.0.0.1:8765。规则与原因写在 `@excel-bro/core/testNetworkGate` 里，
 * 两端（WPS / Excel）读同一份判据。
 */

import { afterEach, beforeEach } from "vitest";
import {
  describeStrayNetworkAttempts,
  installTestNetworkGate,
  resetStrayNetworkAttempts,
  strayNetworkAttempts
} from "@excel-bro/core/testNetworkGate";

installTestNetworkGate();

beforeEach(() => {
  resetStrayNetworkAttempts();
});

afterEach(() => {
  /*
   * 取一份快照再清空：失败信息要留下**这一例**的请求，不能被下一条用例的清理抹掉。
   * 这里抛错而不是 `expect(...).toEqual([])`：`beforeEach` 已经清空过，剩下的
   * 只可能是本用例产生的，直接报出通道、方法与目标更有用。
   */
  const stray = [...strayNetworkAttempts()];
  resetStrayNetworkAttempts();
  if (stray.length > 0) throw new Error(describeStrayNetworkAttempts(stray));
});
