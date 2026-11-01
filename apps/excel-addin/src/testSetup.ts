/**
 * Excel 端测试的 `setupFiles`：与 WPS 端同一份闸门（`@excel-bro/core/testNetworkGate`）。
 *
 * 两端共用一个基地址来源 `api.ts` 的 `API_BASE_URL`（默认 `http://127.0.0.1:8765`），
 * 所以"测试打了用户实例"这种事故不会只发生在一边。判据只有一份，端点各自套上。
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
  const stray = [...strayNetworkAttempts()];
  resetStrayNetworkAttempts();
  if (stray.length > 0) throw new Error(describeStrayNetworkAttempts(stray));
});
