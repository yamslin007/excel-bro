/**
 * L1 纯函数：把可能出现在**用户可见 / 可持久化文本**里的凭据抹掉。
 *
 * 为什么需要它：服务端错误 message 的来源是 Python 侧的 `str(error)`，
 * 而请求里带着 `apiKey`。一旦把 `error.message` 原样写进对话历史，
 * 它就进了宿主 localStorage —— 用户可读、可复制、可被诊断包带走。
 * 时间线等地方已经各自做过一次脱敏，那些实现是模块私有的；这里把它
 * 提到共享层，让“写进任何持久文本之前先过一遍”成为统一动作。
 *
 * 只做文本替换，不解析结构、不判断上下文 —— 上层的取舍（这条消息该
 * 不该保留、要不要整条换成通用文案）仍由调用方决定。
 */

/**
 * 占位符：明确表示“这里原本有内容，被有意隐藏”，而不是静默删除。
 *
 * 用 `【已隐藏】` 这种带分隔符的结构化写法，而不是裸的“已隐藏”三个字，
 * 理由是**要能把它和普通文本区分开**：读取历史时我们只允许还原“这里原本
 * 是凭据”的措辞，如果占位符本身长得和自然语言一样，用户正常说一句
 * “这一列我先已隐藏”就会被误当成脱敏留下的标记再包一层。
 * 方括号是刻意的非自然写法 —— 用户几乎不会手打，JSON 往返也不会被转义，
 * 且不含任何凭据标签，不会被 `containsLikelyCredential` 误判。
 */
export const REDACTED_PLACEHOLDER = "【已隐藏】";

/**
 * 会被当作“凭据键”的标签。同时用于两处：
 * 1. 裸赋值形态（`apiKey=xxx`）；
 * 2. JSON / 类 JSON 的键值形态（`"apiKey": "xxx"`、`'api_key': 'xxx'`）。
 * 两个模式必须用同一份列表，否则新增一种写法时只覆盖一半。
 */
const SECRET_LABELS =
  "api[ _-]?key|apikey|access[ _-]?token|refresh[ _-]?token|auth[ _-]?token|token|secret|client[ _-]?secret|password|passwd|authorization";

const BEARER_PATTERN = /\bBearer\s+[A-Za-z0-9._~+/=-]{8,}/gi;
const BASIC_AUTH_PATTERN = /([A-Za-z][A-Za-z0-9+.-]*:\/\/)([^/\s:@]+):([^@\s/]+)@/g;
const SK_KEY_PATTERN = /\bsk-[A-Za-z0-9_-]{12,}\b/g;

/**
 * JSON / 类 JSON 的键值形态：键与值都可能在引号里，冒号后是值。
 *
 * 覆盖四种组合：`"apiKey":"v"`、`"apiKey": "v"`、`'api_key': 'v'`，以及
 * 值不带引号的 `"apiKey": v`。单引号形式来自配置文件、JS 对象字面量或
 * 用户直接粘贴的片段，JSON.parse 解析不了但同样会进历史和日志。
 *
 * 关键约束：`(?!Bearer\b)` 必须紧跟在**开引号之后**，而不是冒号之后。
 * 放在冒号后会被 `"` 挡掉判断，于是 `"Authorization":"Bearer sk-..."`
 * 里连着方案名的整段值都会被当成本体吞掉，丢掉「Bearer」这条诊断线索。
 * 值部分不含 `=`，避免把 `"key": "a=b=c"` 这类普通值误吞成凭据。
 */
const JSON_LIKE_SECRET_PATTERN = new RegExp(
  `(["']?)(${SECRET_LABELS})\\1(\\s*:\\s*)(["'])(?!Bearer\\b)([^"']*)\\4`,
  "gi"
);

// 注意与 BEARER_PATTERN 的顺序：`Authorization: Bearer <token>` 同时匹配两者，
// Bearer 先跑才能留下「Bearer 已隐藏」这种更清楚的结果；否则会被这里的
// `authorization=` 整段吞掉，变成 `Authorization=已隐藏`（信息更少）。
// `authorization` 只在没有显式 scheme 时才需要兜底，因此排除后面紧跟 Bearer 的情况。
const LABELLED_SECRET_PATTERN = new RegExp(
  `\\b(${SECRET_LABELS})\\s*[:=]\\s*(?!Bearer\\b)[^\\s,;"']+`,
  "gi"
);

/**
 * 抹掉文本里的凭据形态。
 *
 * 覆盖：`Bearer <token>`、URL 内嵌的 `user:pass@`、`sk-` 前缀的 API Key、
 * 裸标签赋值（`apiKey=...` / `token: ...`），以及 JSON / 类 JSON 的
 * 引号键值（`"apiKey":"..."`、`'api_key': '...'`）。
 *
 * 匹配刻意从宽：宁可多隐藏一段普通文本，也不让凭据漏出去。但对
 * 「方案名」与「无标签的普通文本」保持克制 —— 前者保留可诊断信息，
 * 后者不动，避免把错误文案改到用户读不懂。
 *
 * 只做文本替换，不解析结构、不判断上下文；上层的取舍（这条消息该不该
 * 保留、要不要整条换成通用文案）仍由调用方决定。
 */
export function redactCredentials(text: string): string {
  return (
    text
      // 1) 显式 scheme 优先：留下方案名，只换掉 token 本体。
      .replace(BEARER_PATTERN, `Bearer ${REDACTED_PLACEHOLDER}`)
      // 2) URL 内嵌凭据：保留 scheme、主机与路径。
      .replace(BASIC_AUTH_PATTERN, `$1${REDACTED_PLACEHOLDER}@`)
      // 3) 已知前缀的 API Key：无需标签，形态本身就够特征。
      .replace(SK_KEY_PATTERN, REDACTED_PLACEHOLDER)
      // 4) 引号键值：保留键名与引号，只换值。
      .replace(
        JSON_LIKE_SECRET_PATTERN,
        (_match, keyQuote: string, label: string, colon: string, valueQuote: string) =>
          `${keyQuote}${label}${keyQuote}${colon}${valueQuote}${REDACTED_PLACEHOLDER}${valueQuote}`
      )
      // 5) 裸标签赋值兜底：保留标签名，只换值。
      .replace(LABELLED_SECRET_PATTERN, `$1=${REDACTED_PLACEHOLDER}`)
  );
}

/** 文本里是否还残留疑似凭据；用于“脱敏之后再确认一次”的守卫与断言。 */
export function containsLikelyCredential(text: string): boolean {
  return redactCredentials(text) !== text;
}
