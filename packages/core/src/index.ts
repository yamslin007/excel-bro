// 宿主无关共享层。阶段 2 起从 apps/excel-addin/src 逐个迁入纯模块，
// 分层顺序：contracts → pure → semantics → ports → orchestration → ui，不允许反向依赖。
export {};
