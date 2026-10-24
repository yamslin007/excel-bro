import { parse, type Expression, type Statement, type Pattern, type ModuleDeclaration } from "acorn";
import capabilities from "../../../config/capabilities.json";
import { parseExplicitA1Rect } from "./planCore";

// This is an interpreter for a small JS grammar, NOT a JavaScript runtime.
// No evaluated value is a function or a host object. There is no eval fallback.
export type ScriptValue = null | boolean | number | string | ScriptValue[] | { [key: string]: ScriptValue };
export type ScriptAction = { type: string; sheet: string; range: string; [key: string]: ScriptValue };
const limits = capabilities.controlledScript;
const forbidden = new Set(["__proto__", "constructor", "prototype"]);
const fields: Record<string, readonly string[]> = {
  writeValues: ["values"],
  clearRange: ["applyTo"],
  setFont: ["bold", "italic", "name", "size", "color"],
  setAlignment: ["horizontal", "vertical", "wrapText"],
  resizeRange: ["rowHeight"]
};

/** Compiles locally; has no workbook, filesystem, network, or model access. */
export function compileControlledScript(source: string, authorizedSheets: readonly string[]): ScriptAction[] {
  if (!authorizedSheets.length) throw new Error("请先勾选允许操作的工作表");
  if (source.length > limits.maxSourceCharacters) throw new Error("脚本超过长度上限");
  const ast = parse(source, { ecmaVersion: 2022, sourceType: "script" });
  const scopes: Map<string, ScriptValue>[] = [new Map([["authorizedSheets", [...new Set(authorizedSheets)]]])];
  const actions: ScriptAction[] = [];
  let steps = 0;
  let depth = 0;
  let outputSize = 0;
  let cells = 0;
  const tick = () => { if (++steps > limits.maxSteps) throw new Error("脚本超过计算步数上限"); };
  const keySafe = (key: string) => {
    if (forbidden.has(key)) throw new Error("脚本不能访问原型或构造器");
    return key;
  };
  const bind = (id: Pattern, value: ScriptValue) => {
    if (id.type !== "Identifier" || ["authorizedSheets", "emit"].includes(id.name)) throw new Error("只支持普通变量，不能覆盖脚本接口");
    keySafe(id.name);
    const scope = scopes[scopes.length - 1];
    if (scope.has(id.name)) throw new Error("变量重复声明");
    scope.set(id.name, value);
  };
  const scalar = (value: ScriptValue): null | boolean | number | string => {
    if (value !== null && typeof value === "object") throw new Error("此运算只接受常量");
    return value;
  };
  const checked = (value: ScriptValue): ScriptValue => {
    if (typeof value === "number" && !Number.isFinite(value)) throw new Error("计算结果不是有限数值");
    if (typeof value === "string" && value.length > limits.maxStringCharacters) throw new Error("字符串超过长度上限");
    return value;
  };
  const emit = (value: ScriptValue) => {
    if (!value || Array.isArray(value) || typeof value !== "object") throw new Error("emit 需要动作对象");
    const action = value as ScriptAction;
    if (!Object.hasOwn(fields, action.type)) throw new Error("受控脚本尚不支持此动作");
    if (!authorizedSheets.includes(action.sheet)) throw new Error(`工作表「${String(action.sheet)}」未勾选，脚本未生成计划`);
    const rect = typeof action.range === "string" ? parseExplicitA1Rect(action.range) : null;
    if (!rect) throw new Error("范围必须是当前动作工作表内的明确 A1 区域");
    const allowed = new Set(["type", "sheet", "range", ...fields[action.type]]);
    if (Object.keys(action).some(key => !allowed.has(key))) throw new Error("动作含未开放的字段");
    // The only nested action value is a bounded rectangular scalar matrix.
    // Reject DAGs/objects before JSON serialization (which could expand them).
    for (const [key, field] of Object.entries(action)) if (key !== "values") scalar(field);
    if (action.type === "clearRange" && action.applyTo !== "contents") throw new Error("脚本清空操作只支持 contents");
    if (action.type === "writeValues") {
      const matrix = action.values;
      if (rect.rows * rect.columns > limits.maxTargetCells) throw new Error("脚本生成的操作超过范围上限");
      if (!Array.isArray(matrix) || matrix.length !== rect.rows || matrix.some(row => !Array.isArray(row) || row.length !== rect.columns)) throw new Error("写入矩阵与范围尺寸不一致");
      for (const row of matrix as ScriptValue[][]) for (const cell of row) {
        scalar(cell);
        // Prevent literal writes from becoming a formula / external data access.
        if (typeof cell === "string" && /^[\s]*[=+@-]/.test(cell)) throw new Error("脚本写值不支持公式或公式前缀，请使用现有公式计划通道");
      }
    }
    cells += rect.rows * rect.columns;
    if (cells > limits.maxTargetCells || actions.length >= limits.maxActions) throw new Error("脚本生成的操作超过范围或动作数量上限");
    const serialized = JSON.stringify(action);
    outputSize += serialized.length;
    if (outputSize > limits.maxOutputCharacters) throw new Error("脚本生成的计划超过大小上限");
    actions.push(JSON.parse(serialized) as ScriptAction);
  };
  const expr = (node: Expression): ScriptValue => {
    tick();
    if (++depth > limits.maxDepth) throw new Error("脚本嵌套过深");
    try {
      switch (node.type) {
        case "Literal":
          if (node.regex || node.bigint || node.value === undefined) throw new Error("不支持正则或 BigInt");
          return checked(node.value as ScriptValue);
        case "Identifier":
          keySafe(node.name);
          for (let i = scopes.length - 1; i >= 0; i--) if (scopes[i].has(node.name)) return scopes[i].get(node.name)!;
          throw new Error(`未开放的变量：${node.name}`);
        case "ArrayExpression":
          return node.elements.map(item => {
            if (!item || item.type === "SpreadElement") throw new Error("不支持空数组项或展开语法");
            return expr(item);
          });
        case "ObjectExpression": {
          const result: Record<string, ScriptValue> = Object.create(null);
          for (const prop of node.properties) {
            if (prop.type !== "Property" || prop.kind !== "init" || prop.method || prop.computed) throw new Error("只支持普通对象字段");
            const key = prop.key.type === "Identifier" ? prop.key.name : prop.key.type === "Literal" ? String(prop.key.value) : "";
            if (!key || Object.hasOwn(result, keySafe(key))) throw new Error("字段为空或重复");
            result[key] = expr(prop.value as Expression);
          }
          return result;
        }
        case "MemberExpression": {
          if (node.object.type === "Super") throw new Error("不支持宿主对象");
          const object = expr(node.object);
          const key = keySafe(String(node.computed ? scalar(expr(node.property as Expression)) : (node.property as { name: string }).name));
          if (object === null || typeof object !== "object" || !Object.hasOwn(object, key)) throw new Error("只能读取数组或对象自身已有的字段");
          return checked((object as Record<string, ScriptValue>)[key]);
        }
        case "UnaryExpression": {
          const value = scalar(expr(node.argument));
          if (node.operator === "!") return !value;
          if ((node.operator === "+" || node.operator === "-") && typeof value === "number") return checked(node.operator === "-" ? -value : value);
          throw new Error("不支持此一元运算");
        }
        case "BinaryExpression": {
          const a = scalar(expr(node.left as Expression)); const b = scalar(expr(node.right));
          if (node.operator === "===") return a === b;
          if (node.operator === "!==") return a !== b;
          if (node.operator === "+" && (typeof a === "string" || typeof b === "string")) return checked(String(a) + String(b));
          if (typeof a !== "number" || typeof b !== "number") throw new Error("此运算需要数值");
          switch (node.operator) {
            case "+": return checked(a + b); case "-": return checked(a - b); case "*": return checked(a * b);
            case "/": return checked(a / b); case "%": return checked(a % b);
            case "<": return a < b; case ">": return a > b; case "<=": return a <= b; case ">=": return a >= b;
            default: throw new Error("不支持此二元运算");
          }
        }
        case "LogicalExpression": {
          const a = expr(node.left);
          return node.operator === "&&" ? (a ? expr(node.right) : a) : node.operator === "||" ? (a || expr(node.right)) : (a ?? expr(node.right));
        }
        case "ConditionalExpression": return expr(node.test) ? expr(node.consequent) : expr(node.alternate);
        default: throw new Error(`未开放的语法：${node.type}`);
      }
    } finally { depth--; }
  };
  const statement = (node: Statement | ModuleDeclaration): void => {
    tick();
    if (++depth > limits.maxDepth) throw new Error("脚本嵌套过深");
    try {
      switch (node.type) {
        case "EmptyStatement": return;
        case "BlockStatement":
          scopes.push(new Map());
          try { node.body.forEach(statement); } finally { scopes.pop(); }
          return;
        case "VariableDeclaration":
          if (node.kind !== "const") throw new Error("变量请使用 const；不支持可变状态");
          for (const declaration of node.declarations) {
            if (!declaration.init) throw new Error("变量必须初始化");
            bind(declaration.id, expr(declaration.init));
          }
          return;
        case "ForOfStatement": {
          if (node.await || node.left.type !== "VariableDeclaration" || node.left.kind !== "const" || node.left.declarations.length !== 1 || node.left.declarations[0].init) throw new Error("循环只支持 for (const 项 of 数组)");
          const list = expr(node.right);
          if (!Array.isArray(list)) throw new Error("循环只支持有限数组");
          for (const item of list) {
            tick(); scopes.push(new Map());
            try { bind(node.left.declarations[0].id, item); statement(node.body); } finally { scopes.pop(); }
          }
          return;
        }
        case "IfStatement":
          if (expr(node.test)) statement(node.consequent); else if (node.alternate) statement(node.alternate);
          return;
        case "ExpressionStatement": {
          const call = node.expression;
          if (call.type !== "CallExpression" || call.optional || call.callee.type !== "Identifier" || call.callee.name !== "emit" || call.arguments.length !== 1 || call.arguments[0].type === "SpreadElement") throw new Error("只允许调用 emit(动作)，不允许宿主、网络或任意函数");
          emit(expr(call.arguments[0])); return;
        }
        default: throw new Error(`未开放的语法：${node.type}`);
      }
    } finally { depth--; }
  };
  ast.body.forEach(statement);
  if (!actions.length) throw new Error("脚本没有生成任何操作");
  return actions;
}
