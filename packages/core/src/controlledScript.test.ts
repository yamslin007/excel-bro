import { describe, expect, it } from "vitest";
import { compileControlledScript } from "./controlledScript";

const compile = (source: string, sheets = ["One", "Two"]) => compileControlledScript(source, sheets);
const font = (sheet = '"One"') => `emit({type:"setFont", sheet:${sheet}, range:"A1:B2", bold:true});`;

describe("controlled JS compiler", () => {
  it("generates all selected sheets once, with ordinary JS loops and expressions", () => {
    const result = compile(`const row = 1 + 1; for (const sheet of authorizedSheets) {
      emit({type:"setFont", sheet, range:"A" + row + ":B2", bold:sheet === "One"});
    }`);
    expect(result).toEqual([
      {type:"setFont", sheet:"One", range:"A2:B2", bold:true},
      {type:"setFont", sheet:"Two", range:"A2:B2", bold:false}
    ]);
  });
  it("can select a subset, preserving exact names", () => {
    expect(compile(`for (const sheet of authorizedSheets) { if (sheet !== "Two") { ${font("sheet")} } }`)).toHaveLength(1);
    expect(() => compile(font('"one"'))).toThrow("未勾选");
  });
  it("rejects the entire compilation if any emitted target is unauthorized", () => {
    expect(() => compile(font() + font('"Secret"'))).toThrow("未勾选");
    expect(() => compile(font(), [])).toThrow("勾选");
  });
  it("generates bounded literal matrices and contents clearing", () => {
    expect(compile(`emit({type:"writeValues",sheet:"One",range:"A1:B2",values:[[1,2 * 3],[null,"文字"]]});
      emit({type:"clearRange",sheet:"Two",range:"B2",applyTo:"contents"});`)).toHaveLength(2);
  });
  it.each([
    'Application.Run("macro")', 'eval("1")', 'fetch("https://example.com")',
    'new Function("return this")()', 'window.location', 'globalThis.process',
    'const x = authorizedSheets.constructor', 'const x = ({})["__proto__"]',
    'const x = {get sheet(){return "One"}}', 'authorizedSheets.push("Secret")',
    'for (;;) {}', 'while (true) {}', 'import("secret")',
    'const emit = 1', 'const authorizedSheets = ["Secret"]',
    'const x = {"__proto__": {sheet:"Secret"}}',
    'emit({type:"deleteWorksheet",sheet:"One",range:"A1"})',
    'emit({type:"writeFormulas",sheet:"One",range:"A1",formulas:[["=Secret!A1"]]})',
    'emit({type:"resizeRange",sheet:"One",range:"A1",columnWidth:20})',
    'emit({type:"setFont",sheet:"One",range:"Secret!A1",bold:true})',
    'emit({type:"writeValues",sheet:"One",range:"A1",values:[["=INDIRECT(\"Secret!A1\")"]]})',
    'emit({type:"writeValues",sheet:"One",range:"A1",values:[[{}]]})'
  ])("cannot escape through %s", source => { expect(() => compile(source)).toThrow(); });
  it("limits work, output and target size before returning any plan", () => {
    expect(() => compile('for(const a of authorizedSheets) for(const b of authorizedSheets) for(const c of authorizedSheets) {}', Array.from({length:40}, (_,i) => `S${i}`))).toThrow("步数");
    expect(() => compile(font().repeat(101))).toThrow("上限");
    expect(() => compile('emit({type:"setFont",sheet:"One",range:"A1:XFD1048576",bold:true})')).toThrow("上限");
    expect(() => compile(' '.repeat(32769))).toThrow("长度");
    expect(() => compile('const x = 1 / 0')).toThrow("有限数值");
    expect(() => compile(`const text = "${"x".repeat(5000)}"; for (const sheet of authorizedSheets) emit({type:"writeValues",sheet,range:"A1",values:[[text]]});`, Array.from({length:60}, (_,i) => `S${i}`))).toThrow("大小上限");
    expect(() => compile('const a = []; const b = [a,a]; emit({type:"setFont",sheet:"One",range:"A1",bold:b});')).toThrow("常量");
  });
  it("does not retain state or actions across calls", () => {
    expect(compile(font())).toHaveLength(1);
    expect(() => compile("emit(unknown)")).toThrow();
    expect(compile(font('"Two"'))).toHaveLength(1);
  });
});
