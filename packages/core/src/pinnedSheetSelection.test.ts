import { describe, it, expect } from "vitest";
import { restorePinnedSelection, serializePinnedSelection } from "./pinnedSheetSelection";
describe("pinned selection identity", () => {
  const sheets = [{index:1,name:"Source"},{index:2,name:"Target"}];
  const raw = serializePinnedSelection("workbook-a", sheets, [1,2]);
  it("restores only exact metadata and contains no cell values", () => {
    expect(restorePinnedSelection(raw,"workbook-a",sheets)).toEqual({kind:"fixed",indexes:[1,2]});
    expect(Object.keys(JSON.parse(raw))).toEqual(["version","identity","sheets"]);
  });
  it("rejects another workbook, removal, rename and reorder", () => {
    expect(restorePinnedSelection(raw,"workbook-b",sheets)).toEqual({kind:"invalid"});
    for(const changed of [sheets.slice(0,1), [{index:1,name:"Other"},sheets[1]], [{index:2,name:"Source"},{index:1,name:"Target"}]])
      expect(restorePinnedSelection(raw,"workbook-a",changed)).toEqual({kind:"invalid"});
  });
  it("retains explicit empty selection and distinguishes missing from corrupt", () => {
    expect(restorePinnedSelection(serializePinnedSelection("workbook-a",sheets,[]),"workbook-a",sheets)).toEqual({kind:"fixed",indexes:[]});
    expect(restorePinnedSelection(null,"workbook-a",sheets)).toEqual({kind:"none"});
    for(const value of ["bad", "null", '{"version":9}', JSON.stringify({version:1,identity:"workbook-a",sheets:[sheets[0],sheets[0]]})])
      expect(restorePinnedSelection(value,"workbook-a",sheets)).toEqual({kind:"invalid"});
  });
});
