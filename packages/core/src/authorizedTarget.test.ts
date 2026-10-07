import { describe, expect, it } from "vitest";
import { checkAuthorizedTarget, type AuthorizedTarget } from "./authorizedTarget";
const bound: AuthorizedTarget = { fileKey: "file:a", name: "Same.xlsx", sheets: [{ id: "one", name: "Sheet1" }], selected: ["one"] };
describe("early authorized target detection", () => {
  it("does not confuse identical file names with identical targets", () => {
    expect(checkAuthorizedTarget(bound, { ...bound, fileKey: "file:b" })).toBe("file_changed");
  });
  it("detects unavailable host without asserting that data changed", () => {
    expect(checkAuthorizedTarget(bound, null)).toBe("unavailable");
  });
  it.each([
    { ...bound, selected: ["two"] },
    { ...bound, sheets: [{ id: "two", name: "Sheet1" }] },
    { ...bound, sheets: [{ id: "one", name: "Renamed" }] },
    { ...bound, sheets: [] }
  ])("detects selected sheet change/removal/replacement", (current) => {
    expect(checkAuthorizedTarget(bound, current)).toBe("scope_changed");
  });
  it("does not include unselected new sheets in authorization", () => {
    expect(checkAuthorizedTarget(bound, { ...bound, sheets: [...bound.sheets, { id: "two", name: "Extra" }] })).toBe("same");
  });
  it("requires same-session reference for unsaved identity", () => {
    const ref = {};
    const unsaved = { ...bound, fileKey: null, fileRef: ref };
    expect(checkAuthorizedTarget(unsaved, { ...unsaved })).toBe("same");
    expect(checkAuthorizedTarget(unsaved, { ...unsaved, fileRef: {} })).toBe("unavailable");
  });
});
