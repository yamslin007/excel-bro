import { describe, expect, it } from "vitest";
import {
  extractWorkbookDataPeriod,
  workbookNameFromDocumentUrl
} from "./workbookIdentity";

describe("workbook identity", () => {
  it("extracts and decodes the workbook name from an Office document URL", () => {
    expect(
      workbookNameFromDocumentUrl(
        "file:///C:/Users/jia/Desktop/%E5%BD%B1%E9%99%A2%E6%98%8E%E7%BB%86.xlsx"
      )
    ).toBe("影院明细.xlsx");
    expect(workbookNameFromDocumentUrl("")).toBe("未保存的工作簿");
  });

  it("extracts a single report date when both ends are the same", () => {
    expect(
      extractWorkbookDataPeriod(
        "影院放映成绩实时报表(2026-07-26到2026-07-26明细).xlsx"
      )
    ).toBe("2026-07-26");
  });

  it("extracts a report date range", () => {
    expect(
      extractWorkbookDataPeriod("影院报表_2026年7月17日至2026年7月26日.xlsx")
    ).toBe("2026-07-17 至 2026-07-26");
    expect(extractWorkbookDataPeriod("普通工作簿.xlsx")).toBeNull();
  });
});
