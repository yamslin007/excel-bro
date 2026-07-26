import type { WorkbookSnapshot } from "./contracts";

export const demoWorkbook: WorkbookSnapshot = {
  name: "销售分析演示.xlsx",
  capturedAt: new Date().toISOString(),
  activeWorksheet: "7月销售",
  worksheets: [
    {
      name: "6月销售",
      usedRange: "A1:D5",
      rowCount: 5,
      columnCount: 4,
      headers: ["产品编码", "区域", "数量", "销售额"],
      dataRows: [
        ["A001", "华东", 10, 12000],
        ["A002", "华南", 8, 9800],
        ["A003", "华北", 6, 7200],
        ["A004", "华东", 4, 5100]
      ],
      truncated: false
    },
    {
      name: "7月销售",
      usedRange: "A1:D6",
      rowCount: 6,
      columnCount: 4,
      headers: ["产品编码", "区域", "数量", "销售额"],
      dataRows: [
        ["A001", "华东", 12, 14500],
        ["A002", "华南", 5, 6100],
        ["A003", "华北", 9, 10800],
        ["A005", "西南", 7, 8300],
        ["A006", "华东", 3, 3900]
      ],
      truncated: false
    }
  ]
};
