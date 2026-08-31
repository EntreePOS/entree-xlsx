import { createWorkbook } from "@entree_pos/xlsx";

const workbook = createWorkbook("Campaigns");
const sheet = workbook.sheet();

sheet.setData([
  ["Campaign", "Start date", "Conversion", "Budget"],
  ["Lunch launch", new Date("2026-09-01T12:00:00Z"), 0.184, 1250],
  ["Fall catering", new Date("2026-10-15T12:00:00Z"), 0.126, 2400],
  ["Holiday cards", new Date("2026-11-20T12:00:00Z"), 0.219, 980]
]);

sheet.range("A1:D1").style({
  bold: true,
  color: "#FFFFFF",
  fill: "#2457C5"
});
sheet.range("B2:B4").style({ numberFormat: "mmm d, yyyy" });
sheet.range("C2:C4").style({ numberFormat: "0.0%" });
sheet.range("D2:D4").style({ numberFormat: "$#,##0" });
sheet.setColumnWidth("A", 22);
sheet.setColumnWidth("B", 18);
sheet.setColumnWidth("C", 14);
sheet.setColumnWidth("D", 14);

await workbook.save("10-format-dates-and-percentages.xlsx");
