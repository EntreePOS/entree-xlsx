import { createWorkbook } from "@entree_pos/xlsx";

const workbook = createWorkbook("Private Report");
const sheet = workbook.sheet();

sheet.range("A1:D1").merge();
sheet.cell("A1").set("Private sales report").style({
  bold: true,
  fontSize: 18,
  color: "#FFFFFF",
  fill: "#17202A",
  vertical: "center"
});
sheet.setRowHeight(1, 32);
sheet.range("A2:D2").merge();
sheet.cell("A2").set("Yellow cells are editable. Blue cells are protected formulas.").style({
  color: "#5F6B7A",
  italic: true,
  vertical: "center"
});
sheet.setRowHeight(2, 24);

sheet.range("A4:D7").setValues([
  ["Item", "Quantity", "Price", "Total"],
  ["Lunch", 2, 12, null],
  ["Coffee", 3, 3.5, null],
  ["Grand total", null, null, null]
]);
sheet.range("A4:D4").style({
  bold: true,
  color: "#FFFFFF",
  fill: "#2457C5",
  horizontal: "center",
  vertical: "center",
  border: { bottom: { style: "medium", color: "#17202A" } }
});
sheet.range("A4:D7").style({
  border: {
    outline: { style: "thin", color: "#CBD5E1" },
    inside: { style: "thin", color: "#CBD5E1" }
  }
});
sheet.range("A5:C6").style({
  fill: "#FFF2D8",
  protection: { locked: false }
});
sheet.cell("D5").formula("B5*C5", 24).style({
  fill: "#EAF0FF",
  numberFormat: "$#,##0.00"
});
sheet.cell("D6").formula("B6*C6", 10.5).style({
  fill: "#EAF0FF",
  numberFormat: "$#,##0.00"
});
sheet.cell("D7").formula("SUM(D5:D6)", 34.5).style({
  bold: true,
  fill: "#EAF0FF",
  numberFormat: "$#,##0.00"
});
sheet.range("A7:C7").style({ bold: true, fill: "#EAF0FF" });
sheet.protectSheet({
  password: "demo",
  selectUnlockedCells: true,
  formatCells: false
});
workbook.protectStructure({ password: "demo", structure: true });
sheet.setColumnWidth("A", 22);
sheet.setColumnWidth("B", 14);
sheet.setColumnWidth("C", 14);
sheet.setColumnWidth("D", 16);

await workbook.save("14-protect-a-sheet.xlsx");
