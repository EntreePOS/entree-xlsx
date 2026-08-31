import { createWorkbook } from "@entree_pos/xlsx";

const workbook = createWorkbook("Invoice");
const sheet = workbook.sheet();

sheet.range("A1:D1").merge();
sheet.cell("A1").set("Invoice totals").style({
  bold: true,
  fontSize: 18,
  color: "#FFFFFF",
  fill: "#17202A",
  vertical: "center"
});
sheet.setRowHeight(1, 32);
sheet.range("A2:D2").merge();
sheet.cell("A2").set("Formulas remain formulas when the workbook opens in Excel.").style({
  color: "#5F6B7A",
  italic: true,
  vertical: "center"
});
sheet.setRowHeight(2, 24);

sheet.range("A4:D8").setValues([
  ["Item", "Quantity", "Price", "Total"],
  ["Classic Burger", 2, 12.5, null],
  ["Seasoned Fries", 1, 4, null],
  ["Cold Brew", 2, 3.5, null],
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
sheet.range("A4:D8").style({
  border: {
    outline: { style: "thin", color: "#CBD5E1" },
    inside: { style: "thin", color: "#CBD5E1" }
  }
});
sheet.cell("D5").formula("B5*C5", 25);
sheet.cell("D6").formula("B6*C6", 4);
sheet.cell("D7").formula("B7*C7", 7);
sheet.cell("D8").formula("SUM(D5:D7)", 36).style({
  bold: true,
  fill: "#E6F4ED"
});
sheet.range("C5:D8").style({ numberFormat: "$#,##0.00" });
sheet.range("A8:C8").style({ bold: true, fill: "#E6F4ED" });
sheet.setColumnWidth("A", 24);
sheet.setColumnWidth("B", 12);
sheet.setColumnWidth("C", 14);
sheet.setColumnWidth("D", 15);

await workbook.save("07-formulas-and-formats.xlsx");
