import { createWorkbook, openWorkbook } from "@entree_pos/xlsx";

// This first block creates a small starter template so the example is runnable.
// In a real project, use the XLSX or XLSM template your team already owns.
const template = createWorkbook("Invoice");
const templateSheet = template.sheet();

templateSheet.range("A1:D1").merge();
templateSheet.cell("A1").set("INVOICE").style({
  bold: true,
  fontSize: 18,
  color: "#FFFFFF",
  fill: "#17202A",
  vertical: "center"
});
templateSheet.range("A2:D2").merge();
templateSheet.cell("A2").set("Fill the highlighted cells and keep the original design.").style({
  color: "#5F6B7A",
  italic: true,
  vertical: "center"
});

templateSheet.range("A4:D8").setValues([
  ["Item", "Quantity", "Price", "Total"],
  ["", null, null, null],
  ["", null, null, null],
  ["", null, null, null],
  ["Grand total", null, null, null]
]);
templateSheet.range("A4:D4").style({
  bold: true,
  color: "#FFFFFF",
  fill: "#2457C5",
  horizontal: "center",
  vertical: "center",
  border: { bottom: { style: "medium", color: "#17202A" } }
});
templateSheet.range("A4:D8").style({
  border: {
    outline: { style: "thin", color: "#CBD5E1" },
    inside: { style: "thin", color: "#CBD5E1" }
  }
});
templateSheet.range("A5:C7").style({ fill: "#FFF2D8" });
templateSheet.range("C5:D8").style({ numberFormat: "$#,##0.00" });
templateSheet.range("A8:D8").style({ bold: true, fill: "#EAF0FF" });
templateSheet.setColumnWidth("A", 24).setColumnWidth("B", 12);
templateSheet.setColumnWidth("C", 14).setColumnWidth("D", 15);
await template.save("09-invoice-template.xlsx");

// Open the template and change only its data and formulas.
const workbook = await openWorkbook("09-invoice-template.xlsx");
const sheet = workbook.sheet("Invoice");

sheet.range("A5:C7").setValues([
  ["Lunch catering", 12, 18.5],
  ["Coffee service", 12, 3.25],
  ["Delivery", 1, 25]
]);
sheet.cell("D5").formula("B5*C5", 222);
sheet.cell("D6").formula("B6*C6", 39);
sheet.cell("D7").formula("B7*C7", 25);
sheet.cell("D8").formula("SUM(D5:D7)", 286);

await workbook.save("09-edit-a-template.xlsx");
