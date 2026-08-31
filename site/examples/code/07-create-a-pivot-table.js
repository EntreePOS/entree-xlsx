import { createWorkbook } from "@entree_pos/xlsx";

const workbook = createWorkbook("Orders");
const source = workbook.sheet();
const summary = workbook.addSheet("Summary");

source.setData([
  { region: "North", category: "Drinks", sales: 180 },
  { region: "North", category: "Food", sales: 420 },
  { region: "South", category: "Drinks", sales: 220 },
  { region: "South", category: "Food", sales: 510 },
  { region: "West", category: "Drinks", sales: 260 },
  { region: "West", category: "Food", sales: 390 }
]);
source.range("A1:C1").style({
  bold: true,
  color: "#FFFFFF",
  fill: "#2457C5",
  horizontal: "center",
  vertical: "center",
  border: { bottom: { style: "medium", color: "#17202A" } }
});
source.range("A1:C7").style({
  border: {
    outline: { style: "thin", color: "#CBD5E1" },
    inside: { style: "thin", color: "#CBD5E1" }
  }
});
source.range("C2:C7").style({ numberFormat: "$#,##0" });
source.autoFit({ min: 12, max: 22, padding: 3 });

summary.range("A1:F1").merge();
summary.cell("A1").set("Sales by region").style({
  bold: true,
  fontSize: 18,
  color: "#FFFFFF",
  fill: "#17202A",
  vertical: "center"
});
summary.setRowHeight(1, 32);
summary.range("A2:F2").merge();
summary.cell("A2").set("Native PivotTable with cached values and refresh-on-open support.").style({
  color: "#5F6B7A",
  italic: true,
  vertical: "center"
});
summary.setRowHeight(2, 24);

workbook.pivotTables.add({
  name: "SalesByRegion",
  source: { sheet: "Orders", range: "A1:C7" },
  target: { sheet: "Summary", cell: "A4" },
  rows: ["region"],
  columns: ["category"],
  filters: [],
  values: [
    { field: "sales", summarize: "sum", name: "Total sales" }
  ],
  showGrandTotals: true,
  refreshOnLoad: true,
  style: "PivotStyleMedium9"
});

summary.setColumnWidth("A", 28);
summary.setColumnWidth("B", 18);
summary.setColumnWidth("C", 18);
summary.setColumnWidth("D", 22);

await workbook.save("07-create-a-pivot-table.xlsx");
