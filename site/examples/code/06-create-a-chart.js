import { createWorkbook } from "@entree_pos/xlsx";

const workbook = createWorkbook("Sales");
const sheet = workbook.sheet();

sheet.range("A1:H1").merge();
sheet.cell("A1").set("Monthly revenue").style({
  bold: true,
  fontSize: 18,
  color: "#FFFFFF",
  fill: "#17202A",
  vertical: "center"
});
sheet.setRowHeight(1, 32);
sheet.range("A2:H2").merge();
sheet.cell("A2").set("The chart references worksheet cells, so Excel can refresh it after edits.").style({
  color: "#5F6B7A",
  italic: true,
  vertical: "center"
});
sheet.setRowHeight(2, 24);

sheet.range("A4:B9").setValues([
  ["Month", "Revenue"],
  ["January", 18400],
  ["February", 21350],
  ["March", 20100],
  ["April", 24750],
  ["May", 26800]
]);
sheet.range("A4:B4").style({
  bold: true,
  color: "#FFFFFF",
  fill: "#2457C5",
  horizontal: "center",
  vertical: "center",
  border: { bottom: { style: "medium", color: "#17202A" } }
});
sheet.range("A4:B9").style({
  border: {
    outline: { style: "thin", color: "#CBD5E1" },
    inside: { style: "thin", color: "#CBD5E1" }
  }
});
sheet.range("B5:B9").style({ numberFormat: "$#,##0" });
sheet.setColumnWidth("A", 15);
sheet.setColumnWidth("B", 15);

const chart = workbook.charts.add({
  sheet: "Sales",
  name: "RevenueTrend",
  type: "column",
  title: "Monthly revenue",
  range: "A4:B9",
  position: { from: "D4", to: "K18" },
  legend: false
});

// Charts can be edited later without rebuilding their source data.
workbook.charts.update(chart.id, {
  type: "line",
  title: "Revenue trend",
  range: "A4:B9"
});

await workbook.save("06-create-a-chart.xlsx");
