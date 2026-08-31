import { createWorkbook } from "@entree_pos/xlsx";

const workbook = createWorkbook("Orders");
const sheet = workbook.sheet();

sheet.range("A1:E1").merge();
sheet.cell("A1").set("Open orders").style({
  bold: true,
  fontSize: 18,
  color: "#FFFFFF",
  fill: "#17202A",
  vertical: "center"
});
sheet.setRowHeight(1, 32);
sheet.range("A2:E2").merge();
sheet.cell("A2").set("A readable report with merged titles, filters, widths, and links.").style({
  color: "#5F6B7A",
  italic: true,
  vertical: "center"
});
sheet.setRowHeight(2, 24);

sheet.range("A4:E9").setValues([
  ["Order", "Customer", "Status", "Total", "Details"],
  [1041, "Ada Rivera", "Ready", 42.5, null],
  [1042, "Noah Patel", "Preparing", 31, null],
  [1043, "Mina Park", "Ready", 26.75, null],
  [1044, "Owen Brooks", "New", 18.5, null],
  [1045, "Lena Ortiz", "Preparing", 54, null]
]);
sheet.range("A4:E4").style({
  bold: true,
  color: "#FFFFFF",
  fill: "#2457C5",
  horizontal: "center",
  vertical: "center",
  border: { bottom: { style: "medium", color: "#17202A" } }
});
sheet.range("A4:E9").style({
  border: {
    outline: { style: "thin", color: "#CBD5E1" },
    inside: { style: "thin", color: "#CBD5E1" }
  }
});
sheet.range("D5:D9").style({ numberFormat: "$#,##0.00" });

for (let row = 5; row <= 9; row += 1) {
  const orderNumber = sheet.cell(`A${row}`).value;
  sheet.cell(`E${row}`)
    .set("Open order")
    .hyperlink(
      `https://example.com/orders/${orderNumber}`,
      "View order details"
    );
}

sheet.autoFilter("A4:E9");
sheet.setColumnWidth("A", 12);
sheet.setColumnWidth("B", 22);
sheet.setColumnWidth("C", 16);
sheet.setColumnWidth("D", 14);
sheet.setColumnWidth("E", 18);

await workbook.save("08-layout-and-filters.xlsx");
