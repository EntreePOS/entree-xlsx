import { createWorkbook } from "@entree_pos/xlsx";

const workbook = createWorkbook("Menu");
const sheet = workbook.sheet();

sheet.setData([
  ["Item", "Price"],
  ["Classic Burger", 12.5],
  ["Seasoned Fries", 4],
  ["Cold Brew", 3.5]
]);

// Style only the header row.
sheet.range("A1:B1").style({
  bold: true,
  color: "#FFFFFF",
  fill: "#2457C5"
});

// Format prices and make both columns readable.
sheet.range("B2:B4").style({ numberFormat: "$#,##0.00" });
sheet.setColumnWidth("A", 22);
sheet.setColumnWidth("B", 12);

await workbook.save("05-first-style.xlsx");
