import { createWorkbook } from "@entree_pos/xlsx";

const workbook = createWorkbook("Private Balances");
const sheet = workbook.sheet();

sheet.setData([
  ["Account", "Balance"],
  ["River Cafe", 1250],
  ["North Market", 840],
  ["Park Bistro", 2175]
]);
sheet.range("A1:B1").style({
  bold: true,
  color: "#FFFFFF",
  fill: "#2457C5"
});
sheet.range("B2:B4").style({ numberFormat: "$#,##0.00" });
sheet.setColumnWidth("A", 22);
sheet.setColumnWidth("B", 16);

// Excel asks for this password before opening the file.
await workbook.save("15-encrypt-a-workbook.xlsx", {
  password: "demo"
});
