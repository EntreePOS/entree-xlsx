export const samples = {
  data: `import { createWorkbook } from "@entree_pos/xlsx";

const workbook = createWorkbook("Menu");
const sheet = workbook.sheet();

sheet.setData([
  ["Item", "Category", "Price"],
  ["Classic Burger", "Burger", 12.5],
  ["Fries", "Side", 4],
  ["Chocolate Shake", "Drink", 5.25]
]);

await workbook.save("menu.xlsx");`,

  records: `import { createWorkbook } from "@entree_pos/xlsx";

const workbook = createWorkbook("Orders");
const sheet = workbook.sheet();

sheet.setData([
  { order: 1001, customer: "Alex", total: 28.5 },
  { order: 1002, customer: "Jordan", total: 17.25 },
  { order: 1003, customer: "Taylor", total: 42 }
]);

sheet.row(1).style({ bold: true, fill: "#E8EFFF" });
sheet.column("C").style({ numberFormat: "$#,##0.00" });
console.log(sheet.toRecords());

await workbook.save("orders.xlsx");`,

  sheets: `import { createWorkbook } from "@entree_pos/xlsx";

const workbook = createWorkbook("Products");
const products = workbook.sheet();

products.setData([
  ["Item", "Category", "Price"],
  ["Classic Burger", "Burger", 12.5],
  ["Fries", "Side", 4]
]);

const categories = workbook.addSheet("Categories", [
  ["Category", "Taxable"],
  ["Burger", true],
  ["Side", true],
  ["Drink", true]
]);

products.row(1).style({ bold: true });
categories.row(1).style({ bold: true });

await workbook.save("products.xlsx");`,

  styles: `import { createWorkbook } from "@entree_pos/xlsx";

const workbook = createWorkbook("Sales");
const sheet = workbook.sheet();

sheet.setData([
  ["Location", "Revenue"],
  ["North Market", 1840],
  ["Park Bistro", 2175],
  ["Lake Cafe", 1560]
]);

sheet.row(1)
  .style({ bold: true, color: "#FFFFFF", fill: "#2457C5" })
  .height(26);

sheet.column("B")
  .style({ numberFormat: "$#,##0.00" })
  .width(16);

sheet.column("A").width(22);
await workbook.save("sales.xlsx");`,

  namedStyles: `import { createWorkbook } from "@entree_pos/xlsx";

const workbook = createWorkbook("Revenue");
const sheet = workbook.sheet();

workbook.styles
  .define("header", { bold: true, color: "#FFFFFF", fill: "#17324D" })
  .define("money", { numberFormat: "$#,##0.00" });

sheet.setData([
  ["Location", "Revenue"],
  ["North Market", 1840],
  ["Park Bistro", 2175]
]);

sheet.row(1).style("header");
sheet.column("B").style("money");
sheet.autoFit();

await workbook.save("styled-revenue.xlsx");`,

  dates: `import { createWorkbook } from "@entree_pos/xlsx";

const workbook = createWorkbook("Performance");
const sheet = workbook.sheet();

sheet.setData([
  ["Location", "Report date", "Food cost"],
  ["North Market", new Date("2026-08-31T12:00:00Z"), 0.184],
  ["Park Bistro", new Date("2026-08-31T12:00:00Z"), 0.217],
  ["Lake Cafe", new Date("2026-08-31T12:00:00Z"), 0.196]
]);

sheet.row(1).style({ bold: true });
sheet.column("B").style({ numberFormat: "mmm d, yyyy" }).width(16);
sheet.column("C").style({ numberFormat: "0.0%" }).width(14);

await workbook.save("performance.xlsx");`,

  layout: `import { createWorkbook } from "@entree_pos/xlsx";

const workbook = createWorkbook("Weekly Sales");
const sheet = workbook.sheet();

sheet.set("A1", "Weekly Sales");
sheet.range("A1:C1").merge().style({ bold: true, fill: "#E8EFFF" });
sheet.range("A2:C5").setValues([
  ["Location", "Orders", "Revenue"],
  ["North Market", 142, 2840],
  ["Park Bistro", 168, 3375],
  ["Lake Cafe", 119, 2260]
]);

sheet.row(2).style({ bold: true });
sheet.column("C").style({ numberFormat: "$#,##0.00" });
sheet.autoFilter("A2:C5");
sheet.autoFit({ min: 10, max: 24, padding: 3 });

await workbook.save("weekly-sales.xlsx");`,

  helpers: `import { createWorkbook } from "@entree_pos/xlsx";

const workbook = createWorkbook("Inventory");
const sheet = workbook.sheet();

sheet.setData([
  ["Item", "Stock", "Status"],
  ["Classic Burger", 18, "Low"],
  ["Fries", 42, "Ready"],
  ["Chocolate Shake", 12, "Low"]
]);

const burger = sheet.column("A").find("Classic Burger");
console.log("Found at", burger?.address);

sheet.rows("2:4").height(24);
sheet.columns("A:C").width(18);
sheet.column("C").findAll("Low").forEach((cell) => {
  cell.style({ bold: true, color: "#A23B22", fill: "#FFF0E8" });
});

await workbook.save("inventory.xlsx");`,

  formula: `import { createWorkbook } from "@entree_pos/xlsx";

const workbook = createWorkbook("Order");
const sheet = workbook.sheet();

sheet.setData([
  ["Item", "Qty", "Price", "Total"],
  ["Classic Burger", 2, 12.5, null],
  ["Fries", 1, 4, null]
]);

sheet.cell("D2").formula("B2*C2", 25);
sheet.cell("D3").formula("B3*C3", 4);
sheet.cell("D4").formula("SUM(D2:D3)", 29);
sheet.set("C4", "Grand total");

sheet.range("A1:D1").style({ bold: true, fill: "#E8EFFF" });
sheet.range("C2:D4").style({ numberFormat: "$#,##0.00" });
sheet.cell("D4").style({ bold: true });
sheet.autoFit();

await workbook.save("order-total.xlsx");`
};
