import { createWorkbook } from "@entree_pos/xlsx";

const workbook = createWorkbook("Inventory");
const sheet = workbook.sheet();

sheet.setData([
  ["Item", "Stock"],
  ["Classic Burger", 34],
  ["Seasoned Fries", 18]
]);

// Change one existing cell.
sheet.cell("B3").set(24);

// Add one more row.
sheet.cell("A4").set("Cold Brew");
sheet.cell("B4").set(27);

await workbook.save("03-change-cells.xlsx");
