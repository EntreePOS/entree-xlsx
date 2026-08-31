import { createWorkbook } from "@entree_pos/xlsx";

const workbook = createWorkbook("Inventory");
const sheet = workbook.sheet();

const items = [
  { sku: "BK-101", item: "Classic Burger", stock: 34 },
  { sku: "FR-204", item: "Seasoned Fries", stock: 18 },
  { sku: "DR-305", item: "Cold Brew", stock: 27 }
];

sheet.setData(items);

await workbook.save("02-export-records.xlsx");
