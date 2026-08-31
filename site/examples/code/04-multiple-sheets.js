import { createWorkbook } from "@entree_pos/xlsx";

const workbook = createWorkbook("Products");

const products = workbook.sheet();
products.setData([
  ["Product", "Price"],
  ["Classic Burger", 12.5],
  ["Cold Brew", 3.5]
]);

const categories = workbook.addSheet("Categories");
categories.setData([
  ["Category", "Products"],
  ["Food", 1],
  ["Drinks", 1]
]);

// Both sheets are saved in the same XLSX file.
await workbook.save("04-multiple-sheets.xlsx");
