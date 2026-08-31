import { createWorkbook } from "@entree_pos/xlsx";

const workbook = createWorkbook("Style Guide");
const sheet = workbook.sheet();

workbook.styles
  .define("header", {
    bold: true,
    color: "#FFFFFF",
    fill: "#2457C5",
    horizontal: "center"
  })
  .define("money", {
    numberFormat: "$#,##0.00;[Red]-$#,##0.00"
  })
  .define("success", {
    bold: true,
    color: "#2E7D5B",
    fill: "#E6F4ED",
    horizontal: "center"
  })
  .define("warning", {
    bold: true,
    color: "#A15C00",
    fill: "#FFF2D8",
    horizontal: "center"
  });

sheet.range("A1:C5").setValues([
  ["Item", "Price", "Status"],
  ["Classic Burger", 12.5, "Ready"],
  ["Seasoned Fries", 4, "Ready"],
  ["Cold Brew", 3.5, "Low stock"],
  ["Chocolate Cake", 7, "Ready"]
]);
sheet.range("A1:C1").style("header");
sheet.range("B2:B5").style("money");
sheet.cell("C2").style("success");
sheet.cell("C3").style("success");
sheet.cell("C4").style("warning");
sheet.cell("C5").style("success");
sheet.setColumnWidth("A", 24);
sheet.setColumnWidth("B", 14);
sheet.setColumnWidth("C", 16);

await workbook.save("02-reusable-styles.xlsx");
