import { createWorkbook } from "@entree_pos/xlsx";

const workbook = createWorkbook("People");
const sheet = workbook.sheet();

sheet.setData([
  ["Name", "Age"],
  ["Mina", 28],
  ["Noah", 34]
]);

await workbook.save("01-add-data.xlsx");
