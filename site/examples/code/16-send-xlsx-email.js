// npm install @entree_pos/xlsx nodemailer
import { createWorkbook } from "@entree_pos/xlsx";
import nodemailer from "nodemailer";

const workbook = createWorkbook("Daily Sales");
const sheet = workbook.sheet();

sheet.setData([
  ["Item", "Quantity", "Sales"],
  ["Classic Burger", 24, 300],
  ["Seasoned Fries", 18, 72],
  ["Cold Brew", 15, 67.5]
]);
sheet.row(1).style({ bold: true, fill: "#2457C5", color: "#FFFFFF" });
sheet.column("C").style({ numberFormat: "$#,##0.00" });
sheet.autoFit();

// Keep the XLSX in memory. Nodemailer accepts a Buffer as attachment content.
const xlsx = workbook.toBuffer();

const mailer = nodemailer.createTransport({
  host: process.env.SMTP_HOST,
  port: Number(process.env.SMTP_PORT ?? 587),
  secure: process.env.SMTP_PORT === "465",
  auth: {
    user: process.env.SMTP_USER,
    pass: process.env.SMTP_PASSWORD
  }
});

await mailer.sendMail({
  from: process.env.EMAIL_FROM,
  to: "manager@example.com",
  subject: "Daily sales report",
  text: "The daily sales workbook is attached.",
  attachments: [{
    filename: "daily-sales.xlsx",
    content: xlsx,
    contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  }]
});
