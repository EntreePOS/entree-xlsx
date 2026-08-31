import assert from "node:assert/strict";
import test from "node:test";
import { createWorkbook, decryptWorkbookBuffer, encryptWorkbookBuffer, isCompoundFile, parseWorkbook } from "../src/index.js";
import { readCompoundFile } from "../src/cfb.js";

test("encrypts and decrypts an XLSX package using Agile Office encryption", () => {
  const workbook = createWorkbook("Private");
  workbook.sheet().setData([["Secret"], ["restaurant report"]]);
  const plain = workbook.toBuffer();
  const encrypted = encryptWorkbookBuffer(plain, "correct horse", { spinCount: 2_000 });
  assert.equal(isCompoundFile(encrypted), true);
  const streams = readCompoundFile(encrypted).streams;
  assert.ok(streams.has("EncryptionInfo"));
  assert.ok(streams.has("EncryptedPackage"));
  assert.deepEqual(decryptWorkbookBuffer(encrypted, "correct horse"), plain);
  assert.throws(() => decryptWorkbookBuffer(encrypted, "wrong password"), /Incorrect workbook password/);
});

test("uses password options in the friendly workbook API", () => {
  const workbook = createWorkbook("Private");
  workbook.sheet().set("A1", "classified");
  const encrypted = workbook.toBuffer({ password: "open sesame", encryption: { spinCount: 2_000 } });
  assert.throws(() => parseWorkbook(encrypted), /password encrypted/);
  const opened = parseWorkbook(encrypted, { password: "open sesame" });
  assert.equal(opened.sheet().get("A1"), "classified");
});
