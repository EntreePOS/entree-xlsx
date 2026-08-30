import { createCipheriv, createDecipheriv, createHash, createHmac, randomBytes, timingSafeEqual } from "node:crypto";
import { createCompoundFile, isCompoundFile, readCompoundFile } from "./cfb.js";
import { decodeXml, escapeXml, parseAttributes } from "./xml.js";

const BLOCK_VERIFIER = Buffer.from("fea7d2763b4b9e79", "hex");
const BLOCK_VERIFIER_HASH = Buffer.from("d7aa0f6d3061344e", "hex");
const BLOCK_KEY_VALUE = Buffer.from("146e0be7abacd0d6", "hex");
const BLOCK_HMAC_KEY = Buffer.from("5fb2ad010cb9e1f6", "hex");
const BLOCK_HMAC_VALUE = Buffer.from("a0677f02b22c8433", "hex");
const HASH = "sha512";
const BLOCK_SIZE = 16;
const KEY_BYTES = 32;
const SPIN_COUNT = 100_000;
const DATA_SPACES = {
  "\u0006DataSpaces/Version": Buffer.from("3c0000004d006900630072006f0073006f00660074002e0043006f006e007400610069006e00650072002e004400610074006100530070006100630065007300010000000100000001000000", "hex"),
  "\u0006DataSpaces/DataSpaceMap": Buffer.from("08000000010000006800000001000000000000002000000045006e0063007200790070007400650064005000610063006b00610067006500320000005300740072006f006e00670045006e006300720079007000740069006f006e004400610074006100530070006100630065000000", "hex"),
  "\u0006DataSpaces/DataSpaceInfo/StrongEncryptionDataSpace": Buffer.from("0800000001000000320000005300740072006f006e00670045006e006300720079007000740069006f006e005400720061006e00730066006f0072006d000000", "hex"),
  "\u0006DataSpaces/TransformInfo/StrongEncryptionTransform/\u0006Primary": Buffer.from("58000000010000004c0000007b00460046003900410033004600300033002d0035003600450046002d0034003600310033002d0042004400440035002d003500410034003100430031004400300037003200340036007d004e0000004d006900630072006f0073006f00660074002e0043006f006e007400610069006e00650072002e0045006e006300720079007000740069006f006e005400720061006e00730066006f0072006d00000001000000010000000100000000000000000000000000000004000000", "hex")
};

function hash(...parts) {
  const value = createHash(HASH);
  parts.forEach((part) => value.update(part));
  return value.digest();
}

function iteratedPasswordHash(password, salt, spinCount = SPIN_COUNT) {
  let value = hash(salt, Buffer.from(password, "utf16le"));
  const iterator = Buffer.alloc(4);
  for (let index = 0; index < spinCount; index += 1) {
    iterator.writeUInt32LE(index, 0);
    value = hash(iterator, value);
  }
  return value;
}

function derivedKey(passwordHash, blockKey, keyBytes = KEY_BYTES) {
  const value = hash(passwordHash, blockKey);
  if (value.length >= keyBytes) return value.subarray(0, keyBytes);
  return Buffer.concat([value, Buffer.alloc(keyBytes - value.length, 0x36)]);
}

function padded(value, blockSize = BLOCK_SIZE) {
  if (value.length % blockSize === 0) return value;
  return Buffer.concat([value, Buffer.alloc(blockSize - value.length % blockSize)]);
}

function aesEncrypt(key, iv, value) {
  const cipher = createCipheriv(`aes-${key.length * 8}-cbc`, key, iv);
  cipher.setAutoPadding(false);
  return Buffer.concat([cipher.update(padded(value)), cipher.final()]);
}

function aesDecrypt(key, iv, value) {
  const decipher = createDecipheriv(`aes-${key.length * 8}-cbc`, key, iv);
  decipher.setAutoPadding(false);
  return Buffer.concat([decipher.update(value), decipher.final()]);
}

function blockIv(salt, blockKey) {
  return hash(salt, blockKey).subarray(0, BLOCK_SIZE);
}

function encryptPackage(packageBuffer, secretKey, salt) {
  const length = Buffer.alloc(8);
  length.writeBigUInt64LE(BigInt(packageBuffer.length));
  const parts = [length];
  for (let offset = 0, block = 0; offset < packageBuffer.length; offset += 4096, block += 1) {
    const blockKey = Buffer.alloc(4);
    blockKey.writeUInt32LE(block);
    parts.push(aesEncrypt(secretKey, blockIv(salt, blockKey), packageBuffer.subarray(offset, offset + 4096)));
  }
  return Buffer.concat(parts);
}

function findZipEnd(buffer) {
  for (let index = buffer.length - 22; index >= Math.max(0, buffer.length - 65_557); index -= 1) {
    if (buffer.readUInt32LE(index) === 0x06054B50) return index;
  }
  return -1;
}

function padSmallZip(buffer) {
  if (buffer.length >= 4096) return buffer;
  const end = findZipEnd(buffer);
  if (end < 0) return buffer;
  const oldCommentLength = buffer.readUInt16LE(end + 20);
  const marker = Buffer.from(`ENTREEPAD:${buffer.length}:${oldCommentLength}:`);
  const paddingLength = 4096 - buffer.length;
  if (oldCommentLength + paddingLength > 0xFFFF || marker.length > paddingLength) return buffer;
  const output = Buffer.concat([Buffer.from(buffer), marker, Buffer.alloc(paddingLength - marker.length, 0x20)]);
  output.writeUInt16LE(oldCommentLength + paddingLength, end + 20);
  return output;
}

function removeZipPadding(buffer) {
  const end = findZipEnd(buffer);
  if (end < 0) return buffer;
  const commentLength = buffer.readUInt16LE(end + 20);
  const commentStart = end + 22;
  const comment = buffer.subarray(commentStart, commentStart + commentLength).toString("utf8");
  const match = /ENTREEPAD:(\d+):(\d+):/.exec(comment);
  if (!match) return buffer.subarray(0, commentStart + commentLength);
  const originalLength = Number(match[1]);
  const originalCommentLength = Number(match[2]);
  const output = Buffer.from(buffer.subarray(0, originalLength));
  output.writeUInt16LE(originalCommentLength, end + 20);
  return output;
}

function encryptionInfoXml(config) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption" xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password" xmlns:c="http://schemas.microsoft.com/office/2006/keyEncryptor/certificate"><keyData saltSize="16" blockSize="16" keyBits="256" hashSize="64" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" saltValue="${config.dataSalt.toString("base64")}"/><dataIntegrity encryptedHmacKey="${config.encryptedHmacKey.toString("base64")}" encryptedHmacValue="${config.encryptedHmacValue.toString("base64")}"/><keyEncryptors><keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password"><p:encryptedKey spinCount="${config.spinCount}" saltSize="16" blockSize="16" keyBits="256" hashSize="64" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" saltValue="${config.passwordSalt.toString("base64")}" encryptedVerifierHashInput="${config.encryptedVerifier.toString("base64")}" encryptedVerifierHashValue="${config.encryptedVerifierHash.toString("base64")}" encryptedKeyValue="${config.encryptedKey.toString("base64")}"/></keyEncryptor></keyEncryptors></encryption>`;
}

export function encryptWorkbookBuffer(input, password, options = {}) {
  if (typeof password !== "string" || !password.length) throw new TypeError("A non-empty password is required.");
  const packageBuffer = Buffer.isBuffer(input) ? input : Buffer.from(input);
  const packageForEncryption = padSmallZip(packageBuffer);
  const spinCount = options.spinCount ?? SPIN_COUNT;
  const dataSalt = randomBytes(16);
  const passwordSalt = randomBytes(16);
  const secretKey = randomBytes(KEY_BYTES);
  const verifier = randomBytes(16);
  const passwordHash = iteratedPasswordHash(password, passwordSalt, spinCount);
  const encryptedVerifier = aesEncrypt(derivedKey(passwordHash, BLOCK_VERIFIER), passwordSalt, verifier);
  const encryptedVerifierHash = aesEncrypt(derivedKey(passwordHash, BLOCK_VERIFIER_HASH), passwordSalt, hash(verifier));
  const encryptedKey = aesEncrypt(derivedKey(passwordHash, BLOCK_KEY_VALUE), passwordSalt, secretKey);
  const encryptedPackage = encryptPackage(packageForEncryption, secretKey, dataSalt);
  const hmacKey = randomBytes(64);
  const hmacValue = createHmac(HASH, hmacKey).update(encryptedPackage).digest();
  const encryptedHmacKey = aesEncrypt(secretKey, blockIv(dataSalt, BLOCK_HMAC_KEY), hmacKey);
  const encryptedHmacValue = aesEncrypt(secretKey, blockIv(dataSalt, BLOCK_HMAC_VALUE), hmacValue);
  const xml = encryptionInfoXml({ dataSalt, passwordSalt, encryptedVerifier, encryptedVerifierHash, encryptedKey, encryptedHmacKey, encryptedHmacValue, spinCount });
  const header = Buffer.alloc(8);
  header.writeUInt16LE(4, 0);
  header.writeUInt16LE(4, 2);
  header.writeUInt32LE(0x40, 4);
  return createCompoundFile({ ...DATA_SPACES, EncryptionInfo: Buffer.concat([header, Buffer.from(xml)]), EncryptedPackage: encryptedPackage });
}

function encryptionAttributes(xml, tagName) {
  const pattern = new RegExp(`<(?:\\w+:)?${tagName}\\b([^>]*)`, "i");
  const match = pattern.exec(xml);
  return match ? Object.fromEntries(Object.entries(parseAttributes(match[1])).map(([key, value]) => [key, decodeXml(value)])) : undefined;
}

export function decryptWorkbookBuffer(input, password, options = {}) {
  if (typeof password !== "string") throw new TypeError("A password is required to open this encrypted workbook.");
  const compound = readCompoundFile(input);
  const info = compound.streams.get("EncryptionInfo");
  const encryptedPackage = compound.streams.get("EncryptedPackage");
  if (!info || !encryptedPackage) throw new Error("Invalid encrypted Office file: required streams are missing.");
  if (info.readUInt16LE(0) !== 4 || info.readUInt16LE(2) !== 4) throw new Error("Only Agile Office encryption is supported.");
  const xml = info.subarray(8).toString("utf8");
  const keyData = encryptionAttributes(xml, "keyData");
  const encryptedKey = encryptionAttributes(xml, "encryptedKey");
  const integrity = encryptionAttributes(xml, "dataIntegrity");
  if (!keyData || !encryptedKey) throw new Error("Invalid Agile encryption descriptor.");
  const passwordSalt = Buffer.from(encryptedKey.saltValue, "base64");
  const dataSalt = Buffer.from(keyData.saltValue, "base64");
  const passwordHash = iteratedPasswordHash(password, passwordSalt, Number(encryptedKey.spinCount));
  const verifier = aesDecrypt(derivedKey(passwordHash, BLOCK_VERIFIER), passwordSalt, Buffer.from(encryptedKey.encryptedVerifierHashInput, "base64")).subarray(0, 16);
  const verifierHash = aesDecrypt(derivedKey(passwordHash, BLOCK_VERIFIER_HASH), passwordSalt, Buffer.from(encryptedKey.encryptedVerifierHashValue, "base64")).subarray(0, Number(encryptedKey.hashSize));
  const expectedVerifierHash = hash(verifier);
  if (verifierHash.length !== expectedVerifierHash.length || !timingSafeEqual(verifierHash, expectedVerifierHash)) throw new Error("Incorrect workbook password.");
  const secretKey = aesDecrypt(derivedKey(passwordHash, BLOCK_KEY_VALUE), passwordSalt, Buffer.from(encryptedKey.encryptedKeyValue, "base64")).subarray(0, Number(encryptedKey.keyBits) / 8);
  if (integrity && options.verifyIntegrity !== false) {
    const hmacKey = aesDecrypt(secretKey, blockIv(dataSalt, BLOCK_HMAC_KEY), Buffer.from(integrity.encryptedHmacKey, "base64")).subarray(0, Number(keyData.hashSize));
    const hmacValue = aesDecrypt(secretKey, blockIv(dataSalt, BLOCK_HMAC_VALUE), Buffer.from(integrity.encryptedHmacValue, "base64")).subarray(0, Number(keyData.hashSize));
    const expected = createHmac(HASH, hmacKey).update(encryptedPackage).digest();
    if (hmacValue.length !== expected.length || !timingSafeEqual(hmacValue, expected)) throw new Error("Encrypted workbook integrity check failed.");
  }
  const size = Number(encryptedPackage.readBigUInt64LE(0));
  const output = [];
  let offset = 8;
  for (let block = 0; output.reduce((sum, item) => sum + item.length, 0) < size; block += 1) {
    const remaining = size - output.reduce((sum, item) => sum + item.length, 0);
    const plainLength = Math.min(4096, remaining);
    const cipherLength = Math.ceil(plainLength / BLOCK_SIZE) * BLOCK_SIZE;
    const blockKey = Buffer.alloc(4);
    blockKey.writeUInt32LE(block);
    output.push(aesDecrypt(secretKey, blockIv(dataSalt, blockKey), encryptedPackage.subarray(offset, offset + cipherLength)).subarray(0, plainLength));
    offset += cipherLength;
  }
  const decrypted = Buffer.concat(output).subarray(0, size);
  if (decrypted.subarray(0, 4).equals(Buffer.from("504b0304", "hex"))) return removeZipPadding(decrypted);
  return decrypted;
}

export { isCompoundFile };
