export function encodeColumn(column) {
  let value = column + 1;
  let output = "";
  while (value > 0) {
    value -= 1;
    output = String.fromCharCode(65 + (value % 26)) + output;
    value = Math.floor(value / 26);
  }
  return output;
}

export function decodeColumn(column) {
  const value = column.trim().toUpperCase();
  if (!/^[A-Z]+$/.test(value)) throw new TypeError(`Invalid column ${JSON.stringify(column)}.`);
  let result = 0;
  for (const character of value) result = result * 26 + character.charCodeAt(0) - 64;
  return result - 1;
}

export function encodeAddress(address) {
  return `${encodeColumn(address.c)}${address.r + 1}`;
}

export function decodeAddress(address) {
  const match = /^([A-Z]+)([1-9]\d*)$/i.exec(address.trim());
  if (!match) throw new TypeError(`Invalid cell address ${JSON.stringify(address)}.`);
  return { r: Number(match[2]) - 1, c: decodeColumn(match[1]) };
}

export function normalizeAddress(address) {
  if (typeof address === "string") {
    const value = address.trim().toUpperCase();
    if (!/^[A-Z]+[1-9]\d*$/.test(value)) throw new TypeError(`Invalid cell address ${JSON.stringify(address)}.`);
    return value;
  }
  if (!Number.isInteger(address.r) || !Number.isInteger(address.c) || address.r < 0 || address.c < 0) {
    throw new TypeError("Cell coordinates must be non-negative integers.");
  }
  return encodeAddress(address);
}

export function encodeRange(range) {
  const start = encodeAddress(range.s);
  const end = encodeAddress(range.e);
  return start === end ? start : `${start}:${end}`;
}

export function decodeRange(range) {
  const [start, end = start] = range.split(":");
  return orderedRange({ s: decodeAddress(start), e: decodeAddress(end) });
}

export function normalizeRange(range) {
  if (typeof range !== "string") return orderedRange(range);
  try {
    return decodeRange(range.trim().toUpperCase());
  } catch (error) {
    throw new TypeError(`Invalid cell range ${JSON.stringify(range)}.`, { cause: error });
  }
}

function orderedRange(range) {
  const values = [range.s.r, range.s.c, range.e.r, range.e.c];
  if (values.some((value) => !Number.isInteger(value) || value < 0)) {
    throw new TypeError("Range coordinates must be non-negative integers.");
  }
  return {
    s: { r: Math.min(range.s.r, range.e.r), c: Math.min(range.s.c, range.e.c) },
    e: { r: Math.max(range.s.r, range.e.r), c: Math.max(range.s.c, range.e.c) }
  };
}

export function expandWorksheetRef(worksheet, address) {
  const current = worksheet["!ref"] ? decodeRange(worksheet["!ref"]) : { s: address, e: address };
  worksheet["!ref"] = encodeRange({
    s: { r: Math.min(current.s.r, address.r), c: Math.min(current.s.c, address.c) },
    e: { r: Math.max(current.e.r, address.r), c: Math.max(current.e.c, address.c) }
  });
}

export function forEachAddress(range, callback) {
  for (let row = range.s.r; row <= range.e.r; row += 1) {
    for (let column = range.s.c; column <= range.e.c; column += 1) {
      const point = { r: row, c: column };
      callback(encodeAddress(point), point);
    }
  }
}
