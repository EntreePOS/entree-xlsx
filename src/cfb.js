const SIGNATURE = Buffer.from("d0cf11e0a1b11ae1", "hex");
const FREE = 0xFFFFFFFF;
const END = 0xFFFFFFFE;
const FAT = 0xFFFFFFFD;
const DIFAT = 0xFFFFFFFC;

function padded(buffer, multiple) {
  const length = Math.ceil(buffer.length / multiple) * multiple;
  const output = Buffer.alloc(length);
  buffer.copy(output);
  return output;
}

function sectorChain(start, table, limit) {
  if (start === END || start === FREE) return [];
  const result = [];
  const seen = new Set();
  let current = start;
  while (current !== END && current !== FREE) {
    if (current >= limit || seen.has(current)) throw new Error("Invalid compound file sector chain.");
    seen.add(current);
    result.push(current);
    current = table[current];
  }
  return result;
}

function directoryEntry(name, type, color, left, right, child, start, size) {
  const output = Buffer.alloc(128);
  const encoded = Buffer.from(`${name}\0`, "utf16le").subarray(0, 64);
  encoded.copy(output, 0);
  output.writeUInt16LE(Math.min(encoded.length, 64), 64);
  output[66] = type;
  output[67] = color;
  output.writeUInt32LE(left >>> 0, 68);
  output.writeUInt32LE(right >>> 0, 72);
  output.writeUInt32LE(child >>> 0, 76);
  output.writeUInt32LE(start >>> 0, 116);
  output.writeBigUInt64LE(BigInt(size), 120);
  return output;
}

function compareDirectoryNames(left, right) {
  if (left.length !== right.length) return left.length - right.length;
  const a = left.toUpperCase();
  const b = right.toUpperCase();
  for (let index = 0; index < Math.min(a.length, b.length); index += 1) {
    if (a.charCodeAt(index) !== b.charCodeAt(index)) return a.charCodeAt(index) - b.charCodeAt(index);
  }
  return a.length - b.length;
}

function directoryTree(nodes, indexes, links) {
  indexes.sort((left, right) => compareDirectoryNames(nodes[left].name, nodes[right].name));
  indexes.forEach((index) => links.set(index, { left: FREE, right: FREE, parent: FREE, color: 0 }));
  let root = FREE;
  const color = (index) => index === FREE ? 1 : links.get(index).color;
  const rotateLeft = (index) => {
    const right = links.get(index).right;
    links.get(index).right = links.get(right).left;
    if (links.get(right).left !== FREE) links.get(links.get(right).left).parent = index;
    links.get(right).parent = links.get(index).parent;
    if (links.get(index).parent === FREE) root = right;
    else if (index === links.get(links.get(index).parent).left) links.get(links.get(index).parent).left = right;
    else links.get(links.get(index).parent).right = right;
    links.get(right).left = index;
    links.get(index).parent = right;
  };
  const rotateRight = (index) => {
    const left = links.get(index).left;
    links.get(index).left = links.get(left).right;
    if (links.get(left).right !== FREE) links.get(links.get(left).right).parent = index;
    links.get(left).parent = links.get(index).parent;
    if (links.get(index).parent === FREE) root = left;
    else if (index === links.get(links.get(index).parent).right) links.get(links.get(index).parent).right = left;
    else links.get(links.get(index).parent).left = left;
    links.get(left).right = index;
    links.get(index).parent = left;
  };
  for (const index of indexes) {
    let parent = FREE;
    let current = root;
    while (current !== FREE) {
      parent = current;
      current = compareDirectoryNames(nodes[index].name, nodes[current].name) < 0 ? links.get(current).left : links.get(current).right;
    }
    links.get(index).parent = parent;
    if (parent === FREE) root = index;
    else if (compareDirectoryNames(nodes[index].name, nodes[parent].name) < 0) links.get(parent).left = index;
    else links.get(parent).right = index;
    let node = index;
    while (node !== root && color(links.get(node).parent) === 0) {
      const parentIndex = links.get(node).parent;
      const grandParent = links.get(parentIndex).parent;
      if (parentIndex === links.get(grandParent).left) {
        const uncle = links.get(grandParent).right;
        if (color(uncle) === 0) {
          links.get(parentIndex).color = 1;
          links.get(uncle).color = 1;
          links.get(grandParent).color = 0;
          node = grandParent;
        } else {
          if (node === links.get(parentIndex).right) { node = parentIndex; rotateLeft(node); }
          const nextParent = links.get(node).parent;
          const nextGrandParent = links.get(nextParent).parent;
          links.get(nextParent).color = 1;
          links.get(nextGrandParent).color = 0;
          rotateRight(nextGrandParent);
        }
      } else {
        const uncle = links.get(grandParent).left;
        if (color(uncle) === 0) {
          links.get(parentIndex).color = 1;
          links.get(uncle).color = 1;
          links.get(grandParent).color = 0;
          node = grandParent;
        } else {
          if (node === links.get(parentIndex).left) { node = parentIndex; rotateRight(node); }
          const nextParent = links.get(node).parent;
          const nextGrandParent = links.get(nextParent).parent;
          links.get(nextParent).color = 1;
          links.get(nextGrandParent).color = 0;
          rotateLeft(nextGrandParent);
        }
      }
    }
    links.get(root).color = 1;
  }
  return root;
}

export function createCompoundFile(streams) {
  const sectorSize = 512;
  const miniSectorSize = 64;
  const cutoff = 4096;
  const nodes = [{ name: "Root Entry", type: 5, parent: -1, children: [] }];
  for (const [path, content] of Object.entries(streams)) {
    const parts = path.split("/").filter(Boolean);
    let parent = 0;
    parts.forEach((name, partIndex) => {
      const last = partIndex === parts.length - 1;
      let index = nodes[parent].children.find((child) => nodes[child].name === name);
      if (index === undefined) {
        index = nodes.length;
        nodes.push({ name, type: last ? 2 : 1, parent, children: [], ...(last ? { data: Buffer.isBuffer(content) ? content : Buffer.from(content) } : {}) });
        nodes[parent].children.push(index);
      } else if (last) throw new Error(`Duplicate compound file stream: ${path}.`);
      parent = index;
    });
  }
  const entries = nodes.filter((node) => node.type === 2);
  const sectors = [];
  const chains = [];
  const allocate = (content) => {
    const data = padded(content, sectorSize);
    const ids = [];
    for (let offset = 0; offset < data.length; offset += sectorSize) {
      ids.push(sectors.length);
      sectors.push(data.subarray(offset, offset + sectorSize));
    }
    chains.push(ids);
    return ids[0] ?? END;
  };

  const miniParts = [];
  const miniFat = [];
  for (const entry of entries) {
    if (entry.data.length >= cutoff) continue;
    entry.start = miniFat.length;
    const miniCount = Math.max(1, Math.ceil(entry.data.length / miniSectorSize));
    const content = Buffer.alloc(miniCount * miniSectorSize);
    entry.data.copy(content);
    miniParts.push(content);
    for (let index = 0; index < miniCount; index += 1) miniFat.push(index === miniCount - 1 ? END : entry.start + index + 1);
  }
  for (const entry of entries) if (entry.data.length >= cutoff) entry.start = allocate(entry.data);

  const miniStream = Buffer.concat(miniParts);
  const rootStart = miniStream.length ? allocate(miniStream) : END;
  let miniFatStart = END;
  let miniFatSectorCount = 0;
  if (miniFat.length) {
    const miniFatBuffer = Buffer.alloc(Math.ceil(miniFat.length / 128) * sectorSize, 0xFF);
    miniFat.forEach((value, index) => miniFatBuffer.writeUInt32LE(value >>> 0, index * 4));
    miniFatStart = allocate(miniFatBuffer);
    miniFatSectorCount = miniFatBuffer.length / sectorSize;
  }

  const links = new Map();
  const childRoots = new Map();
  nodes.forEach((node, index) => {
    if (node.children.length) childRoots.set(index, directoryTree(nodes, [...node.children], links));
  });
  const directoryParts = nodes.map((node, index) => {
    const link = links.get(index) ?? { left: FREE, right: FREE, color: index === 0 ? 0 : 1 };
    const start = index === 0 ? rootStart : node.type === 2 ? node.start : END;
    const size = index === 0 ? miniStream.length : node.type === 2 ? node.data.length : 0;
    return directoryEntry(node.name, node.type, link.color, link.left, link.right, childRoots.get(index) ?? FREE, start, size);
  });
  while (directoryParts.length % 4) directoryParts.push(Buffer.alloc(128));
  const directoryStart = allocate(Buffer.concat(directoryParts));

  const dataSectorCount = sectors.length;
  let fatSectorCount = 0;
  let difatSectorCount = 0;
  while (true) {
    const nextDifat = Math.max(0, Math.ceil((fatSectorCount - 109) / 127));
    const nextFat = Math.ceil((dataSectorCount + nextDifat + fatSectorCount) / 128);
    if (nextFat === fatSectorCount && nextDifat === difatSectorCount) break;
    fatSectorCount = nextFat;
    difatSectorCount = nextDifat;
  }
  const difatIds = Array.from({ length: difatSectorCount }, (_, index) => dataSectorCount + index);
  const fatIds = Array.from({ length: fatSectorCount }, (_, index) => dataSectorCount + difatSectorCount + index);
  const totalSectors = dataSectorCount + difatSectorCount + fatSectorCount;
  const fatEntries = new Uint32Array(fatSectorCount * 128);
  fatEntries.fill(FREE);
  for (const chain of chains) chain.forEach((sector, index) => { fatEntries[sector] = index === chain.length - 1 ? END : chain[index + 1]; });
  difatIds.forEach((sector) => { fatEntries[sector] = DIFAT; });
  fatIds.forEach((sector) => { fatEntries[sector] = FAT; });

  for (let index = 0; index < difatIds.length; index += 1) {
    const sector = Buffer.alloc(sectorSize, 0xFF);
    const ids = fatIds.slice(109 + index * 127, 109 + (index + 1) * 127);
    ids.forEach((id, offset) => sector.writeUInt32LE(id, offset * 4));
    sector.writeUInt32LE((difatIds[index + 1] ?? END) >>> 0, 127 * 4);
    sectors.push(sector);
  }
  for (let sectorIndex = 0; sectorIndex < fatSectorCount; sectorIndex += 1) {
    const sector = Buffer.alloc(sectorSize, 0xFF);
    for (let index = 0; index < 128; index += 1) sector.writeUInt32LE(fatEntries[sectorIndex * 128 + index] >>> 0, index * 4);
    sectors.push(sector);
  }
  if (sectors.length !== totalSectors) throw new Error("Internal compound file sector allocation error.");

  const header = Buffer.alloc(sectorSize, 0xFF);
  SIGNATURE.copy(header, 0);
  header.fill(0, 8, 24);
  header.writeUInt16LE(0x003E, 24);
  header.writeUInt16LE(3, 26);
  header.writeUInt16LE(0xFFFE, 28);
  header.writeUInt16LE(9, 30);
  header.writeUInt16LE(6, 32);
  header.fill(0, 34, 40);
  header.writeUInt32LE(0, 40);
  header.writeUInt32LE(fatSectorCount, 44);
  header.writeUInt32LE(directoryStart, 48);
  header.writeUInt32LE(0, 52);
  header.writeUInt32LE(cutoff, 56);
  header.writeUInt32LE(miniFatStart >>> 0, 60);
  header.writeUInt32LE(miniFatSectorCount, 64);
  header.writeUInt32LE((difatIds[0] ?? END) >>> 0, 68);
  header.writeUInt32LE(difatSectorCount, 72);
  fatIds.slice(0, 109).forEach((id, index) => header.writeUInt32LE(id, 76 + index * 4));
  return Buffer.concat([header, ...sectors]);
}

export function isCompoundFile(input) {
  const buffer = Buffer.isBuffer(input) ? input : Buffer.from(input);
  return buffer.length >= 8 && buffer.subarray(0, 8).equals(SIGNATURE);
}

export function readCompoundFile(input) {
  const buffer = Buffer.isBuffer(input) ? input : Buffer.from(input);
  if (!isCompoundFile(buffer)) throw new Error("Not a compound file.");
  const major = buffer.readUInt16LE(26);
  const sectorSize = 1 << buffer.readUInt16LE(30);
  const miniSectorSize = 1 << buffer.readUInt16LE(32);
  const cutoff = buffer.readUInt32LE(56);
  const sectorCount = Math.floor(buffer.length / sectorSize) - 1;
  const sector = (id) => {
    if (id < 0 || id >= sectorCount) throw new Error("Invalid compound file sector index.");
    const offset = (id + 1) * sectorSize;
    return buffer.subarray(offset, offset + sectorSize);
  };
  const fatSectorCount = buffer.readUInt32LE(44);
  const fatIds = [];
  for (let index = 0; index < 109 && fatIds.length < fatSectorCount; index += 1) {
    const id = buffer.readUInt32LE(76 + index * 4);
    if (id !== FREE) fatIds.push(id);
  }
  let difat = buffer.readUInt32LE(68);
  const difatCount = buffer.readUInt32LE(72);
  for (let count = 0; count < difatCount && difat !== END; count += 1) {
    const content = sector(difat);
    for (let index = 0; index < sectorSize / 4 - 1 && fatIds.length < fatSectorCount; index += 1) {
      const id = content.readUInt32LE(index * 4);
      if (id !== FREE) fatIds.push(id);
    }
    difat = content.readUInt32LE(sectorSize - 4);
  }
  const fatTable = [];
  for (const id of fatIds) {
    const content = sector(id);
    for (let offset = 0; offset < sectorSize; offset += 4) fatTable.push(content.readUInt32LE(offset));
  }
  const readRegular = (start, size) => Buffer.concat(sectorChain(start, fatTable, sectorCount).map(sector)).subarray(0, size);
  const directoryStart = buffer.readUInt32LE(48);
  const directory = readRegular(directoryStart, sectorChain(directoryStart, fatTable, sectorCount).length * sectorSize);
  const entries = [];
  for (let offset = 0; offset + 128 <= directory.length; offset += 128) {
    const nameLength = directory.readUInt16LE(offset + 64);
    const type = directory[offset + 66];
    if (!type || nameLength < 2) continue;
    entries.push({
      name: directory.subarray(offset, offset + nameLength - 2).toString("utf16le"),
      type,
      start: directory.readUInt32LE(offset + 116),
      size: Number(directory.readBigUInt64LE(offset + 120))
    });
  }
  const root = entries.find((entry) => entry.type === 5);
  const miniStream = root?.size ? readRegular(root.start, root.size) : Buffer.alloc(0);
  const miniFatStart = buffer.readUInt32LE(60);
  const miniFatSectorCount = buffer.readUInt32LE(64);
  const miniFatBuffer = miniFatSectorCount ? readRegular(miniFatStart, miniFatSectorCount * sectorSize) : Buffer.alloc(0);
  const miniFatTable = [];
  for (let offset = 0; offset < miniFatBuffer.length; offset += 4) miniFatTable.push(miniFatBuffer.readUInt32LE(offset));
  const result = new Map();
  for (const entry of entries.filter((item) => item.type === 2)) {
    let content;
    if (entry.size < cutoff) {
      const ids = sectorChain(entry.start, miniFatTable, Math.ceil(miniStream.length / miniSectorSize));
      content = Buffer.concat(ids.map((id) => miniStream.subarray(id * miniSectorSize, (id + 1) * miniSectorSize))).subarray(0, entry.size);
    } else content = readRegular(entry.start, entry.size);
    result.set(entry.name, content);
  }
  return { streams: result, majorVersion: major, sectorSize };
}
