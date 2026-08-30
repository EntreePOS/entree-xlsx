import { posix } from "node:path";
import { allTags, escapeXml } from "./xml.js";

const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";

export function getTextFile(state, path, required = true) {
  const value = state.files.get(path);
  if (!value && required) throw new Error(`Missing OOXML package part: ${path}.`);
  return value?.toString("utf8") ?? "";
}

export function setTextFile(state, path, xml) {
  state.files.set(path, Buffer.from(xml));
  return path;
}

export function relationshipsPath(sourcePath) {
  return sourcePath
    ? posix.join(posix.dirname(sourcePath), "_rels", `${posix.basename(sourcePath)}.rels`)
    : "_rels/.rels";
}

export function resolveTarget(sourcePath, target) {
  if (target.startsWith("/")) return target.slice(1);
  return posix.normalize(posix.join(sourcePath ? posix.dirname(sourcePath) : "", target));
}

export function relativeTarget(sourcePath, targetPath) {
  return posix.relative(posix.dirname(sourcePath), targetPath);
}

export function readRelationships(state, sourcePath) {
  const path = relationshipsPath(sourcePath);
  const xml = getTextFile(state, path, false);
  return {
    path,
    xml,
    items: allTags(xml, "Relationship").map(({ attributes }) => ({
      id: attributes.Id,
      type: attributes.Type,
      target: attributes.Target,
      targetMode: attributes.TargetMode
    }))
  };
}

export function writeRelationships(state, sourcePath, items) {
  const path = relationshipsPath(sourcePath);
  const xml = `${XML_HEADER}<Relationships xmlns="${PACKAGE_REL_NS}">${items.map((item) => `<Relationship Id="${escapeXml(item.id)}" Type="${escapeXml(item.type)}" Target="${escapeXml(item.target)}"${item.targetMode ? ` TargetMode="${escapeXml(item.targetMode)}"` : ""}/>`).join("")}</Relationships>`;
  setTextFile(state, path, xml);
  return path;
}

export function nextRelationshipId(items) {
  const used = new Set(items.map((item) => item.id));
  let number = 1;
  while (used.has(`rId${number}`)) number += 1;
  return `rId${number}`;
}

export function addRelationship(state, sourcePath, type, targetPath, options = {}) {
  const relationships = readRelationships(state, sourcePath);
  const id = nextRelationshipId(relationships.items);
  relationships.items.push({
    id,
    type,
    target: options.external ? targetPath : relativeTarget(sourcePath, targetPath),
    ...(options.external ? { targetMode: "External" } : {})
  });
  writeRelationships(state, sourcePath, relationships.items);
  return id;
}

export function removeRelationship(state, sourcePath, id) {
  const relationships = readRelationships(state, sourcePath);
  const removed = relationships.items.find((item) => item.id === id);
  writeRelationships(state, sourcePath, relationships.items.filter((item) => item.id !== id));
  return removed;
}

export function relationshipTarget(state, sourcePath, id) {
  const relationship = readRelationships(state, sourcePath).items.find((item) => item.id === id);
  return relationship ? resolveTarget(sourcePath, relationship.target) : undefined;
}

export function nextPartPath(state, folder, prefix, extension = ".xml") {
  let number = 1;
  while (state.files.has(`${folder}/${prefix}${number}${extension}`)) number += 1;
  return `${folder}/${prefix}${number}${extension}`;
}

export function addContentTypeOverride(state, partPath, contentType) {
  const path = "[Content_Types].xml";
  let xml = getTextFile(state, path);
  const partName = `/${partPath}`;
  if (allTags(xml, "Override").some(({ attributes }) => attributes.PartName === partName)) return;
  xml = xml.replace(/<\/(?:\w+:)?Types\s*>/i, `<Override PartName="${escapeXml(partName)}" ContentType="${escapeXml(contentType)}"/></Types>`);
  setTextFile(state, path, xml);
}

export function removeContentTypeOverride(state, partPath) {
  const path = "[Content_Types].xml";
  let xml = getTextFile(state, path);
  const escaped = `/${partPath}`.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  xml = xml.replace(new RegExp(`<(?:\\w+:)?Override\\b(?=[^>]*\\bPartName="${escaped}")[^>]*/>`, "gi"), "");
  setTextFile(state, path, xml);
}
