export function escapeXml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

export function decodeXml(value = "") {
  return value.replace(/&(?:#x([0-9a-f]+)|#(\d+)|([a-z]+));/gi, (entity, hex, decimal, named) => {
    if (hex) return String.fromCodePoint(Number.parseInt(hex, 16));
    if (decimal) return String.fromCodePoint(Number(decimal));
    return { amp: "&", lt: "<", gt: ">", quot: '"', apos: "'" }[named.toLowerCase()] ?? entity;
  });
}

export function parseAttributes(source = "") {
  const attributes = {};
  const pattern = /([\w:-]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;
  let match;
  while ((match = pattern.exec(source))) attributes[match[1]] = decodeXml(match[2] ?? match[3] ?? "");
  return attributes;
}

export function tagContent(xml, name) {
  const match = new RegExp(`<(?:\\w+:)?${name}\\b[^>]*>([\\s\\S]*?)<\\/(?:\\w+:)?${name}>`, "i").exec(xml);
  return match?.[1];
}

export function allTags(xml, name) {
  const tags = [];
  const pattern = new RegExp(`<(?:\\w+:)?${name}\\b([^>]*?)(?:\\/>|>([\\s\\S]*?)<\\/(?:\\w+:)?${name}>)`, "gi");
  let match;
  while ((match = pattern.exec(xml))) tags.push({ attributes: parseAttributes(match[1]), content: match[2] ?? "", xml: match[0] });
  return tags;
}
