export function legacyPasswordHash(password) {
  const value = String(password ?? "");
  let hash = 0;
  for (let index = value.length - 1; index >= 0; index -= 1) {
    hash = ((hash >>> 14) & 1) | ((hash << 1) & 0x7FFF);
    hash ^= value.charCodeAt(index);
  }
  hash ^= value.length;
  hash ^= 0xCE4B;
  return (hash & 0xFFFF).toString(16).toUpperCase().padStart(4, "0");
}
