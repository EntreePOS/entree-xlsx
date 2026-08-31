const HORIZONTAL_ALIGNMENTS = new Set(["general", "left", "center", "right", "fill", "justify", "centerContinuous", "distributed"]);
const VERTICAL_ALIGNMENTS = new Set(["top", "center", "bottom", "justify", "distributed"]);
const UNDERLINES = new Set(["single", "double", "singleAccounting", "doubleAccounting"]);
const VERTICAL_ALIGNS = new Set(["baseline", "subscript", "superscript"]);
const GRADIENT_TYPES = new Set(["linear", "path"]);
const BORDER_SIDES = new Set(["left", "right", "top", "bottom", "diagonal", "vertical", "horizontal"]);
const RANGE_BORDER_KEYS = new Set(["all", "outline", "inside", "insideHorizontal", "insideVertical"]);

function isPlainObject(value) {
  return value !== null && typeof value === "object" && !Array.isArray(value) && !(value instanceof Date);
}

function assertObject(value, label) {
  if (!isPlainObject(value)) throw new TypeError(`${label} must be an object.`);
  return value;
}

function normalizeHex(value) {
  const normalized = value.trim().replace(/^#/, "").toUpperCase();
  if (!/^(?:[0-9A-F]{6}|[0-9A-F]{8})$/.test(normalized)) {
    throw new TypeError(`Invalid color ${JSON.stringify(value)}. Use a 6- or 8-digit hex color.`);
  }
  return normalized;
}

function finiteNumber(value, label, { min = -Infinity, max = Infinity } = {}) {
  const number = Number(value);
  if (!Number.isFinite(number) || number < min || number > max) {
    throw new RangeError(`${label} must be a number from ${min} to ${max}.`);
  }
  return number;
}

function optionalBoolean(value, label) {
  if (typeof value !== "boolean") throw new TypeError(`${label} must be true or false.`);
  return value;
}

function copyDefined(source, keys) {
  return Object.fromEntries(keys.filter((key) => source[key] !== undefined).map((key) => [key, source[key]]));
}

export function normalizeColor(color) {
  if (typeof color === "string") return { rgb: normalizeHex(color) };
  if (typeof color === "number" && Number.isInteger(color) && color >= 0 && color <= 0xFFFFFFFF) {
    const digits = color > 0xFFFFFF ? 8 : 6;
    return { rgb: color.toString(16).padStart(digits, "0").toUpperCase() };
  }
  assertObject(color, "Color");
  const fields = ["rgb", "theme", "indexed", "auto"].filter((key) => color[key] !== undefined);
  if (fields.length !== 1) throw new TypeError("A color object must define exactly one of rgb, theme, indexed, or auto.");
  if (color.rgb !== undefined) return { rgb: normalizeHex(String(color.rgb)), ...(color.tint !== undefined ? { tint: finiteNumber(color.tint, "Color tint", { min: -1, max: 1 }) } : {}) };
  if (color.theme !== undefined) return { theme: finiteNumber(color.theme, "Theme color", { min: 0, max: 255 }), ...(color.tint !== undefined ? { tint: finiteNumber(color.tint, "Color tint", { min: -1, max: 1 }) } : {}) };
  if (color.indexed !== undefined) return { indexed: finiteNumber(color.indexed, "Indexed color", { min: 0, max: 65 }) };
  return { auto: optionalBoolean(color.auto, "Automatic color") };
}

function normalizeUnderline(value) {
  if (value === true || value === 1) return "single";
  if (value === false || value === 0) return false;
  if (value === 2) return "double";
  if (!UNDERLINES.has(value)) throw new TypeError(`Invalid underline ${JSON.stringify(value)}.`);
  return value;
}

function normalizeFont(input) {
  const source = input.font === undefined ? {} : assertObject(input.font, "font");
  const font = { ...copyDefined(source, ["bold", "italic", "underline", "strike", "outline", "shadow", "condense", "extend", "name", "size", "family", "charset", "scheme", "verticalAlign", "color"]) };
  const shortcuts = {
    bold: "bold", italic: "italic", underline: "underline", strike: "strike",
    fontOutline: "outline", fontShadow: "shadow", fontName: "name", fontSize: "size",
    fontFamily: "family", fontCharset: "charset", fontScheme: "scheme", fontColor: "color",
    color: "color", verticalAlign: "verticalAlign"
  };
  for (const [shortcut, key] of Object.entries(shortcuts)) if (input[shortcut] !== undefined) font[key] = input[shortcut];
  if (input.subscript === true) font.verticalAlign = "subscript";
  if (input.superscript === true) font.verticalAlign = "superscript";
  for (const key of ["bold", "italic", "strike", "outline", "shadow", "condense", "extend"]) {
    if (font[key] !== undefined) font[key] = optionalBoolean(font[key], `font.${key}`);
  }
  if (font.underline !== undefined) font.underline = normalizeUnderline(font.underline);
  if (font.name !== undefined && (typeof font.name !== "string" || !font.name.trim())) throw new TypeError("font.name must be a non-empty string.");
  if (font.size !== undefined) font.size = finiteNumber(font.size, "font.size", { min: 1, max: 409 });
  if (font.family !== undefined) font.family = finiteNumber(font.family, "font.family", { min: 0, max: 14 });
  if (font.charset !== undefined) font.charset = finiteNumber(font.charset, "font.charset", { min: 0, max: 255 });
  if (font.scheme !== undefined && !["major", "minor", "none"].includes(font.scheme)) throw new TypeError("font.scheme must be major, minor, or none.");
  if (font.verticalAlign !== undefined && !VERTICAL_ALIGNS.has(font.verticalAlign)) throw new TypeError("font.verticalAlign must be baseline, subscript, or superscript.");
  if (font.color !== undefined) font.color = normalizeColor(font.color);
  return font;
}

function looksLikeColor(value) {
  return typeof value === "string" || typeof value === "number" || (isPlainObject(value) && ["rgb", "theme", "indexed", "auto"].some((key) => value[key] !== undefined));
}

function normalizeGradientStop(stop, index) {
  assertObject(stop, `fill.stops[${index}]`);
  const position = stop.position ?? stop.offset ?? stop.v;
  const color = stop.color ?? (looksLikeColor(stop) ? Object.fromEntries(Object.entries(stop).filter(([key]) => ["rgb", "theme", "indexed", "auto", "tint"].includes(key))) : undefined);
  if (position === undefined || color === undefined) throw new TypeError(`fill.stops[${index}] must define position and color.`);
  return { position: finiteNumber(position, `fill.stops[${index}].position`, { min: 0, max: 1 }), color: normalizeColor(color) };
}

function normalizeFill(fill) {
  if (looksLikeColor(fill)) return { type: "pattern", patternType: "solid", foreground: normalizeColor(fill) };
  const source = assertObject(fill, "fill");
  const gradient = source.type === "gradient" || source.stops !== undefined;
  if (gradient) {
    const gradientType = source.gradientType ?? (source.path ? "path" : "linear");
    if (!GRADIENT_TYPES.has(gradientType)) throw new TypeError("fill.gradientType must be linear or path.");
    if (!Array.isArray(source.stops) || source.stops.length < 2) throw new TypeError("A gradient fill requires at least two stops.");
    return {
      type: "gradient",
      gradientType,
      ...(gradientType === "linear" ? { degree: finiteNumber(source.degree ?? source.angle ?? 0, "fill.degree", { min: 0, max: 360 }) } : {}),
      ...copyDefined(source, ["left", "right", "top", "bottom"]),
      stops: source.stops.map(normalizeGradientStop)
    };
  }
  const patternType = source.patternType ?? source.pattern ?? "solid";
  if (typeof patternType !== "string" || !patternType) throw new TypeError("fill.patternType must be a non-empty string.");
  const foreground = source.foreground ?? source.fgColor ?? source.color;
  const background = source.background ?? source.bgColor;
  return {
    type: "pattern",
    patternType,
    ...(foreground !== undefined ? { foreground: normalizeColor(foreground) } : {}),
    ...(background !== undefined ? { background: normalizeColor(background) } : {})
  };
}

function normalizeAlignment(input) {
  const source = input.alignment === undefined ? {} : assertObject(input.alignment, "alignment");
  const alignment = { ...copyDefined(source, ["horizontal", "vertical", "wrapText", "shrinkToFit", "textRotation", "indent", "relativeIndent", "justifyLastLine", "readingOrder"]) };
  for (const key of ["horizontal", "vertical", "wrapText", "shrinkToFit", "textRotation", "indent", "relativeIndent", "justifyLastLine", "readingOrder"]) {
    if (input[key] !== undefined) alignment[key] = input[key];
  }
  if (input.rotation !== undefined) {
    const rotation = finiteNumber(input.rotation, "rotation", { min: -90, max: 90 });
    alignment.textRotation = rotation < 0 ? 90 - rotation : rotation;
  }
  if (input.verticalText === true) alignment.textRotation = 255;
  if (alignment.horizontal !== undefined && !HORIZONTAL_ALIGNMENTS.has(alignment.horizontal)) throw new TypeError(`Invalid horizontal alignment ${JSON.stringify(alignment.horizontal)}.`);
  if (alignment.vertical !== undefined && !VERTICAL_ALIGNMENTS.has(alignment.vertical)) throw new TypeError(`Invalid vertical alignment ${JSON.stringify(alignment.vertical)}.`);
  for (const key of ["wrapText", "shrinkToFit", "justifyLastLine"]) if (alignment[key] !== undefined) alignment[key] = optionalBoolean(alignment[key], `alignment.${key}`);
  if (alignment.textRotation !== undefined) alignment.textRotation = finiteNumber(alignment.textRotation, "alignment.textRotation", { min: 0, max: 255 });
  if (alignment.indent !== undefined) alignment.indent = finiteNumber(alignment.indent, "alignment.indent", { min: 0, max: 250 });
  if (alignment.relativeIndent !== undefined) alignment.relativeIndent = finiteNumber(alignment.relativeIndent, "alignment.relativeIndent", { min: -15, max: 15 });
  if (alignment.readingOrder !== undefined) alignment.readingOrder = finiteNumber(alignment.readingOrder, "alignment.readingOrder", { min: 0, max: 2 });
  return alignment;
}

function normalizeBorderSide(side, label) {
  if (typeof side === "string") return { style: side };
  const source = assertObject(side, label);
  return {
    ...(source.style !== undefined ? { style: String(source.style) } : {}),
    ...(source.color !== undefined ? { color: normalizeColor(source.color) } : {})
  };
}

function normalizeBorder(input) {
  if (input.border === undefined) return { border: {}, rangeBorder: {} };
  const source = assertObject(input.border, "border");
  const border = {};
  const rangeBorder = {};
  for (const [key, value] of Object.entries(source)) {
    if (value === undefined || value === null || value === false) continue;
    if (BORDER_SIDES.has(key)) border[key] = normalizeBorderSide(value, `border.${key}`);
    else if (RANGE_BORDER_KEYS.has(key)) rangeBorder[key] = normalizeBorderSide(value, `border.${key}`);
    else if (["diagonalUp", "diagonalDown", "outline"].includes(key) && typeof value === "boolean") border[key] = value;
    else throw new TypeError(`Unknown border property ${JSON.stringify(key)}.`);
  }
  return { border, rangeBorder };
}

function normalizeProtection(input) {
  const source = input.protection === undefined ? {} : assertObject(input.protection, "protection");
  const protection = { ...copyDefined(source, ["locked", "hidden"]) };
  if (input.locked !== undefined) protection.locked = input.locked;
  if (input.hidden !== undefined) protection.hidden = input.hidden;
  if (input.editable !== undefined) protection.locked = !optionalBoolean(input.editable, "editable");
  for (const key of ["locked", "hidden"]) if (protection[key] !== undefined) protection[key] = optionalBoolean(protection[key], `protection.${key}`);
  return protection;
}

export function normalizeStyle(input) {
  assertObject(input, "Style");
  const raw = input.raw === undefined ? {} : assertObject(input.raw, "style.raw");
  const font = normalizeFont(input);
  const alignment = normalizeAlignment(input);
  const { border, rangeBorder } = normalizeBorder(input);
  const protection = normalizeProtection(input);
  const fill = input.fill === undefined ? undefined : normalizeFill(input.fill);
  const numberFormat = input.numberFormat ?? input.format ?? input.z;
  if (numberFormat !== undefined && typeof numberFormat !== "string") throw new TypeError("numberFormat must be a string.");
  return deepMerge(raw, {
    ...(Object.keys(font).length ? { font } : {}),
    ...(fill ? { fill } : {}),
    ...(Object.keys(alignment).length ? { alignment } : {}),
    ...(numberFormat !== undefined ? { numberFormat } : {}),
    ...(Object.keys(border).length ? { border } : {}),
    ...(Object.keys(rangeBorder).length ? { rangeBorder } : {}),
    ...(Object.keys(protection).length ? { protection } : {})
  });
}

export function deepMerge(base, update) {
  const output = { ...(base ?? {}) };
  for (const [key, value] of Object.entries(update ?? {})) {
    const previous = output[key];
    if (isPlainObject(previous) && isPlainObject(value)) output[key] = deepMerge(previous, value);
    else output[key] = value;
  }
  return output;
}

export function composeStyles(...styles) {
  return styles.filter((style) => style !== undefined).reduce((result, style) => deepMerge(result, normalizeStyle(style)), {});
}

function applyBorder(target, side, value) {
  if (value) target[side] = structuredClone(value);
}

export function styleForCell(style) {
  const output = structuredClone(style ?? {});
  const semantic = output.rangeBorder;
  delete output.rangeBorder;
  if (!semantic) return output;
  const border = (output.border ??= {});
  const value = semantic.all ?? semantic.outline ?? semantic.inside;
  for (const side of ["left", "right", "top", "bottom"]) applyBorder(border, side, value);
  applyBorder(border, "left", semantic.insideVertical);
  applyBorder(border, "right", semantic.insideVertical);
  applyBorder(border, "top", semantic.insideHorizontal);
  applyBorder(border, "bottom", semantic.insideHorizontal);
  if (!Object.keys(border).length) delete output.border;
  return output;
}

export function styleForRangeCell(style, point, range) {
  const output = structuredClone(style ?? {});
  const semantic = output.rangeBorder;
  delete output.rangeBorder;
  if (!semantic) return output;
  const border = (output.border ??= {});
  const { r, c } = point;
  const all = semantic.all;
  if (all) for (const side of ["left", "right", "top", "bottom"]) applyBorder(border, side, all);
  const outline = semantic.outline;
  if (r === range.s.r) applyBorder(border, "top", outline);
  if (r === range.e.r) applyBorder(border, "bottom", outline);
  if (c === range.s.c) applyBorder(border, "left", outline);
  if (c === range.e.c) applyBorder(border, "right", outline);
  const vertical = semantic.insideVertical ?? semantic.inside;
  const horizontal = semantic.insideHorizontal ?? semantic.inside;
  if (c > range.s.c) applyBorder(border, "left", vertical);
  if (c < range.e.c) applyBorder(border, "right", vertical);
  if (r > range.s.r) applyBorder(border, "top", horizontal);
  if (r < range.e.r) applyBorder(border, "bottom", horizontal);
  if (!Object.keys(border).length) delete output.border;
  return output;
}

function styleName(value) {
  if (typeof value !== "string" || !value.trim()) throw new TypeError("Style name must be a non-empty string.");
  return value.trim();
}

function normalizeParents(value) {
  if (value === undefined) return [];
  const parents = Array.isArray(value) ? value : [value];
  return parents.map(styleName);
}

export class StyleCollection {
  constructor(workbook) {
    this.workbook = workbook;
  }

  get names() {
    return Object.keys(this.workbook.DefinedStyles ?? {});
  }

  has(name) {
    return Object.hasOwn(this.workbook.DefinedStyles ?? {}, styleName(name));
  }

  define(name, style, options = {}) {
    const key = styleName(name);
    const definitions = (this.workbook.DefinedStyles ??= {});
    const previous = definitions[key];
    definitions[key] = { style: normalizeStyle(style), extends: normalizeParents(options.extends) };
    try {
      this.get(key);
    } catch (error) {
      if (previous) definitions[key] = previous;
      else delete definitions[key];
      throw error;
    }
    this.workbook.RemovedStyleNames = (this.workbook.RemovedStyleNames ?? []).filter((candidate) => candidate !== key);
    return this;
  }

  defineMany(definitions) {
    assertObject(definitions, "Style definitions");
    const target = (this.workbook.DefinedStyles ??= {});
    const snapshot = structuredClone(target);
    try {
      for (const [name, value] of Object.entries(definitions)) {
        const key = styleName(name);
        const wrapped = isPlainObject(value) && value.style !== undefined;
        target[key] = {
          style: normalizeStyle(wrapped ? value.style : value),
          extends: normalizeParents(wrapped ? value.extends : undefined)
        };
      }
      for (const name of Object.keys(definitions)) this.get(name);
    } catch (error) {
      this.workbook.DefinedStyles = snapshot;
      throw error;
    }
    const names = new Set(Object.keys(definitions));
    this.workbook.RemovedStyleNames = (this.workbook.RemovedStyleNames ?? []).filter((candidate) => !names.has(candidate));
    return this;
  }

  get(name) {
    return structuredClone(this.#resolveName(styleName(name), []));
  }

  getDefinition(name) {
    const key = styleName(name);
    const definition = this.workbook.DefinedStyles?.[key];
    if (!definition) throw new ReferenceError(`Unknown style ${JSON.stringify(key)}. Available styles: ${this.names.join(", ") || "none"}.`);
    return { name: key, extends: [...(definition.extends ?? [])], style: structuredClone(definition.style) };
  }

  list() {
    return this.names.map((name) => this.getDefinition(name));
  }

  remove(name) {
    const key = styleName(name);
    if (!this.has(key)) return false;
    const dependent = this.names.find((candidate) => this.workbook.DefinedStyles[candidate].extends?.includes(key));
    if (dependent) throw new Error(`Cannot remove style ${JSON.stringify(key)} because ${JSON.stringify(dependent)} extends it.`);
    delete this.workbook.DefinedStyles[key];
    if (!(this.workbook.RemovedStyleNames ?? []).includes(key)) (this.workbook.RemovedStyleNames ??= []).push(key);
    return true;
  }

  resolve(input) {
    const items = Array.isArray(input) ? input : [input];
    if (!items.length) throw new TypeError("At least one style must be supplied.");
    return items.reduce((result, item) => {
      const style = typeof item === "string" ? this.#resolveName(styleName(item), []) : normalizeStyle(item);
      return deepMerge(result, style);
    }, {});
  }

  #resolveName(name, stack) {
    const definition = this.workbook.DefinedStyles?.[name];
    if (!definition) throw new ReferenceError(`Unknown style ${JSON.stringify(name)}. Available styles: ${this.names.join(", ") || "none"}.`);
    if (stack.includes(name)) throw new Error(`Circular style inheritance: ${[...stack, name].join(" -> ")}.`);
    const parents = (definition.extends ?? []).map((parent) => this.#resolveName(parent, [...stack, name]));
    return deepMerge(parents.reduce((result, parent) => deepMerge(result, parent), {}), definition.style);
  }
}
