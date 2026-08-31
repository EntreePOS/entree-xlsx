const root = document.documentElement;
const themeButton = document.querySelector(".theme-button");
const savedTheme = localStorage.getItem("entree-xlsx-theme");

if (savedTheme === "light" || savedTheme === "dark") {
  root.dataset.theme = savedTheme;
} else if (window.matchMedia("(prefers-color-scheme: dark)").matches) {
  root.dataset.theme = "dark";
}

function currentTheme() {
  return root.dataset.theme === "dark" ? "dark" : "light";
}

function updateThemeLabel() {
  themeButton.textContent = currentTheme() === "dark" ? "Light" : "Dark";
  themeButton.setAttribute("aria-label", `Use ${themeButton.textContent.toLowerCase()} theme`);
}

updateThemeLabel();

themeButton.addEventListener("click", () => {
  const next = currentTheme() === "dark" ? "light" : "dark";
  root.dataset.theme = next;
  localStorage.setItem("entree-xlsx-theme", next);
  updateThemeLabel();
});

async function copyText(value) {
  try {
    await navigator.clipboard.writeText(value);
    return;
  } catch {
    const textarea = document.createElement("textarea");
    textarea.value = value;
    textarea.setAttribute("readonly", "");
    textarea.style.position = "fixed";
    textarea.style.opacity = "0";
    document.body.append(textarea);
    textarea.select();
    const copied = document.execCommand("copy");
    textarea.remove();
    if (!copied) throw new Error("The browser blocked clipboard access.");
  }
}

document.querySelectorAll(".copy-button").forEach((button) => {
  button.addEventListener("click", async () => {
    const explicitText = button.dataset.copy;
    const code = button.closest(".code-panel")?.querySelector("code")?.textContent;
    const value = explicitText ?? code;
    if (!value) return;

    const original = button.textContent;
    try {
      await copyText(value);
      button.textContent = "Copied";
    } catch {
      button.textContent = "Copy failed";
    }
    window.setTimeout(() => { button.textContent = original; }, 1400);
  });
});

const links = [...document.querySelectorAll(".contents a[href^='#lesson-']")];
const lessons = [...document.querySelectorAll(".lesson")];

const visibleLessons = new Map();
const observer = new IntersectionObserver((entries) => {
  entries.forEach((entry) => {
    if (entry.isIntersecting) visibleLessons.set(entry.target.id, entry);
    else visibleLessons.delete(entry.target.id);
  });
  const visible = [...visibleLessons.values()]
    .sort((a, b) => Math.abs(a.boundingClientRect.top) - Math.abs(b.boundingClientRect.top))[0];
  if (!visible) return;

  links.forEach((link) => {
    const active = link.hash === `#${visible.target.id}`;
    link.classList.toggle("active", active);
    if (active) link.setAttribute("aria-current", "step");
    else link.removeAttribute("aria-current");
  });
}, { rootMargin: "-18% 0px -64%", threshold: [0, .25, .5] });

lessons.forEach((lesson) => observer.observe(lesson));

const parameterHelp = {
  address: "An Excel cell address such as A1 or D12.",
  before: "The one-based row number where new rows will be inserted.",
  bytes: "XLSX file data as a Buffer, Uint8Array, or ArrayBuffer.",
  callback: "A function called once for every cell in the range.",
  changes: "An object containing only the properties you want to update.",
  column: "A column letter, such as B, or a supported column reference.",
  config: "The complete configuration for the new PivotTable.",
  count: "How many rows the operation should insert, delete, or copy.",
  data: "A two-dimensional row array or an array of JavaScript objects.",
  definitions: "An object containing multiple named style definitions.",
  format: "An Excel number-format code such as $#,##0.00 or 0.0%.",
  formula: "An Excel formula expression stored in the cell.",
  height: "The row height in points.",
  mode: "How the new style should combine with existing formatting.",
  name: "The name used for the workbook item being created or requested.",
  options: "Optional settings that change how this method runs.",
  parts: "The style properties to clear while leaving other formatting intact.",
  password: "The password required to open the encrypted workbook.",
  path: "The local file path to read from or save to.",
  records: "An array of objects where object keys map to column headers.",
  range: "An Excel cell range such as A1:D20.",
  reference: "A name, zero-based index, or object ID identifying an existing item.",
  result: "An optional cached value shown before Excel recalculates the formula.",
  row: "The one-based worksheet row number.",
  rows: "A two-dimensional array where each inner array is one worksheet row.",
  sheet: "A worksheet name, zero-based index, or Worksheet object.",
  source: "The workbook input, cell, range, or style being read or copied.",
  start: "The first one-based row affected by the operation.",
  style: "A named style, style object, or list of styles to apply.",
  target: "The destination cell, row, or hyperlink URL used by this method.",
  tooltip: "Optional hover text displayed for the hyperlink in Excel.",
  value: "The JavaScript value to store in the cell.",
  width: "The Excel column width to apply."
};

const signatureHelp = {
  "createWorkbook(name?)": {
    name: "The name of the first worksheet. Excel uses Sheet1 when omitted."
  },
  "openWorkbook(source, options?)": {
    source: "A local path, URL, Buffer, Uint8Array, or ArrayBuffer containing an XLSX file."
  },
  "workbook.addSheet(name, data?)": {
    name: "The unique name shown on the new worksheet tab.",
    data: "Optional rows or object records used to populate the new sheet immediately."
  },
  "styles.define(name, style, options?)": {
    name: "The reusable style name you will pass to cell.style() or range.style()."
  },
  "cell.copyStyleFrom(source, mode?)": {
    source: "The Cell whose formatting should be copied."
  },
  "range.copyStyleFrom(source, options?)": {
    source: "The source Range whose formatting should be copied."
  },
  "cell.hyperlink(target, tooltip?)": {
    target: "The URL, email address, or workbook location opened by the link."
  },
  "charts.list(sheet?)": {
    sheet: "Optional worksheet filter. Omit it to list charts from the entire workbook."
  },
  "pivotTables.list(sheet?)": {
    sheet: "Optional worksheet filter. Omit it to list every PivotTable."
  },
  "workbook.save(path, { password })": {
    password: "The password Excel will require before opening the encrypted file."
  }
};

const parameterTooltip = document.createElement("div");
parameterTooltip.className = "parameter-tooltip";
parameterTooltip.id = "parameter-tooltip";
parameterTooltip.setAttribute("role", "tooltip");
parameterTooltip.setAttribute("aria-hidden", "true");
parameterTooltip.hidden = true;
parameterTooltip.innerHTML = '<strong class="parameter-tooltip-name"></strong><span class="parameter-tooltip-copy"></span>';
document.body.append(parameterTooltip);

const tooltipName = parameterTooltip.querySelector(".parameter-tooltip-name");
const tooltipCopy = parameterTooltip.querySelector(".parameter-tooltip-copy");
let activeParameter;

function positionParameterTooltip() {
  if (!activeParameter || parameterTooltip.hidden) return;
  const target = activeParameter.getBoundingClientRect();
  const tooltip = parameterTooltip.getBoundingClientRect();
  const gap = 10;
  let placement = "top";
  let top = target.top - tooltip.height - gap;
  if (top < 12) {
    placement = "bottom";
    top = target.bottom + gap;
  }
  const left = Math.min(
    window.innerWidth - tooltip.width - 12,
    Math.max(12, target.left + (target.width - tooltip.width) / 2)
  );
  top = Math.min(window.innerHeight - tooltip.height - 12, Math.max(12, top));
  const arrowLeft = Math.min(
    tooltip.width - 16,
    Math.max(16, target.left + target.width / 2 - left)
  );
  parameterTooltip.dataset.placement = placement;
  parameterTooltip.style.setProperty("--tooltip-arrow-x", `${Math.round(arrowLeft)}px`);
  parameterTooltip.style.transform = `translate3d(${Math.round(left)}px, ${Math.round(top)}px, 0)`;
}

function showParameterTooltip(parameter) {
  activeParameter = parameter;
  tooltipName.textContent = parameter.textContent;
  tooltipCopy.textContent = parameter.dataset.tooltip;
  parameterTooltip.hidden = false;
  parameterTooltip.setAttribute("aria-hidden", "false");
  positionParameterTooltip();
}

function hideParameterTooltip(parameter) {
  if (parameter && activeParameter !== parameter) return;
  activeParameter = undefined;
  parameterTooltip.hidden = true;
  parameterTooltip.setAttribute("aria-hidden", "true");
}

document.querySelectorAll(".api-table td:first-child code").forEach((code) => {
  const signature = code.textContent;
  const open = signature.indexOf("(");
  const close = signature.lastIndexOf(")");
  if (open < 0 || close <= open + 1) return;

  const prefix = signature.slice(0, open + 1);
  const parameters = signature.slice(open + 1, close);
  const suffix = signature.slice(close);
  const fragment = document.createDocumentFragment();
  fragment.append(document.createTextNode(prefix));

  let cursor = 0;
  for (const match of parameters.matchAll(/[A-Za-z_$][\w$]*\??/g)) {
    fragment.append(document.createTextNode(parameters.slice(cursor, match.index)));
    const label = match[0];
    const name = label.replace(/\?$/, "");
    const description = signatureHelp[signature]?.[name] ?? parameterHelp[name];
    if (!description) {
      fragment.append(document.createTextNode(label));
    } else {
      const parameter = document.createElement("span");
      parameter.className = "api-param";
      parameter.tabIndex = 0;
      parameter.textContent = label;
      parameter.dataset.tooltip = description;
      parameter.setAttribute("aria-describedby", parameterTooltip.id);
      parameter.setAttribute("aria-label", `${name} parameter: ${description}`);
      parameter.addEventListener("pointerenter", () => showParameterTooltip(parameter));
      parameter.addEventListener("pointerleave", () => {
        if (document.activeElement !== parameter) hideParameterTooltip(parameter);
      });
      parameter.addEventListener("focus", () => showParameterTooltip(parameter));
      parameter.addEventListener("blur", () => hideParameterTooltip(parameter));
      parameter.addEventListener("keydown", (event) => {
        if (event.key === "Escape") parameter.blur();
      });
      fragment.append(parameter);
    }
    cursor = match.index + label.length;
  }

  fragment.append(document.createTextNode(parameters.slice(cursor)));
  fragment.append(document.createTextNode(suffix));
  code.replaceChildren(fragment);
});

window.addEventListener("resize", positionParameterTooltip);
window.addEventListener("scroll", positionParameterTooltip, { passive: true });
