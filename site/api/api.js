const content = document.querySelector("#api-content");
const input = document.querySelector("#api-query");
const count = document.querySelector("#api-count");
const clear = document.querySelector("#clear-search");
const empty = document.querySelector("#api-empty");

function prepareLinks(root) {
  root.querySelectorAll('a[href^="#"]').forEach((link) => {
    link.setAttribute("href", `../${link.getAttribute("href")}`);
  });
}

function filterApi() {
  const query = input.value.trim().toLocaleLowerCase();
  const groups = [...content.querySelectorAll(".api-group")];
  let visible = 0;

  for (const group of groups) {
    let groupMatches = 0;
    for (const row of group.querySelectorAll("tbody tr")) {
      const matches = !query || row.textContent.toLocaleLowerCase().includes(query);
      row.hidden = !matches;
      if (matches) groupMatches += 1;
    }
    group.hidden = groupMatches === 0;
    if (query && groupMatches) group.open = true;
    visible += groupMatches;
  }

  count.textContent = `${visible} ${visible === 1 ? "method" : "methods"}${query ? ` matching “${input.value.trim()}”` : ""}`;
  clear.hidden = !query;
  empty.hidden = visible !== 0;
}

async function loadApi() {
  try {
    const response = await fetch("../index.html");
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    const source = new DOMParser().parseFromString(await response.text(), "text/html");
    const cheatsheet = source.querySelector("#cheatsheet");
    if (!cheatsheet) throw new Error("API list was not found.");
    prepareLinks(cheatsheet);
    content.replaceChildren(document.importNode(cheatsheet, true));
    content.setAttribute("aria-busy", "false");
    filterApi();
  } catch (error) {
    content.setAttribute("aria-busy", "false");
    content.innerHTML = `<p class="api-empty">The API list could not be loaded. <a href="https://github.com/huangxuewu/entree-xlsx/blob/main/docs/API.md">Open the complete reference on GitHub.</a></p>`;
    count.textContent = "API unavailable";
  }
}

input.addEventListener("input", filterApi);
clear.addEventListener("click", () => {
  input.value = "";
  filterApi();
  input.focus();
});
document.addEventListener("keydown", (event) => {
  if (event.key === "/" && document.activeElement !== input) {
    event.preventDefault();
    input.focus();
  }
  if (event.key === "Escape" && document.activeElement === input) {
    input.value = "";
    filterApi();
  }
});

loadApi();
