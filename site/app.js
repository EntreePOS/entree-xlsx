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

document.querySelectorAll(".copy-button").forEach((button) => {
  button.addEventListener("click", async () => {
    const explicitText = button.dataset.copy;
    const code = button.closest(".code-panel")?.querySelector("code")?.textContent;
    const value = explicitText ?? code;
    if (!value) return;

    await navigator.clipboard.writeText(value);
    const original = button.textContent;
    button.textContent = "Copied";
    window.setTimeout(() => { button.textContent = original; }, 1400);
  });
});

const links = [...document.querySelectorAll(".contents a[href^='#lesson-']")];
const lessons = [...document.querySelectorAll(".lesson")];

const observer = new IntersectionObserver((entries) => {
  const visible = entries
    .filter((entry) => entry.isIntersecting)
    .sort((a, b) => b.intersectionRatio - a.intersectionRatio)[0];
  if (!visible) return;

  links.forEach((link) => {
    const active = link.hash === `#${visible.target.id}`;
    link.classList.toggle("active", active);
    if (active) link.setAttribute("aria-current", "step");
    else link.removeAttribute("aria-current");
  });
}, { rootMargin: "-18% 0px -64%", threshold: [0, .25, .5] });

lessons.forEach((lesson) => observer.observe(lesson));
