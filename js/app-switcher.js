/* App switcher — top-left dropdown listing every ScribNet app.
   The menu renders from the APPS array below: one place to edit when apps are added.
   (Statement Analyzer is intentionally listed only on its own page.) */
const APPS = [
  { name: "SE Command Center", glyph: "🎛️", url: "https://scribnetai.github.io/se-command-center/" },
  { name: "Server Sizer", glyph: "🖥️", url: "https://scribnetai.github.io/server-sizer/" },
  { name: "Storage Sizer", glyph: "💾", url: "https://scribnetai.github.io/storage-sizer/" },
  { name: "Network Sizer", glyph: "🌐", url: "https://scribnetai.github.io/network-sizer/" },
  { name: "RVTools Analyzer", glyph: "🔍", url: "https://scribnetai.github.io/rvtools-analyzer/" },
  { name: "Deal Pack", glyph: "🤝", url: "https://scribnetai.github.io/deal-pack/" },
  { name: "Battlecards", glyph: "⚔️", url: "https://scribnetai.github.io/battlecards/" },
  { name: "TCO Calculator", glyph: "🧮", url: "https://scribnetai.github.io/tco-calculator/" },
  { name: "Cargo Foundry", glyph: "🏭", url: "https://scribnetai.github.io/cargo-foundry/" },
  { name: "scribnet.io", glyph: "🏠", url: "https://scribnet.io/" },
];

/* What the "this page" item does: exit any app view back to the landing. */
function appSwitcherHome() {
  ["wizard", "dashboard", "workspace"].forEach((id) => {
    const el = document.getElementById(id);
    if (el) el.hidden = true;
  });
  const landing = document.getElementById("landing");
  if (landing) landing.hidden = false;
  window.scrollTo({ top: 0 });
}

/* This page — rendered as "this page", not a link. */
const APP_SWITCHER_CURRENT = "Network Sizer";

(function initAppSwitcher() {
  const dd = document.getElementById("brandDropdown");
  const btn = document.getElementById("brandHome");
  const menu = document.getElementById("brandMenu");
  if (!dd || !btn || !menu) return;
  const esc = (s) => String(s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  menu.innerHTML = APPS.map((a) => a.name === APP_SWITCHER_CURRENT
    ? `<button type="button" class="brand-item is-current" data-home><span class="bi">${esc(a.glyph)}</span>${esc(a.name)}<span class="cur">this page</span></button>`
    : `<a class="brand-item" href="${esc(a.url)}"><span class="bi">${esc(a.glyph)}</span>${esc(a.name)}</a>`
  ).join("");
  const close = () => { dd.classList.remove("open"); btn.setAttribute("aria-expanded", "false"); };
  btn.addEventListener("click", (e) => {
    e.preventDefault();
    e.stopPropagation();
    const open = dd.classList.toggle("open");
    btn.setAttribute("aria-expanded", String(open));
  });
  document.addEventListener("click", (e) => { if (!dd.contains(e.target)) close(); });
  document.addEventListener("keydown", (e) => { if (e.key === "Escape") close(); });
  const home = menu.querySelector("[data-home]");
  if (home) home.addEventListener("click", () => { close(); appSwitcherHome(); });
})();
