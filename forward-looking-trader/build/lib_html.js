// HTML twin of lib.js: same exports, emits an HTML fragment per document (used for the web edition).
const fs = require("fs"), path = require("path");
const esc = s => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
const runs = t => esc(t).replace(/\*\*([^*]+)\*\*/g, "<strong>$1</strong>").replace(/\*([^*]+)\*/g, "<em>$1</em>");
const P = t => `<p>${runs(t)}</p>`;
const H1 = t => `<h2>${runs(t)}</h2>`, H2 = t => `<h3>${runs(t)}</h3>`, H3 = t => `<h4>${runs(t)}</h4>`;
const B = a => [`<ul>${a.map(t => `<li>${runs(t)}</li>`).join("")}</ul>`];
const N = a => [`<ol>${a.map(t => `<li>${runs(t)}</li>`).join("")}</ol>`];
const F = t => `<div class="formula">${esc(t)}</div>`;
const NOTE = t => `<p class="note">${runs(t)}</p>`;
const PB = () => "", TOC = () => [];
const cellHtml = c => String(c).split("\n").map(runs).join("<br>");
function T(headers, rows, widths) {
  const tot = widths.reduce((a, b) => a + b, 0);
  const cg = `<colgroup>${widths.map(w => `<col style="width:${(100 * w / tot).toFixed(1)}%">`).join("")}</colgroup>`;
  const th = headers ? `<thead><tr>${headers.map(h => `<th>${cellHtml(h)}</th>`).join("")}</tr></thead>` : "";
  const tb = `<tbody>${rows.map(r => `<tr>${r.map(c => `<td>${cellHtml(c)}</td>`).join("")}</tr>`).join("")}</tbody>`;
  return [`<div class="tw"><table class="${headers ? "" : "kv"}${widths.length > 3 ? " wide" : ""}">${cg}${th}${tb}</table></div>`];
}
const CALLOUT = (title, lines) => [`<aside class="callout"><p class="callout-t">${runs(title)}</p>${[].concat(lines).map(P).join("")}</aside>`];
const TITLE = (kicker, title, sub, meta) => [`<header class="doc-head"><p class="kicker">${esc(kicker)}</p><h1>${esc(title)}</h1><p class="doc-sub">${esc(sub)}</p>${meta.map(m => `<p class="meta">${runs(m)}</p>`).join("")}</header>`];
async function build(file, shortTitle, children) {
  const out = path.join(process.env.FLT_HTML, path.basename(file).replace(/\.docx$/, ".html"));
  fs.writeFileSync(out, children.flat(Infinity).join("\n")); console.log("wrote", out);
}
module.exports = { P, H1, H2, H3, B, N, F, NOTE, PB, T, CALLOUT, TITLE, TOC, build, VERSION: "web" };
