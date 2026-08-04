// Gera public/catalogo-zenith.html a partir da planilha de preços e dos assets
// em public/assets/. Rode com: npm run build:catalog

import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import XLSX from "xlsx";
import sharp from "sharp";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(__dirname, "..");
const PUBLIC = path.join(ROOT, "public");
const ASSETS = path.join(PUBLIC, "assets");
const COVER_DIR = path.join(ASSETS, "CAPA");
const FONTS_DIR = path.join(ROOT, "build-assets", "fonts");
const XLSX_PATH = path.join(PUBLIC, "Tabela final atualizada .xlsx");
const OUT_PATH = path.join(PUBLIC, "catalogo-zenith.html");

// ---------- planilha ----------
function parsePrice(v) {
  if (typeof v === "number") return v;
  if (typeof v === "string") {
    const cleaned = v.replace(/[^0-9,.-]/g, "").replace(/\./g, "").replace(",", ".");
    const n = parseFloat(cleaned);
    return isNaN(n) ? null : n;
  }
  return null;
}

function readCatalogData() {
  const wb = XLSX.readFile(XLSX_PATH);
  const data = {};
  for (const sheetName of wb.SheetNames) {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, defval: "" });
    const brands = {};
    const order = [];
    for (let i = 1; i < rows.length; i++) {
      const [grupo, marca, produto, apres, valor] = rows[i];
      if (!marca || !produto) continue;
      const key = marca.toString().trim();
      if (!brands[key]) {
        brands[key] = { grupo: (grupo || "").toString().trim(), produtos: [] };
        order.push(key);
      }
      brands[key].produtos.push({
        nome: produto.toString().trim(),
        apresentacao: (apres || "").toString().trim(),
        preco: parsePrice(valor),
      });
    }
    data[sheetName] = order.map((k) => ({ marca: k, grupo: brands[k].grupo, produtos: brands[k].produtos }));
  }
  return data;
}

// ---------- helpers ----------
function slugify(s) {
  return s.normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-+|-+$/g, "");
}
function esc(s) {
  return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}
const currency = new Intl.NumberFormat("pt-BR", { style: "currency", currency: "BRL" });
function money(v) {
  if (v === null || v === undefined || v === "") return "Consultar";
  return currency.format(v);
}
function cleanText(s) {
  return (s || "").toString().replace(/\s+/g, " ").trim();
}

const BRAND_OVERRIDES = {
  LANDERGOLD: "Landergold",
  "IDN PHARMATECH": "IDN Pharmatech",
  "NEO PEPTIDES": "Neo Peptides",
  "MUSCLE LABS BLACK": "Muscle Labs Black",
};
function brandLabel(name) {
  return BRAND_OVERRIDES[name.toUpperCase()] || name;
}

const GROUP_LABEL_OVERRIDES = {
  EMAGRECEDORES: "CANETAS E EMAGRECEDORES",
  IMPORTADOS: "LINHA IMPORTADA",
};
function groupLabel(name) {
  const cleaned = cleanText(name);
  return GROUP_LABEL_OVERRIDES[cleaned.toUpperCase()] || cleaned;
}

// Marca -> arquivo de capa dentro de public/assets/CAPAS
const PHOTO_FILE_MAP = {
  "FARMÁCIA": "FARMÁCIA.png",
  MANIPULADOS: "MANIPULADOS.png",
  VARIADOS: "VARIADOS.png",
  ALLUVI: "ALLUVI.png",
  "ALPHA PHARMA": "ALPHA-PHARMA.png",
  BIOVANT: "BIOVANT-PEPTIDES.png",
  "CANADA LABS": "canada.png",
  "COOPER PHARMA": "cooper.png",
  "DRAGON ELITE": "dragon-elite.png",
  "EMINENCE LABS": "EMINENCE.png",
  "EMPORIO PHARMA": "EMPORIO-PHARMA.png",
  "ÉTICOS": "ETICOS.png",
  FIOCRUZ: "FIOCRUZ.png",
  "GEN-TIRZ": "GEN-TIRZ.png",
  "HEALTH PEPTÍDEOS": "HEALTH.png",
  "IDN PHARMATECH": "IDN.png",
  INDUFAR: "INDUFAR.png",
  KIGTROPIN: "KING-PHARMA.png",
  "KING PHARMA": "KING-PHARMA.png",
  LANDERGOLD: "LANDER-GOLD.png",
  LANDERLAN: "LANDERLAN.png",
  "MUSCLE LABS": "MUSCLE.png",
  "MUSCLE LABS BLACK": "MUSCLE.png",
  "NEO PEPTIDES": "NEO-PEPTIDES.png",
  NEUROCEPTIX: "NEUROCEPTIX.png",
  NOVAX: "NOVAX.png",
  "OXYGEN KW": "OXYGEN.png",
  "PHARMACOM LABS": "pharmacom.png",
  PRAMIL: "PRAMIL.png",
  PURITY: "PURITY.png",
  QUIMFA: "quimfa.png",
  "ROYAL PHARMACEUTICAL": "royal.png",
  SYNEDICA: "synedica.png",
  "SYNTHOL ALEMANHA": "SYNTHOL-ALEMANHA.png",
  "SYNTHOL USA": "SYNTHOL-USA.png",
  VITAPULMIN: "VITAPULMIN.png",
  ZPHCD: "zphcd.png",
};

// Marca -> arquivo SVG em "MARCAS LOGO/SVG MARCAS" (fallback quando não há capa)
const SVG_FILE_MAP = {
  "ALPHA PHARMA": "ALPHA PHARMA.svg",
  BIOVANT: "BIOVANT.svg",
  "CANADA LABS": "CANADÁ.svg",
  "COOPER PHARMA": "COOPER.svg",
  "DRAGON ELITE": "DRAGON ELITE.svg",
  "EMINENCE LABS": "EMINENCE.svg",
  "ÉTICOS": "ETICOS.svg",
  "GEN-TIRZ": "GEN TIRZ.svg",
  "HEALTH PEPTÍDEOS": "HEALTH.svg",
  "IDN PHARMATECH": "IDN PHARMATECH.svg",
  INDUFAR: "INDUFAR.svg",
  "KING PHARMA": "KING PHARMA.svg",
  LANDERGOLD: "LANDER GOLD.svg",
  LANDERLAN: "LANDELAN.svg",
  "MUSCLE LABS": "MUSCLE.svg",
  "MUSCLE LABS BLACK": "MUSCLE.svg",
  "NEO PEPTIDES": "NEO PEPTIDES.svg",
  NEUROCEPTIX: "NEUROCEPTIX.svg",
  NOVAX: "NOVAX.svg",
  "OXYGEN KW": "OXYGEN.svg",
  "PHARMACOM LABS": "PHARMACOM.svg",
  QUIMFA: "QUIMFA.svg",
  "ROYAL PHARMACEUTICAL": "ROYAL.svg",
  SYNEDICA: "SYNEDICA.svg",
  ZPHCD: "ZPHCD.svg",
};

const CATEGORIES = [
  { key: "EMAGRECECORES", slug: "emagrecedores", title: "Canetas e Emagrecedores" },
  { key: "PEPTIDEOS ", slug: "peptideos", title: "Peptídeos" },
  { key: "MARCAS PREMIUM", slug: "premium", title: "Linha Premium" },
  { key: "MARCAS IMPORTADAS", slug: "importadas", title: "Linha Importada" },
  { key: "GH", slug: "gh", title: "GH" },
  { key: "SARMS + PRODUTOS VARIADOS", slug: "sarms", title: "Sarms &amp; Variados" },
  { key: "FARMÁCIA + MANIPULADOS", slug: "farmacia", title: "Farmácia" },
];

// Mantenha sincronizado com src/data/freight.ts do site principal da Zenith.
const freightTable = [
  { uf: "AC", values: { pac: 130, sedex: 110 } },
  { uf: "AL", values: { transportadora: 120, pac: 100, sedex: 125 } },
  { uf: "AM", values: { transportadora: 120, pac: 100, sedex: 125 } },
  { uf: "AP", values: { transportadora: 120, pac: 100, sedex: 90 } },
  { uf: "BA", values: { transportadora: 80, pac: 80, sedex: 90 } },
  { uf: "CE", values: { transportadora: 80, pac: 72.5, sedex: 105 } },
  { uf: "DF", values: { transportadora: 72, pac: 45, sedex: 60 } },
  { uf: "ES", values: { transportadora: 70, pac: 45, sedex: 70 } },
  { uf: "GO", values: { transportadora: 76, pac: 45, sedex: 70 } },
  { uf: "MA", values: { transportadora: 90, pac: 100, sedex: 125 } },
  { uf: "MG", values: { transportadora: 70, pac: 45, sedex: 70 } },
  { uf: "MS", values: { transportadora: 80, pac: 45, sedex: 85 } },
  { uf: "MT", values: { transportadora: 75, pac: 54, sedex: 90 } },
  { uf: "PA", values: { transportadora: 110, pac: 87, sedex: 105 } },
  { uf: "PB", values: { transportadora: 100, pac: 100, sedex: 125 } },
  { uf: "PE", values: { transportadora: 85, pac: 87, sedex: 115 } },
  { uf: "PI", values: { transportadora: 110, pac: 100, sedex: 125 } },
  { uf: "PR", values: { transportadora: 70, pac: 45, sedex: 60 } },
  { uf: "RJ", values: { transportadora: 70, pac: 45, sedex: 70 } },
  { uf: "RN", values: { transportadora: 100, pac: 100, sedex: 120 } },
  { uf: "RO", values: { transportadora: 170, pac: 130, sedex: 110 } },
  { uf: "RR", values: { transportadora: 120, pac: 120 } },
  { uf: "RS", values: { transportadora: 100, pac: 45, sedex: 70 } },
  { uf: "SC", values: { transportadora: 70, pac: 45, sedex: 70 } },
  { uf: "SE", values: { transportadora: 85, pac: 100, sedex: 125 } },
  { uf: "SP", values: { transportadora: 48, pac: 32, sedex: 40 } },
  { uf: "TO", values: { transportadora: 110, pac: 87, sedex: 105 } },
];

// ---------- assets (imagens/fontes -> base64) ----------
async function toWebpBase64(absPath, width) {
  const buf = await sharp(absPath).resize({ width, withoutEnlargement: true }).webp({ quality: 76 }).toBuffer();
  return buf.toString("base64");
}
function toFontBase64(absPath) {
  return fs.readFileSync(absPath).toString("base64");
}

async function buildImageCache() {
  const cache = {};
  cache.estrela1 = await toWebpBase64(path.join(ASSETS, "estrela1.png"), 240);
  cache.capa = await toWebpBase64(path.join(COVER_DIR, "CAPA-TABELA.png"), 1300);
  cache.carousel = ["FAIXA.svg"].map((file) => ({
    file,
    data: fs.readFileSync(path.join(COVER_DIR, file)).toString("base64"),
  }));

  const photoKeys = [...new Set(Object.values(PHOTO_FILE_MAP))];
  for (const file of photoKeys) {
    cache["photo:" + file] = await toWebpBase64(path.join(ASSETS, "CAPAS", file), 1300);
  }
  return cache;
}

function buildSvgCache() {
  const cache = {};
  const dir = path.join(ASSETS, "MARCAS LOGO", "SVG MARCAS");
  const files = [...new Set(Object.values(SVG_FILE_MAP))];
  for (const file of files) {
    let svg = fs.readFileSync(path.join(dir, file), "utf8");
    svg = svg.replace(/<\?xml[^>]*\?>\s*/, "").trim();
    cache[file] = svg;
  }
  return cache;
}

function buildFontCache() {
  const cache = {};
  for (const w of ["400", "500", "600", "700", "800"]) {
    cache["Barlow-" + w] = toFontBase64(path.join(FONTS_DIR, `Barlow-${w}.woff2`));
  }
  cache["Big-Shoulders-Display-900"] = toFontBase64(path.join(FONTS_DIR, "Big-Shoulders-Display-900.woff2"));
  return cache;
}

// ---------- CSS ----------
function buildFontFaces(fonts) {
  let css = "";
  for (const w of ["400", "500", "600", "700", "800"]) {
    css += `@font-face{font-family:'Barlow';font-style:normal;font-weight:${w};font-display:swap;src:url(data:font/woff2;base64,${fonts["Barlow-" + w]}) format('woff2')}\n`;
  }
  css += `@font-face{font-family:'Big Shoulders Display';font-style:normal;font-weight:700 900;font-display:swap;src:url(data:font/woff2;base64,${fonts["Big-Shoulders-Display-900"]}) format('woff2')}\n`;
  return css;
}

function buildPhotoClasses(images) {
  let css = "";
  const uniqueFiles = [...new Set(Object.values(PHOTO_FILE_MAP))];
  for (const file of uniqueFiles) {
    css += `.photo-${slugify(file)}{background-image:url(data:image/webp;base64,${images["photo:" + file]})}\n`;
  }
  return css;
}

function buildHeroCarousel(images) {
  const items = images.carousel
    .map(({ file, data }) => `<img src="data:image/svg+xml;base64,${data}" alt="${esc(path.basename(file, ".svg"))}">`)
    .join("");
  return `<div class="hero-marquee" aria-label="Marcas em destaque">
      <div class="hero-marquee-track">
        <div class="hero-marquee-set">${items}</div>
        <div class="hero-marquee-set" aria-hidden="true">${items}</div>
      </div>
    </div>`;
}

const CSS_BASE = (images, freightTable) => `
:root{
  --bg:#060606;
  --panel:#101010;
  --panel-2:#171717;
  --line:rgba(255,255,255,.12);
  --line-strong:rgba(255,255,255,.34);
  --plate-hi:#f5f6f8;
  --plate-mid:#c6cad1;
  --plate-lo:#7d818a;
  --silver:#eef0f2;
  --silver-soft:#ffffff;
  --muted:#9a9ea5;
  --muted-2:#5c5f65;
  --radius-lg:20px;
  --radius-md:14px;
  --radius-sm:10px;
  --shadow:0 24px 70px rgba(0,0,0,.65);
  --display:'Big Shoulders Display','Arial Narrow',Impact,sans-serif;
  --body:'Barlow',system-ui,-apple-system,Segoe UI,Roboto,Arial,sans-serif;
  --maxw:640px;
}
@media (prefers-reduced-motion: reduce){
  *{animation-duration:.01ms !important;animation-iteration-count:1 !important;transition-duration:.01ms !important;scroll-behavior:auto !important}
  .hero-marquee-track{animation:none !important;transform:none !important}
}
*{box-sizing:border-box}
html{background:var(--bg);scroll-behavior:smooth}
body{
  margin:0;
  min-height:100vh;
  background:
    radial-gradient(60rem 34rem at 82% -8%, rgba(255,255,255,.05), transparent 60%),
    radial-gradient(50rem 34rem at -12% 18%, rgba(255,255,255,.03), transparent 55%),
    linear-gradient(180deg,#0b0b0b 0%, var(--bg) 26rem);
  color:var(--silver);
  font-family:var(--body);
  -webkit-font-smoothing:antialiased;
  text-rendering:optimizeLegibility;
  padding-bottom:env(safe-area-inset-bottom);
}
img{max-width:100%;display:block}
a{color:inherit}
.page{max-width:var(--maxw);margin:0 auto;padding:0 12px 90px}

/* ---------- steel plate device ---------- */
.plate{
  position:relative;border-radius:10px;border:1px solid rgba(0,0,0,.45);
  background:linear-gradient(128deg, var(--plate-hi) 0%, var(--plate-mid) 52%, var(--plate-lo) 100%);
  box-shadow:inset 0 1px 0 rgba(255,255,255,.6), inset 0 -2px 4px rgba(0,0,0,.25), 0 6px 14px rgba(0,0,0,.45);
}
.plate::after{content:"";position:absolute;inset:0;border-radius:inherit;background:linear-gradient(120deg, rgba(255,255,255,.4), transparent 40%);mix-blend-mode:screen;pointer-events:none}

/* ---------- hero ---------- */
.hero{max-width:var(--maxw);margin:0 auto;position:relative;overflow:hidden;background:#000;line-height:0}
.hero img{width:100%;height:auto;display:block}
.hero-scrim{position:absolute;left:0;right:0;bottom:clamp(54px,10vw,68px);height:22%;background:linear-gradient(180deg, transparent, var(--bg))}
.hero-marquee{position:relative;height:clamp(54px,10vw,68px);display:flex;align-items:center;overflow:hidden;background:#092fd2;border-top:1px solid rgba(255,255,255,.25);box-shadow:0 -10px 28px rgba(0,0,0,.28);line-height:normal}
.hero-marquee::before,.hero-marquee::after{content:"";position:absolute;z-index:2;top:0;bottom:0;width:42px;pointer-events:none}
.hero-marquee::before{left:0;background:linear-gradient(90deg,#092fd2,transparent)}
.hero-marquee::after{right:0;background:linear-gradient(270deg,#092fd2,transparent)}
.hero-marquee-track{display:flex;align-items:center;width:max-content;will-change:transform;animation:hero-marquee 12s linear infinite}
.hero-marquee-set{display:flex;align-items:center;gap:30px;padding-right:30px;flex:none}
.hero-marquee-set img{flex:none;width:auto;height:clamp(31px,6vw,42px);display:block}
.hero-marquee:hover .hero-marquee-track{animation-play-state:paused}
@keyframes hero-marquee{to{transform:translateX(-50%)}}

/* ---------- sticky nav ---------- */
.nav-wrap{position:sticky;top:0;z-index:30;background:rgba(6,6,6,.88);backdrop-filter:blur(14px);border-bottom:1px solid var(--line);padding:10px 0}
.nav{max-width:var(--maxw);margin:0 auto;display:flex;align-items:center;gap:8px;padding:0 12px;overflow-x:auto;scrollbar-width:none}
.nav::-webkit-scrollbar{display:none}
.nav-brand{flex:0 0 auto;width:34px;height:34px;border-radius:50%;background:var(--logo-seal);background-size:cover;box-shadow:0 0 0 1px var(--line-strong)}
.nav a{flex:0 0 auto;font-family:var(--display);font-weight:700;font-size:12.5px;letter-spacing:.03em;text-transform:uppercase;text-decoration:none;color:var(--silver);background:linear-gradient(180deg,rgba(255,255,255,.08),rgba(0,0,0,.24));border:1px solid var(--line);padding:9px 13px;border-radius:999px;white-space:nowrap}
.nav a:active{background:linear-gradient(180deg,rgba(255,255,255,.2),rgba(0,0,0,.24))}
#live-category-nav{display:contents}
.search-wrap{max-width:var(--maxw);margin:9px auto 0;padding:0 12px}
.catalog-search{position:relative;display:flex;align-items:center}
.search-icon{position:absolute;left:13px;width:19px;height:19px;color:var(--muted);pointer-events:none}
.search-input{width:100%;height:44px;padding:0 46px 0 42px;border:1px solid var(--line-strong);border-radius:999px;background:linear-gradient(180deg,rgba(255,255,255,.09),rgba(255,255,255,.035));color:var(--silver-soft);font:600 14px var(--body);outline:none;box-shadow:inset 0 1px 0 rgba(255,255,255,.06)}
.search-input::placeholder{color:var(--muted)}
.search-input::-webkit-search-cancel-button{-webkit-appearance:none;appearance:none;display:none}
.search-input:focus{border-color:rgba(255,255,255,.75);box-shadow:0 0 0 3px rgba(255,255,255,.08)}
.search-clear{position:absolute;right:7px;width:32px;height:32px;border:0;border-radius:50%;background:rgba(255,255,255,.09);color:var(--silver);font-size:20px;line-height:1;cursor:pointer}
.search-clear:hover{background:rgba(255,255,255,.16)}
.search-status{min-height:18px;margin:6px 7px 0;color:var(--muted);font-size:11.5px;font-weight:600}
.search-status.is-empty{color:#e3b7b7}
[hidden]{display:none !important}
@media (max-width:640px){
  .search-input{font-size:16px}
  .hero-marquee-track{animation-duration:7s}
}

/* ---------- frete (100% CSS, sem dependência de JavaScript) ---------- */
.frete-body{padding:4px 12px 16px}
.uf-radio{position:absolute;width:0;height:0;opacity:0;pointer-events:none}
.uf-grid{display:flex;flex-wrap:wrap;gap:7px}
.uf-chip{display:block;cursor:pointer;user-select:none;-webkit-user-select:none;text-align:center;min-width:42px;font-family:var(--display);font-weight:700;font-size:13px;letter-spacing:.02em;padding:9px 4px;border:1px solid var(--line);border-radius:9px;background:rgba(255,255,255,.03);color:var(--silver)}
.uf-radio:checked + .uf-chip{background:linear-gradient(180deg,var(--plate-hi),var(--plate-mid));color:#111214;border-color:var(--plate-hi);font-weight:800}
.uf-radio:focus-visible + .uf-chip{outline:2px solid var(--silver-soft);outline-offset:2px}
.frete-results{margin-top:14px}
.frete-result-panel{display:none;grid-template-columns:repeat(3,1fr);gap:8px}
.frete-result-panel div{border:1px solid var(--line);border-radius:var(--radius-sm);background:rgba(255,255,255,.03);padding:10px 6px;text-align:center}
.frete-result-panel span{display:block;color:var(--muted);font-size:10px;font-weight:800;text-transform:uppercase;letter-spacing:.04em}
.frete-result-panel strong{display:block;margin-top:5px;font-size:16px;font-weight:800;color:var(--silver-soft)}
.frete-hint{margin:12px 2px 0;font-size:11.5px;color:var(--muted-2)}
${freightTable.map((s) => `#uf-${s.uf}:checked ~ .frete-results [data-uf="${s.uf}"]{display:grid}`).join("")}

/* ---------- informações ---------- */
.info-body{padding:4px 14px 16px}
.info-highlight{display:flex;align-items:center;gap:12px;padding:14px 16px;margin-bottom:16px}
.info-highlight strong{display:block;font-family:var(--display);font-weight:800;font-size:15px;text-transform:uppercase;letter-spacing:.01em;color:#141516;line-height:1.25}
.info-list{list-style:none;margin:0;padding:0;display:flex;flex-direction:column;gap:16px}
.info-item{padding-top:14px;border-top:1px solid var(--line)}
.info-item:first-child{padding-top:0;border-top:none}
.info-item .info-label{display:block;font-family:var(--display);font-weight:700;font-size:13px;letter-spacing:.1em;text-transform:uppercase;color:var(--silver-soft);margin-bottom:5px}
.info-item p{margin:0;font-size:13.5px;line-height:1.55;color:var(--muted)}
.info-item p + p{margin-top:6px}

/* ---------- category accordion ---------- */
.category{margin:22px 0;border:1px solid var(--line);border-radius:var(--radius-lg);background:linear-gradient(180deg,rgba(255,255,255,.03),rgba(0,0,0,.4)),var(--panel);box-shadow:var(--shadow);overflow:hidden;scroll-margin-top:70px}
.category summary{list-style:none;cursor:pointer;display:flex;align-items:center;gap:12px;padding:16px 16px;background:linear-gradient(90deg,rgba(255,255,255,.06),rgba(0,0,0,.15))}
.category summary::-webkit-details-marker{display:none}
.cat-icon{flex:0 0 auto;width:46px;height:46px;border-radius:50%;border:1px solid var(--line-strong);background-color:rgba(255,255,255,.05);background-image:var(--logo-seal);background-repeat:no-repeat;background-position:center;background-size:cover}
.cat-icon.is-frete{background-image:none;display:flex;align-items:center;justify-content:center}
.cat-icon.is-frete svg{width:22px;height:22px}
.cat-title{flex:1;min-width:0}
.cat-title h2{margin:0;font-family:var(--display);font-weight:800;font-size:clamp(22px,7.5vw,32px);line-height:1;text-transform:uppercase;color:var(--silver-soft)}
.cat-title small{display:block;margin-top:4px;color:var(--muted);font-size:11.5px;font-weight:700;text-transform:uppercase;letter-spacing:.05em}
.cat-chevron{flex:0 0 auto;width:32px;height:32px;border:1px solid #3158ff;border-radius:50%;display:grid;place-items:center;color:#fff;background:#092fd2;box-shadow:0 0 0 3px rgba(9,47,210,.15),0 6px 16px rgba(9,47,210,.28);font-size:19px;font-weight:900;transition:transform .2s ease,background-color .2s ease,box-shadow .2s ease}
.category[open] .cat-chevron{transform:rotate(45deg);background:#3158ff;box-shadow:0 0 0 4px rgba(49,88,255,.18),0 8px 20px rgba(9,47,210,.34)}
.category-body{padding:6px 12px 14px}

/* ---------- brand card / plate ---------- */
.brand-card{margin:10px 0;border:1px solid var(--line);border-radius:var(--radius-md);background:rgba(255,255,255,.02);overflow:hidden}
.brand-kicker{display:flex;align-items:flex-start;justify-content:space-between;gap:10px;padding:12px 14px 0}
.brand-kicker-left{display:flex;align-items:center;gap:7px;min-width:0}
.brand-kicker-left .mini-seal{flex:0 0 auto;width:16px;height:16px;border-radius:50%;background-image:var(--logo-seal);background-size:cover}
.brand-kicker-left span{font-size:10px;font-weight:800;letter-spacing:.12em;text-transform:uppercase;color:#6f8cff;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.brand-kicker-right{display:flex;flex-direction:column;align-items:flex-end;gap:3px;min-width:0}
.brand-kicker-name{font-family:var(--display);font-weight:700;font-size:13px;letter-spacing:.02em;text-transform:uppercase;color:#6f8cff;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;text-align:right;max-width:100%}
.brand-kicker-count{font-size:9.5px;font-weight:700;letter-spacing:.04em;text-transform:uppercase;color:#4f73ff;white-space:nowrap}
.brand-feature{position:relative;margin:10px 14px 0;min-height:70px;padding:12px 20px;display:flex;align-items:center;justify-content:center}
.brand-feature .logo-svg{width:100%;max-width:220px;height:46px}
.brand-feature .logo-svg svg{width:100%;height:100%;display:block}
.brand-feature .logo-text{font-family:var(--display);font-weight:800;font-size:clamp(17px,6vw,23px);text-transform:uppercase;text-align:center;color:#141516;line-height:1.08;letter-spacing:.01em}

.brand-photo{position:relative;margin:10px 14px 0;aspect-ratio:4.5/1;border-radius:10px;border:1px solid var(--line-strong);background-size:cover;background-position:center;overflow:hidden}

.product-list{list-style:none;margin:12px 0 0;padding:0 12px 10px;border-top:1px solid var(--line)}
.product-row{display:flex;align-items:center;justify-content:space-between;gap:10px;padding:12px 2px;border-bottom:1px solid rgba(255,255,255,.06)}
.product-row:last-child{border-bottom:none}
.product-info{min-width:0}
.product-info .name{display:block;font-size:14.5px;font-weight:600;line-height:1.3}
.product-info .sub{display:block;margin-top:2px;font-size:12px;color:var(--muted-2);line-height:1.3}
.price{flex:0 0 auto;font-weight:800;font-size:14.5px;color:var(--silver-soft);font-variant-numeric:tabular-nums;white-space:nowrap}
.product-row.is-unavailable{opacity:.72}
.product-row.is-unavailable .name,.product-row.is-unavailable .sub{color:var(--muted-2)}
.product-meta{flex:0 0 auto;display:flex;flex-direction:column;align-items:flex-end;gap:4px}
.stock-status{display:inline-flex;align-items:center;border:1px solid rgba(255,107,107,.35);border-radius:999px;padding:2px 7px;background:rgba(160,28,28,.18);color:#ffaaaa;font-size:9px;font-weight:800;letter-spacing:.06em;text-transform:uppercase;white-space:nowrap}
.sync-status{min-height:16px;margin:5px 4px 0;color:var(--muted-2);font-size:10px;font-weight:600;letter-spacing:.02em;text-align:right}
.sync-status.is-live{color:#6f8cff}
.sync-status.is-error{color:#f1b0b0}

/* ---------- footer ---------- */
.site-footer{max-width:var(--maxw);margin:30px auto 0;padding:26px 20px 10px;text-align:center;border-top:1px solid var(--line)}
.site-footer .seal{width:52px;height:52px;margin:0 auto 10px;border-radius:50%;background:var(--logo-seal);background-size:cover}
.site-footer p{margin:0;color:var(--muted-2);font-size:11.5px;letter-spacing:.04em;text-transform:uppercase;font-weight:700}
.site-footer p.credo{color:var(--muted);font-size:12px;margin-top:6px;text-transform:none;letter-spacing:0;font-weight:500}

.top-button{position:fixed;left:50%;bottom:calc(14px + env(safe-area-inset-bottom));transform:translateX(-50%);z-index:40;width:46px;height:46px;border-radius:50%;display:grid;place-items:center;text-decoration:none;background:linear-gradient(180deg,var(--plate-hi),var(--plate-mid));color:#111214;font-weight:900;box-shadow:0 12px 30px rgba(0,0,0,.55);font-size:10px;text-transform:uppercase;font-family:var(--display);letter-spacing:.02em}
`;

// ---------- markup ----------
function buildFreteSection() {
  const radios = freightTable.map((s) => `<input type="radio" name="ufpick" id="uf-${s.uf}" class="uf-radio">`).join("");
  const chips = freightTable.map((s) => `<label for="uf-${s.uf}" class="uf-chip">${s.uf}</label>`).join("");
  const panels = freightTable
    .map((s) => {
      const t = s.values.transportadora != null ? money(s.values.transportadora) : "-";
      const p = s.values.pac != null ? money(s.values.pac) : "-";
      const sx = s.values.sedex != null ? money(s.values.sedex) : "-";
      return `<div class="frete-result-panel" data-uf="${s.uf}"><div><span>Transportadora</span><strong>${t}</strong></div><div><span>PAC</span><strong>${p}</strong></div><div><span>Sedex</span><strong>${sx}</strong></div></div>`;
    })
    .join("\n");

  return `<details class="category" id="frete">
      <summary>
        <span class="cat-icon is-frete" role="img" aria-label="Frete">
          <svg viewBox="0 0 24 24" fill="none" stroke="#eef0f2" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><path d="M3 8.5 12 4l9 4.5-9 4.5-9-4.5Z"/><path d="M3 8.5v7L12 20l9-4.5v-7"/><path d="M12 13v7"/></svg>
        </span>
        <span class="cat-title">
          <h2>Frete</h2>
          <small>Consulte por UF</small>
        </span>
        <span class="cat-chevron">+</span>
      </summary>
      <div class="frete-body">
        ${radios}
        <div class="uf-grid">${chips}</div>
        <div class="frete-results">${panels}</div>
        <p class="frete-hint">Toque na sigla do seu estado para ver o valor de cada modalidade de envio.</p>
      </div>
    </details>`;
}

function buildInfoSection() {
  return `<details class="category" id="informacoes">
      <summary>
        <span class="cat-icon is-frete" role="img" aria-label="Informações">
          <svg viewBox="0 0 24 24" fill="none" stroke="#eef0f2" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="M12 11v6"/><circle cx="12" cy="7.5" r="0.75" fill="#eef0f2" stroke="none"/></svg>
        </span>
        <span class="cat-title">
          <h2>Informações</h2>
          <small>Envio, rastreio e recebimento</small>
        </span>
        <span class="cat-chevron">+</span>
      </summary>
      <div class="info-body">
        <div class="plate info-highlight">
          <strong>Envio em até 48 horas após pagamento</strong>
        </div>
        <ul class="info-list">
          <li class="info-item">
            <span class="info-label">Rastreio</span>
            <p>Somente em envios solicitados via transportadora.</p>
          </li>
          <li class="info-item">
            <span class="info-label">Seguro</span>
            <p>O código de rastreio será fornecido em até 48 horas após a postagem.</p>
            <p>Solicite o código de segunda a sexta-feira, após as 18h.</p>
          </li>
          <li class="info-item">
            <span class="info-label">Recebimento da mercadoria</span>
            <p>Para sua própria segurança, filme a abertura da caixa. O vídeo deve ser 100% nítido, permitindo que a caixa seja visível durante toda a filmagem, além de mostrar o conteúdo ao ser retirado.</p>
            <p>Nosso controle é extremamente rigoroso e tiramos fotos de todos os pedidos antes de embalá-los.</p>
          </li>
          <li class="info-item">
            <span class="info-label">Endereço</span>
            <p>Se o cliente fornecer um endereço ou CEP incorretos e a encomenda for extraviada ou perdida, não nos responsabilizaremos, mesmo que o seguro tenha sido pago.</p>
          </li>
        </ul>
      </div>
    </details>`;
}

function buildBrandCard(categorySlug, brand, images, svgs) {
  const label = brandLabel(brand.marca);
  const brandKey = brand.marca.toUpperCase();
  const id = `${categorySlug}-${slugify(brand.marca)}`;
  const groupName = groupLabel(brand.grupo);
  const count = brand.produtos.length;
  const countLabel = `${count} ${count === 1 ? "produto" : "produtos"}`;

  const rows = brand.produtos
    .map((p) => {
      const sub = cleanText(p.apresentacao);
      return `<li class="product-row">
                <div class="product-info">
                  <span class="name">${esc(cleanText(p.nome))}</span>
                  ${sub ? `<span class="sub">${esc(sub)}</span>` : ""}
                </div>
                <span class="price">${esc(money(p.preco))}</span>
              </li>`;
    })
    .join("\n");

  const photoFile = PHOTO_FILE_MAP[brandKey];
  const svgFile = SVG_FILE_MAP[brandKey];
  let feature;
  if (photoFile) {
    feature = `<div class="brand-photo photo-${slugify(photoFile)}" role="img" aria-label="${esc(label)}"></div>`;
  } else if (svgFile) {
    feature = `<div class="plate brand-feature"><span class="logo-svg" role="img" aria-label="${esc(label)}">${svgs[svgFile]}</span></div>`;
  } else {
    feature = `<div class="plate brand-feature"><span class="logo-text">${esc(label)}</span></div>`;
  }

  return `<section class="brand-card" id="${id}">
          <div class="brand-kicker">
            <span class="brand-kicker-left">
              <span class="mini-seal" role="img" aria-label="Zenith Imports"></span>
              <span>${esc(groupName)}</span>
            </span>
            <span class="brand-kicker-right">
              <span class="brand-kicker-name">${esc(label)}</span>
              <span class="brand-kicker-count">${countLabel}</span>
            </span>
          </div>
          ${feature}
          <ul class="product-list">
            ${rows}
          </ul>
        </section>`;
}

function buildCategory(cat, catalogData, images, svgs) {
  const brands = catalogData[cat.key] || [];
  const totalProducts = brands.reduce((a, b) => a + b.produtos.length, 0);
  const totalBrands = brands.length;
  const cards = brands.map((b) => buildBrandCard(cat.slug, b, images, svgs)).join("\n");

  return `<details class="category catalog-category" id="${cat.slug}">
        <summary>
          <span class="cat-icon" role="img" aria-label="Zenith Imports"></span>
          <span class="cat-title">
            <h2>${cat.title}</h2>
            <small>${totalProducts} produtos &middot; ${totalBrands} marcas</small>
          </span>
          <span class="cat-chevron">+</span>
        </summary>
        <div class="category-body">
          ${cards}
        </div>
      </details>`;
}

function buildNav() {
  const links = CATEGORIES.map((c) => `<a href="#${c.slug}" data-open="${c.slug}">${c.title.replace(/&amp;/g, "&")}</a>`).join("");
  return `<nav class="nav" aria-label="Categorias">
      <span class="nav-brand" role="img" aria-label="Zenith Imports"></span>
      <a href="#frete" data-open="frete">FRETE</a><a href="#informacoes" data-open="informacoes">INFORMAÇÕES</a><span id="live-category-nav">${links}</span>
    </nav>`;
}

function buildSearch() {
  return `<div class="search-wrap">
      <form class="catalog-search" role="search" aria-label="Buscar produtos ou marcas">
        <svg class="search-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" aria-hidden="true"><circle cx="11" cy="11" r="7"></circle><path d="m20 20-4-4"></path></svg>
        <input class="search-input" id="catalog-search" type="search" inputmode="search" autocomplete="off" placeholder="Buscar produto ou marca" aria-describedby="search-status">
        <button class="search-clear" id="search-clear" type="button" aria-label="Limpar busca" hidden>&times;</button>
      </form>
      <div class="search-status" id="search-status" role="status" aria-live="polite"></div>
      <div class="sync-status" id="sync-status" role="status" aria-live="polite">Conectando ao catálogo oficial...</div>
    </div>`;
}

// ---------- main ----------
async function main() {
  console.log("Lendo planilha...");
  const catalogData = readCatalogData();

  console.log("Convertendo imagens...");
  const images = await buildImageCache();
  const svgs = buildSvgCache();
  const fonts = buildFontCache();

  const totalProductsAll = Object.values(catalogData).reduce((a, brands) => a + brands.reduce((x, b) => x + b.produtos.length, 0), 0);
  const totalBrandsAll = Object.values(catalogData).reduce((a, brands) => a + brands.length, 0);

  const html = `<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<title>Catálogo Zenith Imports</title>
<meta name="description" content="Catálogo completo Zenith Imports: emagrecedores, peptídeos, linha premium, importadas, GH, sarms e farmácia. Consulte preços e frete por UF.">
<meta name="theme-color" content="#060606">
<link rel="icon" href="data:image/webp;base64,${images.estrela1}">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
<meta name="robots" content="noindex, nofollow">
<style>
${buildFontFaces(fonts)}
${CSS_BASE(images, freightTable)}
${buildPhotoClasses(images)}
:root{--logo-seal:url(data:image/webp;base64,${images.estrela1})}
</style>
</head>
<body id="topo">
  <figure class="hero">
    <img src="data:image/webp;base64,${images.capa}" alt="Zenith Imports — a escolha de quem busca o melhor. As melhores marcas com entrega garantida.">
    <div class="hero-scrim"></div>
    ${buildHeroCarousel(images)}
  </figure>
  <div class="nav-wrap">
    ${buildNav()}
    ${buildSearch()}
  </div>

  <main class="page">
    ${buildFreteSection()}
    ${buildInfoSection()}
    ${CATEGORIES.map((c) => buildCategory(c, catalogData, images, svgs)).join("\n")}
  </main>

  <footer class="site-footer">
    <span class="seal" role="img" aria-label="Zenith Imports"></span>
    <p>Zenith Imports</p>
    <p class="credo">A escolha de quem busca o melhor</p>
  </footer>

  <a class="top-button" href="#topo" aria-label="Voltar ao topo">Topo</a>

<script>
document.querySelector(".nav").addEventListener("click", function(event){
  var link = event.target.closest("[data-open]");
  if (link) {
    var target = document.getElementById(link.getAttribute("data-open"));
    if (target) target.open = true;
  }
});
function abrirCategoriaDoLink(){
  var id = (window.location.hash || "").replace("#", "");
  if (!id) return;
  var target = document.getElementById(id);
  if (target && target.tagName === "DETAILS") target.open = true;
}
abrirCategoriaDoLink();
window.addEventListener("hashchange", abrirCategoriaDoLink);

var searchInput = document.getElementById("catalog-search");
var searchClear = document.getElementById("search-clear");
var searchStatus = document.getElementById("search-status");
var syncStatus = document.getElementById("sync-status");
var catalogCategories = Array.from(document.querySelectorAll(".catalog-category"));
var searchActive = false;
var liveCatalogSignature = "";
var liveCatalogBusy = false;

var liveCategoryMeta = {
  "CANETAS E EMAGRECEDORES": { slug: "emagrecedores", title: "Canetas e Emagrecedores", order: 10 },
  "EMAGRECEDORES": { slug: "emagrecedores", title: "Canetas e Emagrecedores", order: 10 },
  "PEPTÍDEOS": { slug: "peptideos", title: "Peptídeos", order: 20 },
  "PEPTIDEOS": { slug: "peptideos", title: "Peptídeos", order: 20 },
  "MARCAS PREMIUM": { slug: "premium", title: "Linha Premium", order: 30 },
  "MARCAS IMPORTADAS": { slug: "importadas", title: "Linha Importada", order: 40 },
  "GH": { slug: "gh", title: "GH", order: 50 },
  "SARMS": { slug: "sarms", title: "Sarms & Variados", order: 60 },
  "SARMS + PRODUTOS VARIADOS": { slug: "sarms", title: "Sarms & Variados", order: 60 },
  "FARMÁCIA": { slug: "farmacia", title: "Farmácia", order: 70 },
  "FARMÁCIA + MANIPULADOS": { slug: "farmacia", title: "Farmácia", order: 70 }
};
var livePhotoFiles = ${JSON.stringify(PHOTO_FILE_MAP)};

function normalizeSearch(value){
  return (value || "")
    .normalize("NFD")
    .replace(/[\\u0300-\\u036f]/g, "")
    .toLowerCase()
    .trim();
}

function slugifyLive(value){
  return normalizeSearch(value).replace(/[^a-z0-9]+/g, "-").replace(/^-+|-+$/g, "");
}

function categoryMetaLive(name){
  var normalized = (name || "").trim().toUpperCase();
  return liveCategoryMeta[normalized] || {
    slug: slugifyLive(name) || "categoria",
    title: (name || "Categoria").trim(),
    order: 1000
  };
}

function moneyLive(value){
  return typeof value === "number" && isFinite(value)
    ? new Intl.NumberFormat("pt-BR", { style: "currency", currency: "BRL" }).format(value)
    : "Consultar";
}

function createLiveElement(tag, className, text){
  var element = document.createElement(tag);
  if (className) element.className = className;
  if (text !== undefined && text !== null) element.textContent = text;
  return element;
}

function buildLiveBrandCard(categorySlug, brandName, products){
  var card = createLiveElement("section", "brand-card");
  card.id = categorySlug + "-" + slugifyLive(brandName);

  var kicker = createLiveElement("div", "brand-kicker");
  var kickerLeft = createLiveElement("span", "brand-kicker-left");
  var seal = createLiveElement("span", "mini-seal");
  seal.setAttribute("role", "img");
  seal.setAttribute("aria-label", "Zenith Imports");
  kickerLeft.appendChild(seal);
  var groups = Array.from(new Set(products.map(function(product){ return product.group; }).filter(Boolean)));
  kickerLeft.appendChild(createLiveElement("span", "", groups.join(" • ") || "Catálogo"));

  var kickerRight = createLiveElement("span", "brand-kicker-right");
  kickerRight.appendChild(createLiveElement("span", "brand-kicker-name", brandName));
  kickerRight.appendChild(createLiveElement("span", "brand-kicker-count", products.length + (products.length === 1 ? " produto" : " produtos")));
  kicker.appendChild(kickerLeft);
  kicker.appendChild(kickerRight);
  card.appendChild(kicker);

  var brandKey = (products[0].brand || brandName).toUpperCase();
  var photoFile = livePhotoFiles[brandKey];
  if (photoFile) {
    var photo = createLiveElement("div", "brand-photo photo-" + slugifyLive(photoFile));
    photo.setAttribute("role", "img");
    photo.setAttribute("aria-label", brandName);
    card.appendChild(photo);
  } else {
    var feature = createLiveElement("div", "plate brand-feature");
    feature.appendChild(createLiveElement("span", "logo-text", brandName));
    card.appendChild(feature);
  }

  var list = createLiveElement("ul", "product-list");
  products.forEach(function(product){
    var unavailable = product.status === "out_of_stock";
    var row = createLiveElement("li", "product-row" + (unavailable ? " is-unavailable" : ""));
    row.dataset.productId = product.id;
    var info = createLiveElement("div", "product-info");
    info.appendChild(createLiveElement("span", "name", product.name));
    if (product.presentation) info.appendChild(createLiveElement("span", "sub", product.presentation));
    var meta = createLiveElement("span", "product-meta");
    meta.appendChild(createLiveElement("span", "price", moneyLive(product.finalPrice)));
    if (unavailable) meta.appendChild(createLiveElement("span", "stock-status", "Indisponível"));
    row.appendChild(info);
    row.appendChild(meta);
    list.appendChild(row);
  });
  card.appendChild(list);
  return card;
}

function buildLiveCategory(name, products){
  var meta = categoryMetaLive(name);
  var details = createLiveElement("details", "category catalog-category");
  details.id = meta.slug;
  var summary = createLiveElement("summary");
  var icon = createLiveElement("span", "cat-icon");
  icon.setAttribute("role", "img");
  icon.setAttribute("aria-label", "Zenith Imports");
  var titleWrap = createLiveElement("span", "cat-title");
  titleWrap.appendChild(createLiveElement("h2", "", meta.title));

  var brands = new Map();
  products.forEach(function(product){
    var label = product.displayBrand || product.brand;
    if (!brands.has(label)) brands.set(label, []);
    brands.get(label).push(product);
  });
  titleWrap.appendChild(createLiveElement("small", "", products.length + " produtos · " + brands.size + " marcas"));
  summary.appendChild(icon);
  summary.appendChild(titleWrap);
  summary.appendChild(createLiveElement("span", "cat-chevron", "+"));
  details.appendChild(summary);

  var body = createLiveElement("div", "category-body");
  Array.from(brands.entries())
    .sort(function(a, b){ return a[0].localeCompare(b[0], "pt-BR"); })
    .forEach(function(entry){ body.appendChild(buildLiveBrandCard(meta.slug, entry[0], entry[1])); });
  details.appendChild(body);
  return details;
}

function renderLiveCatalog(products){
  var openCategories = {};
  catalogCategories.forEach(function(category){ openCategories[category.id] = category.open; });

  var groups = new Map();
  products.forEach(function(product){
    if (!groups.has(product.category)) groups.set(product.category, []);
    groups.get(product.category).push(product);
  });

  var entries = Array.from(groups.entries()).sort(function(a, b){
    var metaA = categoryMetaLive(a[0]);
    var metaB = categoryMetaLive(b[0]);
    return metaA.order - metaB.order || metaA.title.localeCompare(metaB.title, "pt-BR");
  });

  var page = document.querySelector("main.page");
  catalogCategories.forEach(function(category){ category.remove(); });
  var fragment = document.createDocumentFragment();
  entries.forEach(function(entry){
    var category = buildLiveCategory(entry[0], entry[1]);
    category.open = openCategories[category.id] === true;
    fragment.appendChild(category);
  });
  page.appendChild(fragment);
  catalogCategories = Array.from(document.querySelectorAll(".catalog-category"));

  var nav = document.getElementById("live-category-nav");
  nav.replaceChildren();
  entries.forEach(function(entry){
    var meta = categoryMetaLive(entry[0]);
    var link = createLiveElement("a", "", meta.title);
    link.href = "#" + meta.slug;
    link.dataset.open = meta.slug;
    nav.appendChild(link);
  });

  filterCatalog();
  abrirCategoriaDoLink();
}

async function syncLiveCatalog(){
  if (liveCatalogBusy || document.visibilityState === "hidden") return;
  liveCatalogBusy = true;
  try {
    var response = await fetch("/api/catalog", { cache: "no-store", headers: { accept: "application/json" } });
    if (!response.ok) throw new Error("HTTP " + response.status);
    var payload = await response.json();
    if (!payload || !Array.isArray(payload.products) || payload.products.length === 0) throw new Error("Catálogo vazio");
    var signature = JSON.stringify(payload.products);
    if (signature !== liveCatalogSignature) {
      renderLiveCatalog(payload.products);
      liveCatalogSignature = signature;
    }
    var unavailableCount = payload.products.filter(function(product){ return product.status === "out_of_stock"; }).length;
    syncStatus.className = "sync-status is-live";
    syncStatus.textContent = "Catálogo oficial sincronizado · " + payload.products.length + " produtos · " + unavailableCount + " indisponíveis";
  } catch (error) {
    syncStatus.className = "sync-status is-error";
    syncStatus.textContent = liveCatalogSignature
      ? "Atualização temporariamente indisponível · mantendo a última versão sincronizada"
      : "Fonte oficial temporariamente indisponível · exibindo a versão de segurança";
  } finally {
    liveCatalogBusy = false;
  }
}

function filterCatalog(){
  var query = normalizeSearch(searchInput.value);
  var isSearching = query.length > 0;
  var totalMatches = 0;

  if (isSearching && !searchActive) {
    catalogCategories.forEach(function(category){
      category.dataset.openBeforeSearch = category.open ? "true" : "false";
    });
  }

  catalogCategories.forEach(function(category){
    var categoryMatches = 0;
    category.querySelectorAll(".brand-card").forEach(function(card){
      var brandMatches = 0;
      var brandCount = card.querySelector(".brand-kicker-count");
      var brandName = card.querySelector(".brand-kicker-name");
      var brandMatchesQuery = isSearching && normalizeSearch(brandName ? brandName.textContent : "").includes(query);
      card.querySelectorAll(".product-row").forEach(function(row){
        var searchable = normalizeSearch(row.textContent);
        var matches = !isSearching || brandMatchesQuery || searchable.includes(query);
        row.hidden = !matches;
        if (matches && isSearching) brandMatches += 1;
      });
      if (isSearching) {
        if (!brandCount.dataset.defaultText) brandCount.dataset.defaultText = brandCount.textContent;
        brandCount.textContent = brandMatches + (brandMatches === 1 ? " resultado" : " resultados");
      } else if (searchActive && brandCount.dataset.defaultText) {
        brandCount.textContent = brandCount.dataset.defaultText;
        delete brandCount.dataset.defaultText;
      }
      card.hidden = isSearching && brandMatches === 0;
      categoryMatches += brandMatches;
    });

    var categoryCount = category.querySelector(".cat-title small");
    if (isSearching) {
      if (!categoryCount.dataset.defaultText) categoryCount.dataset.defaultText = categoryCount.textContent;
      categoryCount.textContent = categoryMatches + (categoryMatches === 1 ? " resultado" : " resultados");
    } else if (searchActive && categoryCount.dataset.defaultText) {
      categoryCount.textContent = categoryCount.dataset.defaultText;
      delete categoryCount.dataset.defaultText;
    }
    category.hidden = isSearching && categoryMatches === 0;
    if (isSearching && categoryMatches > 0) category.open = true;
    if (!isSearching && searchActive) {
      category.open = category.dataset.openBeforeSearch === "true";
      delete category.dataset.openBeforeSearch;
    }
    totalMatches += categoryMatches;
  });

  document.getElementById("frete").hidden = isSearching;
  document.getElementById("informacoes").hidden = isSearching;
  searchClear.hidden = !isSearching;
  searchStatus.classList.toggle("is-empty", isSearching && totalMatches === 0);
  searchStatus.textContent = !isSearching
    ? ""
    : totalMatches === 0
      ? "Nenhum produto encontrado."
      : totalMatches + (totalMatches === 1 ? " produto encontrado." : " produtos encontrados.");
  searchActive = isSearching;
}

searchInput.addEventListener("input", filterCatalog);
searchInput.addEventListener("search", filterCatalog);
searchClear.addEventListener("click", function(){
  searchInput.value = "";
  filterCatalog();
  searchInput.focus();
});
document.querySelector(".catalog-search").addEventListener("submit", function(event){
  event.preventDefault();
});
syncLiveCatalog();
setInterval(syncLiveCatalog, 15000);
window.addEventListener("focus", syncLiveCatalog);
document.addEventListener("visibilitychange", function(){
  if (document.visibilityState === "visible") syncLiveCatalog();
});
</script>
</body>
</html>`;

  fs.writeFileSync(OUT_PATH, html, "utf8");
  console.log("Gerado:", OUT_PATH);
  console.log("Tamanho:", (Buffer.byteLength(html) / 1024 / 1024).toFixed(2), "MB");
  console.log("Produtos:", totalProductsAll, "| Marcas:", totalBrandsAll);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
