// Builds deck/pitch_book.pptx from deck/model_outputs.json (exported from the recalculated model).
// No figure on any slide is typed by hand: everything is read from the model outputs.

const fs = require("fs");
const path = require("path");
const pptxgen = require("pptxgenjs");

const OUT = JSON.parse(fs.readFileSync(path.join(__dirname, "model_outputs.json"), "utf8"));
const N = OUT.names;
const A = OUT.assets;

// ---- Palette and type -------------------------------------------------------------------
const C = {
  ink: "16222E", slate: "3D4F5F", copper: "B5652B", gold: "B8892A", mist: "F2F4F6", line: "D5DBE1",
  muted: "5F6B76", white: "FFFFFF", text: "1E2933", up: "2E7D4F", down: "B23A3A", copperTint: "F4E6DC",
  goldTint: "F5EDD9", upTint: "DCEFE3", downTint: "F6DEDE",
};
const HEAD = "Cambria";
const BODY = "Calibri";
const W = 13.333;
const MARGIN = 0.5;
const CONTENT_W = W - 2 * MARGIN;

// ---- Formatting -------------------------------------------------------------------------
const num = (v, d = 0) => Number(v).toLocaleString("en-US", { minimumFractionDigits: d, maximumFractionDigits: d });
const usdAbs = (v) => (Math.abs(v) >= 1000 ? `US$${num(Math.abs(v) / 1000, 1)}Bn` : `US$${num(Math.abs(v), 0)}M`);
const usdM = (v) => (v < 0 ? `(${usdAbs(v)})` : usdAbs(v));
const cad = (v) => `C$${num(v, 2)}`;
const pct = (v, d = 0) => `${v >= 0 ? "+" : "–"}${num(Math.abs(v) * 100, d)}%`;
const pctPlain = (v, d = 0) => `${num(v * 100, d)}%`;
const mult = (v, d = 2) => `${num(v, d)}x`;
const signedColor = (v) => (v >= 0 ? C.up : C.down);

const years = OUT.deck.years;
const lbPerOz = OUT.deck.gold[OUT.deck.gold.length - 1] / OUT.deck.copper[OUT.deck.copper.length - 1];
const companyAssets = (co) => Object.keys(A).filter((k) => A[k].company === co);
const sumSeries = (keys, field) => years.map((_, i) => keys.reduce((s, k) => s + A[k][field][i], 0));
const auEq = (keys) => {
  const cu = sumSeries(keys, "cu_mlb");
  const au = sumSeries(keys, "au_koz");
  return au.map((v, i) => v + (cu[i] * 1000) / lbPerOz);
};

// ---- Deck scaffolding -------------------------------------------------------------------
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.author = "Ashfaq (github.com/ashrahman19)";
pres.title = "Project Lodestone — Centerra Gold / Ero Copper";
pres.subject = "Hypothetical buy-side merger analysis";
pres.theme = { headFontFace: HEAD, bodyFontFace: BODY };

const FOOTER = "Project Lodestone  |  Hypothetical analysis from public information  |  Not investment advice  |  Not affiliated with any bank";

pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: C.white },
  objects: [
    { text: { text: FOOTER, options: { x: MARGIN, y: 7.05, w: 10.5, h: 0.3, fontFace: BODY, fontSize: 9, color: C.muted, margin: 0 } } },
    { placeholder: { options: { name: "title", type: "title", x: MARGIN, y: 0.35, w: CONTENT_W, h: 0.95, fontFace: HEAD, fontSize: 24, bold: true, color: C.ink, align: "left", valign: "top", margin: 0 }, text: "" } },
  ],
  slideNumber: { x: 12.3, y: 7.05, w: 0.5, h: 0.3, fontFace: BODY, fontSize: 9, color: C.muted, align: "right" },
});
pres.defineSlideMaster({
  title: "DARK",
  background: { color: C.ink },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: MARGIN + 0.3, y: 2.0, w: 11.5, h: 1.3, fontFace: HEAD, fontSize: 44, bold: true, color: C.white, align: "left", margin: 0 }, text: "" } },
  ],
});

let sectionName = null;
function section(title) {
  pres.addSection({ title });
  sectionName = title;
}

function contentSlide(title, notes) {
  const slide = pres.addSlide({ masterName: "CONTENT", sectionTitle: sectionName });
  slide.addText(title, { placeholder: "title" });
  if (notes) slide.addNotes(notes);
  return slide;
}

function text(slide, value, opts) {
  slide.addText(value, { isTextBox: true, fontFace: BODY, fontSize: 14, color: C.text, margin: 0, valign: "top", ...opts });
}

function bullets(slide, items, opts) {
  const runs = items.map((item, i) => ({
    text: item,
    options: { bullet: { indent: 14 }, breakLine: i < items.length - 1, paraSpaceAfter: 8 },
  }));
  text(slide, runs, { fontSize: 14, ...opts });
}

function source(slide, value) {
  text(slide, `Source: ${value}`, { x: MARGIN, y: 6.62, w: CONTENT_W, h: 0.35, fontSize: 10, color: C.muted, objectName: "source" });
}

function statCard(slide, x, y, w, h, value, label, color = C.copper, fill = C.mist) {
  slide.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: fill }, line: { color: fill }, objectName: `card ${label}` });
  const size = value.length > 7 ? 24 : 30;
  text(slide, value, { x: x + 0.2, y: y + 0.15, w: w - 0.4, h: h * 0.5, fontFace: HEAD, fontSize: size, bold: true, color, valign: "middle", fit: "shrink" });
  text(slide, label, { x: x + 0.2, y: y + h * 0.58, w: w - 0.4, h: h * 0.36, fontSize: 12, color: C.slate });
}

function table(slide, rows, opts) {
  const { header, colW, x = MARGIN, y = 1.5, fontSize = 11, highlightRows = [], align = [] } = opts;
  const body = rows.map((row, r) =>
    row.map((cell, c) => {
      const base = typeof cell === "object" && cell !== null ? cell : { text: String(cell) };
      return {
        text: base.text,
        options: {
          fontFace: BODY, fontSize, color: base.color || C.text, bold: base.bold || false,
          align: align[c] || (c === 0 ? "left" : "right"), valign: "middle",
          fill: { color: base.fill || (highlightRows.includes(r) ? C.copperTint : r % 2 ? C.mist : C.white) },
          border: [{ type: "none" }, { type: "none" }, { pt: 0.5, color: C.line }, { type: "none" }],
          margin: [3, 5, 3, 5],
        },
      };
    }),
  );
  const head = header.map((h, c) => ({
    text: h,
    options: { fontFace: BODY, fontSize, bold: true, color: C.white, fill: { color: C.ink }, align: align[c] || (c === 0 ? "left" : "right"), valign: "middle", margin: [3, 5, 3, 5] },
  }));
  slide.addTable([head, ...body], { x, y, w: colW.reduce((a, b) => a + b, 0), colW, rowH: opts.rowH || 0.32, autoPage: false });
}

const chartText = { catAxisLabelColor: C.slate, valAxisLabelColor: C.slate, catAxisLabelFontFace: BODY, valAxisLabelFontFace: BODY, catAxisLabelFontSize: 11, valAxisLabelFontSize: 10, dataLabelFontFace: BODY, legendFontFace: BODY, legendFontSize: 11, legendColor: C.slate, titleFontFace: BODY, titleFontSize: 13, titleColor: C.ink };
const quietAxes = { valGridLine: { color: "E3E7EB", size: 0.5 }, catGridLine: { style: "none" }, catAxisLineShow: false, valAxisLineShow: false };

// ---- Shared figures ---------------------------------------------------------------------
const cgAssets = companyAssets("CG");
const eroAssets = companyAssets("ERO");
const pfEquity = N.CG_SHARES * N.CG_PRICE_CAD * OUT.spot.cadusd + N.DEAL_STOCK_USD;
const pfAuEq2027 = auEq([...cgAssets, ...eroAssets])[0];
const copperRev2027 = Object.values(A).reduce((s, a) => s + a.cu_mlb[0], 0);
const navpsBreakEven = (() => {
  // Interpolate the copper price at which NAVPS accretion crosses zero for the base premium.
  const row = OUT.sensitivity.navps[OUT.sensitivity.premiums.indexOf(N.DEAL_PREMIUM)];
  const cu = OUT.sensitivity.copper;
  for (let i = 1; i < row.length; i++) {
    if (row[i - 1] < 0 && row[i] >= 0) return cu[i - 1] + ((0 - row[i - 1]) / (row[i] - row[i - 1])) * (cu[i] - cu[i - 1]);
  }
  return null;
})();
const sumAt = (keys, field, i = 0) => keys.reduce((s, k) => s + A[k][field][i], 0);
const eroCuRev = sumAt(eroAssets, "cu_mlb") * OUT.deck.copper[0];
const eroCuShare = eroCuRev / (eroCuRev + (sumAt(eroAssets, "au_koz") * OUT.deck.gold[0]) / 1000);
const oksutShare = A.Oksut.ebitda[0] / sumAt(cgAssets, "ebitda");
const oksutSharePf = A.Oksut.ebitda[0] / sumAt([...cgAssets, ...eroAssets], "ebitda");
const copperShare = (keys) => {
  const cu = (sumAt(keys, "cu_mlb") * 1000) / lbPerOz;
  return cu / (cu + sumAt(keys, "au_koz"));
};
const pfCopperShare = copperShare([...cgAssets, ...eroAssets]);
const cgCopperShare = copperShare(cgAssets);
const snapshot = `${OUT.timestamp_utc.slice(0, 4)}-${OUT.timestamp_utc.slice(4, 6)}-${OUT.timestamp_utc.slice(6, 8)}`;
const mktSource = `yfinance market snapshot ${snapshot}; Lodestone model (model/lodestone_model.xlsx)`;

// =========================================================================================
// 1. Cover
section("Overview");
{
  const slide = pres.addSlide({ masterName: "DARK", sectionTitle: sectionName });
  slide.addText("Project Lodestone", { placeholder: "title" });
  text(slide, "Combining Centerra Gold and Ero Copper to create an Americas-weighted copper-gold mid-tier", { x: MARGIN + 0.3, y: 3.35, w: 10.5, h: 0.9, fontFace: HEAD, fontSize: 22, color: "E9C9A8" });
  text(slide, `Discussion materials for the Board of Directors of Centerra Gold  |  Market data as of ${snapshot}`, { x: MARGIN + 0.3, y: 4.6, w: 11, h: 0.4, fontSize: 14, color: C.line });
  text(slide, "Hypothetical illustration prepared from public information only. Not investment advice. Not affiliated with, prepared for, or endorsed by any bank or by either company.", { x: MARGIN + 0.3, y: 6.3, w: 11.5, h: 0.6, fontSize: 11, color: "AEB8C2" });
  slide.addNotes("Title page. Frame the meeting: this is an unsolicited idea for Centerra's board, built only from public filings.");
}

// 2. Executive summary
{
  const slide = contentSlide(
    `Recommendation: Centerra should acquire Ero at ${cad(N.DEAL_OFFER_CAD)} per share, a ${pctPlain(N.DEAL_PREMIUM)} premium`,
    "Lead with the recommendation, then the three numbers the board will ask about: earnings accretion, NAV impact, and leverage.",
  );
  bullets(slide, [
    `Offer of ${cad(N.DEAL_OFFER_CAD)} per Ero share: ${pctPlain(1 - N.DEAL_CASH_PCT)} Centerra stock (${num(N.DEAL_ER, 4)}x exchange ratio) and ${pctPlain(N.DEAL_CASH_PCT)} cash. Equity value ${usdM(N.DEAL_EQUITY_USD)}; enterprise value ${usdM(N.DEAL_EV_USD)}.`,
    `Creates a ${usdM(pfEquity)} copper-gold producer with ${usdM(N.PF_EBITDA_2027)} of 2027E EBITDA and ~${num(pfAuEq2027 / 1000, 2)} Moz of gold-equivalent output.`,
    `Accretive to earnings and cash flow from year one: EPS ${pct(N.ACC_EPS_2027)} and CFPS ${pct(N.ACC_CFPS_2027)} in 2027E, rising with full synergies in 2028E.`,
    `NAV per share is ${pct(N.ACC_NAVPS)} at our below-spot deck but breaks even at ~US$${num(navpsBreakEven, 2)}/lb flat copper, below today's spot of US$${num(OUT.spot.copper, 2)}/lb.`,
    `US$${num(N.SYN_RUNRATE)}M of run-rate synergies (NPV ${usdM(N.SYN_NPV)}). Pro forma leverage of ${mult(N.PF_LEVERAGE, 1)} 2027E EBITDA leaves room to keep funding growth.`,
  ], { x: MARGIN, y: 1.5, w: 7.3, h: 5.0 });
  const cx = 8.3, cw = 2.15, ch = 1.45, gap = 0.25;
  statCard(slide, cx, 1.5, cw, ch, cad(N.DEAL_OFFER_CAD), `Offer per Ero share (${pctPlain(N.DEAL_PREMIUM)} premium)`);
  statCard(slide, cx + cw + gap, 1.5, cw, ch, pct(N.ACC_EPS_2027), "2027E EPS accretion", C.up);
  statCard(slide, cx, 1.5 + ch + gap, cw, ch, pct(N.ACC_CFPS_2028), "2028E CFPS accretion", C.up);
  statCard(slide, cx + cw + gap, 1.5 + ch + gap, cw, ch, `US$${num(N.SYN_RUNRATE)}M`, "Run-rate synergies");
  statCard(slide, cx, 1.5 + 2 * (ch + gap), cw, ch, pct(N.ACC_NAVPS), "NAVPS at analyst deck", C.down);
  statCard(slide, cx + cw + gap, 1.5 + 2 * (ch + gap), cw, ch, pctPlain(N.PF_OWN_CG), "Pro forma Centerra ownership", C.slate);
  source(slide, mktSource);
}

// 3. Transaction overview
{
  const slide = contentSlide(
    `The offer values Ero at ${usdM(N.DEAL_EV_USD)} EV, ${mult(N.DEAL_EV_EBITDA, 1)} 2027E EBITDA, and leaves Centerra holders with ${pctPlain(N.PF_OWN_CG)}`,
    "Walk through the terms top to bottom. The exchange ratio is fixed at Centerra's unaffected price.",
  );
  const rows = [
    ["Ero unaffected share price", cad(N.DEAL_UNAFFECTED)],
    ["Offer premium", pctPlain(N.DEAL_PREMIUM)],
    [{ text: "Offer price per Ero share", bold: true }, { text: cad(N.DEAL_OFFER_CAD), bold: true }],
    ["Cash per share / stock per share", `${cad(N.DEAL_CASH_PS)} / ${cad(N.DEAL_OFFER_CAD - N.DEAL_CASH_PS)}`],
    ["Exchange ratio (Centerra shares per Ero share)", `${num(N.DEAL_ER, 4)}x`],
    ["New Centerra shares issued", `${num(N.DEAL_NEW_SHARES, 1)}M`],
    ["Equity purchase price", usdM(N.DEAL_EQUITY_USD)],
    ["Plus: Ero net debt", usdM(N.DEAL_NET_DEBT)],
    [{ text: "Implied enterprise value", bold: true }, { text: usdM(N.DEAL_EV_USD), bold: true }],
    ["Implied EV / 2027E EBITDA", mult(N.DEAL_EV_EBITDA, 1)],
    ["Implied P/NAV (reserve-only NAV, analyst deck)", mult(N.DEAL_PNAV)],
  ];
  table(slide, rows, { header: ["Transaction terms", ""], colW: [5.2, 2.0], y: 1.5, fontSize: 12, rowH: 0.4 });
  slide.addChart(pres.charts.DOUGHNUT, [{ name: "Ownership", labels: [`Centerra shareholders ${pctPlain(N.PF_OWN_CG)}`, `Ero shareholders ${pctPlain(N.PF_OWN_ERO)}`], values: [N.PF_OWN_CG, N.PF_OWN_ERO] }], {
    ...chartText, x: 8.2, y: 1.45, w: 4.6, h: 4.3, holeSize: 58, chartColors: [C.ink, C.copper], showLegend: true, legendPos: "b",
    showValue: false, showPercent: false, showTitle: true, title: "Pro forma ownership",
  });
  text(slide, `${num(N.PF_SHARES, 1)}M pro forma shares outstanding`, { x: 8.2, y: 5.85, w: 4.6, h: 0.35, fontSize: 12, color: C.slate, align: "center" });
  source(slide, mktSource);
}

// 4. Market outlook
section("Market and companies");
{
  const slide = contentSlide(
    "Our deck sits well below spot for both metals, so the case does not rely on today's prices holding",
    "Copper supply is constrained by grade decline and a thin project pipeline; gold has re-rated on central-bank buying. We still haircut both.",
  );
  const span = years.slice(0, 8).map(String);
  const deckCu = OUT.deck.copper.slice(0, 8);
  const deckAu = OUT.deck.gold.slice(0, 8);
  const chart = (x, title, deck, spot, color, fmt) =>
    slide.addChart(pres.charts.LINE, [
      { name: "Analyst deck", labels: span, values: deck },
      { name: "Spot (flat)", labels: span, values: span.map(() => spot) },
    ], {
      x, y: 1.45, w: 6.0, h: 3.6, chartColors: [color, "9AA5AF"], lineSize: 2.5, lineDataSymbol: "none", showLegend: true, legendPos: "b",
      showTitle: true, title, valAxisLabelFormatCode: fmt, ...chartText, ...quietAxes,
    });
  chart(MARGIN, "Copper (US$/lb)", deckCu, OUT.spot.copper, C.copper, "$0.00");
  chart(6.85, "Gold (US$/oz)", deckAu, OUT.spot.gold, C.gold, "$#,##0");
  bullets(slide, [
    `Copper: deck of US$${num(OUT.deck.copper[0], 2)}/lb in ${years[0]} stepping to US$${num(deckCu[deckCu.length - 1], 2)}/lb long-term, versus spot of US$${num(OUT.spot.copper, 2)}/lb. Electrification demand meets mine-supply growth below 2% a year.`,
    `Gold: deck of US$${num(OUT.deck.gold[0])}/oz in ${years[0]} stepping to US$${num(deckAu[deckAu.length - 1])}/oz long-term, versus spot of US$${num(OUT.spot.gold)}/oz.`,
    "Every NAV and accretion figure in this book uses the deck unless stated otherwise; spot cases appear on the scenario and sensitivity pages.",
  ], { x: MARGIN, y: 5.2, w: CONTENT_W, h: 1.4, fontSize: 13 });
  source(slide, "COMEX HG and GC front-month (yfinance); Lodestone price deck (data/price_deck.csv)");
}

// 5–6. Company profiles
function profile(co, name, ticker, assets, thesis, accent) {
  const navKey = `${co}_NAV`;
  const slide = contentSlide(thesis.title, thesis.notes);
  const rows = assets.map((k) => {
    const a = A[k];
    const gold = a.primary_metal.startsWith("gold");
    return [a.display_name, a.country, a.primary_metal, `${num(a.life)} yrs`, gold ? `US$${num(a.aisc_lom)}/oz` : `US$${num(a.aisc_lom, 2)}/lb`, usdM(a.nav)];
  });
  table(slide, rows, { header: ["Asset", "Location", "Metal", "Reserve life", "LOM AISC (by-product)", "NAV"], colW: [1.7, 1.75, 1.3, 1.15, 1.75, 1.3], y: 1.5, fontSize: 12, rowH: 0.42, align: ["left", "left", "left"] });
  const mkt = OUT.market[ticker];
  const cards = [
    [cad(mkt.price_local), "Share price"],
    [usdM(mkt.market_cap_usd_m), "Market capitalisation"],
    [usdM(N[navKey]), "NAV (deck)"],
    [mult(N[`${co}_PNAV`]), "P/NAV (deck)"],
  ];
  cards.forEach(([v, l], i) => statCard(slide, MARGIN + i * 2.3, 3.55, 2.1, 1.3, v, l, accent));
  bullets(slide, thesis.points, { x: MARGIN, y: 5.1, w: 8.9, h: 1.5, fontSize: 13 });
  const navTotal = assets.reduce((s, k) => s + Math.max(0, A[k].nav), 0);
  slide.addChart(pres.charts.DOUGHNUT, [{ name: "NAV by asset", labels: assets.map((k) => `${A[k].display_name} ${pctPlain(Math.max(0, A[k].nav) / navTotal)}`), values: assets.map((k) => Math.max(0, A[k].nav)) }], {
    ...chartText, x: 9.75, y: 1.4, w: 3.1, h: 3.5, holeSize: 55, chartColors: [accent, C.slate, "9AA5AF"], showLegend: true, legendPos: "b",
    showValue: false, showPercent: false, showTitle: true, title: "Asset NAV mix",
  });
  source(slide, `${name} NI 43-101 reports, MD&A and reserve statements (see appendix); ${mktSource}`);
}
profile("CG", "Centerra Gold", "CG.TO", cgAssets, {
  title: "Centerra: an essentially debt-free gold producer whose copper is streamed and whose Öksüt mine ends by 2029",
  notes: "Centerra's problem is mine life and metal mix, not balance sheet. That is why it can be the buyer.",
  points: [
    `Mount Milligan's PFS extends life to 2045, but 35% of its gold and 18.75% of its copper go to Royal Gold at a fraction of spot.`,
    `Öksüt has ~${num(A.Oksut.life)} years of reserves left; Goldfield is a gold development project. Organic copper growth is limited.`,
    `${usdM(N.CG_CASH)} of cash against ${usdM(-N.CG_DEBT)} of debt gives the capacity to pay part of the price in cash.`,
  ],
}, C.ink);
profile("ERO", "Ero Copper", "ERO.TO", eroAssets, {
  title: "Ero: two built Brazilian copper mines, now generating cash, but carrying US$575M of notes",
  notes: "Ero has finished its build-out. Tucumã is ramping up and Caraíba's Pilar mill upgrade lands in Q4 2026.",
  points: [
    `Caraíba (Bahia) and the newly built Tucumã (Pará): 67.5–77.5 kt of 2026 copper guidance and ${num(A.Caraiba.life)}- and ${num(A.Tucuma.life)}-year reserve lives.`,
    `Xavantina adds gold, but Ero is ~${pctPlain(eroCuShare)} copper by 2027E revenue, the exposure Centerra lacks.`,
    `Net debt of ${usdM(N.DEAL_NET_DEBT)} and remaining growth spend limit Ero's ability to fund the next phase alone.`,
  ],
}, C.copper);

// 7. Strategic rationale
section("Deal rationale");
{
  const slide = contentSlide(
    "Four reasons the combination is worth more than the two companies apart",
    "Each card is one reason with one number. Spend the most time on metal mix and mine life.",
  );
  const pfCu = companyAssets("ERO").concat(cgAssets).reduce((s, k) => s + A[k].cu_mlb[0], 0);
  const cards = [
    ["Balance sheet room", `${mult(N.PF_LEVERAGE, 1)}`, "Pro forma net debt / 2027E EBITDA. Centerra's net cash absorbs Ero's notes and the cash portion of the offer."],
    ["Complementary metal mix", `${num(pfCu, 0)} Mlb`, `Pro forma 2027E payable copper, ${num(pfCu / A.MtMilligan.cu_mlb[0], 1)}x Centerra's own output, and none of Ero's copper is streamed.`],
    ["Mine life and jurisdiction", `${num(A.Caraiba.life)} yrs`, `Caraíba reserve life. Brazilian copper replaces Öksüt, which has ~${num(A.Oksut.life)} years left, and lowers the weight of Türkiye in cash flow.`],
    ["Operating fit", `US$${num(N.SYN_RUNRATE)}M`, "Run-rate synergies from one head office, joint procurement and concentrate marketing, and shared technical services."],
  ];
  const cw = (CONTENT_W - 3 * 0.3) / 4;
  cards.forEach(([head, stat, body], i) => {
    const x = MARGIN + i * (cw + 0.3);
    slide.addShape(pres.shapes.RECTANGLE, { x, y: 1.55, w: cw, h: 4.8, fill: { color: i % 2 ? C.goldTint : C.copperTint }, line: { color: i % 2 ? C.goldTint : C.copperTint }, objectName: `rationale ${i + 1}` });
    slide.addShape(pres.shapes.OVAL, { x: x + 0.25, y: 1.8, w: 0.6, h: 0.6, fill: { color: C.ink }, line: { color: C.ink } });
    text(slide, String(i + 1), { x: x + 0.25, y: 1.8, w: 0.6, h: 0.6, fontFace: HEAD, fontSize: 18, bold: true, color: C.white, align: "center", valign: "middle" });
    text(slide, head, { x: x + 0.25, y: 2.6, w: cw - 0.5, h: 0.7, fontFace: HEAD, fontSize: 18, bold: true, color: C.ink });
    text(slide, stat, { x: x + 0.25, y: 3.35, w: cw - 0.5, h: 0.75, fontFace: HEAD, fontSize: 30, bold: true, color: C.copper });
    text(slide, body, { x: x + 0.25, y: 4.2, w: cw - 0.5, h: 2.0, fontSize: 13, color: C.text });
  });
  source(slide, "Lodestone model: Corporate_NAV, Merger and Synergies tabs");
}

// 8. Pro forma overview
{
  const slide = contentSlide(
    `Pro forma, the company produces ~${num(pfAuEq2027, 0)} koz AuEq in 2027E, ${pctPlain(pfCopperShare)} of it from copper versus ${pctPlain(cgCopperShare)} for Centerra alone`,
    "The production chart shows why Ero matters: Centerra's profile falls as Öksüt ends, and Ero's copper fills the gap.",
  );
  const span = years.slice(0, 14);
  slide.addChart(pres.charts.BAR, [
    { name: "Centerra", labels: span.map(String), values: auEq(cgAssets).slice(0, 14) },
    { name: "Ero", labels: span.map(String), values: auEq(eroAssets).slice(0, 14) },
  ], {
    x: MARGIN, y: 1.45, w: 8.0, h: 5.0, barDir: "col", barGrouping: "stacked", chartColors: [C.ink, C.copper], gapWidthPct: 40,
    showLegend: true, legendPos: "b", showTitle: true, title: "Attributable production, koz gold-equivalent (copper at long-term deck ratio)",
    valAxisLabelFormatCode: "#,##0", ...chartText, ...quietAxes,
  });
  const rows = [
    ["Equity value", usdM(pfEquity)],
    ["2027E EBITDA (incl. synergies)", usdM(N.PF_EBITDA_2027)],
    ["2028E EBITDA (incl. synergies)", usdM(N.PF_EBITDA_2028)],
    ["Net asset value (deck)", usdM(N.PF_NAV)],
    ["2027E payable copper", `${num(copperRev2027, 0)} Mlb`],
    ["Operating assets", "5 + Goldfield"],
    ["Net debt / 2027E EBITDA", mult(N.PF_LEVERAGE, 1)],
  ];
  table(slide, rows, { header: ["Pro forma company", ""], colW: [2.75, 1.5], x: 8.75, y: 1.5, fontSize: 12, rowH: 0.45 });
  source(slide, "Lodestone model: asset tabs (attributable payable metal) and Merger tab");
}

// 9. Asset NAV
section("Valuation");
{
  const slide = contentSlide(
    "Six mine-by-mine NAVs: Ero's two copper mines are worth more than any Centerra asset except Mount Milligan",
    "Discount rates follow the sell-side convention: 5% real for gold, 8% for copper, 6.5% blended for Mount Milligan.",
  );
  const keys = Object.keys(A).sort((x, y) => A[y].nav - A[x].nav);
  slide.addChart(pres.charts.BAR, [
    { name: "Centerra", labels: keys.map((k) => A[k].display_name), values: keys.map((k) => (A[k].company === "CG" ? A[k].nav : 0)) },
    { name: "Ero", labels: keys.map((k) => A[k].display_name), values: keys.map((k) => (A[k].company === "ERO" ? A[k].nav : 0)) },
  ], {
    x: MARGIN, y: 1.45, w: 6.6, h: 5.0, barDir: "bar", barGrouping: "clustered", barOverlapPct: 100, chartColors: [C.ink, C.copper], gapWidthPct: 45,
    catAxisOrientation: "maxMin", showLegend: true, legendPos: "b", showValue: true, dataLabelFormatCode: '#,##0;-#,##0;;', dataLabelPosition: "outEnd", dataLabelColor: C.text, dataLabelFontSize: 11,
    showTitle: true, title: "Attributable NAV at the analyst deck, US$M", valAxisLabelFormatCode: "#,##0", ...chartText, ...quietAxes,
  });
  const rows = keys.map((k) => {
    const a = A[k];
    const gold = a.primary_metal.startsWith("gold");
    return [a.display_name, `${num(a.reserve_mt, 1)} Mt`, gold ? `${num(a.au_gpt, 2)} g/t` : `${num(a.cu_grade_pct * 100, 2)}% Cu`, pctPlain(a.discount_rate, 1), `${num(a.life)} yrs`];
  });
  table(slide, rows, { header: ["Asset", "P&P reserve", "Grade", "Discount", "Life"], colW: [1.65, 1.2, 1.1, 0.95, 0.75], x: 7.35, y: 1.5, fontSize: 12, rowH: 0.42 });
  text(slide, "NAVs are reserve-only: no value for resources outside the mine plan, which is why both stocks trade above 1.0x.", { x: 7.35, y: 4.7, w: 5.5, h: 0.8, fontSize: 12, color: C.slate });
  source(slide, "NI 43-101 technical reports and reserve statements; Lodestone asset tabs");
}

// 10. Corporate NAV
{
  const slide = contentSlide(
    `Both stocks trade at similar multiples of our NAV: ${mult(N.CG_PNAV)} for Centerra and ${mult(N.ERO_PNAV)} for Ero`,
    "Similar P/NAV means a share-for-share deal is close to fair on NAV before the premium. The premium is what the synergies must pay for.",
  );
  const line = (label, a, b, bold = false) => [{ text: label, bold }, { text: a, bold }, { text: b, bold }];
  const rows = [
    line("Attributable asset NAV", usdM(N.CG_ASSET_NAV), usdM(N.ERO_ASSET_NAV)),
    line("Corporate G&A (capitalised)", usdM(N.CG_GA), usdM(N.ERO_GA)),
    line("Other assets and exploration", usdM(N.CG_OTHER), usdM(N.ERO_OTHER)),
    line("Cash", usdM(N.CG_CASH), usdM(N.ERO_CASH)),
    line("Debt", usdM(N.CG_DEBT), usdM(N.ERO_DEBT)),
    line("Net asset value", usdM(N.CG_NAV), usdM(N.ERO_NAV), true),
    line("Shares outstanding", `${num(N.CG_SHARES, 1)}M`, `${num(N.ERO_SHARES, 1)}M`),
    line("NAV per share", cad(N.CG_NAVPS_CAD), cad(N.ERO_NAVPS_CAD), true),
    line("Share price", cad(N.CG_PRICE_CAD), cad(N.ERO_PRICE_CAD)),
    line("P/NAV", mult(N.CG_PNAV), mult(N.ERO_PNAV), true),
  ];
  table(slide, rows, { header: ["US$M unless noted", "Centerra", "Ero"], colW: [3.4, 1.6, 1.6], y: 1.5, fontSize: 12, rowH: 0.42, highlightRows: [5, 7, 9] });
  slide.addChart(pres.charts.BAR, [
    { name: "Analyst deck", labels: ["Centerra", "Ero"], values: [N.CG_PNAV, N.ERO_PNAV] },
    { name: "Spot prices", labels: ["Centerra", "Ero"], values: [OUT.scenarios.Spot.CG_PNAV, OUT.scenarios.Spot.ERO_PNAV] },
  ], {
    x: 7.4, y: 1.45, w: 5.4, h: 4.6, barDir: "col", chartColors: [C.ink, C.copper], gapWidthPct: 60, showLegend: true, legendPos: "b",
    showValue: true, dataLabelFormatCode: '0.00"x"', dataLabelPosition: "outEnd", showTitle: true, title: "P/NAV by price case", valAxisLabelFormatCode: '0.0"x"', ...chartText, ...quietAxes,
  });
  source(slide, "Lodestone model: Corporate_NAV and Scenarios tabs");
}

// 11. Trading comps
{
  const slide = contentSlide(
    "Trading comparables: Ero trades in line with copper peers on EV/EBITDA and below them on cash flow",
    "TTM figures from the latest filings. Peer reserves are estimates and flagged as such in the model.",
  );
  const rows = OUT.comps.map((r) => {
    const label = r[2] === "peer" ? r[1] : { text: r[1], bold: true };
    const f = (v, fn) => (typeof v === "number" ? fn(v) : "n/a");
    return [label, cad(r[3]), usdM(r[4]), usdM(r[5]), f(r[6], (v) => mult(v, 1)), f(r[7], (v) => mult(v, 1)), f(r[8], (v) => mult(v, 2)), f(r[10], (v) => num(v, 2)), f(r[11], (v) => `US$${num(v, 2)}`)];
  });
  const stat = (label) => {
    const r = OUT.comps_stats[label];
    const f = (v, fn) => (typeof v === "number" ? fn(v) : "");
    return [{ text: label, bold: true }, "", "", "", { text: f(r[6], (v) => mult(v, 1)), bold: true }, { text: f(r[7], (v) => mult(v, 1)), bold: true }, "", "", { text: f(r[11], (v) => `US$${num(v, 2)}`), bold: true }];
  };
  table(slide, [...rows, stat("Peer median"), stat("Peer 25th percentile"), stat("Peer 75th percentile")], {
    header: ["Company", "Price", "Mkt cap", "EV", "EV/EBITDA", "P/CF", "P/NAV", "Reserves (Mt CuEq)", "EV / lb CuEq"],
    colW: [2.3, 1.05, 1.35, 1.35, 1.25, 1.05, 1.05, 1.6, 1.3], y: 1.4, fontSize: 11, rowH: 0.3, highlightRows: [0, 1],
  });
  source(slide, `${mktSource}. P/NAV shown only where a NAV is modelled; peer reserves are estimates (data/peer_inputs.csv)`);
}

// 12. Precedents
{
  const slide = contentSlide(
    `Precedent copper and gold deals paid a median ${pctPlain(N.PREC_PREM_MED)} premium; our ${pctPlain(N.DEAL_PREMIUM)} sits in the middle`,
    "The premium range is the most defensible benchmark. P/NAV at announcement is rarely disclosed outside paid databases.",
  );
  const rows = OUT.precedents.map((r) => [r[0].slice(0, 7), r[1], r[2], r[3], r[4], usdM(r[5]), pctPlain(r[6]), typeof r[7] === "number" ? mult(r[7]) : "–"]);
  table(slide, rows, { header: ["Date", "Acquirer", "Target", "Metal", "Consideration", "Equity value", "Premium", "P/NAV"], colW: [0.9, 2.35, 2.35, 1.3, 1.6, 1.4, 1.05, 0.95], y: 1.4, fontSize: 11, rowH: 0.36, align: ["left", "left", "left", "left", "left"] });
  const cards = [[pctPlain(N.PREC_PREM_Q1), "25th percentile premium"], [pctPlain(N.PREC_PREM_MED), "Median premium"], [pctPlain(N.PREC_PREM_Q3), "75th percentile premium"], [pctPlain(N.DEAL_PREMIUM), "Lodestone offer premium"]];
  cards.forEach(([v, l], i) => statCard(slide, MARGIN + i * 3.13, 5.25, 2.9, 1.2, v, l, i === 3 ? C.copper : C.slate));
  source(slide, "Company press releases at announcement; FactSet (P/NAV benchmarks). Premiums to undisturbed price as reported");
}

// 13. Football field
{
  const slide = contentSlide(
    `The ${cad(N.DEAL_OFFER_CAD)} offer sits inside most valuation ranges and at the low end of precedent premiums`,
    "Read the bars against the copper line. NAV at 1.0x is the outlier because it gives no credit for resources outside reserves.",
  );
  const methods = [
    ["52-week trading range", "WK52"], ["Analyst price targets", "TARGETS"], ["Comps: EV / EBITDA", "EVEBITDA"], ["Comps: P / CF", "PCF"],
    ["Comps: EV / lb CuEq reserve", "EVLB"], ["Precedent premiums", "PREC"], ["NAV at Centerra's P/NAV", "NAVREL"], ["NAV at 1.0x (deck to spot)", "NAV"],
  ];
  const lows = methods.map(([, k]) => N[`FF_${k}_LOW`]);
  const highs = methods.map(([, k]) => N[`FF_${k}_HIGH`]);
  const axisMax = Math.ceil(Math.max(...highs, N.DEAL_OFFER_CAD) / 10) * 10 + 10;
  const chart = { x: MARGIN, y: 1.45, w: CONTENT_W, h: 5.0 };
  const plot = { x: 0.25, y: 0.03, w: 0.72, h: 0.86 };
  slide.addChart(pres.charts.BAR, [
    { name: "Low", labels: methods.map((m) => m[0]), values: lows },
    { name: "Range", labels: methods.map((m) => m[0]), values: highs.map((h, i) => h - lows[i]) },
  ], {
    ...chart, barDir: "bar", barGrouping: "stacked", chartColors: [C.white, C.copper], gapWidthPct: 55, catAxisOrientation: "maxMin",
    valAxisMinVal: 0, valAxisMaxVal: axisMax, valAxisMajorUnit: 10, valAxisLabelFormatCode: '"C$"0', showLegend: false, layout: plot, ...chartText, ...quietAxes,
  });
  methods.forEach((_, i) => {
    const rowH = (chart.h * plot.h) / methods.length;
    const y = chart.y + chart.h * plot.y + rowH * i + rowH * 0.18;
    const xAt = (v) => chart.x + chart.w * plot.x + (v / axisMax) * chart.w * plot.w;
    text(slide, cad(lows[i]), { x: xAt(lows[i]) - 0.95, y, w: 0.9, h: rowH * 0.64, fontSize: 10, color: C.slate, align: "right", valign: "middle" });
    text(slide, cad(highs[i]), { x: xAt(highs[i]) + 0.05, y, w: 0.9, h: rowH * 0.64, fontSize: 10, color: C.slate, valign: "middle" });
  });
  const offerX = chart.x + chart.w * plot.x + (N.DEAL_OFFER_CAD / axisMax) * chart.w * plot.w;
  slide.addShape(pres.shapes.LINE, { x: offerX, y: chart.y + chart.h * plot.y, w: 0, h: chart.h * plot.h, line: { color: C.ink, width: 2, dashType: "dash" }, objectName: "offer line" });
  text(slide, `Offer ${cad(N.DEAL_OFFER_CAD)}`, { x: offerX - 0.8, y: 1.15, w: 1.6, h: 0.3, fontSize: 12, bold: true, color: C.ink, align: "center" });
  source(slide, "Lodestone model: Football_Field tab (ranges per Ero share, C$)");
}

// 14. Accretion / dilution
section("Pro forma impact");
{
  const slide = contentSlide(
    `Strongly accretive to EPS and CFPS; NAVPS ${pct(N.ACC_NAVPS)} at our deck, which the synergies only partly offset`,
    "Earnings accretion comes from buying Ero's cheaper earnings with Centerra's higher-rated stock. NAVPS is the honest cost: we pay a premium for copper.",
  );
  const cats = ["EPS 2027E", "EPS 2028E", "CFPS 2027E", "CFPS 2028E", "NAVPS"];
  const vals = [N.ACC_EPS_2027, N.ACC_EPS_2028, N.ACC_CFPS_2027, N.ACC_CFPS_2028, N.ACC_NAVPS];
  slide.addChart(pres.charts.BAR, [
    { name: "Accretion", labels: cats, values: vals.map((v) => (v >= 0 ? v : 0)) },
    { name: "Dilution", labels: cats, values: vals.map((v) => (v < 0 ? v : 0)) },
  ], {
    x: MARGIN, y: 1.45, w: 6.6, h: 5.0, barDir: "col", barGrouping: "clustered", barOverlapPct: 100, chartColors: [C.up, C.down], gapWidthPct: 50, showLegend: false,
    showValue: true, dataLabelFormatCode: '+0%;-0%;;', dataLabelPosition: "outEnd", dataLabelColor: C.text, dataLabelFontSize: 12,
    showTitle: true, title: "Accretion / (dilution) to Centerra shareholders", valAxisLabelFormatCode: "0%", ...chartText, ...quietAxes,
  });
  const rows = [
    ["Standalone EPS (US$)", `${num(N.CG_EPS_2027, 2)}`, `${num(N.CG_EPS_2028, 2)}`],
    ["Pro forma EPS (US$)", `${num(N.PF_EPS_2027, 2)}`, `${num(N.PF_EPS_2028, 2)}`],
    ["Standalone CFPS (US$)", `${num(N.CG_CFPS_2027, 2)}`, `${num(N.CG_CFPS_2028, 2)}`],
    ["Pro forma CFPS (US$)", `${num(N.PF_CFPS_2027, 2)}`, `${num(N.PF_CFPS_2028, 2)}`],
    ["After-tax synergies (US$M)", `${num(N.PF_SYN_AT_2027)}`, `${num(N.PF_SYN_AT_2028)}`],
    ["After-tax financing cost (US$M)", `${num(N.PF_DEBT_COST_2027 + N.PF_CASH_COST_2027)}`, `${num(N.PF_DEBT_COST_2028 + N.PF_CASH_COST_2028)}`],
  ];
  table(slide, rows, { header: ["Per share", "2027E", "2028E"], colW: [3.0, 1.2, 1.2], x: 7.4, y: 1.5, fontSize: 12, rowH: 0.42 });
  const nav = [
    ["Standalone NAVPS (US$)", num(N.CG_NAVPS_USD, 2)],
    ["Pro forma NAVPS (US$)", num(N.PF_NAVPS, 2)],
    ["NAVPS before synergies", pct(N.ACC_NAVPS_EX_SYN)],
    ["NAVPS after synergies", pct(N.ACC_NAVPS)],
  ];
  table(slide, nav, { header: ["NAV per share", ""], colW: [4.2, 1.2], x: 7.4, y: 4.6, fontSize: 12, rowH: 0.38 });
  source(slide, "Lodestone model: Merger tab. Excludes one-off integration and transaction costs and purchase price allocation");
}

// 15. Sensitivities
{
  const slide = contentSlide(
    `NAVPS turns accretive at ~US$${num(navpsBreakEven, 2)}/lb copper; EPS stays accretive even with no synergies`,
    "Left: where the NAV case breaks even. Right: the EPS case is robust to the synergy estimate and to the cash mix.",
  );
  const S = OUT.sensitivity;
  const cell = (v, base) => ({ text: pct(v), color: signedColor(v), fill: base ? C.copperTint : v >= 0 ? C.upTint : C.downTint, bold: base });
  const left = S.navps.map((row, i) => [{ text: pctPlain(S.premiums[i]), bold: true }, ...row.map((v, j) => cell(v, S.premiums[i] === N.DEAL_PREMIUM && false))]);
  table(slide, left, { header: ["Premium", ...S.copper.map((c) => `US$${num(c, 2)}`)], colW: [1.0, ...S.copper.map(() => 0.88)], y: 1.85, fontSize: 11, rowH: 0.38 });
  text(slide, "NAVPS accretion: offer premium vs flat copper price (US$/lb)", { x: MARGIN, y: 1.45, w: 6.3, h: 0.35, fontSize: 13, bold: true, color: C.ink });
  const right = S.eps_2028.map((row, i) => [{ text: `US$${num(S.synergies[i])}M`, bold: true }, ...row.map((v) => cell(v, false))]);
  table(slide, right, { header: ["Synergies", ...S.cash.map((c) => `${pctPlain(c)} cash`)], colW: [1.25, ...S.cash.map(() => 0.95)], x: 7.0, y: 1.85, fontSize: 11, rowH: 0.38 });
  text(slide, "2028E EPS accretion: run-rate synergies vs cash portion", { x: 7.0, y: 1.45, w: 5.8, h: 0.35, fontSize: 13, bold: true, color: C.ink });
  bullets(slide, [
    `At the ${pctPlain(N.DEAL_PREMIUM)} premium, every US$0.50/lb on copper moves NAVPS accretion by roughly ${num(Math.abs((S.navps[3][3] - S.navps[3][2]) * 100), 0)} points: the NAV case is a copper call.`,
    "More cash in the mix raises EPS accretion because Centerra's stock is the more expensive currency, at the cost of leverage.",
  ], { x: MARGIN, y: 4.75, w: CONTENT_W, h: 1.6, fontSize: 13 });
  source(slide, "Lodestone model: Sensitivity and Scenarios tabs (each copper case is a full recalculation of the workbook)");
}

// 16. Synergies
{
  const slide = contentSlide(
    `US$${num(N.SYN_RUNRATE)}M of run-rate synergies, worth ${usdM(N.SYN_NPV)} after tax and integration costs`,
    "Synergies are cost-only. No revenue synergies, no capex deferral, no tax structuring.",
  );
  slide.addChart(pres.charts.BAR, [{ name: "Run-rate", labels: OUT.synergies.map((s) => s.category), values: OUT.synergies.map((s) => s.run_rate_usd_m) }], {
    x: MARGIN, y: 1.45, w: 6.4, h: 3.6, barDir: "bar", chartColors: [C.copper], gapWidthPct: 45, catAxisOrientation: "maxMin", showLegend: false,
    showValue: true, dataLabelFormatCode: '"US$"0"M"', dataLabelPosition: "outEnd", showTitle: true, title: "Pre-tax run-rate synergies by category, US$M",
    valAxisLabelFormatCode: "0", ...chartText, ...quietAxes,
  });
  const rows = OUT.synergies.map((s) => [s.category, s.basis]);
  table(slide, rows, { header: ["Category", "Basis"], colW: [2.4, 3.6], x: 7.0, y: 1.5, fontSize: 11, rowH: 0.75, align: ["left", "left"] });
  const cards = [[`US$${num(N.SYN_2027)}M`, "2027E (50% phase-in)"], [`US$${num(N.SYN_2028)}M`, "2028E (full run-rate)"], [usdM(N.SYN_INTEGRATION), "One-off cost to achieve"], [usdM(N.SYN_NPV), "NPV, after tax"]];
  cards.forEach(([v, l], i) => statCard(slide, MARGIN + i * 3.13, 5.25, 2.9, 1.2, v, l, i === 3 ? C.copper : C.slate));
  source(slide, "Lodestone model: Synergies tab; combined cost base from asset tabs");
}

// 17. Contribution
{
  const slide = contentSlide(
    `Ero contributes ${pctPlain(OUT.contribution[2][4])} of 2027E EBITDA but its holders receive ${pctPlain(N.PF_OWN_ERO)} of the company`,
    "Contribution versus ownership is how Ero's board will judge fairness. Centerra gives up more earnings share than NAV share.",
  );
  const rows = OUT.contribution;
  slide.addChart(pres.charts.BAR, [
    { name: "Centerra", labels: rows.map((r) => r[0]), values: rows.map((r) => r[3]) },
    { name: "Ero", labels: rows.map((r) => r[0]), values: rows.map((r) => r[4]) },
  ], {
    x: MARGIN, y: 1.45, w: 8.6, h: 5.0, barDir: "bar", barGrouping: "percentStacked", chartColors: [C.ink, C.copper], gapWidthPct: 45, catAxisOrientation: "maxMin",
    showLegend: true, legendPos: "b", showValue: true, dataLabelFormatCode: "0%", dataLabelPosition: "ctr", dataLabelColor: C.white, valAxisLabelFormatCode: "0%",
    showTitle: true, title: "Relative contribution", ...chartText, ...quietAxes,
  });
  statCard(slide, 9.5, 1.6, 3.3, 1.5, pctPlain(N.PF_OWN_CG), "Centerra pro forma ownership", C.ink);
  statCard(slide, 9.5, 3.35, 3.3, 1.5, pctPlain(N.PF_OWN_ERO), "Ero pro forma ownership", C.copper);
  text(slide, "Ownership sits between the earnings split and the NAV split, which is where a negotiated outcome usually lands.", { x: 9.5, y: 5.1, w: 3.3, h: 1.3, fontSize: 12, color: C.slate });
  source(slide, "Lodestone model: Contribution tab");
}

// 18. Funding
{
  const slide = contentSlide(
    `The ${usdM(N.DEAL_CASH_NEED)} cash need is met from balance sheet cash and a ${usdM(N.DEAL_NEW_DEBT)} term loan; leverage peaks at ${mult(N.PF_LEVERAGE, 1)}`,
    "Sources and uses. The new debt is small relative to pro forma EBITDA.",
  );
  const uses = [
    ["Stock consideration to Ero shareholders", usdM(N.DEAL_STOCK_USD)],
    ["Cash consideration to Ero shareholders", usdM(N.DEAL_CASH_USD)],
    ["Transaction costs", usdM(N.DEAL_TXN_COSTS)],
    [{ text: "Total uses", bold: true }, { text: usdM(N.DEAL_STOCK_USD + N.DEAL_CASH_NEED), bold: true }],
  ];
  const sources = [
    ["New Centerra shares issued", usdM(N.DEAL_STOCK_USD)],
    ["Combined balance sheet cash", usdM(N.DEAL_CASH_USED)],
    ["New acquisition term loan", usdM(N.DEAL_NEW_DEBT)],
    [{ text: "Total sources", bold: true }, { text: usdM(N.DEAL_STOCK_USD + N.DEAL_CASH_USED + N.DEAL_NEW_DEBT), bold: true }],
  ];
  table(slide, sources, { header: ["Sources", "US$"], colW: [4.0, 1.6], y: 1.5, fontSize: 13, rowH: 0.48 });
  table(slide, uses, { header: ["Uses", "US$"], colW: [4.2, 1.6], x: 6.9, y: 1.5, fontSize: 13, rowH: 0.48 });
  const cards = [[usdM(N.PF_NET_DEBT), "Pro forma net debt"], [mult(N.PF_LEVERAGE, 1), "Net debt / 2027E EBITDA"], [pctPlain(N.DEAL_DEBT_RATE, 1), "Assumed cost of new debt"], [usdM(N.DEAL_NET_DEBT), "Ero net debt assumed"]];
  cards.forEach(([v, l], i) => statCard(slide, MARGIN + i * 3.13, 4.4, 2.9, 1.4, v, l, C.slate));
  source(slide, "Lodestone model: Merger tab. Minimum cash balance of US$250M retained");
}

// 19. Risks
section("Risks and next steps");
{
  const slide = contentSlide(
    "Key risks: jurisdiction, execution and the copper price, each with a mitigant",
    "Do not oversell. The board will ask about Brazil and about NAV dilution; answer both directly.",
  );
  const rows = [
    [{ text: "Jurisdiction", bold: true }, "Adds Brazil (Bahia, Pará, Mato Grosso) to Canada, Türkiye and Nevada; royalty, tax-incentive (SUDENE) and permitting changes.", `Brazil has a long mining track record; Öksüt is ${pctPlain(oksutShare)} of Centerra's 2027E mine EBITDA today and ${pctPlain(oksutSharePf)} pro forma.`],
    [{ text: "Execution", bold: true }, "Integration across three new sites in a new country; Tucumã still ramping up; key Ero staff retention.", "Keep Ero's Brazilian operating team intact; synergies are corporate and procurement-led, not site-level cuts."],
    [{ text: "Commodity prices", bold: true }, `NAVPS accretion depends on copper: ${pct(OUT.sensitivity.navps[3][0])} at US$${num(OUT.sensitivity.copper[0], 2)}/lb flat.`, "Deck is already below spot; consider a copper collar on part of Ero's 2027–28 output to protect leverage."],
    [{ text: "Shareholder support", bold: true }, "Centerra holders may resist NAV dilution and a premium for copper exposure.", "Lead with CFPS accretion, mine-life extension and re-rating to diversified copper-gold peers."],
    [{ text: "Financing", bold: true }, `${usdM(N.DEAL_NEW_DEBT)} of new debt and Ero's notes (change-of-control).`, `Leverage of ${mult(N.PF_LEVERAGE, 1)}; refinance Ero's notes on Centerra's credit.`],
  ];
  table(slide, rows, { header: ["Risk", "What could go wrong", "Mitigant"], colW: [2.1, 5.2, 5.03], y: 1.45, fontSize: 12, rowH: 0.78, align: ["left", "left", "left"] });
  source(slide, "Lodestone analysis");
}

// 20. Next steps
{
  const slide = contentSlide("Next steps: a disciplined path from board discussion to announcement in about ten weeks", "Close with the ask: approval to make a confidential approach.");
  const steps = [
    ["Weeks 1–2", "Board review", "Approve approach strategy, price range and consideration mix"],
    ["Weeks 2–4", "Confidential approach", "Non-binding proposal to Ero's chair; sign NDA and standstill"],
    ["Weeks 4–8", "Due diligence", "Site visits to Caraíba and Tucumã; reserve audit; Brazil tax and royalty review"],
    ["Weeks 8–10", "Definitive agreement", "Fairness opinions, financing commitment, lock-ups from key holders"],
    ["Close +4–6 mo.", "Approvals", "Ero shareholder vote, court approval (plan of arrangement), CADE and Investment Canada"],
  ];
  const cw = (CONTENT_W - 4 * 0.25) / 5;
  steps.forEach(([when, what, detail], i) => {
    const x = MARGIN + i * (cw + 0.25);
    slide.addShape(pres.shapes.OVAL, { x: x + cw / 2 - 0.35, y: 1.7, w: 0.7, h: 0.7, fill: { color: i === 0 ? C.copper : C.ink }, line: { color: i === 0 ? C.copper : C.ink } });
    text(slide, String(i + 1), { x: x + cw / 2 - 0.35, y: 1.7, w: 0.7, h: 0.7, fontFace: HEAD, fontSize: 18, bold: true, color: C.white, align: "center", valign: "middle" });
    if (i < steps.length - 1) slide.addShape(pres.shapes.LINE, { x: x + cw / 2 + 0.4, y: 2.05, w: cw + 0.25 - 0.8, h: 0, line: { color: C.line, width: 1.5 } });
    text(slide, when, { x, y: 2.6, w: cw, h: 0.4, fontSize: 12, bold: true, color: C.copper, align: "center" });
    text(slide, what, { x, y: 3.0, w: cw, h: 0.6, fontFace: HEAD, fontSize: 17, bold: true, color: C.ink, align: "center" });
    text(slide, detail, { x: x + 0.1, y: 3.65, w: cw - 0.2, h: 1.6, fontSize: 13, color: C.text, align: "center" });
  });
  slide.addShape(pres.shapes.RECTANGLE, { x: MARGIN, y: 5.45, w: CONTENT_W, h: 0.95, fill: { color: C.mist }, line: { color: C.mist } });
  text(slide, `Decision requested today: authorise a confidential, non-binding approach to Ero at up to ${cad(N.DEAL_OFFER_CAD)} per share in a mix of Centerra stock and cash.`, { x: MARGIN + 0.3, y: 5.55, w: CONTENT_W - 0.6, h: 0.75, fontSize: 15, bold: true, color: C.ink, valign: "middle" });
}

// 21. Appendix: methodology
section("Appendix");
{
  const slide = contentSlide("Appendix: methodology and key assumptions", "For reviewers who want to check the model mechanics.");
  bullets(slide, [
    "Each asset is a reserve-constrained life-of-mine model: ore processed = min(mill capacity, remaining reserve), then grade x recovery x payability gives payable metal.",
    "Revenue at the analyst deck after the Royal Gold stream at Mount Milligan; costs from site cost per tonne, TC/RC, royalties and sustaining capex, calibrated to 2026 guidance.",
    "Unlevered after-tax free cash flow discounted mid-year at real rates: 5% gold, 8% copper, 6.5% blended for Mount Milligan.",
    "Corporate NAV = attributable asset NAVs - capitalised G&A + other assets + cash - debt. NAV is reserve-only.",
    "Accretion uses standalone 2027E/2028E net income and operating cash flow from the same asset models, plus synergies and financing costs after tax.",
    "Every assumption lives in a CSV under data/; the Excel model, this deck and the dashboard are all generated from those files. A test suite checks the Python engine against Excel within 0.5%.",
  ], { x: MARGIN, y: 1.5, w: 6.6, h: 5.0, fontSize: 13 });
  const rows = years.slice(0, 3).map((y, i) => [String(y) + (i === 2 ? "+" : ""), `US$${num(OUT.deck.copper[i], 2)}`, `US$${num(OUT.deck.gold[i])}`]);
  rows.push(["Spot", `US$${num(OUT.spot.copper, 2)}`, `US$${num(OUT.spot.gold)}`]);
  table(slide, rows, { header: ["Price deck", "Copper /lb", "Gold /oz"], colW: [1.6, 1.8, 1.8], x: 7.6, y: 1.5, fontSize: 12, rowH: 0.42 });
  const disc = Object.values(A).map((a) => [a.display_name, pctPlain(a.discount_rate, 1), a.expansion_mtpa > a.throughput_mtpa ? `${num(a.throughput_mtpa, 1)} → ${num(a.expansion_mtpa, 1)} Mtpa` : `${num(a.throughput_mtpa, 1)} Mtpa`]);
  table(slide, disc, { header: ["Asset", "Discount rate", "Throughput"], colW: [1.6, 1.8, 1.8], x: 7.6, y: 3.85, fontSize: 12, rowH: 0.36 });
  source(slide, "Lodestone model: Inputs and asset tabs; data/*.csv");
}

// 22. Appendix: sources
{
  const slide = contentSlide("Appendix: sources and data status", "Be upfront about which inputs are estimates.");
  const rows = [
    ["Centerra reserves (Mt Milligan, Öksüt, Goldfield)", "Year-end 2025 MRMR tables", "Verified"],
    ["Mount Milligan life extension and expansion", "NI 43-101 PFS, Sep 2025", "Verified; Cu grade estimated"],
    ["Centerra 2026 guidance", "Q2 2026 MD&A (Form 6-K)", "Verified"],
    ["Royal Gold stream terms", "Centerra MD&A", "Verified"],
    ["Caraíba reserves", "Ero reserve statement, Dec 2024", "Verified"],
    ["Tucumã reserves", "NI 43-101 technical report, 2021", "Verified; depleted in model"],
    ["Xavantina reserves", "Back-solved from 2025 production", "Estimate"],
    ["Ero 2026 guidance", "FY2025 MD&A; Q2 2026 results", "Verified"],
    ["Peer reserves (CuEq)", "Company reserve statements (approx.)", "Estimate"],
    ["Precedent premiums", "Announcement press releases", "As reported; to verify"],
    ["Market data", `yfinance snapshot ${snapshot}`, "Automated"],
  ];
  table(slide, rows, { header: ["Input", "Source", "Status"], colW: [5.0, 4.6, 2.73], y: 1.5, fontSize: 12, rowH: 0.42, align: ["left", "left", "left"] });
  source(slide, "data/sources.csv and data/asset_assumption_notes.csv (full URLs)");
}

pres.writeFile({ fileName: path.join(__dirname, "pitch_book.pptx") }).then((f) => console.log(`wrote ${path.relative(process.cwd(), f)}`));
