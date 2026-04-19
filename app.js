let charts = {};
let raw = {};

const fmtInt = (n) => (Number(n) || 0).toLocaleString("en-IN");
const fmtMoney = (n) =>
  `₹${(Number(n) || 0).toLocaleString("en-IN", { maximumFractionDigits: 2, minimumFractionDigits: 2 })}`;
const fmtPct = (x) => `${((Number(x) || 0) * 100).toFixed(2)}%`;

function parseMoney(v) {
  if (v == null) return 0;
  if (typeof v === "number") return v;
  return Number(String(v).replace(/[₹,\s]/g, "")) || 0;
}
function parseNumber(v) {
  if (v == null) return 0;
  if (typeof v === "number") return v;
  return Number(String(v).replace(/[,]/g, "")) || 0;
}
function parsePct(v) {
  if (v == null) return 0;
  if (typeof v === "number") return v;
  const s = String(v).trim();
  if (s.endsWith("%")) return (Number(s.replace("%", "")) || 0) / 100;
  return Number(s) || 0;
}
function excelSerialToDate(n) {
  const ms = Math.round((n - 25569) * 86400 * 1000);
  const d = new Date(ms);
  if (Number.isNaN(d.getTime())) return String(n);
  return d.toISOString().slice(0, 10);
}

function destroyChart(id) {
  if (charts[id]) { charts[id].destroy(); charts[id] = null; }
}
function findSheetStartsWith(wb, prefix) {
  return wb.SheetNames.find(s => s.toLowerCase().startsWith(prefix.toLowerCase())) || null;
}
function sheetToJson(wb, sheetName) {
  const ws = wb.Sheets[sheetName];
  return XLSX.utils.sheet_to_json(ws, { defval: null });
}
function escapeHtml(s) {
  return String(s)
    .replaceAll("&","&amp;")
    .replaceAll("<","&lt;")
    .replaceAll(">","&gt;")
    .replaceAll('"',"&quot;")
    .replaceAll("'","&#039;");
}

// Robust column getter (handles spaces/case/SheetJS weird headers)
function getCol(row, wanted) {
  const w = String(wanted).trim().toLowerCase();
  for (const k of Object.keys(row)) {
    if (String(k).trim().toLowerCase() === w) return row[k];
  }
  return null;
}

/* ---------- Load Excel ---------- */

async function loadExcelFromUrl(url = "data.xlsx") {
  const res = await fetch(url);
  if (!res.ok) throw new Error(`Cannot fetch ${url}. Put data.xlsx next to index.html.`);
  const ab = await res.arrayBuffer();
  const wb = XLSX.read(ab, { type: "array" });
  parseWorkbook(wb);
}

async function loadExcelFromFile(file) {
  const ab = await file.arrayBuffer();
  const wb = XLSX.read(ab, { type: "array" });
  parseWorkbook(wb);
}

function parseWorkbook(wb) {
  const sBig = findSheetStartsWith(wb, "Biggest_changes");
  const sTime = findSheetStartsWith(wb, "Time_series");
  const sWord = wb.SheetNames.find(s => s.toLowerCase().startsWith("searches(word_")) || null;
  const sSearch = wb.SheetNames.find(s => s.toLowerCase().startsWith("searches(search_")) || null;
  const sOpt = findSheetStartsWith(wb, "Optimisation_score");
  const sNet = findSheetStartsWith(wb, "Networks");
  const sDev = findSheetStartsWith(wb, "Devices");
  const sDay = findSheetStartsWith(wb, "Day_&_Hour(by-day)");
  const sHour = findSheetStartsWith(wb, "Day_&_Hour(by hour)");
  const sEach = findSheetStartsWith(wb, "Day_&_Hour(by-each-hour)");

  const missing = [];
  if (!sBig) missing.push("Biggest_changes...");
  if (!sTime) missing.push("Time_series...");
  if (!sWord) missing.push("Searches(Word_...)");
  if (!sSearch) missing.push("Searches(Search_...)");
  if (!sOpt) missing.push("Optimisation_score...");
  if (!sNet) missing.push("Networks...");
  if (!sDev) missing.push("Devices...");
  if (!sDay) missing.push("Day_&_Hour(by-day)");
  if (!sHour) missing.push("Day_&_Hour(by hour)");
  if (!sEach) missing.push("Day_&_Hour(by-each-hour)");

  if (missing.length) {
    alert("Missing sheets:\n- " + missing.join("\n- "));
    console.log("Available sheets:", wb.SheetNames);
    return;
  }

  raw.biggest = sheetToJson(wb, sBig).map(r => ({
    campaign: getCol(r, "Campaign Name"),
    cost: parseMoney(getCol(r, "Cost")),
    costComp: parseMoney(getCol(r, "Cost (Comparison)")),
    clicks: parseNumber(getCol(r, "Clicks")),
    clicksComp: parseNumber(getCol(r, "Clicks (Comparison)")),
    interactions: parseNumber(getCol(r, "Interactions")),
    interactionsComp: parseNumber(getCol(r, "Interactions (Comparison)")),
  })).filter(x => x.campaign);

  raw.time = sheetToJson(wb, sTime).map(r => {
    let d = getCol(r, "Date");
    if (typeof d === "number") d = excelSerialToDate(d);
    else if (d instanceof Date) d = d.toISOString().slice(0,10);
    else d = String(d);

    return {
      date: d,
      clicks: parseNumber(getCol(r, "Clicks")),
      impressions: parseNumber(getCol(r, "Impressions")),
      ctr: parsePct(getCol(r, "CTR")),
      avgCpc: parseMoney(getCol(r, "Avg. CPC")),
    };
  }).filter(x => x.date);

console.log("sWord sheet picked:", sWord);
console.log("wordRows first row:", sheetToJson(wb, sWord)[0]);

  const wordRows = sheetToJson(wb, sWord);

 // --- WORD sheet (Top Keywords) - robust even when headers are broken
const wordWS = wb.Sheets[sWord];

// Read sheet as a 2D array (rows/columns), not as objects
const wordAOA = XLSX.utils.sheet_to_json(wordWS, { header: 1, defval: "" });

// Find the real header row by locating "Clicks" and "Impressions"
let headerRowIndex = 0;
for (let i = 0; i < Math.min(wordAOA.length, 15); i++) {
  const row = wordAOA[i].map(x => String(x).trim().toLowerCase());
  if (row.includes("clicks") && row.includes("impressions")) {
    headerRowIndex = i;
    break;
  }
}

// Build a column index map from that header row
const header = wordAOA[headerRowIndex].map(x => String(x).trim().toLowerCase());
const idx = (name) => header.indexOf(String(name).trim().toLowerCase());

// Column positions (fallbacks if names slightly differ)
const colWord = idx("word") !== -1 ? idx("word") : 0;
const colClicks = idx("clicks");
const colImpr = idx("impressions");
const colCost = idx("cost");
const colConv = idx("conversions");
const colTopQ = idx("top containing queries");

raw.word = wordAOA
  .slice(headerRowIndex + 1)
  .map(r => ({
    keyword: r[colWord],
    clicks: parseNumber(colClicks !== -1 ? r[colClicks] : 0),
    impressions: parseNumber(colImpr !== -1 ? r[colImpr] : 0),
    cost: parseMoney(colCost !== -1 ? r[colCost] : 0),
    conversions: parseNumber(colConv !== -1 ? r[colConv] : 0),
    topQueries: colTopQ !== -1 ? r[colTopQ] : "",
  }))
  .filter(x => x.keyword && String(x.keyword).trim().length > 0);



  raw.search = sheetToJson(wb, sSearch).map(r => ({
    search: getCol(r, "Search"),
    cost: parseMoney(getCol(r, "Cost")),
    clicks: parseNumber(getCol(r, "Clicks")),
    impressions: parseNumber(getCol(r, "Impressions")),
    conversions: parseNumber(getCol(r, "Conversions")),
  })).filter(x => x.search);

  raw.opt = sheetToJson(wb, sOpt).map(r => ({
    campaign: getCol(r, "Campaign Name"),
    score: parsePct(getCol(r, "Optimisation score")),
  })).filter(x => x.campaign);

  raw.networks = sheetToJson(wb, sNet).map(r => ({
    network: getCol(r, "Network"),
    clicks: parseNumber(getCol(r, "Clicks")),
    cost: parseMoney(getCol(r, "Cost")),
    avgCpc: parseMoney(getCol(r, "Avg. CPC")),
  })).filter(x => x.network);

  raw.devices = sheetToJson(wb, sDev).map(r => ({
    device: getCol(r, "Device"),
    cost: parseMoney(getCol(r, "Cost")),
    impressions: parseNumber(getCol(r, "Impressions")),
    clicks: parseNumber(getCol(r, "Clicks")),
  })).filter(x => x.device);

  raw.day = sheetToJson(wb, sDay).map(r => ({
    day: getCol(r, "Day"),
    impressions: parseNumber(getCol(r, "Impressions")),
  })).filter(x => x.day);

  raw.hour = sheetToJson(wb, sHour).map(r => ({
    hour: parseNumber(getCol(r, "Start Hour")),
    impressions: parseNumber(getCol(r, "Impressions")),
  }));

  raw.each = sheetToJson(wb, sEach).map(r => ({
    day: getCol(r, "Day"),
    hour: parseNumber(getCol(r, "Start Hour")),
    impressions: parseNumber(getCol(r, "Impressions")),
  })).filter(x => x.day && Number.isFinite(x.hour));

  initTabs();
  initFilters();
  renderAll();
}

/* ---------- Tabs ---------- */

function initTabs() {
  document.querySelectorAll(".tabBtn").forEach(btn => {
    btn.addEventListener("click", () => {
      document.querySelectorAll(".tabBtn").forEach(b => b.classList.remove("active"));
      btn.classList.add("active");

      const tab = btn.dataset.tab;
      document.querySelectorAll(".section").forEach(s => s.classList.remove("active"));
      document.getElementById(tab).classList.add("active");
    });
  });
}

/* ---------- Filters ---------- */

function initFilters() {
  const campaignSelect = document.getElementById("campaignSelect");
  const networkSelect = document.getElementById("networkSelect");
  const deviceSelect = document.getElementById("deviceSelect");

  const campaigns = [...new Set(raw.biggest.map(x => x.campaign))].sort();

  campaignSelect.innerHTML =
    `<option value="__all__">All campaigns</option>` +
    campaigns.map(c => `<option value="${escapeHtml(c)}">${escapeHtml(c)}</option>`).join("");

  networkSelect.innerHTML =
    `<option value="__all__">All networks</option>` +
    raw.networks.map(n => `<option value="${escapeHtml(n.network)}">${escapeHtml(n.network)}</option>`).join("");

  deviceSelect.innerHTML =
    `<option value="__all__">All devices</option>` +
    raw.devices.map(d => `<option value="${escapeHtml(d.device)}">${escapeHtml(d.device)}</option>`).join("");

  [campaignSelect, networkSelect, deviceSelect].forEach(sel => sel.addEventListener("change", renderAll));
}

function getSel() {
  return {
    campaign: document.getElementById("campaignSelect").value,
    network: document.getElementById("networkSelect").value,
    device: document.getElementById("deviceSelect").value,
  };
}

/* ---------- KPIs ---------- */

function overallTotals() {
  const clicks = raw.time.reduce((s,r)=>s+r.clicks,0);
  const impressions = raw.time.reduce((s,r)=>s+r.impressions,0);
  const cost = raw.networks.reduce((s,r)=>s+r.cost,0);
  return { clicks, impressions, cost };
}

function computeKpis(sel) {
  let { clicks, impressions, cost } = overallTotals();

  if (sel.campaign !== "__all__") {
    const c = raw.biggest.find(x => x.campaign === sel.campaign);
    if (c) { clicks = c.clicks; cost = c.cost; }
  }

  if (sel.network !== "__all__") {
    const n = raw.networks.find(x => x.network === sel.network);
    if (n) { clicks = n.clicks; cost = n.cost; }
  }

  if (sel.device !== "__all__") {
    const d = raw.devices.find(x => x.device === sel.device);
    if (d) { clicks = d.clicks; impressions = d.impressions; cost = d.cost; }
  }

  const ctr = impressions > 0 ? clicks / impressions : 0;
  const avgCpc = clicks > 0 ? cost / clicks : 0;
  return { clicks, impressions, ctr, avgCpc, cost };
}

/* ---------- Render ---------- */

function renderAll() {
  const sel = getSel();
  const k = computeKpis(sel);

  document.getElementById("kpiClicks").textContent = fmtInt(k.clicks);
  document.getElementById("kpiImpr").textContent = fmtInt(k.impressions);
  document.getElementById("kpiCtr").textContent = fmtPct(k.ctr);
  document.getElementById("kpiCpc").textContent = fmtMoney(k.avgCpc);

  renderOverviewCharts();
  renderOpt();
  renderKeywordAndSearch();
  renderCampaignTable();
  renderTimePatterns();
  renderInsightsAndLimits(sel, k);
}

function renderOverviewCharts() {
  const ts = raw.time.slice().sort((a,b)=>a.date.localeCompare(b.date));
  const step = Math.max(1, Math.floor(ts.length / 12));
  const sampled = ts.filter((_,i)=> i % step === 0).slice(0, 12);

  const labels = sampled.map(x=>x.date);

  destroyChart("impressionsLine");
  charts.impressionsLine = new Chart(document.getElementById("impressionsLine"), {
    type:"line",
    data:{ labels, datasets:[{ label:"Impressions", data: sampled.map(x=>x.impressions) }] },
    options:{ responsive:true, plugins:{ legend:{ display:true } } }
  });

  destroyChart("clicksLine");
  charts.clicksLine = new Chart(document.getElementById("clicksLine"), {
    type:"line",
    data:{ labels, datasets:[{ label:"Clicks", data: sampled.map(x=>x.clicks) }] },
    options:{ responsive:true, plugins:{ legend:{ display:true } } }
  });

  destroyChart("networkBar");
  charts.networkBar = new Chart(document.getElementById("networkBar"), {
    type:"bar",
    data:{
      labels: raw.networks.map(x=>x.network),
      datasets:[{ label:"Clicks", data: raw.networks.map(x=>x.clicks) }]
    },
    options:{ responsive:true, plugins:{ legend:{ display:true } } }
  });

  destroyChart("deviceBar");
  charts.deviceBar = new Chart(document.getElementById("deviceBar"), {
    type:"bar",
    data:{
      labels: raw.devices.map(x=>x.device),
      datasets:[{ label:"Clicks", data: raw.devices.map(x=>x.clicks) }]
    },
    options:{ responsive:true, plugins:{ legend:{ display:true } } }
  });
}

function renderOpt() {
  destroyChart("optScoreBar");
  charts.optScoreBar = new Chart(document.getElementById("optScoreBar"), {
    type:"bar",
    data:{
      labels: raw.opt.map(x=>x.campaign),
      datasets:[{ label:"Optimization score", data: raw.opt.map(x=>x.score) }]
    },
    options:{
      responsive:true,
      scales:{ y:{ min:0, max:1 } },
      plugins:{ legend:{ display:true } }
    }
  });
}

function renderKeywordAndSearch() {
  const topKw = raw.word.slice().sort((a,b)=>b.clicks-a.clicks).slice(0,10);
  destroyChart("keywordBar");
  charts.keywordBar = new Chart(document.getElementById("keywordBar"), {
    type:"bar",
    data:{ labels: topKw.map(x=>x.keyword), datasets:[{ label:"Clicks", data: topKw.map(x=>x.clicks) }] },
    options:{ responsive:true, plugins:{ legend:{ display:true } } }
  });

  const topS = raw.search.slice().sort((a,b)=>b.clicks-a.clicks).slice(0,10);
  destroyChart("searchBar");
  charts.searchBar = new Chart(document.getElementById("searchBar"), {
    type:"bar",
    data:{ labels: topS.map(x=>x.search), datasets:[{ label:"Clicks", data: topS.map(x=>x.clicks) }] },
    options:{ responsive:true, plugins:{ legend:{ display:true } } }
  });

  fillTable(
    "keywordTable",
    ["Word","Clicks","Impressions","Cost","Conversions","Top Containing Queries"],
    raw.word.slice().sort((a,b)=>b.clicks-a.clicks).slice(0,50).map(x=>[
      x.keyword, fmtInt(x.clicks), fmtInt(x.impressions), fmtMoney(x.cost), fmtInt(x.conversions), x.topQueries ?? ""
    ])
  );

  fillTable(
    "searchTable",
    ["Search","Clicks","Impressions","Cost","Conversions"],
    raw.search.slice().sort((a,b)=>b.clicks-a.clicks).slice(0,50).map(x=>[
      x.search, fmtInt(x.clicks), fmtInt(x.impressions), fmtMoney(x.cost), fmtInt(x.conversions)
    ])
  );
}

function renderCampaignTable() {
  const rows = raw.biggest.slice().sort((a,b)=>b.cost-a.cost);
  fillTable(
    "campaignTable",
    ["Campaign","Cost","Cost (Comp)","Clicks","Clicks (Comp)","Interactions","Interactions (Comp)"],
    rows.map(x=>[
      x.campaign, fmtMoney(x.cost), fmtMoney(x.costComp), fmtInt(x.clicks), fmtInt(x.clicksComp),
      fmtInt(x.interactions), fmtInt(x.interactionsComp)
    ])
  );
}

function renderTimePatterns() {
  destroyChart("dayBar");
  charts.dayBar = new Chart(document.getElementById("dayBar"), {
    type:"bar",
    data:{ labels: raw.day.map(x=>x.day), datasets:[{ label:"Impressions", data: raw.day.map(x=>x.impressions) }] },
    options:{ responsive:true, plugins:{ legend:{ display:true } } }
  });

  const byHour = Array.from({length:24}, (_,h)=> {
    const r = raw.hour.find(x=>x.hour===h);
    return r ? r.impressions : 0;
  });

  destroyChart("hourBar");
  charts.hourBar = new Chart(document.getElementById("hourBar"), {
    type:"bar",
    data:{ labels: Array.from({length:24}, (_,h)=>String(h)), datasets:[{ label:"Impressions", data: byHour }] },
    options:{ responsive:true, plugins:{ legend:{ display:true } } }
  });

  renderHeatmap();
}

function renderHeatmap() {
  const daysOrder = ["Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"];
  const grid = {};
  let max = 0;

  for (const r of raw.each) {
    const key = `${r.day}__${r.hour}`;
    grid[key] = r.impressions;
    if (r.impressions > max) max = r.impressions;
  }

  let html = `<table class="heat"><thead><tr><th class="rowHead">Day</th>`;
  for (let h=0; h<24; h++) html += `<th>${h}</th>`;
  html += `</tr></thead><tbody>`;

  for (const day of daysOrder) {
    html += `<tr><td class="rowHead">${day}</td>`;
    for (let h=0; h<24; h++) {
      const v = grid[`${day}__${h}`] ?? 0;
      const intensity = max ? (v / max) : 0;
      const alpha = 0.10 + 0.75 * intensity;
      const bg = `rgba(0, 123, 255, ${alpha.toFixed(3)})`;
      const fg = intensity > 0.55 ? "#fff" : "#111";
      html += `<td style="background:${bg}; color:${fg};">${v ? fmtInt(v) : ""}</td>`;
    }
    html += `</tr>`;
  }

  html += `</tbody></table>`;
  document.getElementById("heatmapBox").innerHTML = html;
}

function fillTable(tableId, headers, rows) {
  const table = document.getElementById(tableId);
  const thead = `<thead><tr>${headers.map(h=>`<th>${escapeHtml(h)}</th>`).join("")}</tr></thead>`;
  const tbody = `<tbody>${rows.map(r=>`<tr>${r.map(c=>`<td>${escapeHtml(c)}</td>`).join("")}</tr>`).join("")}</tbody>`;
  table.innerHTML = thead + tbody;
}

function renderInsightsAndLimits(sel, k) {
  const topNet = raw.networks.slice().sort((a,b)=>b.clicks-a.clicks)[0];
  const topDev = raw.devices.slice().sort((a,b)=>b.clicks-a.clicks)[0];
  const topKw = raw.word.slice().sort((a,b)=>b.clicks-a.clicks)[0];
  const bestOpt = raw.opt.slice().sort((a,b)=>b.score-a.score)[0];
  const worstOpt = raw.opt.slice().sort((a,b)=>a.score-b.score)[0];

  const insights = [];
  insights.push(`Selected: Campaign <b>${sel.campaign==="__all__"?"All":escapeHtml(sel.campaign)}</b>, Network <b>${sel.network==="__all__"?"All":escapeHtml(sel.network)}</b>, Device <b>${sel.device==="__all__"?"All":escapeHtml(sel.device)}</b>.`);
  insights.push(`Performance: <b>${fmtInt(k.clicks)}</b> clicks, <b>${fmtInt(k.impressions)}</b> impressions, CTR <b>${fmtPct(k.ctr)}</b>, Avg CPC <b>${fmtMoney(k.avgCpc)}</b>.`);

  if (topNet) insights.push(`Top network by clicks: <b>${escapeHtml(topNet.network)}</b> (${fmtInt(topNet.clicks)} clicks).`);
  if (topDev) insights.push(`Top device by clicks: <b>${escapeHtml(topDev.device)}</b> (${fmtInt(topDev.clicks)} clicks).`);
  if (topKw) insights.push(`Top keyword by clicks: <b>${escapeHtml(topKw.keyword)}</b> (${fmtInt(topKw.clicks)} clicks).`);
  if (bestOpt && worstOpt) insights.push(`Optimization range: best <b>${escapeHtml(bestOpt.campaign)}</b> (${(bestOpt.score*100).toFixed(1)}%), lowest <b>${escapeHtml(worstOpt.campaign)}</b> (${(worstOpt.score*100).toFixed(1)}%).`);

  insights.push(`<br/><b>Recommended actions:</b>`);
  insights.push(`• Prioritize <b>${escapeHtml(topDev?.device || "Mobile")}</b> experience (landing speed, mobile-first layout).`);
  insights.push(`• Review search terms around the top keyword to add <b>negative keywords</b> and improve relevance/CTR.`);
  insights.push(`• Improve low optimization campaigns by applying recommendations that reduce wasted spend (avoid CPC-increasing changes unless justified).`);

  document.getElementById("insightsBox").innerHTML = insights.join("<br/>");

  const limitations = [];
  limitations.push(`Conversions are 0 in keyword/search tables, so insights focus on traffic metrics (clicks, impressions, CTR, CPC) rather than ROI.`);
  limitations.push(`Campaign filter affects KPIs using Biggest_changes sheet. For charts to fully update by Campaign/Network/Device, Excel needs campaign-level breakdown tables (device-by-campaign, network-by-campaign, keyword-by-campaign, time-series-by-campaign).`);
  document.getElementById("limitationsBox").innerHTML = limitations.map(x=>`• ${x}`).join("<br/>");
}

/* ---------- Boot ---------- */

document.getElementById("fileInput").addEventListener("change", async (e) => {
  const file = e.target.files?.[0];
  if (!file) return;
  try { await loadExcelFromFile(file); }
  catch (err) { console.error(err); alert("Failed to load Excel file."); }
});

loadExcelFromUrl("data.xlsx").catch(() => {
  document.getElementById("limitationsBox").innerHTML =
    `• Could not auto-load <b>data.xlsx</b>. Use “Load Excel” on top-right to upload the Excel file.`;
});
