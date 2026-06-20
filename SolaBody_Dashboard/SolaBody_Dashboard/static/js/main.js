/* ─────────────────────────────────────────────────────────────
   Sola Body — Revenue Command Center
   Frontend JavaScript
───────────────────────────────────────────────────────────── */

"use strict";

// ── State ─────────────────────────────────────────────────────
const state = {
  sessionId:    null,
  selectedDate: null,
  trendDays:    30,
  charts:       {},
};

// ── Helpers ───────────────────────────────────────────────────
const peso = (n) =>
  "₱" + Number(n).toLocaleString("en-PH", { minimumFractionDigits: 0, maximumFractionDigits: 0 });

const pct = (n, inverse = false) => {
  if (n === null || n === undefined) return null;
  const v   = parseFloat(n);
  const dir = inverse ? -v : v;
  const cls = dir > 0 ? "up" : dir < 0 ? "down" : "flat";
  const sym = dir > 0 ? "▲" : dir < 0 ? "▼" : "—";
  return { text: `${sym} ${Math.abs(v).toFixed(1)}%`, cls };
};

const fmt = (n, decimals = 2) => Number(n).toFixed(decimals);

function setDelta(elId, delta, inverse = false) {
  const el = document.getElementById(elId);
  if (!el) return;
  const d = pct(delta, inverse);
  if (!d) { el.textContent = ""; return; }
  el.textContent = d.text;
  el.className   = "kpi-delta " + d.cls;
}

function destroyChart(key) {
  if (state.charts[key]) {
    state.charts[key].destroy();
    delete state.charts[key];
  }
}

// ── Upload Logic ──────────────────────────────────────────────
const uploadScreen  = document.getElementById("upload-screen");
const dashboard     = document.getElementById("dashboard");
const dropZone      = document.getElementById("drop-zone");
const fileInput     = document.getElementById("file-input");
const uploadError   = document.getElementById("upload-error");
const uploadLoading = document.getElementById("upload-loading");

dropZone.addEventListener("click", () => fileInput.click());

dropZone.addEventListener("dragover", (e) => {
  e.preventDefault();
  dropZone.classList.add("active");
});
dropZone.addEventListener("dragleave", () => dropZone.classList.remove("active"));
dropZone.addEventListener("drop", (e) => {
  e.preventDefault();
  dropZone.classList.remove("active");
  const file = e.dataTransfer.files[0];
  if (file) processFile(file);
});

fileInput.addEventListener("change", () => {
  if (fileInput.files[0]) processFile(fileInput.files[0]);
});

async function processFile(file) {
  uploadError.style.display   = "none";
  uploadLoading.style.display = "flex";
  document.getElementById("drop-label").style.opacity = "0.4";

  const formData = new FormData();
  formData.append("file", file);

  try {
    const res  = await fetch("/api/upload", { method: "POST", body: formData });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Upload failed");

    state.sessionId = data.session_id;

    // Populate date selector
    const sel = document.getElementById("date-select");
    sel.innerHTML = "";
    [...data.dates].reverse().forEach((d) => {
      const opt = document.createElement("option");
      opt.value       = d;
      opt.textContent = formatDateLabel(d);
      sel.appendChild(opt);
    });

    // Update header
    document.getElementById("header-rows").textContent  = `${data.rows} days`;
    document.getElementById("header-range").textContent = `${formatShort(data.start)} — ${formatShort(data.end)}`;
    document.getElementById("header-meta").style.display = "flex";

    // Switch screens
    uploadScreen.style.display = "none";
    dashboard.style.display    = "flex";

    // Load first date
    state.selectedDate = [...data.dates].at(-1);
    sel.value          = state.selectedDate;
    await loadDay(state.selectedDate);

  } catch (err) {
    uploadError.textContent  = "Error: " + err.message;
    uploadError.style.display = "block";
  } finally {
    uploadLoading.style.display = "none";
    document.getElementById("drop-label").style.opacity = "1";
  }
}

function formatDateLabel(dateStr) {
  const d = new Date(dateStr + "T00:00:00");
  return d.toLocaleDateString("en-PH", { weekday: "short", month: "short", day: "numeric", year: "numeric" });
}
function formatShort(dateStr) {
  const d = new Date(dateStr + "T00:00:00");
  return d.toLocaleDateString("en-PH", { month: "short", day: "numeric", year: "numeric" });
}

// ── Date / Trend selector ─────────────────────────────────────
document.getElementById("date-select").addEventListener("change", async (e) => {
  state.selectedDate = e.target.value;
  await loadDay(state.selectedDate);
});

document.getElementById("trend-days").addEventListener("change", async (e) => {
  state.trendDays = parseInt(e.target.value);
  await loadTrend(state.selectedDate);
});

// ── Load daily data ───────────────────────────────────────────
async function loadDay(dateStr) {
  const loading = document.getElementById("day-loading");
  loading.style.display = "flex";

  try {
    const [dayRes, trendRes] = await Promise.all([
      fetch(`/api/daily?session_id=${state.sessionId}&date=${dateStr}`),
      fetch(`/api/trend?session_id=${state.sessionId}&date=${dateStr}&days=${state.trendDays}`),
    ]);
    const day   = await dayRes.json();
    const trend = await trendRes.json();

    if (!dayRes.ok)   throw new Error(day.error);
    if (!trendRes.ok) throw new Error(trend.error);

    renderDay(day);
    renderTrend(trend);

  } catch (err) {
    console.error(err);
  } finally {
    loading.style.display = "none";
  }
}

async function loadTrend(dateStr) {
  const res   = await fetch(`/api/trend?session_id=${state.sessionId}&date=${dateStr}&days=${state.trendDays}`);
  const trend = await res.json();
  if (res.ok) renderTrend(trend);
}

// ── Render KPIs + platform/ads charts ────────────────────────
function renderDay(data) {
  // Day header
  const d = new Date(data.date + "T00:00:00");
  document.getElementById("day-title").textContent =
    d.toLocaleDateString("en-PH", { weekday: "long", month: "long", day: "numeric", year: "numeric" });

  const campEl = document.getElementById("campaign-tag");
  if (data.remarks) {
    campEl.textContent    = "📌 " + data.remarks;
    campEl.style.display  = "inline-block";
  } else {
    campEl.style.display  = "none";
  }

  // KPIs
  document.getElementById("kpi-revenue").textContent    = peso(data.kpis.total_revenue);
  document.getElementById("kpi-exvat").textContent      = peso(data.kpis.total_revenue_ex_vat);
  document.getElementById("kpi-ads").textContent        = peso(data.kpis.total_ads_spent);
  document.getElementById("kpi-roas").textContent       = fmt(data.kpis.roas) + "x";
  document.getElementById("kpi-mtd").textContent        = peso(data.kpis.mtd_revenue);
  document.getElementById("kpi-mtd-roas").textContent   = fmt(data.kpis.mtd_roas) + "x";

  setDelta("kpi-revenue-delta", data.deltas.rev_vs_yesterday);
  setDelta("kpi-ads-delta",     data.deltas.ads_vs_yesterday, true); // inverse: lower spend = good

  // Platform bar chart
  const platLabels  = ["Shopify", "Shopee", "Lazada", "TikTok Shop"];
  const platColors  = ["#5C832F", "#EE4D2D", "#0F146D", "#555555"];
  const platValues  = [
    data.platforms.shopify, data.platforms.shopee,
    data.platforms.lazada,  data.platforms.tiktok,
  ];

  destroyChart("platform");
  state.charts.platform = new Chart(document.getElementById("chart-platform"), {
    type: "bar",
    data: {
      labels: platLabels,
      datasets: [{ data: platValues, backgroundColor: platColors, borderRadius: 4, borderSkipped: false }],
    },
    options: {
      indexAxis: "y",
      responsive: true,
      plugins: {
        legend: { display: false },
        tooltip: { callbacks: { label: (ctx) => " " + peso(ctx.raw) } },
      },
      scales: {
        x: { ticks: { callback: (v) => "₱" + (v / 1000).toFixed(0) + "K" }, grid: { color: "#F1F5F9" } },
        y: { grid: { display: false } },
      },
    },
  });

  // Ad spend bar chart
  const adLabels = Object.keys(data.ads);
  const adValues = Object.values(data.ads).map((v) => parseFloat(v));
  const combined = adLabels.map((l, i) => ({ l, v: adValues[i] }))
    .filter((x) => x.v > 0)
    .sort((a, b) => b.v - a.v);

  destroyChart("ads");
  state.charts.ads = new Chart(document.getElementById("chart-ads"), {
    type: "bar",
    data: {
      labels: combined.map((x) => x.l),
      datasets: [{ data: combined.map((x) => x.v), backgroundColor: "#457B9D", borderRadius: 4, borderSkipped: false }],
    },
    options: {
      indexAxis: "y",
      responsive: true,
      plugins: {
        legend: { display: false },
        tooltip: { callbacks: { label: (ctx) => " " + peso(ctx.raw) } },
      },
      scales: {
        x: { ticks: { callback: (v) => "₱" + (v / 1000).toFixed(0) + "K" }, grid: { color: "#F1F5F9" } },
        y: { grid: { display: false } },
      },
    },
  });

  // Platform share doughnut
  destroyChart("share");
  state.charts.share = new Chart(document.getElementById("chart-share"), {
    type: "doughnut",
    data: {
      labels: platLabels,
      datasets: [{ data: platValues, backgroundColor: platColors, borderWidth: 2, borderColor: "#fff" }],
    },
    options: {
      responsive: true,
      cutout: "68%",
      plugins: {
        legend: { position: "bottom", labels: { font: { size: 11 }, boxWidth: 12, padding: 8 } },
        tooltip: { callbacks: { label: (ctx) => " " + peso(ctx.raw) } },
      },
    },
  });
}

// ── Render trend charts ───────────────────────────────────────
function renderTrend(data) {
  const labels = data.dates.map((d) => {
    const dt = new Date(d + "T00:00:00");
    return dt.toLocaleDateString("en-PH", { month: "short", day: "numeric" });
  });

  // Campaign annotation points
  const campPoints = data.dates.map((d, i) => {
    const c = data.campaigns.find((c) => c.date === d);
    return c ? data.revenue[i] : null;
  });

  // Revenue + Ads trend
  destroyChart("trend");
  state.charts.trend = new Chart(document.getElementById("chart-trend"), {
    type: "line",
    data: {
      labels,
      datasets: [
        {
          label: "Total Revenue",
          data: data.revenue,
          borderColor:     "#1D3557",
          backgroundColor: "rgba(29,53,87,0.07)",
          borderWidth: 2,
          fill: true,
          tension: 0.3,
          pointRadius: 0,
          pointHoverRadius: 4,
        },
        {
          label: "Ad Spend",
          data: data.ads,
          borderColor:     "#E63946",
          borderWidth: 1.5,
          borderDash: [5, 3],
          fill: false,
          tension: 0.3,
          pointRadius: 0,
          pointHoverRadius: 4,
        },
        {
          label: "Campaign",
          data: campPoints,
          type: "scatter",
          pointStyle: "star",
          pointRadius: 8,
          pointHoverRadius: 10,
          backgroundColor: "#F59E0B",
          showLine: false,
        },
      ],
    },
    options: {
      responsive: true,
      interaction: { mode: "index", intersect: false },
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            label: (ctx) => {
              if (ctx.dataset.label === "Campaign") return "📌 Campaign day";
              return ctx.dataset.label + ": " + peso(ctx.raw);
            },
          },
        },
      },
      scales: {
        x: { grid: { color: "#F1F5F9" }, ticks: { font: { size: 11 }, maxRotation: 0 } },
        y: { grid: { color: "#F1F5F9" }, ticks: { callback: (v) => "₱" + (v / 1000).toFixed(0) + "K" } },
      },
    },
  });

  // Platform stacked area
  const platConfig = [
    { key: "shopify", label: "Shopify",   color: "#5C832F" },
    { key: "shopee",  label: "Shopee",    color: "#EE4D2D" },
    { key: "lazada",  label: "Lazada",    color: "#0F146D" },
    { key: "tiktok",  label: "TikTok",    color: "#888888" },
  ];

  destroyChart("platformTrend");
  state.charts.platformTrend = new Chart(document.getElementById("chart-platform-trend"), {
    type: "line",
    data: {
      labels,
      datasets: platConfig.map((p) => ({
        label:           p.label,
        data:            data[p.key],
        borderColor:     p.color,
        backgroundColor: p.color + "22",
        borderWidth: 1.5,
        fill: "stack",
        tension: 0.3,
        pointRadius: 0,
        pointHoverRadius: 4,
      })),
    },
    options: {
      responsive: true,
      interaction: { mode: "index", intersect: false },
      plugins: {
        legend: { position: "bottom", labels: { font: { size: 11 }, boxWidth: 12, padding: 10 } },
        tooltip: { callbacks: { label: (ctx) => ctx.dataset.label + ": " + peso(ctx.raw) } },
      },
      scales: {
        x: { grid: { color: "#F1F5F9" }, ticks: { font: { size: 11 }, maxRotation: 0 } },
        y: { grid: { color: "#F1F5F9" }, stacked: true, ticks: { callback: (v) => "₱" + (v / 1000).toFixed(0) + "K" } },
      },
    },
  });
}

// ── Download buttons ──────────────────────────────────────────
function downloadReport(type) {
  if (!state.sessionId) return;
  const base  = `/download/${type}?session_id=${state.sessionId}`;
  const extra = type === "daily" && state.selectedDate ? `&date=${state.selectedDate}` : "";
  window.location.href = base + extra;
}

document.getElementById("dl-daily").addEventListener("click",   () => downloadReport("daily"));
document.getElementById("dl-weekly").addEventListener("click",  () => downloadReport("weekly"));
document.getElementById("dl-monthly").addEventListener("click", () => downloadReport("monthly"));
document.getElementById("dl-audit").addEventListener("click",   () => downloadReport("audit"));

// ── Reset ─────────────────────────────────────────────────────
document.getElementById("reset-btn").addEventListener("click", () => {
  state.sessionId    = null;
  state.selectedDate = null;
  Object.keys(state.charts).forEach(destroyChart);
  fileInput.value              = "";
  uploadError.style.display    = "none";
  document.getElementById("header-meta").style.display = "none";
  dashboard.style.display      = "none";
  uploadScreen.style.display   = "flex";
});
