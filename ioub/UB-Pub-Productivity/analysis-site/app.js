const DATA_PATHS = {
  master: "../data/master/lecturers.csv",
  yearly: "../data/history/author_metrics_yearly.csv",
  snapshots: [
    "../data/snapshots/author_metrics_2025-04.csv",
    "../data/snapshots/author_metrics_2025-05.csv",
    "../data/snapshots/author_metrics_2025-06.csv",
    "../data/snapshots/author_metrics_2025-07.csv",
    "../data/snapshots/author_metrics_2025-08.csv",
    "../data/snapshots/author_metrics_2025-09.csv",
    "../data/snapshots/author_metrics_2025-10.csv",
    "../data/snapshots/author_metrics_2025-11.csv",
    "../data/snapshots/author_metrics_2025-12.csv",
    "../data/snapshots/author_metrics_2026-01.csv",
    "../data/snapshots/author_metrics_2026-02.csv",
    "../data/snapshots/author_metrics_2026-03.csv",
  ],
};

const state = {
  datasets: { master: [], yearly: [], monthly: [] },
  filters: {
    granularity: "monthly",
    scope: "ALL",
    period: "ALL",
    comparisonMode: "selected",
    rankingMetric: "publications",
    rankingSearch: "",
    detailSearch: "",
  },
  tables: {
    ranking: { sortKey: "primaryValue", sortDir: "desc", page: 1, pageSize: 12 },
    detail: { sortKey: "period", sortDir: "desc", page: 1, pageSize: 14 },
  },
  charts: {},
};

const dom = {
  granularitySelect: document.getElementById("granularitySelect"),
  scopeSelect: document.getElementById("scopeSelect"),
  periodSelect: document.getElementById("periodSelect"),
  comparisonModeSelect: document.getElementById("comparisonModeSelect"),
  rankingMetricSelect: document.getElementById("rankingMetricSelect"),
  rankingSearch: document.getElementById("rankingSearch"),
  detailSearch: document.getElementById("detailSearch"),
  kpiGrid: document.getElementById("kpiGrid"),
  rankingTable: document.getElementById("rankingTable"),
  detailTable: document.getElementById("detailTable"),
  topMovers: document.getElementById("topMovers"),
  coverageNotes: document.getElementById("coverageNotes"),
  statusText: document.getElementById("statusText"),
  downloadButton: document.getElementById("downloadButton"),
};

const indexes = {
  masterByLecturerId: new Map(),
  masterByScopusId: new Map(),
  masterByNip: new Map(),
};

function parseCsv(url) {
  return new Promise((resolve, reject) => {
    Papa.parse(url, {
      download: true,
      header: true,
      skipEmptyLines: true,
      complete: (results) => resolve(results.data),
      error: reject,
    });
  });
}

function normalize(value) {
  return (value ?? "").toString().trim();
}

function normalizeKey(value) {
  return normalize(value).replace(/[^0-9A-Za-z]+/g, "").toUpperCase();
}

function slugify(value) {
  return normalize(value).replace(/[^0-9A-Za-z]+/g, "_").replace(/^_+|_+$/g, "").toUpperCase() || "UNKNOWN";
}

function toNumber(value) {
  const text = normalize(value);
  if (!text) return null;
  const parsed = Number(text);
  return Number.isFinite(parsed) ? parsed : null;
}

function formatNumber(value) {
  if (value === null || value === undefined || Number.isNaN(value)) return "-";
  return new Intl.NumberFormat("en-US").format(value);
}

function createOption(value, label) {
  const option = document.createElement("option");
  option.value = value;
  option.textContent = label;
  return option;
}

function loadData() {
  dom.statusText.textContent = "Loading CSV data...";
  return Promise.all([
    parseCsv(DATA_PATHS.master),
    parseCsv(DATA_PATHS.yearly),
    Promise.all(DATA_PATHS.snapshots.map(parseCsv)),
  ]).then(([master, yearly, snapshotLists]) => {
    state.datasets.master = master;
    state.datasets.yearly = yearly;
    state.datasets.monthly = snapshotLists.flat();
    buildIndexes();
    initializeFilters();
    render();
  }).catch((error) => {
    dom.statusText.textContent = "Failed to load data. Serve this folder through a web server or GitHub Pages.";
    console.error(error);
  });
}

function buildIndexes() {
  indexes.masterByLecturerId.clear();
  indexes.masterByScopusId.clear();
  indexes.masterByNip.clear();

  state.datasets.master.forEach((row) => {
    const lecturerId = normalize(row.lecturer_id);
    const scopusId = normalize(row.scopus_author_id || row["SCOPUS ID"]);
    const nip = normalize(row["NIP/NIK"] || row["NIK/NIP"]);
    if (lecturerId) indexes.masterByLecturerId.set(lecturerId, row);
    if (scopusId) indexes.masterByScopusId.set(normalizeKey(scopusId), row);
    if (nip) indexes.masterByNip.set(normalizeKey(nip), row);
  });
}

function resolveMaster(row) {
  const lecturerId = normalize(row.lecturer_id);
  const scopusId = normalize(row.scopus_author_id);
  const nip = normalize(row.nip_nik || row["NIP/NIK"] || row["NIK/NIP"]);
  return indexes.masterByLecturerId.get(lecturerId)
    || indexes.masterByScopusId.get(normalizeKey(scopusId))
    || indexes.masterByNip.get(normalizeKey(nip))
    || null;
}

function resolveRecordId(row, master) {
  const lecturerId = normalize(row.lecturer_id);
  if (lecturerId) return lecturerId;
  if (master?.lecturer_id) return normalize(master.lecturer_id);
  if (normalize(row.scopus_author_id)) return `SCOPUS_${slugify(row.scopus_author_id)}`;
  return `${slugify(row.faculty_code)}_${slugify(row.lecturer_name)}`;
}

function enrichYearlyRow(row) {
  const master = resolveMaster(row);
  return {
    id: `${normalize(row.year)}-${resolveRecordId(row, master)}`,
    lecturerId: resolveRecordId(row, master),
    lecturerName: normalize(row.lecturer_name) || normalize(master?.lecturer_name) || normalize(master?.Nama),
    facultyCode: normalize(row.faculty_code) || normalize(master?.faculty_code) || normalize(master?.Homebase),
    facultyName: normalize(row.faculty_name) || normalize(master?.faculty_name) || normalize(master?.Homebase),
    jabatan: normalize(master?.JABATAN),
    scopusAuthorId: normalize(row.scopus_author_id) || normalize(master?.scopus_author_id) || normalize(master?.["SCOPUS ID"]),
    activeStatus: normalize(row.active_status) || normalize(master?.active_status) || normalize(master?.Keterangan),
    period: normalize(row.year),
    periodLabel: normalize(row.year),
    periodSort: Number(normalize(row.year)),
    publications: toNumber(row.publication_count),
    publicationTotal: toNumber(row.publication_count_total_cum),
    citations: toNumber(row.citation_count),
    citationTotal: null,
    hIndex: toNumber(row.h_index),
    apiStatus: normalize(row.api_status),
    notes: normalize(row.notes),
    type: "yearly",
  };
}

function enrichMonthlyRow(row) {
  const master = resolveMaster(row);
  return {
    id: `${normalize(row.snapshot_month)}-${resolveRecordId(row, master)}`,
    lecturerId: resolveRecordId(row, master),
    lecturerName: normalize(row.lecturer_name) || normalize(master?.lecturer_name) || normalize(master?.Nama),
    facultyCode: normalize(row.faculty_code) || normalize(master?.faculty_code) || normalize(master?.Homebase),
    facultyName: normalize(row.faculty_name) || normalize(master?.faculty_name) || normalize(master?.Homebase),
    jabatan: normalize(master?.JABATAN),
    scopusAuthorId: normalize(row.scopus_author_id) || normalize(master?.scopus_author_id) || normalize(master?.["SCOPUS ID"]),
    activeStatus: normalize(master?.active_status) || normalize(master?.Keterangan),
    period: normalize(row.snapshot_month),
    periodLabel: normalize(row.snapshot_month),
    periodSort: Number(normalize(row.snapshot_month).replace("-", "")),
    publications: toNumber(row.publication_count_month),
    publicationTotal: toNumber(row.publication_count_total),
    citations: toNumber(row.citation_count_month),
    citationTotal: toNumber(row.citation_count_total),
    hIndex: toNumber(row.h_index),
    apiStatus: normalize(row.api_status),
    notes: normalize(row.notes),
    type: "monthly",
  };
}

function allRecords(granularity) {
  return granularity === "yearly"
    ? state.datasets.yearly.map(enrichYearlyRow)
    : state.datasets.monthly.map(enrichMonthlyRow);
}

function initializeFilters() {
  dom.granularitySelect.value = state.filters.granularity;
  dom.comparisonModeSelect.value = state.filters.comparisonMode;
  dom.rankingMetricSelect.value = state.filters.rankingMetric;

  dom.granularitySelect.addEventListener("change", () => {
    state.filters.granularity = dom.granularitySelect.value;
    state.filters.period = "ALL";
    state.tables.ranking.page = 1;
    state.tables.detail.page = 1;
    rebuildScopeOptions();
    rebuildPeriodOptions();
    render();
  });

  dom.scopeSelect.addEventListener("change", () => {
    state.filters.scope = dom.scopeSelect.value;
    state.tables.ranking.page = 1;
    state.tables.detail.page = 1;
    render();
  });

  dom.periodSelect.addEventListener("change", () => {
    state.filters.period = dom.periodSelect.value;
    state.tables.ranking.page = 1;
    state.tables.detail.page = 1;
    render();
  });

  dom.comparisonModeSelect.addEventListener("change", () => {
    state.filters.comparisonMode = dom.comparisonModeSelect.value;
    render();
  });

  dom.rankingMetricSelect.addEventListener("change", () => {
    state.filters.rankingMetric = dom.rankingMetricSelect.value;
    state.tables.ranking.page = 1;
    render();
  });

  dom.rankingSearch.addEventListener("input", () => {
    state.filters.rankingSearch = dom.rankingSearch.value.toLowerCase();
    state.tables.ranking.page = 1;
    render();
  });

  dom.detailSearch.addEventListener("input", () => {
    state.filters.detailSearch = dom.detailSearch.value.toLowerCase();
    state.tables.detail.page = 1;
    render();
  });

  dom.downloadButton.addEventListener("click", downloadFilteredDetailCsv);

  rebuildScopeOptions();
  rebuildPeriodOptions();
}

function rebuildScopeOptions() {
  const records = allRecords(state.filters.granularity);
  const faculties = Array.from(new Set(records.map((row) => row.facultyCode).filter(Boolean))).sort();
  dom.scopeSelect.innerHTML = "";
  dom.scopeSelect.appendChild(createOption("ALL", "All UB"));
  faculties.forEach((faculty) => dom.scopeSelect.appendChild(createOption(faculty, faculty)));
  if (![...dom.scopeSelect.options].some((option) => option.value === state.filters.scope)) {
    state.filters.scope = "ALL";
  }
  dom.scopeSelect.value = state.filters.scope;
}

function rebuildPeriodOptions() {
  const records = allRecords(state.filters.granularity);
  const periods = Array.from(new Set(records.map((row) => row.period))).sort();
  dom.periodSelect.innerHTML = "";
  dom.periodSelect.appendChild(createOption("ALL", "All Periods"));
  periods.forEach((period) => dom.periodSelect.appendChild(createOption(period, period)));
  if (![...dom.periodSelect.options].some((option) => option.value === state.filters.period)) {
    state.filters.period = "ALL";
  }
  dom.periodSelect.value = state.filters.period;
}

function filteredRecords() {
  return allRecords(state.filters.granularity).filter((row) => {
    if (state.filters.scope !== "ALL" && row.facultyCode !== state.filters.scope) return false;
    if (state.filters.period !== "ALL" && row.period !== state.filters.period) return false;
    return true;
  });
}

function visibleRecordsForTrend() {
  return allRecords(state.filters.granularity).filter((row) => (
    state.filters.scope === "ALL" || row.facultyCode === state.filters.scope
  ));
}

function sumMetric(values) {
  return values.reduce((total, value) => total + (value ?? 0), 0);
}

function meanMetric(values) {
  const present = values.filter((value) => value !== null && value !== undefined);
  if (!present.length) return null;
  return present.reduce((total, value) => total + value, 0) / present.length;
}

function buildKpis(records) {
  const lecturers = new Set(records.map((row) => row.lecturerId));
  const publicationActive = new Set(records.filter((row) => (row.publications ?? 0) > 0).map((row) => row.lecturerId));
  const publicationTotal = sumMetric(records.map((row) => row.publications));
  const citationTotal = sumMetric(records.map((row) => row.citations));
  const avgHIndex = meanMetric(records.map((row) => row.hIndex));
  return [
    { label: "Visible Lecturers", value: formatNumber(lecturers.size), foot: state.filters.scope === "ALL" ? "Across UB" : `In ${state.filters.scope}` },
    { label: "Publications", value: formatNumber(publicationTotal), foot: state.filters.granularity === "monthly" ? "Month-specific counts" : "Year-specific counts" },
    { label: "Citations", value: formatNumber(citationTotal), foot: "Visible rows only" },
    { label: "Avg H-Index", value: avgHIndex === null ? "-" : avgHIndex.toFixed(2), foot: "Across visible records" },
    { label: "Lecturers With Publications", value: formatNumber(publicationActive.size), foot: `${lecturers.size ? ((publicationActive.size / lecturers.size) * 100).toFixed(1) : "0.0"}% of visible lecturers` },
  ];
}

function renderKpis(records) {
  const kpis = buildKpis(records);
  dom.kpiGrid.innerHTML = "";
  kpis.forEach((kpi) => {
    const card = document.createElement("div");
    card.className = "kpi";
    card.innerHTML = `
      <div class="kpi-label">${kpi.label}</div>
      <div class="kpi-value">${kpi.value}</div>
      <div class="kpi-foot">${kpi.foot}</div>
    `;
    dom.kpiGrid.appendChild(card);
  });
}

function aggregateByPeriod(records) {
  const grouped = new Map();
  records.forEach((row) => {
    if (!grouped.has(row.period)) grouped.set(row.period, { publications: 0, citations: 0 });
    const bucket = grouped.get(row.period);
    bucket.publications += row.publications ?? 0;
    bucket.citations += row.citations ?? 0;
  });
  return [...grouped.entries()]
    .sort(([a], [b]) => a.localeCompare(b))
    .map(([period, metrics]) => ({ period, ...metrics }));
}

function aggregateActiveByFacultyAndPeriod(records) {
  const periods = Array.from(new Set(records.map((row) => row.period))).sort();
  const faculties = Array.from(new Set(records.map((row) => row.facultyCode).filter(Boolean))).sort();
  const datasets = faculties.map((faculty) => ({
    label: faculty,
    data: periods.map((period) => (
      new Set(
        records
          .filter((row) => row.period === period && row.facultyCode === faculty && (row.publications ?? 0) > 0)
          .map((row) => row.lecturerId)
      ).size
    )),
  }));
  return { periods, datasets };
}

function aggregateFacultyImpact(records, selectedOnly) {
  const base = selectedOnly && state.filters.period !== "ALL"
    ? records.filter((row) => row.period === state.filters.period)
    : records;

  const grouped = new Map();
  base.forEach((row) => {
    if (!row.facultyCode) return;
    if (!grouped.has(row.facultyCode)) grouped.set(row.facultyCode, { publications: 0, citations: 0 });
    const bucket = grouped.get(row.facultyCode);
    bucket.publications += row.publications ?? 0;
    bucket.citations += row.citations ?? 0;
  });

  return [...grouped.entries()]
    .sort((a, b) => (b[1].publications + b[1].citations) - (a[1].publications + a[1].citations))
    .map(([faculty, metrics]) => ({ faculty, ...metrics }));
}

function aggregateJabatan(records) {
  const base = state.filters.period === "ALL" ? records : records.filter((row) => row.period === state.filters.period);
  const grouped = new Map();
  base.forEach((row) => {
    const jabatan = row.jabatan || "Unknown";
    if (!grouped.has(jabatan)) grouped.set(jabatan, { publications: 0, citations: 0 });
    const bucket = grouped.get(jabatan);
    bucket.publications += row.publications ?? 0;
    bucket.citations += row.citations ?? 0;
  });

  return [...grouped.entries()]
    .sort((a, b) => (b[1].publications + b[1].citations) - (a[1].publications + a[1].citations))
    .slice(0, 10)
    .map(([jabatan, metrics]) => ({ jabatan, ...metrics }));
}

function groupRankingRows(records) {
  const map = new Map();
  records.forEach((row) => {
    if (!map.has(row.lecturerId)) {
      map.set(row.lecturerId, {
        lecturerId: row.lecturerId,
        lecturerName: row.lecturerName,
        facultyCode: row.facultyCode,
        jabatan: row.jabatan || "Unknown",
        scopusAuthorId: row.scopusAuthorId,
        publications: 0,
        citations: 0,
        hIndex: 0,
      });
    }
    const item = map.get(row.lecturerId);
    item.publications += row.publications ?? 0;
    item.citations += row.citations ?? 0;
    item.hIndex = Math.max(item.hIndex, row.hIndex ?? 0);
  });
  return [...map.values()];
}

function rankingRows(records) {
  const grouped = groupRankingRows(records);
  return grouped
    .map((row) => ({
      ...row,
      primaryValue:
        state.filters.rankingMetric === "citations" ? row.citations :
        state.filters.rankingMetric === "hIndex" ? row.hIndex :
        row.publications,
    }))
    .filter((row) => {
      const q = state.filters.rankingSearch;
      if (!q) return true;
      return [row.lecturerName, row.facultyCode, row.jabatan, row.scopusAuthorId]
        .some((value) => normalize(value).toLowerCase().includes(q));
    });
}

function topMovers(records) {
  if (state.filters.period === "ALL") return [];
  return records
    .filter((row) => row.period === state.filters.period)
    .map((row) => ({
      lecturerName: row.lecturerName,
      facultyCode: row.facultyCode,
      publications: row.publications ?? 0,
      citations: row.citations ?? 0,
    }))
    .sort((a, b) => (b.publications + b.citations) - (a.publications + a.citations))
    .slice(0, 6);
}

function coverageSummary(records) {
  const unmatched = records.filter((row) => !normalize(row.scopusAuthorId) && !normalize(row.lecturerId));
  const withJabatan = records.filter((row) => normalize(row.jabatan)).length;
  const withHIndex = records.filter((row) => row.hIndex !== null).length;
  return [
    `Visible rows: ${formatNumber(records.length)}`,
    `Rows with jabatan mapping: ${formatNumber(withJabatan)}`,
    `Rows with h-index value: ${formatNumber(withHIndex)}`,
    `Rows still unmatched to master data: ${formatNumber(unmatched.length)}`,
  ];
}

function destroyChart(name) {
  if (state.charts[name]) state.charts[name].destroy();
}

function createChart(name, canvasId, config) {
  destroyChart(name);
  state.charts[name] = new Chart(document.getElementById(canvasId), config);
}

function facultyPalette(index, alpha) {
  const palette = [
    `rgba(181, 74, 45, ${alpha})`,
    `rgba(32, 79, 109, ${alpha})`,
    `rgba(120, 95, 55, ${alpha})`,
    `rgba(85, 140, 94, ${alpha})`,
    `rgba(135, 81, 147, ${alpha})`,
    `rgba(208, 124, 60, ${alpha})`,
    `rgba(88, 114, 188, ${alpha})`,
    `rgba(74, 157, 157, ${alpha})`,
    `rgba(142, 88, 88, ${alpha})`,
    `rgba(99, 99, 99, ${alpha})`,
  ];
  return palette[index % palette.length];
}

function chartOptions(extra = {}) {
  return {
    responsive: true,
    maintainAspectRatio: false,
    interaction: { intersect: false, mode: "index" },
    scales: {
      x: { stacked: !!extra.stacked, grid: { display: false } },
      y: {
        stacked: !!extra.stacked,
        beginAtZero: true,
        ticks: { callback: (value) => formatNumber(value) },
      },
    },
    plugins: {
      legend: { labels: { usePointStyle: true, boxWidth: 10 } },
      tooltip: {
        callbacks: {
          label: (context) => `${context.dataset.label}: ${formatNumber(context.parsed.y)}`,
        },
      },
    },
  };
}

function renderCharts(records) {
  const trendSeries = aggregateByPeriod(visibleRecordsForTrend());
  createChart("trendChart", "trendChart", {
    type: "line",
    data: {
      labels: trendSeries.map((row) => row.period),
      datasets: [
        {
          label: "Publications",
          data: trendSeries.map((row) => row.publications),
          borderColor: "#b54a2d",
          backgroundColor: "rgba(181, 74, 45, 0.16)",
          tension: 0.25,
          fill: true,
        },
        {
          label: "Citations",
          data: trendSeries.map((row) => row.citations),
          borderColor: "#204f6d",
          backgroundColor: "rgba(32, 79, 109, 0.12)",
          tension: 0.25,
          fill: true,
        },
      ],
    },
    options: chartOptions(),
  });

  const activeData = aggregateActiveByFacultyAndPeriod(visibleRecordsForTrend());
  createChart("activeFacultyChart", "activeFacultyChart", {
    type: "bar",
    data: {
      labels: activeData.periods,
      datasets: activeData.datasets.slice(0, 10).map((dataset, index) => ({
        label: dataset.label,
        data: dataset.data,
        backgroundColor: facultyPalette(index, 0.82),
        borderRadius: 6,
        stack: "active",
      })),
    },
    options: chartOptions({ stacked: true }),
  });

  const facultyImpact = aggregateFacultyImpact(visibleRecordsForTrend(), state.filters.comparisonMode === "selected");
  createChart("facultyImpactChart", "facultyImpactChart", {
    type: "bar",
    data: {
      labels: facultyImpact.map((row) => row.faculty),
      datasets: [
        { label: "Publications", data: facultyImpact.map((row) => row.publications), backgroundColor: "rgba(181, 74, 45, 0.82)", borderRadius: 6 },
        { label: "Citations", data: facultyImpact.map((row) => row.citations), backgroundColor: "rgba(32, 79, 109, 0.78)", borderRadius: 6 },
      ],
    },
    options: chartOptions(),
  });

  const jabatanData = aggregateJabatan(records);
  createChart("jabatanChart", "jabatanChart", {
    type: "bar",
    data: {
      labels: jabatanData.map((row) => row.jabatan),
      datasets: [
        { label: "Publications", data: jabatanData.map((row) => row.publications), backgroundColor: "rgba(241, 135, 92, 0.88)", borderRadius: 6 },
        { label: "Citations", data: jabatanData.map((row) => row.citations), backgroundColor: "rgba(54, 116, 156, 0.78)", borderRadius: 6 },
      ],
    },
    options: chartOptions(),
  });
}

function renderInsights(records) {
  const movers = topMovers(records);
  dom.topMovers.innerHTML = movers.length
    ? `<div class="insight-list">${movers.map((row) => `
        <div class="insight-item">
          <div>
            <strong>${row.lecturerName}</strong>
            <div class="muted">${row.facultyCode}</div>
          </div>
          <div>
            <div>Pub: <strong>${formatNumber(row.publications)}</strong></div>
            <div>Cite: <strong>${formatNumber(row.citations)}</strong></div>
          </div>
        </div>`).join("")}</div>`
    : `<p class="muted">Choose a single period to see current-period movers.</p>`;

  const notes = coverageSummary(records);
  dom.coverageNotes.innerHTML = `
    <div class="insight-list">
      ${notes.map((note) => `<div class="insight-item"><span>${note}</span></div>`).join("")}
    </div>
  `;
}

function sortRows(rows, sortKey, sortDir) {
  const direction = sortDir === "asc" ? 1 : -1;
  return [...rows].sort((a, b) => {
    const left = a[sortKey];
    const right = b[sortKey];
    if (typeof left === "number" && typeof right === "number") return (left - right) * direction;
    return normalize(left).localeCompare(normalize(right)) * direction;
  });
}

function paginate(rows, page, pageSize) {
  const start = (page - 1) * pageSize;
  return rows.slice(start, start + pageSize);
}

function toggleSort(tableState, key) {
  if (tableState.sortKey === key) {
    tableState.sortDir = tableState.sortDir === "asc" ? "desc" : "asc";
  } else {
    tableState.sortKey = key;
    tableState.sortDir = ["lecturerName", "facultyCode", "jabatan", "period"].includes(key) ? "asc" : "desc";
  }
  render();
}

function renderTable(targetEl, config) {
  const { rows, columns, tableState, onSort } = config;
  const sortedRows = sortRows(rows, tableState.sortKey, tableState.sortDir);
  const pageCount = Math.max(1, Math.ceil(sortedRows.length / tableState.pageSize));
  tableState.page = Math.min(tableState.page, pageCount);
  const pageRows = paginate(sortedRows, tableState.page, tableState.pageSize);

  const wrapper = document.createElement("div");
  wrapper.className = "table-wrapper";
  const table = document.createElement("table");
  const thead = document.createElement("thead");
  const headRow = document.createElement("tr");

  columns.forEach((column) => {
    const th = document.createElement("th");
    const isActive = tableState.sortKey === column.key;
    const arrow = isActive ? (tableState.sortDir === "asc" ? " ↑" : " ↓") : " ↕";
    th.textContent = `${column.label}${arrow}`;
    th.addEventListener("click", () => onSort(column.key));
    headRow.appendChild(th);
  });

  thead.appendChild(headRow);
  table.appendChild(thead);

  const tbody = document.createElement("tbody");
  pageRows.forEach((row) => {
    const tr = document.createElement("tr");
    columns.forEach((column) => {
      const td = document.createElement("td");
      const value = typeof column.render === "function" ? column.render(row) : row[column.key];
      td.innerHTML = value instanceof HTMLElement ? "" : value;
      if (value instanceof HTMLElement) td.appendChild(value);
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });

  table.appendChild(tbody);
  wrapper.appendChild(table);

  const footer = document.createElement("div");
  footer.className = "table-footer";
  footer.innerHTML = `<div class="muted">${formatNumber(sortedRows.length)} rows</div>`;

  const pagination = document.createElement("div");
  pagination.className = "pagination";
  const prev = document.createElement("button");
  prev.textContent = "Prev";
  prev.disabled = tableState.page <= 1;
  prev.addEventListener("click", () => { tableState.page -= 1; render(); });
  const next = document.createElement("button");
  next.textContent = "Next";
  next.disabled = tableState.page >= pageCount;
  next.addEventListener("click", () => { tableState.page += 1; render(); });
  const pageLabel = document.createElement("span");
  pageLabel.textContent = `Page ${tableState.page} / ${pageCount}`;
  pagination.append(prev, pageLabel, next);
  footer.appendChild(pagination);

  targetEl.innerHTML = "";
  targetEl.append(wrapper, footer);
}

function renderRankingTable(records) {
  const rows = rankingRows(records);
  renderTable(dom.rankingTable, {
    rows,
    columns: [
      { key: "lecturerName", label: "Lecturer" },
      { key: "facultyCode", label: "Faculty" },
      { key: "jabatan", label: "Jabatan" },
      { key: "publications", label: "Publications", render: (row) => formatNumber(row.publications) },
      { key: "citations", label: "Citations", render: (row) => formatNumber(row.citations) },
      { key: "hIndex", label: "H-Index", render: (row) => formatNumber(row.hIndex) },
      { key: "primaryValue", label: "Ranking Value", render: (row) => `<span class="chip">${formatNumber(row.primaryValue)}</span>` },
    ],
    tableState: state.tables.ranking,
    onSort: (key) => toggleSort(state.tables.ranking, key),
  });
}

function renderDetailTable(records) {
  const query = state.filters.detailSearch;
  const filtered = records.filter((row) => {
    if (!query) return true;
    return [row.lecturerName, row.facultyCode, row.jabatan, row.scopusAuthorId, row.period]
      .some((value) => normalize(value).toLowerCase().includes(query));
  });

  renderTable(dom.detailTable, {
    rows: filtered,
    columns: [
      { key: "period", label: state.filters.granularity === "monthly" ? "Month" : "Year" },
      { key: "lecturerName", label: "Lecturer" },
      { key: "facultyCode", label: "Faculty" },
      { key: "jabatan", label: "Jabatan" },
      { key: "publications", label: state.filters.granularity === "monthly" ? "Publication Count (Month)" : "Publication Count (Year)", render: (row) => formatNumber(row.publications) },
      { key: "citations", label: state.filters.granularity === "monthly" ? "Citation Count (Month)" : "Citation Count (Year)", render: (row) => formatNumber(row.citations) },
      { key: "publicationTotal", label: "Publication Total", render: (row) => formatNumber(row.publicationTotal) },
      { key: "citationTotal", label: "Citation Total", render: (row) => formatNumber(row.citationTotal) },
      { key: "hIndex", label: "H-Index", render: (row) => formatNumber(row.hIndex) },
    ],
    tableState: state.tables.detail,
    onSort: (key) => toggleSort(state.tables.detail, key),
  });
}

function downloadFilteredDetailCsv() {
  const rows = filteredRecords();
  const csv = Papa.unparse(rows.map((row) => ({
    period: row.period,
    lecturer_id: row.lecturerId,
    lecturer_name: row.lecturerName,
    faculty_code: row.facultyCode,
    jabatan: row.jabatan,
    scopus_author_id: row.scopusAuthorId,
    publication_count: row.publications ?? "",
    publication_total: row.publicationTotal ?? "",
    citation_count: row.citations ?? "",
    citation_total: row.citationTotal ?? "",
    h_index: row.hIndex ?? "",
    api_status: row.apiStatus,
    notes: row.notes,
  })));

  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `ub-publication-analysis-${state.filters.granularity}-${state.filters.scope}-${state.filters.period}.csv`;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

function render() {
  const records = filteredRecords();
  dom.statusText.textContent = `${formatNumber(records.length)} visible rows`;
  renderKpis(records);
  renderCharts(records);
  renderInsights(records);
  renderRankingTable(records);
  renderDetailTable(records);
}

loadData();
