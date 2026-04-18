const IRN_THRESHOLD = 3;
const TABLE_PAGE_SIZE = 10;
const COUNTRY_REFERENCE_LIST = [
  "United States", "India", "United Kingdom", "Iran", "Japan", "Australia", "China", "Italy", "Malaysia",
  "Saudi Arabia", "Germany", "Pakistan", "Taiwan", "Nigeria", "Ethiopia", "South Korea", "Poland", "Canada",
  "Russian Federation", "Brazil", "Spain", "Egypt", "Netherlands", "Thailand", "Turkey", "Portugal", "Romania",
  "France", "United Arab Emirates", "Bangladesh", "Iraq", "Singapore", "Austria", "Viet Nam", "South Africa",
  "Jordan", "Norway", "Switzerland", "Hungary", "Hong Kong", "Greece", "Sweden", "Chile", "Colombia", "Finland",
  "Ireland", "Ghana", "Philippines", "Belgium", "Czech Republic", "New Zealand", "Argentina", "Mexico", "Kenya",
  "Oman", "Sri Lanka", "Yemen", "Denmark", "Croatia", "Kuwait", "Bulgaria", "Brunei Darussalam", "Serbia", "Qatar",
  "Morocco", "Estonia", "Nepal", "Bahrain", "Lithuania", "Kazakhstan", "Peru", "Tunisia", "Cameroon", "Ukraine",
  "Kyrgyzstan", "Israel", "Zambia", "Latvia", "Algeria", "Jamaica", "Georgia", "Lebanon", "Botswana", "Iceland",
  "Sudan", "Luxembourg", "Libya", "Gambia", "Guatemala", "Uganda", "Ecuador", "Cyprus", "Cambodia", "Palestine",
  "Tanzania", "Slovakia", "Rwanda", "Bosnia and Herzegovina", "Mongolia", "Syrian Arab Republic", "Malta", "Laos",
  "Sierra Leone", "Uruguay", "Costa Rica", "Benin", "Armenia", "Albania", "Antigua and Barbuda", "Azerbaijan",
  "Afghanistan", "Fiji", "Slovenia", "Venezuela", "Cuba", "Macao", "Timor-Leste", "Barbados", "Somalia", "Moldova",
  "Uzbekistan", "Bolivia", "El Salvador", "Paraguay", "Mozambique", "Namibia", "Democratic Republic Congo",
  "Madagascar", "Côte d'Ivoire", "American Samoa", "Andorra", "Angola", "Anguilla", "Antarctica", "Aruba",
  "Bahamas", "Belarus", "Belize", "Bermuda", "Bhutan", "Burkina Faso", "Burundi", "Cape Verde", "Cayman Islands",
  "Central African Republic", "Chad", "Comoros", "Congo", "Cook Islands", "Curacao", "Djibouti", "Dominica",
  "Dominican Republic", "Equatorial Guinea", "Eritrea", "Falkland Islands (Malvinas)", "Faroe Islands",
  "Federated States of Micronesia", "French Guiana", "French Polynesia", "Gabon", "Gibraltar", "Greenland",
  "Grenada", "Guadeloupe", "Guam", "Guinea", "Guinea-Bissau", "Guyana", "Haïti", "Honduras", "Jersey", "Kiribati",
  "Lesotho", "Liberia", "Liechtenstein", "Malawi", "Maldives", "Mali", "Marshall Islands", "Martinique",
  "Mauritania", "Mauritius", "Mayotte", "Monaco", "Montenegro", "Montserrat", "Myanmar", "Nauru",
  "Netherlands Antilles", "New Caledonia", "Nicaragua", "Niger", "Niue", "Norfolk Island", "North Korea",
  "North Macedonia", "Northern Mariana Islands", "Palau", "Panama", "Papua New Guinea", "Puerto Rico", "Reunion",
  "Saint Helena", "Saint Kitts and Nevis", "Saint Lucia", "Saint Vincent and the Grenadines", "Samoa", "San Marino",
  "Sao Tome and Principe", "Senegal", "Seychelles", "Solomon Islands", "South Georgia and the South Sandwich Islands",
  "South Sudan", "Suriname", "Svalbard and Jan Mayen", "Swaziland", "Tajikistan", "Togo", "Tonga",
  "Trinidad and Tobago", "Turkmenistan", "Turks and Caicos Islands", "Tuvalu", "United States Minor Outlying Islands",
  "Vanuatu", "Vatican City State", "Virgin Islands (British)", "Virgin Islands (U.S.)", "Wallis and Futuna",
  "Western Sahara", "Zimbabwe"
];

// Extend these maps over time to improve institution and country normalization.
const INSTITUTION_ALIAS_MAP = {
  "brawijaya university": "Universitas Brawijaya",
  "universitas brawijaya": "Universitas Brawijaya",
  "univ brawijaya": "Universitas Brawijaya",
  "universiti malaya": "Universiti Malaya",
  "university of malaya": "Universiti Malaya",
  "universitas indonesia": "Universitas Indonesia",
  "seoul national university": "Seoul National University",
  "the university of tokyo": "The University of Tokyo",
  "nanyang technological university": "Nanyang Technological University",
};

const COUNTRY_ALIAS_MAP = {
  indonesia: "Indonesia",
  malaysia: "Malaysia",
  singapore: "Singapore",
  thailand: "Thailand",
  japan: "Japan",
  china: "China",
  "hong kong": "Hong Kong",
  korea: "South Korea",
  "south korea": "South Korea",
  australia: "Australia",
  iran: "Iran",
  germany: "Germany",
  france: "France",
  netherlands: "Netherlands",
  uk: "United Kingdom",
  "united kingdom": "United Kingdom",
  england: "United Kingdom",
  scotland: "United Kingdom",
  usa: "United States",
  "united states": "United States",
  canada: "Canada",
  taiwan: "Taiwan",
  india: "India",
  pakistan: "Pakistan",
  saudi: "Saudi Arabia",
  "saudi arabia": "Saudi Arabia",
  qatar: "Qatar",
  "united arab emirates": "United Arab Emirates",
  vietnam: "Viet Nam",
  "viet nam": "Viet Nam",
  philippines: "Philippines",
  brunei: "Brunei Darussalam",
  russia: "Russian Federation",
  "russian federation": "Russian Federation",
  italy: "Italy",
  spain: "Spain",
  belgium: "Belgium",
  sweden: "Sweden",
  norway: "Norway",
  denmark: "Denmark",
  switzerland: "Switzerland",
  finland: "Finland",
  poland: "Poland",
  portugal: "Portugal",
  turkey: "Turkey",
  mexico: "Mexico",
  brazil: "Brazil",
  egypt: "Egypt",
};

const COUNTRY_BY_INSTITUTION_ALIAS = {
  "Universitas Brawijaya": "Indonesia",
  "Universiti Malaya": "Malaysia",
  "Nanyang Technological University": "Singapore",
  "Seoul National University": "South Korea",
  "The University of Tokyo": "Japan",
};

const UB_NAME_PATTERNS = [
  /universitas brawijaya/i,
  /\bbrawijaya university\b/i,
  /\buniv(?:ersity)? brawijaya\b/i,
];

const DOM = {
  reloadDataBtn: document.getElementById("reloadDataBtn"),
  loadStatus: document.getElementById("loadStatus"),
  dataSummaryBadge: document.getElementById("dataSummaryBadge"),
  scopeSelect: document.getElementById("scopeSelect"),
  startYearSelect: document.getElementById("startYearSelect"),
  endYearSelect: document.getElementById("endYearSelect"),
  irnWindowMode: document.getElementById("irnWindowMode"),
  globalSearchInput: document.getElementById("globalSearchInput"),
  excludeIndonesiaCheckbox: document.getElementById("excludeIndonesiaCheckbox"),
  detailCountrySelect: document.getElementById("detailCountrySelect"),
  kpiGrid: document.getElementById("kpiGrid"),
  irnUniversityTable: document.getElementById("irnUniversityTable"),
  countryTable: document.getElementById("countryTable"),
  universityTable: document.getElementById("universityTable"),
  lecturerUniversityTable: document.getElementById("lecturerUniversityTable"),
  lecturerCountryTable: document.getElementById("lecturerCountryTable"),
  unmatchedAuthorsTable: document.getElementById("unmatchedAuthorsTable"),
  unmatchedInstitutionsTable: document.getElementById("unmatchedInstitutionsTable"),
  duplicateVariantsTable: document.getElementById("duplicateVariantsTable"),
  detailPanel: document.getElementById("detailPanel"),
  loadingOverlay: document.getElementById("loadingOverlay"),
  loadingMessage: document.getElementById("loadingMessage"),
};

const state = {
  raw: null,
  filters: {
    scope: "all",
    startYear: null,
    endYear: null,
    irnWindowMode: "selected",
    search: "",
    excludeIndonesiaPartners: false,
  },
  derived: null,
  sort: {
    universityTable: { key: "publicationCount", direction: "desc" },
    countryTable: { key: "irnUniversityCount", direction: "desc" },
    irnUniversityTable: { key: "publicationCount", direction: "desc" },
  },
  detail: null,
  charts: {},
  tableUi: {},
};

function init() {
  bindEvents();
  renderEmptyDashboard();
  loadBundledData();
}

function bindEvents() {
  DOM.reloadDataBtn.addEventListener("click", loadBundledData);
  DOM.scopeSelect.addEventListener("change", onFilterChange);
  DOM.startYearSelect.addEventListener("change", onFilterChange);
  DOM.endYearSelect.addEventListener("change", onFilterChange);
  DOM.irnWindowMode.addEventListener("change", onFilterChange);
  DOM.globalSearchInput.addEventListener("input", onFilterChange);
  DOM.excludeIndonesiaCheckbox.addEventListener("change", onFilterChange);
  DOM.detailCountrySelect.addEventListener("change", () => {
    if (DOM.detailCountrySelect.value) {
      state.detail = { type: "country", key: DOM.detailCountrySelect.value };
      renderDetailPanel(state.derived);
    }
  });

  document.querySelectorAll("th[data-sort]").forEach((header) => {
    header.addEventListener("click", () => {
      const tbodyId = header.closest("table").querySelector("tbody").id;
      toggleSort(tbodyId, header.dataset.sort);
      renderDashboard();
    });
  });

  document.querySelectorAll("[data-export]").forEach((button) => {
    button.addEventListener("click", () => exportDataset(button.dataset.export));
  });

  window.addEventListener("resize", () => {
    Object.values(state.charts).forEach((chart) => chart?.resize());
  });

  [
    "IrnUniversityTable",
    "CountryTable",
    "UniversityTable",
    "LecturerUniversityTable",
    "LecturerCountryTable",
    "UnmatchedAuthorsTable",
    "UnmatchedInstitutionsTable",
    "DuplicateVariantsTable",
  ].forEach((name) => {
    const input = document.getElementById(`search${name}`);
    if (input) {
      const tableId = name.charAt(0).toLowerCase() + name.slice(1);
      state.tableUi[tableId] = { page: 1, search: "" };
      input.addEventListener("input", () => {
        state.tableUi[tableId].search = input.value.trim().toLowerCase();
        state.tableUi[tableId].page = 1;
        renderDashboard();
      });
    }
  });
}

async function loadBundledData() {
  try {
    showLoading("Reading bundled Excel workbooks...");
    const [lecturerWorkbook, publicationWorkbook, institutionWorkbook] = await Promise.all([
      readBundledWorkbook("lecturer"),
      readBundledWorkbook("publication"),
      readBundledWorkbook("institution"),
    ]);

    showLoading("Parsing lecturer mapping...");
    const lecturerRows = XLSX.utils.sheet_to_json(
      lecturerWorkbook.Sheets[lecturerWorkbook.SheetNames[0]],
      { defval: "" }
    );

    showLoading("Parsing publication records...");
    const publicationRows = XLSX.utils.sheet_to_json(
      publicationWorkbook.Sheets[publicationWorkbook.SheetNames[0]],
      { defval: "" }
    );

    showLoading("Parsing institution-country mapping...");
    const institutionRows = XLSX.utils.sheet_to_json(
      institutionWorkbook.Sheets[institutionWorkbook.SheetNames[0]],
      { defval: "" }
    );

    state.raw = buildDataModel(lecturerRows, publicationRows, institutionRows);
    initializeFilters(state.raw);
    renderDashboard();

    setStatus(
      `Loaded ${state.raw.publications.length.toLocaleString()} publications, ${state.raw.lecturers.length.toLocaleString()} lecturer mappings, and ${state.raw.institutionCountryMap.size.toLocaleString()} institution-country mappings.`,
      "success"
    );
  } catch (error) {
    console.error(error);
    setStatus(`Failed to load bundled files: ${error.message}`, "error");
  } finally {
    hideLoading();
  }
}

function clearDashboard() {
  state.raw = null;
  state.derived = null;
  state.detail = null;
  DOM.globalSearchInput.value = "";
  DOM.excludeIndonesiaCheckbox.checked = false;
  DOM.detailCountrySelect.innerHTML = '<option value="">Select country</option>';
  document.querySelectorAll(".table-search-input").forEach((input) => {
    input.value = "";
  });
  Object.keys(state.tableUi).forEach((key) => {
    state.tableUi[key] = { page: 1, search: "" };
  });
  [DOM.startYearSelect, DOM.endYearSelect, DOM.irnWindowMode].forEach((element) => {
    element.disabled = true;
  });
  DOM.scopeSelect.value = "all";
  disposeCharts();
  renderEmptyDashboard();
  setStatus("Dashboard cleared. Reload bundled data to begin again.", "info");
}

function buildDataModel(lecturerRows, publicationRows, institutionRows = []) {
  const lecturers = lecturerRows
    .map((row) => ({
      scopusId: normalizeScopusId(row["SCOPUS ID"]),
      faculty: normalizeText(row["FACULTY"]),
      name: normalizeText(row["NAME"]),
    }))
    .filter((row) => row.scopusId && row.faculty && row.name);

  const lecturerById = new Map();
  const facultySet = new Set();
  lecturers.forEach((lecturer) => {
    lecturerById.set(lecturer.scopusId, lecturer);
    facultySet.add(lecturer.faculty);
  });

  const unmatchedAuthorCounts = new Map();
  const institutionVariantMap = new Map();
  const institutionCountryMap = buildInstitutionCountryMap(institutionRows);

  const publications = publicationRows
    .map((row, index) => {
      const authorIds = splitPipeValues(row["Scopus Author Ids"]).map(normalizeScopusId).filter(Boolean);
      const affiliations = splitPipeValues(row["Institutions"] || row["Scopus Affiliation names"]);
      const affiliationIds = splitPipeValues(row["Scopus Affiliation IDs"]);
      const ubLecturers = [];
      const facultySetForPaper = new Set();

      authorIds.forEach((authorId) => {
        const lecturer = lecturerById.get(authorId);
        if (lecturer) {
          ubLecturers.push(lecturer);
          facultySetForPaper.add(lecturer.faculty);
        } else {
          unmatchedAuthorCounts.set(authorId, (unmatchedAuthorCounts.get(authorId) || 0) + 1);
        }
      });

      const partnerInstitutionMap = new Map();
      affiliations.forEach((rawInstitution, institutionIndex) => {
        const normalizedInstitution = normalizeInstitution(rawInstitution);
        if (!normalizedInstitution || isUBInstitution(normalizedInstitution)) {
          return;
        }

        const country =
          institutionCountryMap.get(normalizedInstitution) ||
          inferCountry(rawInstitution, normalizedInstitution) ||
          "Unknown";
        if (!institutionVariantMap.has(normalizedInstitution)) {
          institutionVariantMap.set(normalizedInstitution, new Set());
        }
        institutionVariantMap.get(normalizedInstitution).add(String(rawInstitution).trim());

        const key = `${normalizedInstitution}__${country}`;
        if (!partnerInstitutionMap.has(key)) {
          partnerInstitutionMap.set(key, {
            name: normalizedInstitution,
            country,
            affiliationId: normalizeText(affiliationIds[institutionIndex]),
          });
        }
      });

      const year = Number.parseInt(row["Year"], 10);
      const countryRegionCount = Number.parseInt(row["Number of Countries/Regions"], 10) || 0;

      return {
        id: buildPaperId(row, index),
        year: Number.isFinite(year) ? year : null,
        countryRegionCount,
        isInternational: countryRegionCount > 1,
        authorIds,
        ubLecturers: dedupeByKey(ubLecturers, "scopusId"),
        faculties: Array.from(facultySetForPaper).sort(),
        partnerInstitutions: Array.from(partnerInstitutionMap.values()),
      };
    })
    .filter((paper) => paper.year);

  return {
    lecturers,
    lecturerById,
    faculties: Array.from(facultySet).sort(),
    publications,
    unmatchedAuthorCounts,
    institutionVariantMap,
    institutionCountryMap,
  };
}

function initializeFilters(rawData) {
  const years = Array.from(new Set(rawData.publications.map((paper) => paper.year))).sort((a, b) => a - b);
  const minYear = years[0];
  const maxYear = years[years.length - 1];
  const currentYear = new Date().getFullYear();
  const defaultEndYear = years.includes(currentYear) ? currentYear : maxYear;
  const defaultStartYear = Math.max(minYear, defaultEndYear - 4);

  state.filters = {
    scope: "all",
    startYear: defaultStartYear,
    endYear: defaultEndYear,
    irnWindowMode: "selected",
    search: "",
    excludeIndonesiaPartners: false,
  };

  populateSelect(DOM.scopeSelect, ["All UB", ...rawData.faculties], null, "All UB");
  DOM.scopeSelect.value = "All UB";
  DOM.startYearSelect.disabled = false;
  DOM.endYearSelect.disabled = false;
  DOM.irnWindowMode.disabled = false;

  populateSelect(DOM.startYearSelect, years, null);
  populateSelect(DOM.endYearSelect, years, null);
  DOM.startYearSelect.value = String(defaultStartYear);
  DOM.endYearSelect.value = String(defaultEndYear);
}

function populateSelect(selectElement, values, placeholder, allValue = null) {
  const options = [];
  if (placeholder !== null) {
    options.push(`<option value="">${placeholder}</option>`);
  }
  values.forEach((value) => {
    const optionValue = allValue !== null && value === allValue ? "all" : String(value);
    options.push(`<option value="${escapeHtml(optionValue)}">${escapeHtml(String(value))}</option>`);
  });
  selectElement.innerHTML = options.join("");
}

function onFilterChange() {
  state.filters.scope = DOM.scopeSelect.value;
  state.filters.startYear = Number.parseInt(DOM.startYearSelect.value, 10) || null;
  state.filters.endYear = Number.parseInt(DOM.endYearSelect.value, 10) || null;
  state.filters.irnWindowMode = DOM.irnWindowMode.value;
  state.filters.search = DOM.globalSearchInput.value.trim().toLowerCase();
  state.filters.excludeIndonesiaPartners = DOM.excludeIndonesiaCheckbox.checked;

  if (state.filters.startYear && state.filters.endYear && state.filters.startYear > state.filters.endYear) {
    [state.filters.startYear, state.filters.endYear] = [state.filters.endYear, state.filters.startYear];
    DOM.startYearSelect.value = String(state.filters.startYear);
    DOM.endYearSelect.value = String(state.filters.endYear);
  }

  renderDashboard();
}

function renderDashboard() {
  if (!state.raw) {
    renderEmptyDashboard();
    return;
  }

  state.derived = computeFilteredAnalytics(state.raw, state.filters);
  syncDetailCountryOptions(state.derived.countryStats);
  renderKpis(state.derived);
  renderCharts(state.derived);
  renderTables(state.derived);
  renderDetailPanel(state.derived);
  DOM.dataSummaryBadge.textContent = `${state.derived.filteredPublications.length.toLocaleString()} filtered publications`;
}

function computeFilteredAnalytics(rawData, filters) {
  const filteredPublications = rawData.publications.filter((paper) => {
    const withinYearRange =
      (!filters.startYear || paper.year >= filters.startYear) &&
      (!filters.endYear || paper.year <= filters.endYear);
    if (!withinYearRange) {
      return false;
    }
    if (filters.scope !== "all") {
      return paper.faculties.includes(filters.scope);
    }
    return true;
  });

  const irnWindowStart =
    filters.irnWindowMode === "rolling5" && filters.endYear
      ? Math.max(filters.startYear || filters.endYear - 4, filters.endYear - 4)
      : filters.startYear;

  const irnPapers = filteredPublications.filter((paper) => !irnWindowStart || paper.year >= irnWindowStart);
  const yearlyStats = buildYearlyStats(filteredPublications);
  const distributionStats = buildCountryRegionDistribution(filteredPublications);
  const partnerFilteredPublications = filters.excludeIndonesiaPartners
    ? filteredPublications.map((paper) => ({
        ...paper,
        partnerInstitutions: paper.partnerInstitutions.filter((institution) => institution.country !== "Indonesia"),
      }))
    : filteredPublications;
  const partnerFilteredIrnPapers = filters.excludeIndonesiaPartners
    ? irnPapers.map((paper) => ({
        ...paper,
        partnerInstitutions: paper.partnerInstitutions.filter((institution) => institution.country !== "Indonesia"),
      }))
    : irnPapers;

  const universityStats = buildUniversityStats(partnerFilteredPublications, partnerFilteredIrnPapers);
  const countryStats = buildCountryStats(universityStats, COUNTRY_REFERENCE_LIST);
  const facultyIntlStats = buildFacultyStats(filteredPublications, rawData.faculties);
  const dataQuality = buildDataQuality(rawData, partnerFilteredPublications);
  const lecturerUniversityRows = buildLecturerUniversityRows(universityStats);
  const lecturerCountryRows = buildLecturerCountryRows(countryStats);
  const searchText = filters.search;
  const searchFilter = (text) => !searchText || text.toLowerCase().includes(searchText);
  const filteredUniversityStats = universityStats.filter((row) =>
    searchFilter(`${row.name} ${row.country} ${row.lecturerNames.join(" ")} ${row.faculties.join(" ")}`)
  );
  const filteredCountryStats = countryStats.filter((row) =>
    searchFilter(`${row.name} ${row.universities.join(" ")} ${row.lecturerNames.join(" ")}`)
  );

  return {
    filteredPublications,
    irnWindowStart,
    yearlyStats,
    distributionStats,
    universityStats: filteredUniversityStats,
    irnUniversityStats: filteredUniversityStats.filter((row) => row.irnQualified).filter((row) =>
      searchFilter(`${row.name} ${row.country} ${row.lecturerNames.join(" ")} ${row.faculties.join(" ")}`)
    ),
    countryStats: filteredCountryStats,
    facultyIntlStats,
    dataQuality,
    topUniversities: filteredUniversityStats.slice(0, 10),
    topCountries: filteredCountryStats.slice(0, 10),
    lecturerUniversityRows: lecturerUniversityRows.filter((row) =>
      searchFilter(`${row.university} ${row.country} ${row.lecturers.join(" ")}`)
    ),
    lecturerCountryRows: lecturerCountryRows.filter((row) =>
      searchFilter(`${row.country} ${row.universities.join(" ")} ${row.lecturers.join(" ")}`)
    ),
  };
}

function buildYearlyStats(publications) {
  const map = new Map();
  publications.forEach((paper) => {
    if (!map.has(paper.year)) {
      map.set(paper.year, { year: paper.year, total: 0, domestic: 0, international: 0 });
    }
    const bucket = map.get(paper.year);
    bucket.total += 1;
    if (paper.isInternational) {
      bucket.international += 1;
    } else {
      bucket.domestic += 1;
    }
  });
  return Array.from(map.values()).sort((a, b) => a.year - b.year);
}

function buildCountryRegionDistribution(publications) {
  const exactMap = new Map();
  let domestic = 0;
  let international = 0;

  publications.forEach((paper) => {
    exactMap.set(paper.countryRegionCount, (exactMap.get(paper.countryRegionCount) || 0) + 1);
    if (paper.isInternational) {
      international += 1;
    } else {
      domestic += 1;
    }
  });

  const total = publications.length || 1;
  return {
    exact: Array.from(exactMap.entries())
      .map(([bucket, count]) => ({ bucket, count, percentage: (count / total) * 100 }))
      .sort((a, b) => a.bucket - b.bucket),
    grouped: [
      { name: "Domestic / single country", count: domestic, percentage: (domestic / total) * 100 },
      { name: "International / multi-country", count: international, percentage: (international / total) * 100 },
    ],
  };
}

function buildUniversityStats(publications, irnPapers) {
  const map = new Map();
  const irnPaperIdsByUniversity = new Map();

  irnPapers.forEach((paper) => {
    paper.partnerInstitutions.forEach((institution) => {
      if (!irnPaperIdsByUniversity.has(institution.name)) {
        irnPaperIdsByUniversity.set(institution.name, new Set());
      }
      irnPaperIdsByUniversity.get(institution.name).add(paper.id);
    });
  });

  publications.forEach((paper) => {
    paper.partnerInstitutions.forEach((institution) => {
      if (!map.has(institution.name)) {
        map.set(institution.name, {
          name: institution.name,
          country: institution.country,
          publicationIds: new Set(),
          lecturerNames: new Set(),
          lecturerPaperCounts: new Map(),
          faculties: new Set(),
          yearlyCounts: new Map(),
        });
      }

      const bucket = map.get(institution.name);
      bucket.publicationIds.add(paper.id);
      bucket.yearlyCounts.set(paper.year, (bucket.yearlyCounts.get(paper.year) || 0) + 1);
      paper.ubLecturers.forEach((lecturer) => {
        bucket.lecturerNames.add(lecturer.name);
        bucket.lecturerPaperCounts.set(
          lecturer.name,
          (bucket.lecturerPaperCounts.get(lecturer.name) || 0) + 1
        );
      });
      paper.faculties.forEach((faculty) => bucket.faculties.add(faculty));
    });
  });

  return Array.from(map.values())
    .map((bucket) => {
      const irnCount = irnPaperIdsByUniversity.get(bucket.name)?.size || 0;
      return {
        name: bucket.name,
        country: bucket.country,
        publicationCount: bucket.publicationIds.size,
        lecturerNames: Array.from(bucket.lecturerNames).sort(),
        lecturerCount: bucket.lecturerNames.size,
        lecturerPaperCounts: mapToSortedPairs(bucket.lecturerPaperCounts),
        faculties: Array.from(bucket.faculties).sort(),
        yearlyCounts: mapToSortedPairs(bucket.yearlyCounts),
        irnPublicationCount: irnCount,
        irnQualified: irnCount >= IRN_THRESHOLD,
      };
    })
    .sort((a, b) => b.publicationCount - a.publicationCount || a.name.localeCompare(b.name));
}

function buildCountryStats(universityStats, referenceCountries = []) {
  const map = new Map();
  referenceCountries.forEach((country) => {
    map.set(country, {
      name: country,
      publicationCount: 0,
      universities: [],
      irnUniversities: [],
      lecturerNames: new Set(),
      yearlyCounts: new Map(),
    });
  });

  universityStats.forEach((university) => {
    const country = university.country || "Unknown";
    if (!map.has(country)) {
      map.set(country, {
        name: country,
        publicationCount: 0,
        universities: [],
        irnUniversities: [],
        lecturerNames: new Set(),
        yearlyCounts: new Map(),
      });
    }

    const bucket = map.get(country);
    bucket.publicationCount += university.publicationCount;
    bucket.universities.push(university.name);
    if (university.irnQualified) {
      bucket.irnUniversities.push(university.name);
    }
    university.lecturerNames.forEach((name) => bucket.lecturerNames.add(name));
    university.yearlyCounts.forEach(([year, count]) => {
      bucket.yearlyCounts.set(year, (bucket.yearlyCounts.get(year) || 0) + count);
    });
  });

  return Array.from(map.values())
    .map((bucket) => ({
      name: bucket.name,
      publicationCount: bucket.publicationCount,
      universities: bucket.universities.sort(),
      irnUniversities: bucket.irnUniversities.sort(),
      irnUniversityCount: bucket.irnUniversities.length,
      lecturerNames: Array.from(bucket.lecturerNames).sort(),
      yearlyCounts: mapToSortedPairs(bucket.yearlyCounts),
    }))
    .sort((a, b) => b.publicationCount - a.publicationCount || a.name.localeCompare(b.name));
}

function buildInstitutionCountryMap(institutionRows) {
  const map = new Map();
  institutionRows.forEach((row) => {
    const institution = normalizeInstitution(row["Institution"]);
    const country = normalizeCountryValue(row["Country/Region"]);
    const otherNames = String(row["Other name"] ?? "")
      .split(";")
      .map((value) => normalizeInstitution(value))
      .filter(Boolean);

    if (!country) {
      return;
    }

    [institution, ...otherNames].filter(Boolean).forEach((name) => {
      map.set(name, country);
    });
  });
  return map;
}

function buildFacultyStats(publications, faculties) {
  return faculties
    .map((faculty) => {
      const papers = publications.filter((paper) => paper.faculties.includes(faculty));
      const internationalCount = papers.filter((paper) => paper.isInternational).length;
      return {
        faculty,
        totalCount: papers.length,
        internationalCount,
      };
    })
    .filter((row) => row.totalCount > 0)
    .sort((a, b) => b.internationalCount - a.internationalCount || a.faculty.localeCompare(b.faculty))
    .slice(0, 10);
}

function buildDataQuality(rawData, filteredPublications) {
  const unknownInstitutionCounts = new Map();

  filteredPublications.forEach((paper) => {
    paper.partnerInstitutions.forEach((institution) => {
      if (institution.country === "Unknown") {
        unknownInstitutionCounts.set(
          institution.name,
          (unknownInstitutionCounts.get(institution.name) || 0) + 1
        );
      }
    });
  });

  const duplicateVariants = Array.from(rawData.institutionVariantMap.entries())
    .map(([canonical, variants]) => ({
      canonical,
      variants: Array.from(variants).filter(Boolean).sort(),
    }))
    .filter((row) => row.variants.length > 1)
    .sort((a, b) => b.variants.length - a.variants.length || a.canonical.localeCompare(b.canonical))
    .slice(0, 100);

  return {
    unmatchedAuthors: Array.from(rawData.unmatchedAuthorCounts.entries())
      .map(([authorId, occurrences]) => ({ authorId, occurrences }))
      .sort((a, b) => b.occurrences - a.occurrences || a.authorId.localeCompare(b.authorId))
      .slice(0, 200),
    unknownInstitutions: Array.from(unknownInstitutionCounts.entries())
      .map(([institution, occurrences]) => ({ institution, occurrences }))
      .sort((a, b) => b.occurrences - a.occurrences || a.institution.localeCompare(b.institution))
      .slice(0, 200),
    duplicateVariants,
  };
}

function buildLecturerUniversityRows(universityStats) {
  return universityStats.map((row) => ({
    university: row.name,
    country: row.country,
    lecturers: row.lecturerNames,
    lecturerCounts: row.lecturerPaperCounts.map(([name, count]) => `${name} (${count})`),
  }));
}

function buildLecturerCountryRows(countryStats) {
  return countryStats.map((row) => ({
    country: row.name,
    universities: row.universities,
    lecturers: row.lecturerNames,
  }));
}

function renderEmptyDashboard() {
  DOM.kpiGrid.innerHTML = createPlaceholderCards();
  [
    DOM.irnUniversityTable,
    DOM.countryTable,
    DOM.universityTable,
    DOM.lecturerUniversityTable,
    DOM.lecturerCountryTable,
    DOM.unmatchedAuthorsTable,
    DOM.unmatchedInstitutionsTable,
    DOM.duplicateVariantsTable,
  ].forEach((tbody) => {
    tbody.innerHTML = '<tr><td colspan="5" class="muted">Load bundled data to populate this table.</td></tr>';
  });

  DOM.detailPanel.className = "detail-body empty-state";
  DOM.detailPanel.textContent =
    "Click a university or country once data is loaded to inspect yearly trends, collaborating lecturers, and faculty involvement.";
  disposeCharts();
  renderEmptyCharts();
  DOM.dataSummaryBadge.textContent = "No data loaded";
}

function createPlaceholderCards() {
  const titles = [
    "Total Publications",
    "International Collaboration",
    "International Share",
    "Partner Universities",
    "IRN Universities",
    "Partner Countries",
  ];

  return titles
    .map(
      (title) => `
      <article class="panel kpi-card">
        <h3>${title}</h3>
        <strong>--</strong>
        <span>Load bundled data to calculate this metric.</span>
      </article>
    `
    )
    .join("");
}

function renderKpis(derived) {
  const totalPublications = derived.filteredPublications.length;
  const internationalPublications = derived.filteredPublications.filter((paper) => paper.isInternational).length;
  const internationalShare = totalPublications ? (internationalPublications / totalPublications) * 100 : 0;

  const cards = [
    {
      title: "Total Publications",
      value: formatNumber(totalPublications),
      note: "Distinct filtered papers",
    },
    {
      title: "International Collaboration",
      value: formatNumber(internationalPublications),
      note: "Publications with more than 1 country/region",
    },
    {
      title: "International Share",
      value: `${internationalShare.toFixed(1)}%`,
      note: "International publications / total filtered publications",
    },
    {
      title: "Partner Universities",
      value: formatNumber(derived.universityStats.length),
      note: "Excluding UB after normalization",
    },
    {
      title: "IRN Universities",
      value: formatNumber(derived.irnUniversityStats.length),
      note: `At least ${IRN_THRESHOLD} joint publications in the IRN window`,
    },
    {
      title: "Partner Countries",
      value: formatNumber(derived.countryStats.length),
      note: "Countries inferred from affiliation strings and aliases",
    },
  ];

  DOM.kpiGrid.innerHTML = cards
    .map(
      (card) => `
      <article class="panel kpi-card">
        <h3>${card.title}</h3>
        <strong>${card.value}</strong>
        <span>${card.note}</span>
      </article>
    `
    )
    .join("");
}

function renderCharts(derived) {
  renderYearTrendChart(derived.yearlyStats);
  renderIntlTrendChart(derived.yearlyStats);
  renderCountryDistributionChart(derived.distributionStats.exact);
  renderDomesticInternationalChart(derived.distributionStats.grouped);
  renderTopUniversitiesChart(derived.topUniversities);
  renderTopCountriesChart(derived.topCountries);
  renderFacultyIntlChart(derived.facultyIntlStats);
}

function renderYearTrendChart(yearlyStats) {
  const chart = getChart("yearTrendChart");
  chart.clear();
  chart.setOption({
    tooltip: { trigger: "axis" },
    grid: { left: 40, right: 20, top: 30, bottom: 40 },
    xAxis: { type: "category", data: yearlyStats.map((row) => row.year) },
    yAxis: { type: "value" },
    series: [
      {
        type: "line",
        smooth: true,
        data: yearlyStats.map((row) => row.total),
        symbolSize: 10,
        areaStyle: {
          color: new echarts.graphic.LinearGradient(0, 0, 0, 1, [
            { offset: 0, color: "rgba(0,95,115,0.42)" },
            { offset: 1, color: "rgba(0,95,115,0.04)" },
          ]),
        },
        lineStyle: { color: "#005f73", width: 4, shadowBlur: 12, shadowColor: "rgba(0,95,115,0.2)" },
        itemStyle: { color: "#ffffff", borderColor: "#005f73", borderWidth: 2 },
      },
    ],
  });
}

function renderIntlTrendChart(yearlyStats) {
  const chart = getChart("intlTrendChart");
  chart.clear();
  const domesticSeries = yearlyStats.map((row) => {
    const hasInternational = row.international > 0;
    return {
      value: row.domestic,
      itemStyle: {
        color: "#94a9c7",
        borderRadius: hasInternational ? [0, 0, 0, 0] : [12, 12, 0, 0],
      },
    };
  });
  const internationalSeries = yearlyStats.map((row) => ({
    value: row.international,
    itemStyle: {
      color: "#c06c2e",
      borderRadius: row.international > 0 ? [12, 12, 0, 0] : [0, 0, 0, 0],
    },
  }));

  chart.setOption({
    tooltip: { trigger: "axis", axisPointer: { type: "shadow" } },
    legend: {
      bottom: 0,
      data: [
        { name: "Domestic", itemStyle: { color: "#94a9c7" } },
        { name: "International", itemStyle: { color: "#c06c2e" } },
      ],
    },
    grid: { left: 40, right: 20, top: 20, bottom: 50 },
    xAxis: { type: "category", data: yearlyStats.map((row) => row.year) },
    yAxis: { type: "value" },
    series: [
      {
        name: "Domestic",
        type: "bar",
        stack: "total",
        barWidth: 58,
        barGap: "-100%",
        data: domesticSeries,
      },
      {
        name: "International",
        type: "bar",
        stack: "total",
        barWidth: 58,
        data: internationalSeries,
      },
    ],
  });
}

function renderCountryDistributionChart(distribution) {
  const chart = getChart("countryDistributionChart");
  chart.clear();
  chart.setOption({
    tooltip: {
      trigger: "axis",
      formatter: (params) => `${params[0].axisValue} countries/regions<br>${params[0].data} publications`,
    },
    grid: { left: 40, right: 20, top: 20, bottom: 40 },
    xAxis: { type: "category", data: distribution.map((row) => row.bucket) },
    yAxis: { type: "value" },
    series: [
      {
        type: "bar",
        data: distribution.map((row) => row.count),
        itemStyle: {
          color: (params) =>
            Number(params.name) > 1
              ? new echarts.graphic.LinearGradient(0, 0, 0, 1, [
                  { offset: 0, color: "#0f8795" },
                  { offset: 1, color: "#005f73" },
                ])
              : new echarts.graphic.LinearGradient(0, 0, 0, 1, [
                  { offset: 0, color: "#b8c9db" },
                  { offset: 1, color: "#8aa4bf" },
                ]),
          borderRadius: [12, 12, 0, 0],
        },
      },
    ],
  });
}

function renderDomesticInternationalChart(grouped) {
  const chart = getChart("domesticInternationalChart");
  chart.clear();
  chart.setOption({
    tooltip: {
      trigger: "item",
      formatter: (params) =>
        `${params.name}<br>${params.value} publications (${Number(params.percent).toFixed(1)}%)`,
    },
    series: [
      {
        type: "pie",
        radius: ["48%", "72%"],
        center: ["50%", "48%"],
        label: { formatter: "{b}\n{d}%" },
        labelLine: { length: 16, length2: 12 },
        data: grouped.map((row, index) => ({
          name: row.name,
          value: row.count,
          itemStyle: {
            color: index === 0 ? "#91a8c2" : "#005f73",
            shadowBlur: 18,
            shadowColor: index === 0 ? "rgba(145,168,194,0.22)" : "rgba(0,95,115,0.22)",
          },
        })),
      },
    ],
  });
}

function renderTopUniversitiesChart(universities) {
  const chart = getChart("topUniversitiesChart");
  chart.off("click");
  chart.clear();
  chart.setOption({
    tooltip: { trigger: "axis", axisPointer: { type: "shadow" } },
    grid: { left: 160, right: 24, top: 16, bottom: 24 },
    xAxis: { type: "value" },
    yAxis: {
      type: "category",
      data: universities.map((row) => row.name).reverse(),
      axisLabel: { width: 140, overflow: "truncate" },
    },
    series: [
      {
        type: "bar",
        data: universities.map((row) => row.publicationCount).reverse(),
        itemStyle: {
          color: new echarts.graphic.LinearGradient(1, 0, 0, 0, [
            { offset: 0, color: "#0f8795" },
            { offset: 1, color: "#005f73" },
          ]),
          borderRadius: [0, 12, 12, 0],
        },
      },
    ],
  });

  chart.on("click", (params) => {
    const university = universities[universities.length - 1 - params.dataIndex];
    if (university) {
      state.detail = { type: "university", key: university.name };
      renderDetailPanel(state.derived);
    }
  });
}

function renderTopCountriesChart(countries) {
  const chart = getChart("topCountriesChart");
  chart.off("click");
  chart.clear();
  chart.setOption({
    tooltip: { trigger: "axis", axisPointer: { type: "shadow" } },
    grid: { left: 120, right: 24, top: 16, bottom: 24 },
    xAxis: { type: "value" },
    yAxis: { type: "category", data: countries.map((row) => row.name).reverse() },
    series: [
      {
        type: "bar",
        data: countries.map((row) => row.publicationCount).reverse(),
        itemStyle: {
          color: new echarts.graphic.LinearGradient(1, 0, 0, 0, [
            { offset: 0, color: "#d8853a" },
            { offset: 1, color: "#b35f22" },
          ]),
          borderRadius: [0, 12, 12, 0],
        },
      },
    ],
  });

  chart.on("click", (params) => {
    const country = countries[countries.length - 1 - params.dataIndex];
    if (country) {
      state.detail = { type: "country", key: country.name };
      renderDetailPanel(state.derived);
    }
  });
}

function renderFacultyIntlChart(rows) {
  const chart = getChart("facultyIntlChart");
  chart.clear();
  chart.setOption({
    tooltip: { trigger: "axis", axisPointer: { type: "shadow" } },
    grid: { left: 180, right: 24, top: 16, bottom: 24 },
    xAxis: { type: "value" },
    yAxis: {
      type: "category",
      data: rows.map((row) => row.faculty).reverse(),
      axisLabel: { width: 160, overflow: "truncate" },
    },
    series: [
      {
        type: "bar",
        data: rows.map((row) => row.internationalCount).reverse(),
        itemStyle: {
          color: new echarts.graphic.LinearGradient(1, 0, 0, 0, [
            { offset: 0, color: "#33a266" },
            { offset: 1, color: "#1f7a4a" },
          ]),
          borderRadius: [0, 12, 12, 0],
        },
      },
    ],
  });
}

function renderEmptyCharts() {
  [
    "yearTrendChart",
    "intlTrendChart",
    "countryDistributionChart",
    "domesticInternationalChart",
    "topUniversitiesChart",
    "topCountriesChart",
    "facultyIntlChart",
  ].forEach((id) => {
    const chart = getChart(id);
    chart.clear();
    chart.setOption({
      title: {
        text: "Waiting for data",
        left: "center",
        top: "middle",
        textStyle: { color: "#6d7c90", fontSize: 16, fontWeight: 600 },
      },
      xAxis: { show: false },
      yAxis: { show: false },
      series: [],
    });
  });
}

function renderTables(derived) {
  renderUniversityTable(derived.universityStats);
  renderIrnUniversityTable(derived.universityStats);
  renderCountryTable(derived.countryStats);
  renderLecturerUniversityTable(derived.lecturerUniversityRows);
  renderLecturerCountryTable(derived.lecturerCountryRows);
  renderDataQualityTables(derived.dataQuality);
}

function syncDetailCountryOptions(countryStats) {
  const currentValue = state.detail?.type === "country" ? state.detail.key : DOM.detailCountrySelect.value;
  DOM.detailCountrySelect.innerHTML = ['<option value="">Select country</option>']
    .concat(
      countryStats.map(
        (row) => `<option value="${escapeAttribute(row.name)}">${escapeHtml(row.name)}</option>`
      )
    )
    .join("");
  if (currentValue && countryStats.some((row) => row.name === currentValue)) {
    DOM.detailCountrySelect.value = currentValue;
  }
}

function renderUniversityTable(rows) {
  const sorted = sortRows("universityTable", rows);
  const paged = filterAndPaginateRows("universityTable", sorted, (row) =>
    `${row.name} ${row.country} ${row.publicationCount} ${row.lecturerNames.join(" ")} ${row.faculties.join(" ")}`
  );
  DOM.universityTable.innerHTML = renderRowsOrEmpty(
    paged.rows.map(
      (row) => `
      <tr>
        <td class="clickable-cell" data-detail-type="university" data-detail-key="${escapeAttribute(row.name)}">${escapeHtml(row.name)}</td>
        <td class="clickable-cell" data-detail-type="country" data-detail-key="${escapeAttribute(row.country)}">${escapeHtml(row.country)}</td>
        <td>${formatNumber(row.publicationCount)}</td>
        <td>${formatNumber(row.lecturerCount)}</td>
        <td>${renderTagList(row.faculties)}</td>
      </tr>
    `
    ),
    5
  );
  renderPager("universityTable", paged);
  bindDetailCells(DOM.universityTable);
}

function renderIrnUniversityTable(rows) {
  const sorted = sortRows("irnUniversityTable", rows);
  const paged = filterAndPaginateRows("irnUniversityTable", sorted, (row) =>
    `${row.name} ${row.country} ${row.publicationCount} ${row.lecturerNames.join(" ")} ${row.faculties.join(" ")}`
  );
  DOM.irnUniversityTable.innerHTML = renderRowsOrEmpty(
    paged.rows.map(
      (row) => `
      <tr>
        <td class="clickable-cell" data-detail-type="university" data-detail-key="${escapeAttribute(row.name)}">${escapeHtml(row.name)}</td>
        <td class="clickable-cell" data-detail-type="country" data-detail-key="${escapeAttribute(row.country)}">${escapeHtml(row.country)}</td>
        <td>${formatNumber(row.publicationCount)}</td>
        <td>${escapeHtml(row.lecturerNames.join(", ")) || '<span class="muted">No matched lecturers</span>'}</td>
        <td>${renderTagList(row.faculties)}</td>
      </tr>
    `
    ),
    5
  );
  renderPager("irnUniversityTable", paged);
  bindDetailCells(DOM.irnUniversityTable);
}

function renderCountryTable(rows) {
  const sorted = sortRows("countryTable", rows);
  const paged = filterAndPaginateRows("countryTable", sorted, (row) =>
    `${row.name} ${row.irnUniversities.join(" ")} ${row.universities.join(" ")} ${row.publicationCount}`
  );
  DOM.countryTable.innerHTML = renderRowsOrEmpty(
    paged.rows.map(
      (row) => `
      <tr>
        <td class="clickable-cell" data-detail-type="country" data-detail-key="${escapeAttribute(row.name)}">${escapeHtml(row.name)}</td>
        <td>${formatNumber(row.irnUniversityCount)}</td>
        <td>${formatNumber(row.publicationCount)}</td>
        <td>${escapeHtml(row.irnUniversities.join(", ")) || '<span class="muted">No IRN universities in current window</span>'}</td>
      </tr>
    `
    ),
    4
  );
  renderPager("countryTable", paged);
  bindDetailCells(DOM.countryTable);
}

function renderLecturerUniversityTable(rows) {
  const paged = filterAndPaginateRows("lecturerUniversityTable", rows, (row) =>
    `${row.university} ${row.country} ${row.lecturerCounts.join(" ")}`
  );
  DOM.lecturerUniversityTable.innerHTML = renderRowsOrEmpty(
    paged.rows.map(
      (row) => `
      <tr>
        <td class="clickable-cell" data-detail-type="university" data-detail-key="${escapeAttribute(row.university)}">${escapeHtml(row.university)}</td>
        <td class="clickable-cell" data-detail-type="country" data-detail-key="${escapeAttribute(row.country)}">${escapeHtml(row.country)}</td>
        <td>${escapeHtml(row.lecturerCounts.join(", ")) || '<span class="muted">No paper counts</span>'}</td>
      </tr>
    `
    ),
    3
  );
  renderPager("lecturerUniversityTable", paged);
  bindDetailCells(DOM.lecturerUniversityTable);
}

function renderLecturerCountryTable(rows) {
  const paged = filterAndPaginateRows("lecturerCountryTable", rows, (row) =>
    `${row.country} ${row.universities.join(" ")} ${row.lecturers.join(" ")}`
  );
  DOM.lecturerCountryTable.innerHTML = renderRowsOrEmpty(
    paged.rows.map(
      (row) => `
      <tr>
        <td class="clickable-cell" data-detail-type="country" data-detail-key="${escapeAttribute(row.country)}">${escapeHtml(row.country)}</td>
        <td>${escapeHtml(row.universities.join(", "))}</td>
        <td>${escapeHtml(row.lecturers.join(", "))}</td>
      </tr>
    `
    ),
    3
  );
  renderPager("lecturerCountryTable", paged);
  bindDetailCells(DOM.lecturerCountryTable);
}

function renderDataQualityTables(dataQuality) {
  const pagedAuthors = filterAndPaginateRows("unmatchedAuthorsTable", dataQuality.unmatchedAuthors, (row) =>
    `${row.authorId} ${row.occurrences}`
  );
  DOM.unmatchedAuthorsTable.innerHTML = renderRowsOrEmpty(
    pagedAuthors.rows.map(
      (row) => `
      <tr>
        <td>${escapeHtml(row.authorId)}</td>
        <td>${formatNumber(row.occurrences)}</td>
      </tr>
    `
    ),
    2
  );
  renderPager("unmatchedAuthorsTable", pagedAuthors);

  const pagedInstitutions = filterAndPaginateRows("unmatchedInstitutionsTable", dataQuality.unknownInstitutions, (row) =>
    `${row.institution} ${row.occurrences}`
  );
  DOM.unmatchedInstitutionsTable.innerHTML = renderRowsOrEmpty(
    pagedInstitutions.rows.map(
      (row) => `
      <tr>
        <td>${escapeHtml(row.institution)}</td>
        <td>${formatNumber(row.occurrences)}</td>
      </tr>
    `
    ),
    2
  );
  renderPager("unmatchedInstitutionsTable", pagedInstitutions);

  const pagedVariants = filterAndPaginateRows("duplicateVariantsTable", dataQuality.duplicateVariants, (row) =>
    `${row.canonical} ${row.variants.join(" ")}`
  );
  DOM.duplicateVariantsTable.innerHTML = renderRowsOrEmpty(
    pagedVariants.rows.map(
      (row) => `
      <tr>
        <td>${escapeHtml(row.canonical)}</td>
        <td>${escapeHtml(row.variants.join(" | "))}</td>
      </tr>
    `
    ),
    2
  );
  renderPager("duplicateVariantsTable", pagedVariants);
}

function renderDetailPanel(derived) {
  if (!state.detail) {
    DOM.detailPanel.className = "detail-body empty-state";
    DOM.detailPanel.textContent =
      "Click a university or country in the tables or charts to inspect yearly trends, involved faculties, and collaborating lecturers.";
    return;
  }

  DOM.detailPanel.className = "detail-body";

  if (state.detail.type === "university") {
    const row = derived.universityStats.find((item) => item.name === state.detail.key);
    if (!row) {
      DOM.detailPanel.className = "detail-body empty-state";
      DOM.detailPanel.textContent = "That university is outside the current filters.";
      return;
    }

    DOM.detailPanel.innerHTML = `
      <h4>${escapeHtml(row.name)}</h4>
      <div class="detail-meta">
        <div class="detail-stat"><span>Country</span><strong>${escapeHtml(row.country)}</strong></div>
        <div class="detail-stat"><span>Joint publications</span><strong>${formatNumber(row.publicationCount)}</strong></div>
        <div class="detail-stat"><span>IRN publications in window</span><strong>${formatNumber(row.irnPublicationCount)}</strong></div>
      </div>
      <p><strong>Collaborating lecturers:</strong> ${escapeHtml(row.lecturerNames.join(", ")) || "No matched lecturers"}</p>
      <p><strong>Involved faculties:</strong> ${escapeHtml(row.faculties.join(", ")) || "No matched faculties"}</p>
      <p><strong>Yearly trend:</strong> ${escapeHtml(row.yearlyCounts.map(([year, count]) => `${year}: ${count}`).join(" | ")) || "No yearly data"}</p>
      <p><strong>Paper counts per lecturer:</strong> ${escapeHtml(row.lecturerPaperCounts.map(([name, count]) => `${name} (${count})`).join(", ")) || "No lecturer counts"}</p>
    `;
    return;
  }

  const row = derived.countryStats.find((item) => item.name === state.detail.key);
  if (!row) {
    DOM.detailPanel.className = "detail-body empty-state";
    DOM.detailPanel.textContent = "That country is outside the current filters.";
    return;
  }

  DOM.detailPanel.innerHTML = `
    <h4>${escapeHtml(row.name)}</h4>
    <div class="detail-meta">
      <div class="detail-stat"><span>Partner universities</span><strong>${formatNumber(row.universities.length)}</strong></div>
      <div class="detail-stat"><span>IRN universities</span><strong>${formatNumber(row.irnUniversityCount)}</strong></div>
      <div class="detail-stat"><span>Joint publications</span><strong>${formatNumber(row.publicationCount)}</strong></div>
    </div>
    <p><strong>Universities:</strong> ${escapeHtml(row.universities.join(", ")) || "No partner universities"}</p>
    <p><strong>Collaborating lecturers:</strong> ${escapeHtml(row.lecturerNames.join(", ")) || "No matched lecturers"}</p>
    <p><strong>Yearly trend:</strong> ${escapeHtml(row.yearlyCounts.map(([year, count]) => `${year}: ${count}`).join(" | ")) || "No yearly data"}</p>
  `;
}

function bindDetailCells(container) {
  container.querySelectorAll("[data-detail-type]").forEach((element) => {
    element.addEventListener("click", () => {
      state.detail = {
        type: element.dataset.detailType,
        key: element.dataset.detailKey,
      };
      renderDetailPanel(state.derived);
    });
  });
}

function exportDataset(name) {
  if (!state.derived) {
    setStatus("Load the two Excel files before exporting.", "warning");
    return;
  }

  const exporters = {
    universities: () =>
      state.derived.universityStats.map((row) => ({
        University: row.name,
        Country: row.country,
        Publications: row.publicationCount,
        UBLecturers: row.lecturerNames.join("; "),
        Faculties: row.faculties.join("; "),
        IRNPublicationCount: row.irnPublicationCount,
        IRNQualified: row.irnQualified ? "Yes" : "No",
      })),
    partnerUniversities: () =>
      state.derived.universityStats.map((row) => ({
        University: row.name,
        Country: row.country,
        JointPublications: row.publicationCount,
        UBLecturers: row.lecturerNames.join("; "),
        Faculties: row.faculties.join("; "),
      })),
    countries: () =>
      state.derived.countryStats.map((row) => ({
        Country: row.name,
        Publications: row.publicationCount,
        PartnerUniversities: row.universities.join("; "),
        IRNUniversities: row.irnUniversities.join("; "),
        UBLecturers: row.lecturerNames.join("; "),
      })),
    irnCountries: () =>
      state.derived.countryStats
        .filter((row) => row.irnUniversityCount > 0)
        .map((row) => ({
          Country: row.name,
          IRNUniversityCount: row.irnUniversityCount,
          IRNUniversities: row.irnUniversities.join("; "),
          Publications: row.publicationCount,
        })),
    lecturerUniversity: () =>
      state.derived.lecturerUniversityRows.map((row) => ({
        University: row.university,
        Country: row.country,
        Lecturers: row.lecturers.join("; "),
        PapersPerLecturer: row.lecturerCounts.join("; "),
      })),
    lecturerCountry: () =>
      state.derived.lecturerCountryRows.map((row) => ({
        Country: row.country,
        Universities: row.universities.join("; "),
        Lecturers: row.lecturers.join("; "),
      })),
  };

  const rows = exporters[name]?.();
  if (!rows?.length) {
    setStatus("No rows available for export in the current filter state.", "warning");
    return;
  }

  const worksheet = XLSX.utils.json_to_sheet(rows);
  const csv = XLSX.utils.sheet_to_csv(worksheet);
  downloadFile(`${name}_${buildFilterStamp()}.csv`, csv, "text/csv;charset=utf-8;");
}

function sortRows(tableId, rows) {
  const sortState = state.sort[tableId];
  if (!sortState) {
    return rows;
  }

  const direction = sortState.direction === "asc" ? 1 : -1;
  return [...rows].sort((a, b) => {
    const aValue = a[sortState.key];
    const bValue = b[sortState.key];

    if (typeof aValue === "number" && typeof bValue === "number") {
      return (aValue - bValue) * direction;
    }
    return String(aValue).localeCompare(String(bValue)) * direction;
  });
}

function toggleSort(tableId, key) {
  const current = state.sort[tableId];
  if (!current) {
    state.sort[tableId] = { key, direction: "desc" };
    return;
  }

  if (current.key === key) {
    current.direction = current.direction === "asc" ? "desc" : "asc";
  } else {
    current.key = key;
    current.direction = key.toLowerCase().includes("name") || key === "country" ? "asc" : "desc";
  }
}

async function readBundledWorkbook(key) {
  const bundle = window.BUNDLED_WORKBOOKS?.[key];
  if (!bundle?.base64) {
    throw new Error(`Missing bundled workbook: ${key}`);
  }
  const binary = atob(bundle.base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i += 1) {
    bytes[i] = binary.charCodeAt(i);
  }
  const arrayBuffer = bytes.buffer;
  return XLSX.read(arrayBuffer, { type: "array" });
}

function normalizeScopusId(value) {
  const cleaned = String(value ?? "").trim();
  if (!cleaned) {
    return "";
  }
  if (/^\d+(\.0+)?$/.test(cleaned)) {
    return cleaned.replace(/\.0+$/, "");
  }
  return cleaned;
}

function normalizeInstitution(value) {
  const raw = normalizeText(value);
  if (!raw) {
    return "";
  }
  const lower = raw.toLowerCase();
  if (INSTITUTION_ALIAS_MAP[lower]) {
    return INSTITUTION_ALIAS_MAP[lower];
  }
  return toTitleCase(
    lower
      .replace(/\s+/g, " ")
      .replace(/\buniv\b/g, "university")
      .replace(/\bdept\b/g, "department")
      .replace(/\bfac\b/g, "faculty")
      .trim()
  );
}

function normalizeCountryValue(value) {
  const raw = normalizeText(value);
  if (!raw) {
    return "";
  }
  const lower = raw.toLowerCase();
  return COUNTRY_ALIAS_MAP[lower] || toTitleCase(lower);
}

function inferCountry(rawInstitution, canonicalInstitution) {
  if (COUNTRY_BY_INSTITUTION_ALIAS[canonicalInstitution]) {
    return COUNTRY_BY_INSTITUTION_ALIAS[canonicalInstitution];
  }

  const parts = normalizeText(rawInstitution)
    .split(",")
    .map((part) => part.trim().toLowerCase())
    .filter(Boolean);

  for (let i = parts.length - 1; i >= 0; i -= 1) {
    const candidate = parts[i];
    if (COUNTRY_ALIAS_MAP[candidate]) {
      return COUNTRY_ALIAS_MAP[candidate];
    }
    const matched = Object.keys(COUNTRY_ALIAS_MAP).find((key) => candidate.includes(key));
    if (matched) {
      return COUNTRY_ALIAS_MAP[matched];
    }
  }

  const canonicalLower = canonicalInstitution.toLowerCase();
  const matchedByInstitution = Object.keys(COUNTRY_ALIAS_MAP).find((key) => canonicalLower.includes(key));
  return matchedByInstitution ? COUNTRY_ALIAS_MAP[matchedByInstitution] : "";
}

function splitPipeValues(value) {
  return String(value ?? "")
    .split("|")
    .map((part) => part.trim())
    .filter(Boolean);
}

function normalizeText(value) {
  return String(value ?? "").replace(/\s+/g, " ").trim();
}

function buildPaperId(row, index) {
  const candidates = [
    row["EID"],
    row["Scopus EID"],
    row["DOI"],
    row["Title"],
    row["Document Title"],
  ]
    .map(normalizeText)
    .filter(Boolean);
  return candidates[0] || `paper-${index + 1}`;
}

function isUBInstitution(value) {
  return UB_NAME_PATTERNS.some((pattern) => pattern.test(value));
}

function dedupeByKey(items, key) {
  const seen = new Set();
  return items.filter((item) => {
    if (seen.has(item[key])) {
      return false;
    }
    seen.add(item[key]);
    return true;
  });
}

function mapToSortedPairs(map) {
  return Array.from(map.entries()).sort((a, b) => {
    if (typeof a[0] === "number" && typeof b[0] === "number") {
      return a[0] - b[0];
    }
    return String(a[0]).localeCompare(String(b[0]));
  });
}

function renderTagList(values) {
  if (!values.length) {
    return '<span class="muted">-</span>';
  }
  return `<div class="tag-list">${values
    .slice(0, 6)
    .map((value) => `<span class="tag">${escapeHtml(value)}</span>`)
    .join("")}${values.length > 6 ? `<span class="tag">+${values.length - 6}</span>` : ""}</div>`;
}

function renderRowsOrEmpty(rows, colspan) {
  return rows.length
    ? rows.join("")
    : `<tr><td colspan="${colspan}" class="muted">No records match the current filters.</td></tr>`;
}

function filterAndPaginateRows(tableId, rows, searchTextBuilder) {
  const ui = state.tableUi[tableId] || { page: 1, search: "" };
  state.tableUi[tableId] = ui;

  const filteredRows = ui.search
    ? rows.filter((row) => searchTextBuilder(row).toLowerCase().includes(ui.search))
    : rows;

  const totalPages = Math.max(1, Math.ceil(filteredRows.length / TABLE_PAGE_SIZE));
  ui.page = Math.min(ui.page, totalPages);
  const startIndex = (ui.page - 1) * TABLE_PAGE_SIZE;
  return {
    rows: filteredRows.slice(startIndex, startIndex + TABLE_PAGE_SIZE),
    totalRows: filteredRows.length,
    page: ui.page,
    totalPages,
    startIndex,
  };
}

function renderPager(tableId, pageState) {
  const pagerId = `pager${tableId.charAt(0).toUpperCase()}${tableId.slice(1)}`;
  const container = document.getElementById(pagerId);
  if (!container) {
    return;
  }

  const pageButtons = buildCompactPageList(pageState.page, pageState.totalPages).map((item) =>
    item === "ellipsis"
      ? '<span class="pager-ellipsis">...</span>'
      : `<button type="button" data-page-number="${item}" ${item === pageState.page ? 'disabled' : ''}>${item}</button>`
  );

  container.innerHTML = `
    <span>${pageState.totalRows ? `${pageState.startIndex + 1}-${Math.min(pageState.startIndex + TABLE_PAGE_SIZE, pageState.totalRows)}` : "0"} of ${pageState.totalRows}</span>
    <button type="button" data-page-action="prev" ${pageState.page <= 1 ? "disabled" : ""}>Prev</button>
    ${pageButtons.join("")}
    <button type="button" data-page-action="next" ${pageState.page >= pageState.totalPages ? "disabled" : ""}>Next</button>
  `;

  container.querySelectorAll("button").forEach((button) => {
    button.addEventListener("click", () => {
      const ui = state.tableUi[tableId];
      if (button.dataset.pageNumber) {
        ui.page = Number(button.dataset.pageNumber);
      } else {
        ui.page += button.dataset.pageAction === "prev" ? -1 : 1;
      }
      renderDashboard();
    });
  });
}

function buildCompactPageList(currentPage, totalPages) {
  if (totalPages <= 7) {
    return Array.from({ length: totalPages }, (_, index) => index + 1);
  }

  const pages = new Set([1, totalPages]);
  for (let page = currentPage - 1; page <= currentPage + 1; page += 1) {
    if (page > 1 && page < totalPages) {
      pages.add(page);
    }
  }

  if (currentPage <= 3) {
    pages.add(2);
    pages.add(3);
    pages.add(4);
  }

  if (currentPage >= totalPages - 2) {
    pages.add(totalPages - 1);
    pages.add(totalPages - 2);
    pages.add(totalPages - 3);
  }

  const sortedPages = Array.from(pages).filter((page) => page >= 1 && page <= totalPages).sort((a, b) => a - b);
  const compact = [];

  sortedPages.forEach((page, index) => {
    if (index > 0 && page - sortedPages[index - 1] > 1) {
      compact.push("ellipsis");
    }
    compact.push(page);
  });

  return compact;
}

function formatNumber(value) {
  return Number(value || 0).toLocaleString();
}

function toTitleCase(value) {
  return value.replace(/\w\S*/g, (word) => word.charAt(0).toUpperCase() + word.slice(1));
}

function getChart(id) {
  if (!state.charts[id]) {
    state.charts[id] = echarts.init(document.getElementById(id));
  }
  return state.charts[id];
}

function disposeCharts() {
  Object.values(state.charts).forEach((chart) => chart?.dispose());
  state.charts = {};
}

function setStatus(message, level) {
  DOM.loadStatus.textContent = message;
  const styles = {
    success: "rgba(31, 122, 74, 0.18)",
    error: "rgba(163, 63, 52, 0.18)",
    warning: "rgba(192, 108, 46, 0.2)",
    info: "rgba(255, 255, 255, 0.06)",
  };
  DOM.loadStatus.style.background = styles[level] || styles.info;
}

function showLoading(message) {
  DOM.loadingMessage.textContent = message;
  DOM.loadingOverlay.classList.remove("hidden");
}

function hideLoading() {
  DOM.loadingOverlay.classList.add("hidden");
}

function buildFilterStamp() {
  return `${state.filters.scope || "all"}_${state.filters.startYear || "min"}-${state.filters.endYear || "max"}`;
}

function downloadFile(filename, content, mimeType) {
  const blob = new Blob([content], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = filename;
  anchor.click();
  URL.revokeObjectURL(url);
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function escapeAttribute(value) {
  return escapeHtml(value).replace(/"/g, "&quot;");
}

init();
