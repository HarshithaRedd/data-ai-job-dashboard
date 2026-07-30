const state = {
  jobs: [],
  meta: null,
};

const elements = {
  lastUpdated: document.getElementById("last-updated"),
  totalJobs: document.getElementById("total-jobs"),
  newJobs: document.getElementById("new-jobs"),
  companiesWithJobs: document.getElementById("companies-with-jobs"),
  companiesChecked: document.getElementById("companies-checked"),
  resultCount: document.getElementById("result-count"),
  searchInput: document.getElementById("search-input"),
  companyFilter: document.getElementById("company-filter"),
  modeFilter: document.getElementById("mode-filter"),
  sortFilter: document.getElementById("sort-filter"),
  newOnlyFilter: document.getElementById("new-only-filter"),
  jobsTableBody: document.getElementById("jobs-table-body"),
  message: document.getElementById("message"),
  healthToggle: document.getElementById("health-toggle"),
  healthPanel: document.getElementById("health-panel"),
  healthList: document.getElementById("health-list"),
  healthOk: document.getElementById("health-ok"),
  healthZero: document.getElementById("health-zero"),
  healthError: document.getElementById("health-error"),
};

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function safeUrl(value) {
  try {
    const url = new URL(value);
    return ["http:", "https:"].includes(url.protocol) ? url.href : "";
  } catch {
    return "";
  }
}

function formatDate(value) {
  if (!value) return "See posting";
  const date = new Date(`${value}T00:00:00Z`);
  if (Number.isNaN(date.getTime())) return value;
  return new Intl.DateTimeFormat("en-US", {
    month: "short",
    day: "numeric",
    year: "numeric",
    timeZone: "UTC",
  }).format(date);
}

function formatTimestamp(value) {
  if (!value) return "Not available";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return value;
  return new Intl.DateTimeFormat("en-US", {
    month: "short",
    day: "numeric",
    year: "numeric",
    hour: "numeric",
    minute: "2-digit",
    timeZoneName: "short",
  }).format(date);
}

function populateCompanyFilter() {
  const companies = [...new Set(state.jobs.map((job) => job.company).filter(Boolean))]
    .sort((a, b) => a.localeCompare(b));
  const currentValue = elements.companyFilter.value;
  elements.companyFilter.innerHTML = '<option value="">All companies</option>';
  for (const company of companies) {
    const option = document.createElement("option");
    option.value = company;
    option.textContent = company;
    elements.companyFilter.append(option);
  }
  elements.companyFilter.value = companies.includes(currentValue) ? currentValue : "";
}

function getFilteredJobs() {
  const query = elements.searchInput.value.trim().toLowerCase();
  const company = elements.companyFilter.value;
  const mode = elements.modeFilter.value;
  const newOnly = elements.newOnlyFilter.checked;
  const sort = elements.sortFilter.value;

  const filtered = state.jobs.filter((job) => {
    const haystack = [
      job.job_title,
      job.company,
      job.location,
      job.work_mode,
      job.employment_type,
      job.description,
    ].join(" ").toLowerCase();

    return (
      (!query || haystack.includes(query)) &&
      (!company || job.company === company) &&
      (!mode || job.work_mode === mode) &&
      (!newOnly || Boolean(job.is_new))
    );
  });

  filtered.sort((a, b) => {
    if (sort === "company") {
      return `${a.company} ${a.job_title}`.localeCompare(`${b.company} ${b.job_title}`);
    }
    if (sort === "title") {
      return `${a.job_title} ${a.company}`.localeCompare(`${b.job_title} ${b.company}`);
    }
    if (Boolean(a.is_new) !== Boolean(b.is_new)) {
      return Number(Boolean(b.is_new)) - Number(Boolean(a.is_new));
    }
    return String(b.date_posted || "").localeCompare(String(a.date_posted || ""));
  });

  return filtered;
}

function renderJobs() {
  const jobs = getFilteredJobs();
  elements.resultCount.textContent = jobs.length.toLocaleString("en-US");
  elements.jobsTableBody.innerHTML = "";

  if (!jobs.length) {
    elements.message.textContent = state.jobs.length
      ? "No jobs match the current filters."
      : "No matching Data, AI, or analyst jobs were found in the latest run. Open Source Health to see which sites returned zero results or errors.";
    return;
  }

  elements.message.textContent = "";

  for (const job of jobs) {
    const row = document.createElement("tr");
    const url = safeUrl(job.application_url);
    const newBadge = job.is_new ? '<span class="new-pill">NEW</span>' : "";
    const employment = job.employment_type && job.employment_type !== "See posting"
      ? escapeHtml(job.employment_type)
      : "";

    row.innerHTML = `
      <td class="role-cell">
        <div class="role-title">${escapeHtml(job.job_title)} ${newBadge}</div>
        <span class="role-meta">${employment || escapeHtml(job.source || "Official career site")}</span>
      </td>
      <td>
        <strong>${escapeHtml(job.company)}</strong>
        <span class="company-meta">Verified detail page</span>
      </td>
      <td>${escapeHtml(job.location || "See posting")}</td>
      <td><span class="mode-pill">${escapeHtml(job.work_mode || "See posting")}</span></td>
      <td>${escapeHtml(formatDate(job.date_posted))}</td>
      <td>
        ${url ? `<a class="apply-button" href="${escapeHtml(url)}" target="_blank" rel="noopener noreferrer">View & Apply</a>` : "Unavailable"}
      </td>
    `;
    elements.jobsTableBody.append(row);
  }
}

function renderMeta() {
  const meta = state.meta || {};
  elements.lastUpdated.textContent = formatTimestamp(meta.last_updated);
  elements.totalJobs.textContent = Number(meta.total_jobs || state.jobs.length).toLocaleString("en-US");
  elements.newJobs.textContent = Number(meta.new_jobs || 0).toLocaleString("en-US");
  elements.companiesWithJobs.textContent = Number(meta.companies_with_jobs || 0).toLocaleString("en-US");
  elements.companiesChecked.textContent = Number(meta.companies_configured || 0).toLocaleString("en-US");
}

function renderHealth() {
  const health = Array.isArray(state.meta?.health) ? state.meta.health : [];
  const ok = health.filter((item) => item.status === "ok").length;
  const zero = health.filter((item) => item.status === "zero").length;
  const error = health.filter((item) => item.status === "error").length;
  elements.healthOk.textContent = ok;
  elements.healthZero.textContent = zero;
  elements.healthError.textContent = error;
  elements.healthList.innerHTML = "";

  const statusRank = { error: 0, zero: 1, ok: 2 };
  health
    .slice()
    .sort((a, b) => (statusRank[a.status] ?? 9) - (statusRank[b.status] ?? 9) || a.company.localeCompare(b.company))
    .forEach((item) => {
      const card = document.createElement("div");
      card.className = "health-item";
      const statusLabel = item.status === "ok" ? `${item.jobs_found} jobs` : item.status;
      const statusClass = item.status === "ok" ? "ok" : item.status === "error" ? "error" : "zero";
      card.innerHTML = `
        <div>
          <div class="health-item__name">${escapeHtml(item.company)}</div>
          <div class="health-item__meta">${escapeHtml(item.method || "auto")} · ${escapeHtml(item.duration_seconds || 0)}s${item.message ? ` · ${escapeHtml(item.message)}` : ""}</div>
        </div>
        <span class="status-pill status-pill--${statusClass}">${escapeHtml(statusLabel)}</span>
      `;
      elements.healthList.append(card);
    });
}

async function loadData() {
  try {
    const cacheBust = Date.now();
    const [jobsResponse, metaResponse] = await Promise.all([
      fetch(`data/jobs.json?v=${cacheBust}`, { cache: "no-store" }),
      fetch(`data/meta.json?v=${cacheBust}`, { cache: "no-store" }),
    ]);

    if (!jobsResponse.ok || !metaResponse.ok) {
      throw new Error(`Data request failed (${jobsResponse.status}/${metaResponse.status})`);
    }

    const [jobs, meta] = await Promise.all([jobsResponse.json(), metaResponse.json()]);
    state.jobs = Array.isArray(jobs) ? jobs : [];
    state.meta = meta && typeof meta === "object" ? meta : {};

    populateCompanyFilter();
    renderMeta();
    renderHealth();
    renderJobs();
  } catch (error) {
    elements.message.textContent = `Could not load dashboard data: ${error.message}`;
    console.error(error);
  }
}

for (const element of [
  elements.searchInput,
  elements.companyFilter,
  elements.modeFilter,
  elements.sortFilter,
  elements.newOnlyFilter,
]) {
  element.addEventListener("input", renderJobs);
  element.addEventListener("change", renderJobs);
}

elements.healthToggle.addEventListener("click", () => {
  const isHidden = elements.healthPanel.hasAttribute("hidden");
  if (isHidden) {
    elements.healthPanel.removeAttribute("hidden");
  } else {
    elements.healthPanel.setAttribute("hidden", "");
  }
  elements.healthToggle.setAttribute("aria-expanded", String(isHidden));
  elements.healthToggle.textContent = isHidden ? "Hide source health" : "View source health";
});

loadData();
setInterval(loadData, 5 * 60 * 1000);
