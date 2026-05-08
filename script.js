(() => {
  "use strict";

  const DEFAULT_CONFIG = {
    siteUrl: "https://trolleykw.sharepoint.com/sites/BusinessDevelopment",
    projectListName: "Project Schedule",
    costListName: "Cost Performance",
    scheduleListName: "Schedule Performance"
  };

  const state = {
    config: loadConfig(),
    projects: [],
    costs: [],
    schedules: []
  };

  document.addEventListener("DOMContentLoaded", () => {
    bindUi();
    setConfigInputs();
    loadDashboardData();
  });

  function bindUi() {
    document.querySelectorAll(".tab").forEach(button => {
      button.addEventListener("click", () => switchTab(button.dataset.tab));
    });

    document.getElementById("saveConfigBtn").addEventListener("click", () => {
      state.config = readConfigInputs();
      localStorage.setItem("bdDashboardConfig", JSON.stringify(state.config));
      loadDashboardData();
    });

    document.getElementById("refreshBtn").addEventListener("click", loadDashboardData);
    document.getElementById("projectSearch").addEventListener("input", renderProjects);
    document.getElementById("costSearch").addEventListener("input", renderCosts);
    document.getElementById("scheduleSearch").addEventListener("input", renderSchedules);
  }

  function loadConfig() {
    try {
      return { ...DEFAULT_CONFIG, ...JSON.parse(localStorage.getItem("bdDashboardConfig") || "{}") };
    } catch {
      return DEFAULT_CONFIG;
    }
  }

  function setConfigInputs() {
    document.getElementById("siteUrl").value = state.config.siteUrl;
    document.getElementById("projectListName").value = state.config.projectListName;
    document.getElementById("costListName").value = state.config.costListName;
    document.getElementById("scheduleListName").value = state.config.scheduleListName;
  }

  function readConfigInputs() {
    return {
      siteUrl: document.getElementById("siteUrl").value.trim().replace(/\/$/, ""),
      projectListName: document.getElementById("projectListName").value.trim(),
      costListName: document.getElementById("costListName").value.trim(),
      scheduleListName: document.getElementById("scheduleListName").value.trim()
    };
  }

  async function loadDashboardData() {
    setStatus("offline", "Connecting...");
    try {
      const [projects, costs, schedules] = await Promise.all([
        getSharePointListItems(state.config.projectListName),
        getSharePointListItems(state.config.costListName),
        getSharePointListItems(state.config.scheduleListName)
      ]);

      state.projects = projects.map(normalizeProjectItem);
      state.costs = costs.map(normalizeCostItem);
      state.schedules = schedules.map(normalizeScheduleItem);

      renderAll();
      setStatus("online", "Connected");
      document.getElementById("lastUpdated").textContent = new Date().toLocaleString();
    } catch (error) {
      console.error(error);
      setStatus("offline", "Connection failed");
      renderEmptyAll();
      alert(`Unable to load SharePoint data.\n\n${error.message}\n\nCheck: Site URL, list names, permissions, and column internal names.`);
    }
  }

  async function getSharePointListItems(listName) {
    const encodedList = listName.replace(/'/g, "''");
    const url = `${state.config.siteUrl}/_api/web/lists/getbytitle('${encodedList}')/items?$top=5000`;
    const response = await fetch(url, {
      method: "GET",
      credentials: "include",
      headers: {
        "Accept": "application/json;odata=nometadata"
      }
    });

    if (!response.ok) {
      const text = await response.text().catch(() => "");
      throw new Error(`${listName}: ${response.status} ${response.statusText}. ${text.slice(0, 250)}`);
    }

    const data = await response.json();
    return data.value || [];
  }

  function normalizeProjectItem(item) {
    return {
      id: pick(item, ["ActivityID", "Activity_x0020_ID", "Title", "ID"]),
      activityName: pick(item, ["ActivityName", "Activity_x0020_Name", "Title"]),
      duration: toNumber(pick(item, ["Duration"])),
      start: formatDate(pick(item, ["Start", "StartDate", "Start_x0020_Date"])),
      end: formatDate(pick(item, ["End", "EndDate", "End_x0020_Date"])),
      actualDate: formatDate(pick(item, ["ActualDate", "Actual_x0020_Date"])),
      status: pick(item, ["Status"]),
      percent: toNumber(pick(item, ["PercentComplete", "Percent_x0020_Complete", "Progress", "Pct", "pct"])) || inferPercent(pick(item, ["Status"])),
      notes: pick(item, ["Notes", "Description", "Comments"]),
      blocker: pick(item, ["Blocker", "Risk", "Issue"])
    };
  }

  function normalizeCostItem(item) {
    const planned = toNumber(pick(item, ["Planned", "PlannedCost", "Planned_x0020_Cost"]));
    const actual = toNumber(pick(item, ["Actual", "ActualCost", "Actual_x0020_Cost"]));
    const vo = toNumber(pick(item, ["VO", "VariationOrder", "Variation_x0020_Order"])) || 0;
    const total = toNumber(pick(item, ["Total"])) || actual + vo;
    const variance = toNumber(pick(item, ["Var", "Variance"])) || total - planned;
    const cpi = toNumber(pick(item, ["CPI"])) || (total ? planned / total : null);
    return {
      storeName: pick(item, ["StoreName", "Store_x0020_Name", "Title"]),
      storeType: pick(item, ["StoreType", "Store_x0020_Type"]),
      meter: pick(item, ["Meter"]),
      planned,
      actual,
      vo,
      total,
      variance,
      cpi
    };
  }

  function normalizeScheduleItem(item) {
    const plannedDate = pick(item, ["PlannedDate", "Planned_x0020_Date"]);
    const actualDate = pick(item, ["ActualDate", "Actual_x0020_Date"]);
    return {
      storeName: pick(item, ["StoreName", "Store_x0020_Name", "Title"]),
      storeType: pick(item, ["StoreType", "Store_x0020_Type"]),
      plannedDate: formatDate(plannedDate),
      actualDate: formatDate(actualDate),
      status: pick(item, ["Status"]),
      variance: pick(item, ["Var", "Variance", "Delay", "DelayDays", "Delay_x0020_Days"])
    };
  }

  function pick(item, names) {
    for (const name of names) {
      if (item[name] !== undefined && item[name] !== null && item[name] !== "") return item[name];
    }
    return "";
  }

  function toNumber(value) {
    if (value === null || value === undefined || value === "") return 0;
    const number = Number(String(value).replace(/,/g, ""));
    return Number.isFinite(number) ? number : 0;
  }

  function formatDate(value) {
    if (!value) return "";
    const parsed = new Date(value);
    if (Number.isNaN(parsed.getTime())) return value;
    return parsed.toLocaleDateString(undefined, { year: "numeric", month: "short", day: "numeric" });
  }

  function inferPercent(status) {
    const s = String(status || "").toLowerCase();
    if (s.includes("complete")) return 100;
    if (s.includes("progress")) return 50;
    return 0;
  }

  function renderAll() {
    renderKpis();
    renderProjects();
    renderCosts();
    renderSchedules();
  }

  function renderKpis() {
    const total = state.projects.length;
    const completed = state.projects.filter(p => /complete|done|closed/i.test(p.status)).length;
    const inProgress = state.projects.filter(p => /progress|finalizing|review/i.test(p.status)).length;
    const pending = state.projects.filter(p => /pending|hold|pre|contract/i.test(p.status)).length;

    setText("totalProjects", total);
    setText("completedProjects", completed);
    setText("inProgressProjects", inProgress);
    setText("pendingProjects", pending);
  }

  function renderProjects() {
    const container = document.getElementById("projectCards");
    const query = document.getElementById("projectSearch").value.toLowerCase();
    const projects = state.projects.filter(p => JSON.stringify(p).toLowerCase().includes(query));

    if (!projects.length) return renderEmpty(container);

    container.innerHTML = projects.map(p => `
      <article class="project-card">
        <h3>${escapeHtml(p.activityName || p.id || "Untitled Activity")}</h3>
        <span class="badge ${badgeClass(p.status)}">${escapeHtml(p.status || "No Status")}</span>
        <div class="progress-bar"><span style="width:${Math.max(0, Math.min(100, p.percent || 0))}%"></span></div>
        <div class="meta">
          <div><small>Activity ID</small>${escapeHtml(p.id)}</div>
          <div><small>Duration</small>${escapeHtml(p.duration)}</div>
          <div><small>Start</small>${escapeHtml(p.start)}</div>
          <div><small>End</small>${escapeHtml(p.end)}</div>
          <div><small>Actual Date</small>${escapeHtml(p.actualDate)}</div>
          <div><small>Progress</small>${escapeHtml(p.percent)}%</div>
        </div>
        ${p.notes ? `<p class="notes"><strong>Notes:</strong> ${escapeHtml(p.notes)}</p>` : ""}
        ${p.blocker ? `<p class="notes"><strong>Blocker:</strong> ${escapeHtml(p.blocker)}</p>` : ""}
      </article>
    `).join("");
  }

  function renderCosts() {
    const body = document.getElementById("costTableBody");
    const query = document.getElementById("costSearch").value.toLowerCase();
    const rows = state.costs.filter(c => JSON.stringify(c).toLowerCase().includes(query));
    if (!rows.length) return renderEmptyRow(body, 9);
    body.innerHTML = rows.map(c => `
      <tr>
        <td>${escapeHtml(c.storeName)}</td>
        <td>${escapeHtml(c.storeType)}</td>
        <td>${escapeHtml(c.meter)}</td>
        <td class="number">${formatNumber(c.planned)}</td>
        <td class="number">${formatNumber(c.actual)}</td>
        <td class="number">${formatNumber(c.vo)}</td>
        <td class="number">${formatNumber(c.total)}</td>
        <td class="number">${formatNumber(c.variance)}</td>
        <td class="number">${c.cpi ? Number(c.cpi).toFixed(2) : ""}</td>
      </tr>
    `).join("");
  }

  function renderSchedules() {
    const body = document.getElementById("scheduleTableBody");
    const query = document.getElementById("scheduleSearch").value.toLowerCase();
    const rows = state.schedules.filter(s => JSON.stringify(s).toLowerCase().includes(query));
    if (!rows.length) return renderEmptyRow(body, 6);
    body.innerHTML = rows.map(s => `
      <tr>
        <td>${escapeHtml(s.storeName)}</td>
        <td>${escapeHtml(s.storeType)}</td>
        <td>${escapeHtml(s.plannedDate)}</td>
        <td>${escapeHtml(s.actualDate)}</td>
        <td>${escapeHtml(s.status)}</td>
        <td>${escapeHtml(s.variance)}</td>
      </tr>
    `).join("");
  }

  function renderEmptyAll() {
    renderEmpty(document.getElementById("projectCards"));
    renderEmptyRow(document.getElementById("costTableBody"), 9);
    renderEmptyRow(document.getElementById("scheduleTableBody"), 6);
  }

  function renderEmpty(container) {
    container.innerHTML = document.getElementById("emptyTemplate").innerHTML;
  }

  function renderEmptyRow(body, colspan) {
    body.innerHTML = `<tr><td colspan="${colspan}"><div class="empty-state"><h3>No data loaded</h3><p>Refresh the dashboard after checking SharePoint configuration.</p></div></td></tr>`;
  }

  function badgeClass(status) {
    const s = String(status || "").toLowerCase();
    if (s.includes("complete") || s.includes("done") || s.includes("closed")) return "completed";
    if (s.includes("progress") || s.includes("finalizing") || s.includes("review")) return "progress";
    return "pending";
  }

  function switchTab(tabId) {
    document.querySelectorAll(".tab").forEach(t => t.classList.toggle("active", t.dataset.tab === tabId));
    document.querySelectorAll(".panel").forEach(p => p.classList.toggle("active", p.id === tabId));
  }

  function setStatus(className, text) {
    const el = document.getElementById("connectionStatus");
    el.className = `status ${className}`;
    el.textContent = text;
  }

  function setText(id, value) {
    document.getElementById(id).textContent = value;
  }

  function formatNumber(value) {
    if (value === null || value === undefined || value === "") return "";
    const number = Number(value);
    return Number.isFinite(number) ? number.toLocaleString() : escapeHtml(value);
  }

  function escapeHtml(value) {
    return String(value ?? "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }
})();
