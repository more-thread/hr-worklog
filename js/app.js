/* HR Weekly Report Generator — client-side, no backend.
   Built for the ONEHRMS weekly monitoring template:
   No. | Assigned To | Task Name | Module | Request Type | Progress |
   Started Development | Ended Development | Completed Date | Status
   Handles both the single-sheet template (qwe.xlsx) and the multi-sheet
   filled workbook (Summary / Previous Week Progress / Current Week Goals /
   Issues & Concerns). */

(function () {
  "use strict";

  // ---- Canonical fields + fuzzy header aliases -------------------------
  const FIELDS = {
    developer:     { label: "Assigned To (developer)", required: true,
      aliases: ["assigned to", "assignee", "developer", "dev", "owner", "name", "resource", "member", "person"] },
    task:          { label: "Task Name", required: false,
      aliases: ["task name", "task", "description", "item", "summary", "title", "work item", "activity"] },
    module:        { label: "Module / System", required: false,
      aliases: ["module", "system", "application", "app", "product", "project", "component"] },
    requestType:   { label: "Request Type", required: false,
      aliases: ["request type", "type", "task type", "category", "work type", "req type"] },
    progress:      { label: "Progress (workflow stage)", required: false,
      aliases: ["progress", "stage", "workflow", "dev stage", "step"] },
    status:        { label: "Status", required: false,
      aliases: ["status", "state", "task status"] },
    startedDate:   { label: "Started Development", required: false,
      aliases: ["started development", "start date", "started", "date started", "start"] },
    endedDate:     { label: "Ended Development", required: false,
      aliases: ["ended development", "end date", "ended", "date ended", "end"] },
    completedDate: { label: "Completed Date", required: false,
      aliases: ["completed date", "completed", "completion date", "date completed", "done date"] }
  };

  // ---- State -----------------------------------------------------------
  const state = {
    workbook: null,
    sheetName: "",
    rawRows: [],
    headers: [],
    mapping: {},
    rows: [],
    agg: null,
    charts: {}
  };

  const $ = (id) => document.getElementById(id);
  const css = (name) => getComputedStyle(document.documentElement).getPropertyValue(name).trim();

  // ---- Persisted API key ----------------------------------------------
  const savedKey = localStorage.getItem("hrwr_key");
  if (savedKey) { $("apiKey").value = savedKey; $("rememberKey").checked = true; }
  $("rememberKey").addEventListener("change", syncKey);
  $("apiKey").addEventListener("change", syncKey);
  function syncKey() {
    if ($("rememberKey").checked && $("apiKey").value.trim())
      localStorage.setItem("hrwr_key", $("apiKey").value.trim());
    else localStorage.removeItem("hrwr_key");
  }

  // ---- File upload -----------------------------------------------------
  const dz = $("dropzone"), fileInput = $("fileInput");
  dz.addEventListener("dragover", (e) => { e.preventDefault(); dz.classList.add("drag"); });
  dz.addEventListener("dragleave", () => dz.classList.remove("drag"));
  dz.addEventListener("drop", (e) => {
    e.preventDefault(); dz.classList.remove("drag");
    if (e.dataTransfer.files.length) handleFile(e.dataTransfer.files[0]);
  });
  fileInput.addEventListener("change", (e) => { if (e.target.files.length) handleFile(e.target.files[0]); });
  $("sampleBtn").addEventListener("click", loadSample);

  function notice(elId, msg, kind) {
    $(elId).innerHTML = msg ? `<div class="notice ${kind || "info"}">${msg}</div>` : "";
  }

  function handleFile(file) {
    $("fileName").textContent = file.name;
    notice("uploadNotice", "");
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: "array", cellDates: true });
        loadWorkbook(wb);
      } catch (err) {
        notice("uploadNotice", "Could not read that file: " + err.message, "err");
      }
    };
    reader.onerror = () => notice("uploadNotice", "Failed to read the file.", "err");
    reader.readAsArrayBuffer(file);
  }

  // ---- Workbook / sheet selection -------------------------------------
  function sheetToMatrix(ws) {
    return XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", blankrows: false, raw: true });
  }

  // score a candidate header row by how many known fields its cells match
  function headerScore(cells) {
    let score = 0;
    for (const key of Object.keys(FIELDS)) {
      const al = FIELDS[key].aliases;
      if (cells.some((c) => { const n = norm(c); return n && al.some((a) => n === a || n.includes(a) || a.includes(n)); }))
        score++;
    }
    return score;
  }

  // find the best header row within the first few rows of a matrix
  function detectHeaderRow(matrix) {
    let best = { idx: 0, score: -1 };
    for (let i = 0; i < Math.min(matrix.length, 8); i++) {
      const s = headerScore(matrix[i] || []);
      if (s > best.score) best = { idx: i, score: s };
    }
    return best;
  }

  function loadWorkbook(wb) {
    state.workbook = wb;
    // score every sheet, pick the one whose best header row matches the most fields
    let bestSheet = null;
    for (const name of wb.SheetNames) {
      const m = sheetToMatrix(wb.Sheets[name]);
      const h = detectHeaderRow(m);
      if (!bestSheet || h.score > bestSheet.score) bestSheet = { name, score: h.score, headerIdx: h.idx };
    }
    // sheet picker
    const picker = $("sheetPicker");
    picker.innerHTML = wb.SheetNames.map((n) => `<option value="${escapeAttr(n)}"${n === bestSheet.name ? " selected" : ""}>${escapeHtml(n)}</option>`).join("");
    $("sheetPickerWrap").classList.toggle("hidden", wb.SheetNames.length <= 1);
    picker.onchange = () => selectSheet(picker.value);

    selectSheet(bestSheet.name);
    $("mapCard").classList.remove("hidden");
    $("mapCard").scrollIntoView({ behavior: "smooth", block: "start" });
  }

  function selectSheet(name) {
    state.sheetName = name;
    const matrix = sheetToMatrix(state.workbook.Sheets[name]);
    const { idx } = detectHeaderRow(matrix);
    const headers = (matrix[idx] || []).map((h, i) => String(h).trim() || ("Column " + (i + 1)));
    const rows = [];
    for (let r = idx + 1; r < matrix.length; r++) {
      const cells = matrix[r] || [];
      if (!cells.some((c) => c !== "" && c != null)) continue;
      const obj = {};
      headers.forEach((h, i) => { obj[h] = cells[i] != null ? cells[i] : ""; });
      rows.push(obj);
    }
    state.headers = headers;
    state.rawRows = rows;
    autoMap();
    renderMapping();
    notice("uploadNotice", `Loaded <strong>${rows.length}</strong> rows from sheet “<strong>${escapeHtml(name)}</strong>”.`, "ok");
  }

  // ---- Auto column mapping --------------------------------------------
  function norm(s) { return String(s == null ? "" : s).toLowerCase().replace(/[_\-.]/g, " ").replace(/\s+/g, " ").trim(); }

  function autoMap() {
    const used = new Set();
    state.mapping = {};
    for (const key of Object.keys(FIELDS)) {
      const al = FIELDS[key].aliases;
      let best = "";
      for (const h of state.headers) { if (used.has(h)) continue; if (al.some((a) => norm(h) === a)) { best = h; break; } }
      if (!best) for (const h of state.headers) { if (used.has(h)) continue; const nh = norm(h); if (al.some((a) => nh.includes(a) || a.includes(nh))) { best = h; break; } }
      state.mapping[key] = best;
      if (best) used.add(best);
    }
  }

  function renderMapping() {
    const grid = $("mapGrid");
    grid.innerHTML = "";
    for (const key of Object.keys(FIELDS)) {
      const f = FIELDS[key];
      const opts = ['<option value="">— none —</option>'].concat(
        state.headers.map((h) => `<option value="${escapeAttr(h)}"${state.mapping[key] === h ? " selected" : ""}>${escapeHtml(h)}</option>`)
      ).join("");
      const div = document.createElement("div");
      div.className = "map-item";
      div.innerHTML = `<label>${escapeHtml(f.label)}${f.required ? ' <span class="req">*</span>' : ""}</label><select data-field="${key}">${opts}</select>`;
      grid.appendChild(div);
    }
    grid.querySelectorAll("select").forEach((sel) =>
      sel.addEventListener("change", (e) => { state.mapping[e.target.dataset.field] = e.target.value; }));
  }

  $("generateBtn").addEventListener("click", () => {
    if (!state.mapping.developer) { notice("mapNotice", "Please map the <strong>Assigned To</strong> column — it's required.", "err"); return; }
    notice("mapNotice", "");
    normalize();
    aggregate();
    renderReport();
    $("report").classList.remove("hidden");
    $("report").scrollIntoView({ behavior: "smooth", block: "start" });
  });

  // ---- Normalization ---------------------------------------------------
  function raw(row, key) { const h = state.mapping[key]; return h ? row[h] : ""; }
  function clean(v) { const s = v == null ? "" : String(v).trim(); return (s === "" || s.toUpperCase() === "NULL") ? "" : s; }

  function toDate(v) {
    if (v == null || v === "") return null;
    if (v instanceof Date && !isNaN(v)) return v;
    const s = String(v).trim();
    if (!s || s.toUpperCase() === "NULL") return null;
    const n = Number(s);
    if (isFinite(n) && n > 20000 && n < 80000) return new Date(Math.round((n - 25569) * 86400 * 1000)); // Excel serial → JS date
    const d = new Date(s);
    return isNaN(d) ? null : d;
  }

  // "LAST,FIRST,MIDDLE" -> "First Last" (title-cased)
  function fmtName(v) {
    const s = clean(v);
    if (!s) return "Unassigned";
    if (s.includes(",")) {
      const p = s.split(",").map((x) => x.trim()).filter(Boolean);
      const last = p[0] || "", first = p[1] || "";
      return titleCase((first + " " + last).trim() || s);
    }
    return s;
  }
  function titleCase(s) { return String(s).toLowerCase().replace(/\b([a-zà-ÿ])/g, (m) => m.toUpperCase()); }

  const STATUS_LABEL = { done: "Done Development", ongoing: "Ongoing", todo: "Not Yet Started" };
  function normStatus(v) {
    const s = norm(v);
    if (!s) return null;
    if (/done|complete|finish|closed|deploy|resolved|shipped/.test(s)) return "done";
    if (/ongoing|in progress|progress|wip|doing|active|started/.test(s)) return "ongoing";
    if (/not yet|not started|todo|to do|backlog|new|pending|assigned|planned/.test(s)) return "todo";
    return "todo";
  }
  const isRework = (progress) => /fail/i.test(progress || "");
  const isPaused = (progress) => /paus|hold/i.test(progress || "");

  function normalize() {
    state.rows = state.rawRows.map((r) => {
      const status = normStatus(raw(r, "status")) || "todo";
      const progress = clean(raw(r, "progress"));
      return {
        developer: fmtName(raw(r, "developer")),
        task: clean(raw(r, "task")),
        module: clean(raw(r, "module")) || "—",
        requestType: (clean(raw(r, "requestType")) || "Uncategorized").toUpperCase(),
        progress,
        status,
        started: toDate(raw(r, "startedDate")),
        ended: toDate(raw(r, "endedDate")),
        completed: toDate(raw(r, "completedDate")),
        rework: isRework(progress),
        paused: isPaused(progress)
      };
    }).filter((r) => r.developer && r.developer !== "Unassigned" || r.task);
  }

  // ---- Aggregation -----------------------------------------------------
  function aggregate() {
    const rows = state.rows;
    const devs = {}, types = {}, statusTotals = { done: 0, ongoing: 0, todo: 0 };
    const rework = [], paused = [];

    for (const r of rows) {
      statusTotals[r.status]++;
      const d = devs[r.developer] || (devs[r.developer] = { name: r.developer, total: 0, done: 0, ongoing: 0, todo: 0, rework: 0 });
      d.total++; d[r.status]++; if (r.rework) d.rework++;
      const t = types[r.requestType] || (types[r.requestType] = { name: r.requestType, total: 0, done: 0 });
      t.total++; if (r.status === "done") t.done++;
      if (r.rework) rework.push(r);
      if (r.paused) paused.push(r);
    }

    const devList = Object.values(devs)
      .map((d) => ({ ...d, pctDone: d.total ? Math.round((d.done / d.total) * 100) : 0 }))
      .sort((a, b) => b.total - a.total);
    const typeList = Object.values(types).sort((a, b) => b.total - a.total);

    const totalTasks = rows.length;
    const avgLoad = devList.length ? totalTasks / devList.length : 0;
    const overloaded = devList.filter((d) => avgLoad > 0 && d.total > avgLoad * 1.4);
    const notStartedHeavy = devList.filter((d) => d.todo >= 3).sort((a, b) => b.todo - a.todo);

    state.agg = {
      devList, typeList, statusTotals, totalTasks, avgLoad, overloaded, notStartedHeavy,
      rework, paused,
      completionRate: totalTasks ? Math.round((statusTotals.done / totalTasks) * 100) : 0,
      goals: readGoals(), issues: readIssues(), period: readPeriod()
    };
  }

  // ---- Optional sheets: Current Week Goals / Issues & Concerns ---------
  function findSheet(re) {
    if (!state.workbook) return null;
    const name = state.workbook.SheetNames.find((n) => re.test(n));
    return name ? sheetToMatrix(state.workbook.Sheets[name]) : null;
  }
  function matrixToObjs(m) {
    if (!m || !m.length) return [];
    const { idx } = detectHeaderRow(m);
    const headers = (m[idx] || []).map((h) => String(h).trim());
    const out = [];
    for (let r = idx + 1; r < m.length; r++) {
      const cells = m[r] || [];
      if (!cells.some((c) => clean(c))) continue;
      const o = {}; headers.forEach((h, i) => { o[h] = cells[i]; }); out.push(o);
    }
    return out;
  }
  function pick(o, re) { const k = Object.keys(o).find((k) => re.test(k)); return k ? clean(o[k]) : ""; }

  function readGoals() {
    const m = findSheet(/goal|next week|coming week/i);
    return matrixToObjs(m).map((o) => ({
      task: pick(o, /task|goal|item|name/i),
      assignee: fmtName(pick(o, /assign|owner|dev/i)),
      start: (() => { const d = toDate(pick(o, /start|tentative|date/i)); return d ? d.toLocaleDateString() : ""; })()
    })).filter((g) => g.task || (g.assignee && g.assignee !== "Unassigned"));
  }
  function readIssues() {
    const m = findSheet(/issue|concern|risk/i);
    return matrixToObjs(m).map((o) => ({
      issue: pick(o, /issue|concern|risk|problem/i),
      recommendation: pick(o, /recommend|action|resolution|mitigat/i)
    })).filter((x) => x.issue || x.recommendation);
  }

  // reporting period: find "Start Date"/"End Date" cells anywhere in the workbook
  function firstDate(cands) { for (const c of cands) { const d = toDate(c); if (d) return d; } return null; }
  function readPeriod() {
    if (!state.workbook) return null;
    for (const name of state.workbook.SheetNames) {
      const m = sheetToMatrix(state.workbook.Sheets[name]);
      let start = null, end = null;
      for (let i = 0; i < m.length; i++) {
        const row = m[i] || [];
        for (let j = 0; j < row.length; j++) {
          const n = norm(row[j]);
          if (n === "start date") start = start || firstDate([row[j + 1], (m[i + 1] || [])[j]]);
          if (n === "end date") end = end || firstDate([row[j + 1], (m[i + 1] || [])[j]]);
        }
      }
      if (start && end) return { start, end };
    }
    return null;
  }

  // ---- Render report ---------------------------------------------------
  function renderReport() {
    const a = state.agg;
    const fmtD = (d) => d.toLocaleDateString(undefined, { month: "short", day: "numeric" });
    const period = a.period
      ? `Reporting period ${fmtD(a.period.start)} – ${a.period.end.toLocaleDateString(undefined, { month: "short", day: "numeric", year: "numeric" })} · `
      : "";
    $("reportMeta").textContent =
      `${period}${a.totalTasks} tasks · ${a.devList.length} team members · ${a.typeList.length} request type(s) · generated ${new Date().toLocaleDateString(undefined, { year: "numeric", month: "long", day: "numeric" })}`;
    renderKPIs(a);
    renderCharts(a);
    renderDevTable(a);
    renderGoalsIssues(a);
    renderAnalysis(builtInAnalysis(a), "Built-in analysis (no AI). Add an API key above and click “Generate with Claude” for a written narrative.");
  }

  function renderKPIs(a) {
    const cards = [
      { num: a.totalTasks, lbl: "Total tasks", cls: "" },
      { num: a.completionRate + "%", lbl: "Done development", cls: "good" },
      { num: a.statusTotals.ongoing, lbl: "Ongoing", cls: "" },
      { num: a.statusTotals.todo, lbl: "Not yet started", cls: a.statusTotals.todo ? "warn" : "good" },
      { num: a.rework.length, lbl: "Failed QA / rework", cls: a.rework.length ? "bad" : "good" },
      { num: a.devList.length, lbl: "Team members", cls: "" }
    ];
    $("kpiGrid").innerHTML = cards.map((c) => `<div class="kpi ${c.cls}"><div class="num">${c.num}</div><div class="lbl">${c.lbl}</div></div>`).join("");
  }

  function renderDevTable(a) {
    const rows = a.devList.map((d) => `
      <tr>
        <td>${escapeHtml(d.name)}</td>
        <td class="num">${d.total}</td>
        <td><span class="pill done">${d.done} done</span> <span class="pill progress">${d.ongoing} ongoing</span> <span class="pill todo">${d.todo} not started</span>${d.rework ? ` <span class="pill blocked">${d.rework} rework</span>` : ""}</td>
        <td><div style="display:flex;align-items:center;gap:8px"><div class="bar-mini" style="flex:1"><span style="width:${d.pctDone}%"></span></div><span class="num" style="min-width:34px">${d.pctDone}%</span></div></td>
      </tr>`).join("");
    $("devTable").innerHTML = `
      <thead><tr><th>Developer</th><th class="num">Tasks</th><th>Status split</th><th style="min-width:150px">% Done</th></tr></thead>
      <tbody>${rows}</tbody>`;
  }

  // ---- Charts ----------------------------------------------------------
  function renderCharts(a) {
    Object.values(state.charts).forEach((c) => c.destroy());
    state.charts = {};
    Chart.defaults.color = css("--text-muted") || "#5b6b7c";
    Chart.defaults.font.family = "-apple-system, BlinkMacSystemFont, Segoe UI, Roboto, sans-serif";
    const gridColor = css("--border") || "#e2e5ea";
    const C = { done: css("--good"), ongoing: css("--accent"), todo: css("--warn"), rework: css("--bad"), muted: css("--text-muted") };

    const grid = $("chartGrid");
    grid.innerHTML = "";
    const defs = [
      { id: "status", title: "Overall status", capId: "cap_status" },
      { id: "tasks", title: "Tasks per developer (by status)", capId: "cap_tasks" },
      { id: "type", title: "Tasks by request type (total vs done)", capId: "cap_type" },
      { id: "completion", title: "Completion by developer", capId: "cap_completion" }
    ];
    for (const d of defs) {
      const box = document.createElement("div");
      box.className = "chart-box";
      box.innerHTML = `<h3>${escapeHtml(d.title)}</h3><div class="chart-canvas-holder"><canvas id="chart_${d.id}"></canvas></div><div class="cap" id="${d.capId}"></div>`;
      grid.appendChild(box);
    }

    const st = a.statusTotals;
    state.charts.status = new Chart($("chart_status"), {
      type: "doughnut",
      data: { labels: ["Done Development", "Ongoing", "Not Yet Started"],
        datasets: [{ data: [st.done, st.ongoing, st.todo], backgroundColor: [C.done, C.ongoing, C.todo], borderWidth: 0 }] },
      options: { responsive: true, maintainAspectRatio: false, cutout: "62%",
        plugins: { legend: { position: "bottom", labels: { boxWidth: 12, padding: 12, font: { size: 11 } } } } }
    });

    const devNames = a.devList.map((d) => d.name);
    state.charts.tasks = new Chart($("chart_tasks"), {
      type: "bar",
      data: { labels: devNames, datasets: [
        { label: "Done", data: a.devList.map((d) => d.done), backgroundColor: C.done },
        { label: "Ongoing", data: a.devList.map((d) => d.ongoing), backgroundColor: C.ongoing },
        { label: "Not started", data: a.devList.map((d) => d.todo), backgroundColor: C.todo }
      ] },
      options: baseOpts(gridColor, { scales: { x: { stacked: true, grid: { color: gridColor } }, y: { stacked: true, beginAtZero: true, grid: { color: gridColor } } } })
    });

    const types = a.typeList;
    state.charts.type = new Chart($("chart_type"), {
      type: "bar",
      data: { labels: types.map((t) => t.name), datasets: [
        { label: "Total", data: types.map((t) => t.total), backgroundColor: C.muted },
        { label: "Done", data: types.map((t) => t.done), backgroundColor: C.done }
      ] },
      options: baseOpts(gridColor)
    });

    const byPct = a.devList.slice().sort((x, y) => x.pctDone - y.pctDone);
    state.charts.completion = new Chart($("chart_completion"), {
      type: "bar",
      data: { labels: byPct.map((d) => d.name), datasets: [{ label: "% done", data: byPct.map((d) => d.pctDone),
        backgroundColor: byPct.map((d) => d.pctDone >= 75 ? C.done : d.pctDone >= 40 ? C.ongoing : C.todo) }] },
      options: baseOpts(gridColor, { indexAxis: "y", plugins: { legend: { display: false } },
        scales: { x: { beginAtZero: true, max: 100, grid: { color: gridColor } }, y: { grid: { color: gridColor } } } })
    });
  }
  function baseOpts(gridColor, extra) {
    return Object.assign({
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { labels: { boxWidth: 12, boxHeight: 12, font: { size: 11 } } } },
      scales: { x: { grid: { color: gridColor } }, y: { grid: { color: gridColor }, beginAtZero: true } }
    }, extra || {});
  }
  function setCaptions(caps) {
    ["status", "tasks", "type", "completion"].forEach((k) => {
      const el = $("cap_" + k); if (el && caps && caps[k]) el.textContent = caps[k];
    });
  }

  // ---- Goals & Issues rendering ---------------------------------------
  function renderGoalsIssues(a) {
    if (a.goals && a.goals.length) {
      $("goalsCard").classList.remove("hidden");
      $("goalsBody").innerHTML = "<ul>" + a.goals.map((g) =>
        `<li>${escapeHtml(g.task || "(unnamed)")}${g.assignee && g.assignee !== "Unassigned" ? ` — <strong>${escapeHtml(g.assignee)}</strong>` : ""}${g.start ? ` <span class="muted">(start ${escapeHtml(g.start)})</span>` : ""}</li>`).join("") + "</ul>";
    } else $("goalsCard").classList.add("hidden");

    if (a.issues && a.issues.length) {
      $("issuesCard").classList.remove("hidden");
      $("issuesBody").innerHTML = "<ul>" + a.issues.map((x) =>
        `<li>${escapeHtml(x.issue || "")}${x.recommendation ? ` <span class="muted">→ ${escapeHtml(x.recommendation)}</span>` : ""}</li>`).join("") + "</ul>";
    } else $("issuesCard").classList.add("hidden");
  }

  // ---- Built-in (non-AI) analysis -------------------------------------
  function builtInAnalysis(a) {
    const s = a.statusTotals;
    const summary =
      `The team logged ${a.totalTasks} tasks across ${a.devList.length} developer(s) this week — ${s.done} done development (${a.completionRate}%), ${s.ongoing} ongoing, and ${s.todo} not yet started. ` +
      (a.rework.length ? `${a.rework.length} task(s) failed QA and need rework. ` : `No tasks failed QA. `) +
      (a.typeList.length ? `Most work was ${a.typeList[0].name.toLowerCase()} (${a.typeList[0].total} task(s)).` : "");

    const recs = [];
    if (a.overloaded.length) recs.push(`Rebalance workload — ${a.overloaded.map((d) => d.name).join(", ")} carried well above the team average (${a.avgLoad.toFixed(1)} tasks/dev).`);
    if (a.rework.length) recs.push(`Prioritize ${a.rework.length} failed-QA task(s) for rework before starting new items.`);
    if (a.paused.length) recs.push(`Resume ${a.paused.length} paused task(s) or confirm they are intentionally on hold.`);
    if (a.notStartedHeavy.length) recs.push(`Kick off backlog — ${a.notStartedHeavy.slice(0, 3).map((d) => `${d.name} (${d.todo} not started)`).join(", ")}.`);
    if (!recs.length) recs.push("Workload and throughput look balanced — maintain the current cadence.");

    const risks = [];
    a.rework.slice(0, 6).forEach((r) => risks.push(`Failed QA: “${short(r.task)}” (${r.developer})`));
    a.paused.slice(0, 4).forEach((r) => risks.push(`Paused: “${short(r.task)}” (${r.developer})`));
    a.overloaded.forEach((d) => risks.push(`Heavy load: ${d.name} has ${d.total} tasks (team avg ${a.avgLoad.toFixed(1)}).`));

    const caps = {
      status: `${a.completionRate}% of tasks reached Done Development; ${s.ongoing} ongoing, ${s.todo} not yet started.`,
      tasks: `Task load and status split across ${a.devList.length} developer(s).`,
      type: `Volume by request type — grey is total, green is completed.`,
      completion: `Each developer's share of tasks reaching Done Development.`
    };
    return { summary, recommendations: recs, risks, captions: caps };
  }
  function short(s) { s = String(s || "").trim(); return s.length > 90 ? s.slice(0, 90) + "…" : (s || "(unnamed task)"); }

  // ---- Analysis rendering ---------------------------------------------
  function renderAnalysis(data, sourceLabel) {
    $("aiSourceLabel").textContent = sourceLabel || "";
    const list = (arr) => arr && arr.length ? `<ul>${arr.map((x) => `<li>${escapeHtml(x)}</li>`).join("")}</ul>` : `<p class="muted">None flagged.</p>`;
    $("aiContent").innerHTML = `
      <div class="ai-block"><h3><span class="badge summary">Summary</span> Weekly summary</h3><div class="ai-body"><p>${escapeHtml(data.summary || "")}</p></div></div>
      <div class="ai-block"><h3><span class="badge recs">Recommend</span> Improvement recommendations</h3><div class="ai-body">${list(data.recommendations)}</div></div>
      <div class="ai-block"><h3><span class="badge risks">Risks</span> Risk &amp; blocker callouts</h3><div class="ai-body">${list(data.risks)}</div></div>`;
    setCaptions(data.captions);
  }

  // ---- Claude integration (bring-your-own-key) ------------------------
  $("aiBtn").addEventListener("click", generateWithClaude);
  async function generateWithClaude() {
    const key = $("apiKey").value.trim();
    if (!key) { notice("aiNotice", 'No API key set. Open <strong>Claude AI settings</strong> above and paste your Anthropic key, or keep using the built-in analysis.', "err"); $("settingsBox").open = true; return; }
    if (!state.agg) return;
    const btn = $("aiBtn"); btn.disabled = true; const original = btn.innerHTML;
    btn.innerHTML = '<span class="spinner"></span> Asking Claude…'; notice("aiNotice", "");
    try {
      const data = await callClaude(key, $("model").value, state.agg);
      renderAnalysis(data, `Written by ${$("model").value} · for HR stakeholders.`);
      notice("aiNotice", "Analysis generated by Claude.", "ok");
    } catch (err) {
      notice("aiNotice", "Claude request failed: " + escapeHtml(err.message) + " — showing built-in analysis instead.", "err");
      renderAnalysis(builtInAnalysis(state.agg), "Built-in analysis (Claude call failed).");
    } finally { btn.disabled = false; btn.innerHTML = original; }
  }

  function buildContext(a) {
    return {
      totals: { tasks: a.totalTasks, developers: a.devList.length, doneDevelopment: a.statusTotals.done,
        ongoing: a.statusTotals.ongoing, notYetStarted: a.statusTotals.todo, completedPct: a.completionRate,
        failedQaRework: a.rework.length, paused: a.paused.length },
      perDeveloper: a.devList.map((d) => ({ name: d.name, tasks: d.total, done: d.done, ongoing: d.ongoing, notStarted: d.todo, rework: d.rework, pctDone: d.pctDone })),
      byRequestType: a.typeList.map((t) => ({ type: t.name, total: t.total, done: t.done })),
      failedQaTasks: a.rework.slice(0, 15).map((r) => ({ task: short(r.task), developer: r.developer, module: r.module, stage: r.progress })),
      pausedTasks: a.paused.slice(0, 15).map((r) => ({ task: short(r.task), developer: r.developer })),
      currentWeekGoals: (a.goals || []).slice(0, 20),
      issuesRaised: (a.issues || []).slice(0, 20)
    };
  }

  async function callClaude(key, model, agg) {
    const ctx = buildContext(agg);
    const prompt =
`You are an HR operations analyst writing a weekly status report for HR stakeholders (non-technical managers).
The data below covers the workload, tasks, and progress of an internal development team working on HR-system projects (ONEHRMS) for the previous week, as JSON.
Statuses are: "Done Development", "Ongoing", "Not Yet Started". "failedQaRework" tasks failed QA and need rework. Tasks are grouped by request type (system enhancement, bug, initiative, support, group task).

Write a concise, professional report. Respond with ONLY a JSON object (no markdown, no code fences) with exactly these keys:
- "summary": string, 2-4 sentences summarizing throughput, workload balance and standouts, in plain business language.
- "recommendations": array of 3-6 short, specific, actionable strings (rebalancing workload, clearing failed-QA rework, resuming paused work, starting the not-yet-started backlog, at-risk request types).
- "risks": array of strings flagging rework, paused work, overloaded developers, or large not-started backlogs worth a manager's attention (empty array if none).
- "captions": object with keys "status", "tasks", "type", "completion" — each a single short sentence describing that chart for a stakeholder.

Data:
${JSON.stringify(ctx, null, 2)}`;

    const res = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: { "content-type": "application/json", "x-api-key": key, "anthropic-version": "2023-06-01", "anthropic-dangerous-direct-browser-access": "true" },
      body: JSON.stringify({ model, max_tokens: 2048, messages: [{ role: "user", content: prompt }] })
    });
    if (!res.ok) {
      let detail = res.status + " " + res.statusText;
      try { const j = await res.json(); if (j.error && j.error.message) detail = j.error.message; } catch (_) {}
      if (res.status === 401) detail = "Invalid API key (401).";
      throw new Error(detail);
    }
    const json = await res.json();
    const text = (json.content || []).filter((b) => b.type === "text").map((b) => b.text).join("\n");
    return parseModelJson(text);
  }

  function parseModelJson(text) {
    let t = text.trim().replace(/^```(?:json)?/i, "").replace(/```$/, "").trim();
    const first = t.indexOf("{"), last = t.lastIndexOf("}");
    if (first !== -1 && last !== -1) t = t.slice(first, last + 1);
    const obj = JSON.parse(t);
    return {
      summary: obj.summary || "",
      recommendations: Array.isArray(obj.recommendations) ? obj.recommendations : [],
      risks: Array.isArray(obj.risks) ? obj.risks : [],
      captions: obj.captions || {}
    };
  }

  // ---- Sample data (mirrors the ONEHRMS template) ---------------------
  function loadSample() {
    const devs = ["DELA CRUZ,JUAN,SANTOS", "REYES,MARIA,LOPEZ", "TAN,PAOLO,GARCIA", "SANTOS,ANGELA,CRUZ"];
    const types = ["SYSTEM ENHANCEMENT", "SYSTEM BUG", "INITIATIVE", "SUPPORT", "GROUP TASK"];
    const stages = { done: ["FOR QA", "PASSED BY QA AND FOR SRA REVIEW", "END", "FAILED BY QA"], ongoing: ["RESUME", "TASK GRABBED", "PAUSE"], todo: ["ASSIGNED", "TASK GRABBED"] };
    const statuses = ["Done Development", "Done Development", "Ongoing", "Not Yet Started", "Ongoing"];
    const rows = []; let id = 1;
    const wsToSerial = (d) => Math.round(d.getTime() / 86400000 + 25569);
    for (const dev of devs) {
      const n = 4 + Math.floor(Math.random() * 6);
      for (let i = 0; i < n; i++) {
        const status = statuses[Math.floor(Math.random() * statuses.length)];
        const key = status === "Done Development" ? "done" : status === "Ongoing" ? "ongoing" : "todo";
        const stage = stages[key][Math.floor(Math.random() * stages[key].length)];
        const start = new Date(); start.setDate(start.getDate() - Math.floor(Math.random() * 6));
        rows.push({
          "No.": id,
          "Assigned To": dev,
          "Task Name": `000${100 + id}-0726-00${(id % 3) + 1} Form: SAMPLE TASK ${id} — enhancement / fix work item description.`,
          "Module": Math.random() > 0.4 ? "ONEHRMS" : "NULL",
          "Request Type": types[Math.floor(Math.random() * types.length)],
          "Progress": stage,
          "Started Development": key !== "todo" ? wsToSerial(start) : "NULL",
          "Ended Development": key === "done" ? wsToSerial(new Date()) : "NULL",
          "Completed Date": key === "done" && stage === "END" ? wsToSerial(new Date()) : "NULL",
          "Status": status
        });
        id++;
      }
    }
    const goals = [
      { "No.": 1, "Task Name": "Continue HeyBuzz media handling", "Assigned To": "DELA CRUZ,JUAN,SANTOS", "Tentative Start Date": wsToSerial(new Date()) },
      { "No.": 2, "Task Name": "Finish Probationary Evaluation processing", "Assigned To": "REYES,MARIA,LOPEZ", "Tentative Start Date": "" },
      { "No.": 3, "Task Name": "Banned Words maintenance in HeyBuzz", "Assigned To": "TAN,PAOLO,GARCIA", "Tentative Start Date": wsToSerial(new Date()) }
    ];
    const issues = [
      { "No.": 1, "Issues & Concerns": "ABS server event logs fill up quickly, blocking user access.", "Recommendation": "Add token-expiry validation and rotate logs on a schedule." }
    ];
    const summary = [{ "Start Date": wsToSerial(new Date()), "End Date": wsToSerial(new Date()) }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(summary), "Summary");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), "Previous Week Progress");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(goals), "Current Week Goals");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(issues), "Issues & Concerns");
    $("fileName").textContent = "sample-onehrms-week.xlsx (generated)";
    loadWorkbook(wb);
  }

  // ---- Print & escaping ------------------------------------------------
  $("printBtn").addEventListener("click", () => window.print());
  function escapeHtml(s) { return String(s == null ? "" : s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c])); }
  function escapeAttr(s) { return escapeHtml(s); }
})();
