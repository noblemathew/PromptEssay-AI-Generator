<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Extraction Results · Location Discovery Portal</title>
<link rel="stylesheet" href="css/style.css">
</head>
<body>
<script src="js/auth.js"></script>
<script>requireAuth();</script>

<div class="shell">
  <aside class="sidebar">
    <div class="brand">
      <div class="mark">LD</div>
      <div class="name">Location Discovery<small>Portal</small></div>
    </div>
    <nav>
      <div class="nav-group">
        <a class="nav-link" href="home.html"><span class="ico">&#8962;</span> Home</a>
        <a class="nav-link active" href="results.html"><span class="ico">&#9776;</span> Extraction Results</a>
        <a class="nav-link" href="workflow.html"><span class="ico">&#8635;</span> Workflow</a>
        <a class="nav-link" href="agent.html"><span class="ico">&#9881;</span> Agent Control</a>
      </div>
    </nav>
    <div class="sidebar-foot">
      <a class="logout-link" href="#" onclick="logout(); return false;"><span class="ico">&#8630;</span> Sign out</a>
    </div>
  </aside>

  <div class="main">
    <div class="topbar">
      <div>
        <h2>Extraction Results</h2>
        <div class="crumbs" id="resultsFor">Results</div>
      </div>
      <div class="user"><div class="avatar">U</div> Signed in</div>
    </div>

    <div class="content">
      <div class="card">
        <div class="card-head">
          <div>
            <h3 id="resultsTitle">Results</h3>
            <p>Locations, titles and contact details found for the selected site</p>
          </div>
          <div style="display:flex; gap:8px;">
            <span class="pill success">Extraction Complete</span>
            <button class="btn primary" id="exportBtn">Export CSV</button>
          </div>
        </div>

        <div style="overflow-x:auto;">
          <table>
            <thead>
              <tr>
                <th>Location</th>
                <th>Title</th>
                <th>Country</th>
                <th>Phone</th>
                <th>Source URL</th>
                <th>Status</th>
              </tr>
            </thead>
            <tbody id="resultsBody"></tbody>
          </table>
        </div>
        <div class="empty" id="resultsEmpty" style="display:none;">
          No extraction has been run yet. Go to <a href="home.html" style="color:var(--accent); font-weight:600;">Home</a> and select a website to extract.
        </div>
      </div>
    </div>
  </div>
</div>

<div class="toast" id="toast"></div>

<script>
  /* -----------------------------------------------------------
     Flexible field mapping — tries several likely key names per
     column, case-insensitively, since the flow's exact field
     names may vary.
     ----------------------------------------------------------- */
  const FIELD_MAP = {
    location: ["location", "name", "branch", "locationname", "locationName"],
    title:    ["title", "role", "label"],
    country:  ["country", "region"],
    phone:    ["phone", "phonenumber", "phoneNumber", "contactnumber", "contactNumber"],
    src:      ["sourceurl", "sourceUrl", "source", "src", "url", "website"],
    status:   ["status", "confidence"]
  };

  function pick(row, keys) {
    const lowerMap = {};
    Object.keys(row).forEach(k => { lowerMap[k.toLowerCase()] = row[k]; });
    for (const k of keys) {
      const v = lowerMap[k.toLowerCase()];
      if (v !== undefined && v !== null && v !== "") return v;
    }
    return "—";
  }

  /** Try to JSON.parse a string; return null if it isn't valid JSON. */
  function tryParseJson(str) {
    if (typeof str !== "string") return null;
    const trimmed = str.trim();
    if (!(trimmed.startsWith("{") || trimmed.startsWith("["))) return null;
    try { return JSON.parse(trimmed); } catch { return null; }
  }

  /**
   * Recursively searches the response for the first array made up of
   * objects (as opposed to an array of strings/numbers). Handles two
   * common real-world shapes: the array nested a few levels deep
   * (e.g. { output: { locations: [...] } }), and the array encoded
   * as a JSON *string* inside a field (common with chat/agent-style
   * responses, e.g. { text: "[{...}, {...}]" }).
   */
  function findRowsDeep(node, depth = 0) {
    if (depth > 6 || node === null || node === undefined) return null;

    if (Array.isArray(node)) {
      if (node.length > 0 && typeof node[0] === "object" && node[0] !== null && !Array.isArray(node[0])) {
        return node;
      }
      return null;
    }

    if (typeof node === "string") {
      const parsed = tryParseJson(node);
      if (parsed) return findRowsDeep(parsed, depth + 1);
      return null;
    }

    if (typeof node === "object") {
      // Prefer keys that sound like a results list
      const priorityKeys = ["locations", "data", "results", "rows", "items", "value", "output", "outputs"];
      for (const key of priorityKeys) {
        if (key in node) {
          const found = findRowsDeep(node[key], depth + 1);
          if (found) return found;
        }
      }
      for (const key of Object.keys(node)) {
        const found = findRowsDeep(node[key], depth + 1);
        if (found) return found;
      }
    }
    return null;
  }

  /** Flatten a single object's own properties into a flat {key: value} map for display/export when no row list was found. */
  function flattenObject(obj, prefix = "") {
    const out = {};
    Object.entries(obj).forEach(([k, v]) => {
      const key = prefix ? `${prefix}.${k}` : k;
      if (v !== null && typeof v === "object" && !Array.isArray(v)) {
        Object.assign(out, flattenObject(v, key));
      } else {
        out[key] = Array.isArray(v) ? JSON.stringify(v) : v;
      }
    });
    return out;
  }

  function csvEscape(val) {
    const s = String(val ?? "");
    if (/[",\n]/.test(s)) return '"' + s.replace(/"/g, '""') + '"';
    return s;
  }

  const site = localStorage.getItem("ldp_last_extraction");
  const rawData = localStorage.getItem("ldp_last_extraction_data");
  const body = document.getElementById("resultsBody");
  const empty = document.getElementById("resultsEmpty");
  const table = document.querySelector("table");
  const statusBadge = document.querySelector(".pill.success");
  const exportBtn = document.getElementById("exportBtn");

  let rows = null;          // array of row objects, when we found one
  let flatFallback = null;  // { key: value } map, when we only got a single object
  let payload = null;

  function showEmpty(message) {
    table.style.display = "none";
    statusBadge.style.display = "none";
    empty.style.display = "block";
    empty.textContent = message;
    document.getElementById("resultsTitle").textContent = "No results yet";
    document.getElementById("resultsFor").textContent = "Run an extraction from Home first";
  }

  function renderFallbackCard(obj) {
    flatFallback = flattenObject(obj);
    table.style.display = "none";
    statusBadge.textContent = "Response received";
    statusBadge.className = "pill amber";

    const wrap = document.createElement("div");
    wrap.innerHTML = `
      <p style="font-size:13px; color:var(--text-dim); margin:0 0 14px;">
        The flow responded, but it didn't look like a list of multiple locations —
        here's everything it returned. You can still export this as CSV below.
      </p>`;

    const kvCard = document.createElement("div");
    kvCard.style.cssText = "border:1px solid var(--border); border-radius:8px; overflow:hidden;";
    const kvTable = document.createElement("table");
    kvTable.innerHTML = "<tbody></tbody>";
    const tbody = kvTable.querySelector("tbody");
    Object.entries(flatFallback).forEach(([k, v]) => {
      const tr = document.createElement("tr");
      tr.innerHTML = `<td style="font-weight:600; width:220px; color:var(--text-dim);">${k}</td><td>${String(v)}</td>`;
      tbody.appendChild(tr);
    });
    kvCard.appendChild(kvTable);
    wrap.appendChild(kvCard);

    empty.innerHTML = "";
    empty.appendChild(wrap);
    empty.style.display = "block";
  }

  if (!site || !rawData) {
    showEmpty("No extraction has been run yet. Go to Home and select a website to extract.");
  } else {
    payload = JSON.parse(rawData);
    rows = findRowsDeep(payload);

    document.getElementById("resultsTitle").textContent = "Results for " + site;
    document.getElementById("resultsFor").textContent = site;

    if (rows && rows.length > 0) {
      rows.forEach(r => {
        const tr = document.createElement("tr");
        const statusVal = pick(r, FIELD_MAP.status);
        const pillClass = /verif|high|confirm/i.test(String(statusVal)) ? "success" : "amber";
        tr.innerHTML = `
          <td>${pick(r, FIELD_MAP.location)}</td>
          <td>${pick(r, FIELD_MAP.title)}</td>
          <td>${pick(r, FIELD_MAP.country)}</td>
          <td>${pick(r, FIELD_MAP.phone)}</td>
          <td style="color:var(--accent);">${pick(r, FIELD_MAP.src)}</td>
          <td><span class="pill ${pillClass}">${statusVal === "—" ? "Received" : statusVal}</span></td>`;
        body.appendChild(tr);
      });
    } else if (payload && typeof payload === "object") {
      renderFallbackCard(payload);
    } else {
      showEmpty("The flow responded, but with something we couldn't read as data.");
    }
  }

  exportBtn.addEventListener("click", () => {
    let csv;
    if (rows && rows.length > 0) {
      const header = ["Location", "Title", "Country", "Phone", "Source URL", "Status"].map(csvEscape).join(",");
      const lines = rows.map(r => [
        pick(r, FIELD_MAP.location), pick(r, FIELD_MAP.title), pick(r, FIELD_MAP.country),
        pick(r, FIELD_MAP.phone), pick(r, FIELD_MAP.src), pick(r, FIELD_MAP.status)
      ].map(csvEscape).join(","));
      csv = [header, ...lines].join("\n");
    } else if (flatFallback) {
      const header = "Field,Value";
      const lines = Object.entries(flatFallback).map(([k, v]) => [csvEscape(k), csvEscape(v)].join(","));
      csv = [header, ...lines].join("\n");
    } else {
      showToast("Nothing to export yet — run an extraction first");
      return;
    }

    const blob = new Blob([csv], { type: "text/csv" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = "extraction_results.csv";
    a.click();
    showToast("CSV downloaded");
  });

  function showToast(msg) {
    const t = document.getElementById("toast");
    t.textContent = msg;
    t.classList.add("show");
    setTimeout(() => t.classList.remove("show"), 2600);
  }
</script>
</body>
</html>
