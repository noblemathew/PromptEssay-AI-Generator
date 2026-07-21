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
                <th>Location Name</th>
                <th>Address</th>
                <th>City</th>
                <th>State</th>
                <th>Country</th>
                <th>Phone</th>
                <th>Source URL</th>
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
    location: ["location_name", "location", "name", "branch", "locationname"],
    address:  ["address", "street", "streetaddress"],
    city:     ["city", "town"],
    state:    ["state", "province"],
    country:  ["country", "region"],
    phone:    ["phone", "phonenumber", "contactnumber"],
    src:      ["source_url", "sourceurl", "source", "src", "url", "website"]
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

  /**
   * Pulls JSON out of a string that may be:
   *   - plain JSON ("[{...}]" or "{...}")
   *   - JSON wrapped in a markdown code fence (```json ... ```),
   *     possibly with prose text before and/or after it — this is
   *     how chat/agent-style responses commonly answer.
   * Returns the parsed value, or null if nothing parseable was found.
   */
  function tryParseJson(str) {
    if (typeof str !== "string") return null;

    // 1) A ```json ... ``` (or plain ``` ... ```) fenced block anywhere in the text
    const fenced = str.match(/```(?:json)?\s*([\s\S]*?)```/i);
    if (fenced) {
      try { return JSON.parse(fenced[1].trim()); } catch { /* fall through */ }
    }

    // 2) The whole string is JSON on its own
    const trimmed = str.trim();
    if (trimmed.startsWith("{") || trimmed.startsWith("[")) {
      try { return JSON.parse(trimmed); } catch { /* fall through */ }
    }

    // 3) A bracket-balanced [ ... ] array embedded anywhere in the text,
    //    even without a code fence around it
    const start = str.indexOf("[");
    if (start !== -1) {
      let depth = 0;
      for (let i = start; i < str.length; i++) {
        if (str[i] === "[") depth++;
        else if (str[i] === "]") {
          depth--;
          if (depth === 0) {
            const candidate = str.slice(start, i + 1);
            try { return JSON.parse(candidate); } catch { /* keep looking */ }
            break;
          }
        }
      }
    }

    return null;
  }

  /**
   * Recursively searches the response for the first array made up of
   * objects. Also returns any leftover prose text found alongside the
   * JSON (e.g. a "Coverage Notice" some agents append) so it isn't
   * silently thrown away.
   */
  function findRowsDeep(node, depth = 0) {
    if (depth > 6 || node === null || node === undefined) return { rows: null, note: null };

    if (Array.isArray(node)) {
      if (node.length > 0 && typeof node[0] === "object" && node[0] !== null && !Array.isArray(node[0])) {
        return { rows: node, note: null };
      }
      return { rows: null, note: null };
    }

    if (typeof node === "string") {
      const parsed = tryParseJson(node);
      if (parsed) {
        const inner = findRowsDeep(parsed, depth + 1);
        if (inner.rows) {
          // Anything in the string after the JSON block is likely a trailing note
          const fenceEnd = node.lastIndexOf("```");
          let note = null;
          if (fenceEnd !== -1) {
            const after = node.slice(fenceEnd + 3).trim();
            if (after.length > 5) note = after.replace(/^-+\s*/, "").replace(/^#+\s*/, "");
          }
          return { rows: inner.rows, note };
        }
      }
      return { rows: null, note: null };
    }

    if (typeof node === "object") {
      const priorityKeys = ["locations", "data", "results", "rows", "items", "value", "output", "outputs"];
      for (const key of priorityKeys) {
        if (key in node) {
          const found = findRowsDeep(node[key], depth + 1);
          if (found.rows) return found;
        }
      }
      for (const key of Object.keys(node)) {
        const found = findRowsDeep(node[key], depth + 1);
        if (found.rows) return found;
      }
    }
    return { rows: null, note: null };
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
  let note = null;          // any leftover context text (e.g. "Coverage Notice")
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
    const found = findRowsDeep(payload);
    rows = found.rows;
    note = found.note;

    document.getElementById("resultsTitle").textContent = "Results for " + site + (rows ? ` (${rows.length} found)` : "");
    document.getElementById("resultsFor").textContent = site;

    if (rows && rows.length > 0) {
      rows.forEach(r => {
        const tr = document.createElement("tr");
        tr.innerHTML = `
          <td style="font-weight:600;">${pick(r, FIELD_MAP.location)}</td>
          <td>${pick(r, FIELD_MAP.address)}</td>
          <td>${pick(r, FIELD_MAP.city)}</td>
          <td>${pick(r, FIELD_MAP.state)}</td>
          <td>${pick(r, FIELD_MAP.country)}</td>
          <td>${pick(r, FIELD_MAP.phone)}</td>
          <td style="color:var(--accent);">${pick(r, FIELD_MAP.src)}</td>`;
        body.appendChild(tr);
      });

      if (note) {
        const noteBox = document.createElement("div");
        noteBox.style.cssText = "margin-top:14px; padding:12px 14px; background:var(--amber-dim); border-radius:8px; font-size:12.5px; color:#7a5a12; line-height:1.5;";
        noteBox.textContent = note;
        document.querySelector(".content .card").appendChild(noteBox);
      }
    } else if (payload && typeof payload === "object") {
      renderFallbackCard(payload);
    } else {
      showEmpty("The flow responded, but with something we couldn't read as data.");
    }
  }

  exportBtn.addEventListener("click", () => {
    let csv;
    if (rows && rows.length > 0) {
      const header = ["Location Name", "Address", "City", "State", "Country", "Phone", "Source URL"].map(csvEscape).join(",");
      const lines = rows.map(r => [
        pick(r, FIELD_MAP.location), pick(r, FIELD_MAP.address), pick(r, FIELD_MAP.city),
        pick(r, FIELD_MAP.state), pick(r, FIELD_MAP.country), pick(r, FIELD_MAP.phone), pick(r, FIELD_MAP.src)
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
