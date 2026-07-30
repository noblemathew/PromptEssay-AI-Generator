/* ===========================================================
   server.js — One Portal backend (SQL Server edition).
   =========================================================== */

import "dotenv/config";
import express from "express";
import cors from "cors";
import cookieParser from "cookie-parser";
import path from "path";
import { fileURLToPath } from "url";
import { exec } from "child_process";
import { promisify } from "util";

import { query } from "./db.js";
import { registerAuthRoutes, requireSession } from "./auth.js";
import { ConfidentialClientApplication } from "@azure/msal-node";

const execAsync = promisify(exec);
const __dirname = path.dirname(fileURLToPath(import.meta.url));

const app = express();
const PORT = process.env.PORT || 3000;
const FRONTEND_ORIGIN = process.env.FRONTEND_ORIGIN || "http://localhost:3000";

app.use(cors({ origin: FRONTEND_ORIGIN, credentials: true }));
app.use(express.json());
app.use(cookieParser());

registerAuthRoutes(app);

function callerManagerId(user) {
  return user.role === "manager" ? user.id : user.m_id;
}

/* -----------------------------------------------------------
   [NO SQL — HARDCODED] In-memory data used in place of the
   database while SQL access is unavailable. This resets every
   time the server restarts — it's not persisted anywhere.
   Matches the hardcoded admin identity in auth.js (id: 1).
   ----------------------------------------------------------- */
let HARDCODED_TOOLS = [
  { id: 1, slug: "webscraper",      name: "Location Discovery",      description: "Extracts business location data from company websites" },
  { id: 2, slug: "interview-agent", name: "Interview Assessment",    description: "Runs a structured technical interview and evaluates the transcript" },
  { id: 3, slug: "ai-advisor",      name: "AI Advisor",              description: "Chat-based assistant for quick questions and guidance" },
  { id: 4, slug: "doc-summarizer",  name: "Document Summarizer",     description: "Summarizes long documents into key points" }
];

let HARDCODED_PROJECTS = [
  { id: 1, name: "Location Mapping Pilot",  m_id: 1 },
  { id: 2, name: "Candidate Screening Q3",  m_id: 1 },
  { id: 3, name: "Client Advisory Desk",    m_id: 1 },
  { id: 4, name: "Policy Document Review",  m_id: 1 }
];

// Every project gets all 4 tools, for now.
let HARDCODED_PROJECT_TOOLS = HARDCODED_PROJECTS.flatMap(p =>
  HARDCODED_TOOLS.map(t => ({ project_id: p.id, tool_id: t.id }))
);

let HARDCODED_EMPLOYEES = [
  { id: 1, email: "priya.s@example.com",  display_name: "Priya Sharma",  m_id: 1 },
  { id: 2, email: "james.k@example.com",  display_name: "James Kim",     m_id: 1 },
  { id: 3, email: "sofia.m@example.com",  display_name: "Sofia Martins", m_id: 1 },
  { id: 4, email: "arjun.n@example.com",  display_name: "Arjun Nair",    m_id: 1 }
];

let nextProjectId = 5;
let nextEmployeeId = 5;

/* ---------------- Projects ---------------- */

/* [SQL — DISABLED, no DB permission] Original query. Restore by
   uncommenting and removing the hardcoded version below it.

app.get("/api/projects", requireSession, async (req, res) => {
  const managerId = callerManagerId(req.user);
  const rows = await query(
    `SELECT p.id, p.name,
            (SELECT COUNT(*) FROM project_tools pt WHERE pt.project_id = p.id) AS tool_count
     FROM projects p
     WHERE p.m_id = @managerId
     ORDER BY p.id DESC`,
    { managerId }
  );
  res.json(rows);
});
*/

// [NO SQL — HARDCODED]
app.get("/api/projects", requireSession, async (req, res) => {
  const managerId = callerManagerId(req.user);
  const rows = HARDCODED_PROJECTS
    .filter(p => p.m_id === managerId)
    .map(p => ({
      id: p.id,
      name: p.name,
      tool_count: HARDCODED_PROJECT_TOOLS.filter(pt => pt.project_id === p.id).length
    }))
    .sort((a, b) => b.id - a.id);
  res.json(rows);
});

/* [SQL — DISABLED, no DB permission] Original query.

app.get("/api/projects/:id", requireSession, async (req, res) => {
  const managerId = callerManagerId(req.user);
  const rows = await query(
    "SELECT * FROM projects WHERE id = @id AND m_id = @managerId",
    { id: req.params.id, managerId }
  );
  if (rows.length === 0) return res.status(404).json({ error: "Project not found" });
  res.json(rows[0]);
});
*/

// [NO SQL — HARDCODED]
app.get("/api/projects/:id", requireSession, async (req, res) => {
  const managerId = callerManagerId(req.user);
  const project = HARDCODED_PROJECTS.find(p => String(p.id) === String(req.params.id) && p.m_id === managerId);
  if (!project) return res.status(404).json({ error: "Project not found" });
  res.json(project);
});

/* [SQL — DISABLED, no DB permission] Original query.

app.post("/api/projects", requireSession, async (req, res) => {
  if (req.user.role !== "manager") return res.status(403).json({ error: "Only managers can create projects" });
  const { name } = req.body || {};
  if (!name?.trim()) return res.status(400).json({ error: "Project name is required" });

  const rows = await query(
    "INSERT INTO projects (name, m_id) OUTPUT INSERTED.id AS id VALUES (@name, @managerId)",
    { name: name.trim(), managerId: req.user.id }
  );
  res.status(201).json({ id: rows[0].id, name: name.trim() });
});
*/

// [NO SQL — HARDCODED]
app.post("/api/projects", requireSession, async (req, res) => {
  if (req.user.role !== "manager") return res.status(403).json({ error: "Only managers can create projects" });
  const { name } = req.body || {};
  if (!name?.trim()) return res.status(400).json({ error: "Project name is required" });

  const project = { id: nextProjectId++, name: name.trim(), m_id: req.user.id };
  HARDCODED_PROJECTS.push(project);
  res.status(201).json({ id: project.id, name: project.name });
});

/* [SQL — DISABLED, no DB permission] Original query.

app.delete("/api/projects/:id", requireSession, async (req, res) => {
  if (req.user.role !== "manager") return res.status(403).json({ error: "Only managers can delete projects" });

  const rows = await query(
    "SELECT * FROM projects WHERE id = @id AND m_id = @managerId",
    { id: req.params.id, managerId: req.user.id }
  );
  if (rows.length === 0) return res.status(404).json({ error: "Project not found" });

  await query("DELETE FROM project_tools WHERE project_id = @id", { id: req.params.id });
  await query("DELETE FROM projects WHERE id = @id", { id: req.params.id });
  res.status(204).send();
});
*/

// [NO SQL — HARDCODED]
app.delete("/api/projects/:id", requireSession, async (req, res) => {
  if (req.user.role !== "manager") return res.status(403).json({ error: "Only managers can delete projects" });

  const project = HARDCODED_PROJECTS.find(p => String(p.id) === String(req.params.id) && p.m_id === req.user.id);
  if (!project) return res.status(404).json({ error: "Project not found" });

  HARDCODED_PROJECT_TOOLS = HARDCODED_PROJECT_TOOLS.filter(pt => String(pt.project_id) !== String(req.params.id));
  HARDCODED_PROJECTS = HARDCODED_PROJECTS.filter(p => String(p.id) !== String(req.params.id));
  res.status(204).send();
});

/* ---------------- Tools (per project) ---------------- */

/* [SQL — DISABLED, no DB permission] Original query.

app.get("/api/projects/:id/tools", requireSession, async (req, res) => {
  const managerId = callerManagerId(req.user);

  const owns = await query("SELECT id FROM projects WHERE id = @id AND m_id = @managerId", { id: req.params.id, managerId });
  if (owns.length === 0) return res.status(404).json({ error: "Project not found" });

  const tools = await query(
    `SELECT t.id, t.slug, t.name, t.description
     FROM project_tools pt
     JOIN tools t ON t.id = pt.tool_id
     WHERE pt.project_id = @id
     ORDER BY t.id`,
    { id: req.params.id }
  );
  res.json(tools);
});
*/

// [NO SQL — HARDCODED]
app.get("/api/projects/:id/tools", requireSession, async (req, res) => {
  const managerId = callerManagerId(req.user);
  const owns = HARDCODED_PROJECTS.find(p => String(p.id) === String(req.params.id) && p.m_id === managerId);
  if (!owns) return res.status(404).json({ error: "Project not found" });

  const toolIds = HARDCODED_PROJECT_TOOLS.filter(pt => String(pt.project_id) === String(req.params.id)).map(pt => pt.tool_id);
  const tools = HARDCODED_TOOLS.filter(t => toolIds.includes(t.id));
  res.json(tools);
});

/* ---------------- Employees (team members) ---------------- */

/* [SQL — DISABLED, no DB permission] Original query.

app.get("/api/employees", requireSession, async (req, res) => {
  if (req.user.role !== "manager") return res.status(403).json({ error: "Only managers can view team lists" });
  const rows = await query("SELECT id, email, display_name FROM employees WHERE m_id = @managerId ORDER BY id", { managerId: req.user.id });
  res.json(rows);
});
*/

// [NO SQL — HARDCODED]
app.get("/api/employees", requireSession, async (req, res) => {
  if (req.user.role !== "manager") return res.status(403).json({ error: "Only managers can view team lists" });
  const rows = HARDCODED_EMPLOYEES.filter(e => e.m_id === req.user.id);
  res.json(rows);
});

/* [SQL — DISABLED, no DB permission] Original query.

app.post("/api/employees", requireSession, async (req, res) => {
  if (req.user.role !== "manager") return res.status(403).json({ error: "Only managers can add team members" });
  const { email, display_name } = req.body || {};
  if (!email?.trim()) return res.status(400).json({ error: "Email is required" });

  const existing = await query("SELECT id FROM employees WHERE email = @email", { email: email.trim() });
  if (existing.length > 0) return res.status(409).json({ error: "That email is already added" });

  // New employees need a password too — a manager-set temporary one for now,
  // since there's no self-service invite flow yet. Worth swapping for a
  // proper "set your own password on first login" flow later.
  const bcrypt = (await import("bcryptjs")).default;
  const tempPasswordHash = await bcrypt.hash("ChangeMe123!", 10);

  const rows = await query(
    `INSERT INTO employees (email, display_name, password_hash, m_id)
     OUTPUT INSERTED.id AS id
     VALUES (@email, @displayName, @passwordHash, @managerId)`,
    { email: email.trim(), displayName: display_name || null, passwordHash: tempPasswordHash, managerId: req.user.id }
  );
  res.status(201).json({ id: rows[0].id, email: email.trim(), temp_password: "ChangeMe123!" });
});
*/

// [NO SQL — HARDCODED]
app.post("/api/employees", requireSession, async (req, res) => {
  if (req.user.role !== "manager") return res.status(403).json({ error: "Only managers can add team members" });
  const { email, display_name } = req.body || {};
  if (!email?.trim()) return res.status(400).json({ error: "Email is required" });

  const existing = HARDCODED_EMPLOYEES.find(e => e.email === email.trim());
  if (existing) return res.status(409).json({ error: "That email is already added" });

  const employee = { id: nextEmployeeId++, email: email.trim(), display_name: display_name || null, m_id: req.user.id };
  HARDCODED_EMPLOYEES.push(employee);
  res.status(201).json({ id: employee.id, email: employee.email, temp_password: "admin123" });
});

/* [SQL — DISABLED, no DB permission] Original query.

app.delete("/api/employees/:id", requireSession, async (req, res) => {
  if (req.user.role !== "manager") return res.status(403).json({ error: "Only managers can remove team members" });

  const rows = await query("SELECT * FROM employees WHERE id = @id AND m_id = @managerId", { id: req.params.id, managerId: req.user.id });
  if (rows.length === 0) return res.status(404).json({ error: "Team member not found" });

  await query("DELETE FROM employees WHERE id = @id", { id: req.params.id });
  res.status(204).send();
});
*/

// [NO SQL — HARDCODED]
app.delete("/api/employees/:id", requireSession, async (req, res) => {
  if (req.user.role !== "manager") return res.status(403).json({ error: "Only managers can remove team members" });

  const employee = HARDCODED_EMPLOYEES.find(e => String(e.id) === String(req.params.id) && e.m_id === req.user.id);
  if (!employee) return res.status(404).json({ error: "Team member not found" });

  HARDCODED_EMPLOYEES = HARDCODED_EMPLOYEES.filter(e => String(e.id) !== String(req.params.id));
  res.status(204).send();
});

/* -----------------------------------------------------------
   Webscraper tool proxy — unchanged from the original build.
   ----------------------------------------------------------- */

const FLOW_URL = process.env.POWER_AUTOMATE_URL;
const TOKEN_RESOURCE = process.env.TOKEN_RESOURCE || "https://service.flow.microsoft.com/";

async function getFlowToken() {
  const { stdout } = await execAsync(`az account get-access-token --resource ${TOKEN_RESOURCE}`);
  const parsed = JSON.parse(stdout);
  if (!parsed.accessToken) throw new Error("No accessToken in az CLI output");
  return parsed.accessToken;
}

// [NO SQL] Talks only to Power Automate — no database access.
app.post("/api/extract", requireSession, async (req, res) => {
  const { company, url } = req.body || {};
  if (!company || !url) return res.status(400).json({ error: "Both 'company' and 'url' are required." });
  if (!FLOW_URL) return res.status(500).json({ error: "Server is missing POWER_AUTOMATE_URL — set it in server/.env" });

  let token;
  try {
    token = await getFlowToken();
  } catch (err) {
    console.error("Token fetch failed:", err.message);
    return res.status(401).json({
      error: "Could not get an Azure AD token. Run `az login` in this server's terminal, then try again.",
      details: err.message
    });
  }

  try {
    const flowRes = await fetch(FLOW_URL, {
      method: "POST",
      headers: { "Content-Type": "application/json", "Authorization": `Bearer ${token}` },
      body: JSON.stringify({ company, url })
    });

    const text = await flowRes.text();
    let body;
    try { body = JSON.parse(text); } catch { body = { raw: text }; }

    if (!flowRes.ok) return res.status(502).json({ error: `Flow responded with ${flowRes.status}`, details: body });
    res.json(body);
  } catch (err) {
    console.error(err);
    res.status(502).json({ error: "Could not reach the Power Automate flow", details: String(err) });
  }
});

app.get("/api/health", async (req, res) => {
  /* [SQL — DISABLED, no DB permission] Original check.
  let dbOk = false;
  try { await query("SELECT 1 AS ok"); dbOk = true; } catch { leave false }
  */
  res.json({ ok: true, dbOk: "not applicable — running without a database", flowConfigured: Boolean(FLOW_URL) });
});

/* -----------------------------------------------------------
   Interview Assessment tool proxy — same auth pattern as the
   webscraper's /api/extract above (reuses getFlowToken()), just
   pointed at this tool's own Power Automate flow. The tool's
   own page sends whatever payload it already builds (candidate
   details, question bank, transcript, etc.) — forwarded as-is.
   NOTE: this route does NOT touch SQL at all.
   ----------------------------------------------------------- */

const INTERVIEW_FLOW_URL = process.env.INTERVIEW_FLOW_URL;

// [NO SQL] Talks only to Power Automate — no database access.
app.post("/api/interview-evaluate", requireSession, async (req, res) => {
  if (!INTERVIEW_FLOW_URL) {
    return res.status(500).json({ error: "Server is missing INTERVIEW_FLOW_URL — set it in server/.env" });
  }

  let token;
  try {
    token = await getFlowToken();
  } catch (err) {
    console.error("Token fetch failed:", err.message);
    return res.status(401).json({
      error: "Could not get an Azure AD token. Run `az login` in this server's terminal, then try again.",
      details: err.message
    });
  }

  try {
    const flowRes = await fetch(INTERVIEW_FLOW_URL, {
      method: "POST",
      headers: { "Content-Type": "application/json", "Authorization": `Bearer ${token}` },
      body: JSON.stringify(req.body)
    });

    const text = await flowRes.text();
    let body;
    try { body = JSON.parse(text); } catch { body = { raw: text }; }

    if (!flowRes.ok) return res.status(502).json({ error: `Flow responded with ${flowRes.status}`, details: body });
    res.json(body);
  } catch (err) {
    console.error(err);
    res.status(502).json({ error: "Could not reach the Power Automate flow", details: String(err) });
  }
});

/* -----------------------------------------------------------
   AI Advisor chat tool proxy. This tool is NOT registered in the
   database (no permission to write to SQL for it) — its tile is
   hardcoded client-side instead, in project.html. This route does
   NOT touch SQL at all.

   Token here is fetched via MSAL client-credentials (App ID +
   Tenant ID + Client Secret) — a Web-registered app, secret only
   ever used here, server-side. This is separate from getFlowToken()
   (the az-login-based function) still used by webscraper and
   interview-agent above — nothing about those two changed.
   ----------------------------------------------------------- */

const ADVISOR_FLOW_URL = process.env.ADVISOR_FLOW_URL;

const advisorMsalClient = new ConfidentialClientApplication({
  auth: {
    clientId: process.env.ENTRA_CLIENT_ID,
    authority: `https://login.microsoftonline.com/${process.env.ENTRA_TENANT_ID}`,
    clientSecret: process.env.ENTRA_CLIENT_SECRET
  }
});

async function getFlowTokenViaMsal() {
  const result = await advisorMsalClient.acquireTokenByClientCredential({
    scopes: ["https://service.flow.microsoft.com/.default"]
  });
  if (!result?.accessToken) throw new Error("MSAL returned no access token");
  return result.accessToken;
}

// [NO SQL] Talks only to Power Automate — no database access. Not registered in the database at all; its frontend tile is hardcoded in project.html.
app.post("/api/advisor-chat", requireSession, async (req, res) => {
  if (!ADVISOR_FLOW_URL) {
    return res.status(500).json({ error: "Server is missing ADVISOR_FLOW_URL — set it in server/.env" });
  }

  let token;
  try {
    token = await getFlowTokenViaMsal();
  } catch (err) {
    console.error("MSAL token fetch failed:", err.message);
    return res.status(401).json({
      error: "Could not get an Azure AD token via MSAL. Check ENTRA_CLIENT_ID / ENTRA_TENANT_ID / ENTRA_CLIENT_SECRET in server/.env.",
      details: err.message
    });
  }

  try {
    const flowRes = await fetch(ADVISOR_FLOW_URL, {
      method: "POST",
      headers: { "Content-Type": "application/json", "Authorization": `Bearer ${token}` },
      body: JSON.stringify(req.body)
    });

    const text = await flowRes.text();
    if (!flowRes.ok) {
      return res.status(502).json({ error: `Flow responded with ${flowRes.status}`, details: text });
    }
    res.send(text);
  } catch (err) {
    console.error(err);
    res.status(502).json({ error: "Could not reach the Power Automate flow", details: String(err) });
  }
});

/* ---------------- Serve the static frontend ---------------- */
app.use(express.static(path.join(__dirname, "..")));

app.listen(PORT, () => {
  console.log(`One Portal backend running at http://localhost:${PORT}`);
});
