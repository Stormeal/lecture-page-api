import { google } from "googleapis";

type Role = "student" | "teacher" | "admin";

function normalizeRole(value: unknown): Role {
  const v = String(value ?? "")
    .trim()
    .toLowerCase();
  if (v === "admin" || v === "teacher" || v === "student") return v;
  return "student";
}

function parseIsoMs(value: unknown): number | null {
  const ms = Date.parse(String(value ?? ""));
  return Number.isFinite(ms) ? ms : null;
}

async function getSheets() {
  const missing = ["GOOGLE_CLIENT_EMAIL", "GOOGLE_PRIVATE_KEY", "GOOGLE_SHEET_ID"].filter((k) => !process.env[k]);
  if (missing.length) {
    return { ok: false as const, missing };
  }

  const auth = new google.auth.GoogleAuth({
    credentials: {
      client_email: process.env.GOOGLE_CLIENT_EMAIL,
      private_key: process.env.GOOGLE_PRIVATE_KEY?.replace(/\\n/g, "\n"),
    },
    scopes: ["https://www.googleapis.com/auth/spreadsheets"],
  });

  const sheets = google.sheets({ version: "v4", auth });
  const spreadsheetId = process.env.GOOGLE_SHEET_ID as string;

  return { ok: true as const, sheets, spreadsheetId };
}

async function requireAdminSession(sheets: any, spreadsheetId: string, sessionId: string) {
  const sessionsResp = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: "AuthSessions!A2:G",
  });

  const rows: any[] = sessionsResp.data.values ?? [];
  const row = rows.find((r) => String(r?.[0] ?? "") === sessionId);
  if (!row) return { ok: false as const, status: 401, message: "Invalid session" };

  const role = normalizeRole(row?.[2]);
  const expiresAtMs = parseIsoMs(row?.[4]);
  const revokedAt = String(row?.[5] ?? "").trim();

  if (revokedAt) return { ok: false as const, status: 401, message: "Session revoked" };
  if (expiresAtMs === null || expiresAtMs <= Date.now()) {
    return { ok: false as const, status: 401, message: "Session expired" };
  }
  if (role !== "admin") return { ok: false as const, status: 403, message: "Forbidden" };

  return { ok: true as const };
}

function requireBearerSessionId(req: any) {
  const authHeader = String(req.headers?.authorization ?? "");
  return authHeader.startsWith("Bearer ") ? authHeader.slice(7).trim() : "";
}

function getGitHubConfig() {
  const token = String(process.env.GITHUB_TOKEN ?? "").trim();
  const owner = String(process.env.GITHUB_OWNER ?? "Stormeal").trim();
  const repo = String(process.env.GITHUB_REPO ?? "lecture-page").trim();
  const workflowId = String(process.env.GITHUB_WORKFLOW_ID ?? "playwright-admin-run.yml").trim();
  const ref = String(process.env.GITHUB_WORKFLOW_REF ?? "main").trim();

  const missing = [];
  if (!token) missing.push("GITHUB_TOKEN");
  if (!owner) missing.push("GITHUB_OWNER");
  if (!repo) missing.push("GITHUB_REPO");
  if (!workflowId) missing.push("GITHUB_WORKFLOW_ID");
  if (!ref) missing.push("GITHUB_WORKFLOW_REF");

  return {
    token,
    owner,
    repo,
    workflowId,
    ref,
    missing,
  };
}

function getGitHubHeaders(token: string) {
  return {
    Accept: "application/vnd.github+json",
    Authorization: `Bearer ${token}`,
    "X-GitHub-Api-Version": "2022-11-28",
  };
}

export default async function handler(req: any, res: any) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");

  if (req.method === "OPTIONS") {
    return res.status(204).end();
  }

  if (req.method !== "GET") {
    return res.status(405).json({ success: false, message: "Method not allowed" });
  }

  try {
    const sheetsInit = await getSheets();
    if (!sheetsInit.ok) {
      return res.status(500).json({
        success: false,
        message: "Missing environment variables",
        missing: sheetsInit.missing,
      });
    }

    const { sheets, spreadsheetId } = sheetsInit;

    const sessionId = requireBearerSessionId(req);
    if (!sessionId) {
      return res.status(401).json({ success: false, message: "Missing session" });
    }

    const adminCheck = await requireAdminSession(sheets, spreadsheetId, sessionId);
    if (!adminCheck.ok) {
      return res.status(adminCheck.status).json({ success: false, message: adminCheck.message });
    }

    const githubConfig = getGitHubConfig();
    if (githubConfig.missing.length > 0) {
      return res.status(500).json({
        success: false,
        message: "Missing GitHub environment variables",
        missing: githubConfig.missing,
      });
    }

    const response = await fetch(
      `https://api.github.com/repos/${githubConfig.owner}/${githubConfig.repo}/actions/workflows/${githubConfig.workflowId}/runs?branch=${encodeURIComponent(githubConfig.ref)}&per_page=10`,
      {
        headers: getGitHubHeaders(githubConfig.token),
      },
    );

    if (!response.ok) {
      const text = await response.text();
      throw new Error(`Failed to list workflow runs (${response.status}): ${text}`);
    }

    const data = (await response.json()) as {
      workflow_runs?: Array<{
        id: number;
        html_url: string;
        run_number: number;
        status: string | null;
        conclusion: string | null;
        created_at: string | null;
        updated_at: string | null;
        actor?: { login?: string | null } | null;
        triggering_actor?: { login?: string | null } | null;
        name?: string | null;
        display_title?: string | null;
        event?: string | null;
      }>;
    };

    const runs = (data.workflow_runs ?? []).map((run) => ({
      runId: run.id,
      runNumber: run.run_number,
      runUrl: run.html_url,
      workflowName: run.name ?? "Playwright Admin Run",
      title: run.display_title ?? `Run #${run.run_number}`,
      status: run.status ?? "unknown",
      conclusion: run.conclusion,
      createdAt: run.created_at,
      updatedAt: run.updated_at,
      actor:
        run.triggering_actor?.login?.trim() ||
        run.actor?.login?.trim() ||
        null,
      event: run.event ?? null,
    }));

    return res.status(200).json({
      success: true,
      runs,
    });
  } catch (error: any) {
    console.error(error);
    return res.status(500).json({
      success: false,
      message: "Server error",
      error: error?.message ?? String(error),
    });
  }
}
