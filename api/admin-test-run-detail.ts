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

  const missing = [];
  if (!token) missing.push("GITHUB_TOKEN");
  if (!owner) missing.push("GITHUB_OWNER");
  if (!repo) missing.push("GITHUB_REPO");

  return {
    token,
    owner,
    repo,
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

    const runId = Number.parseInt(String(req.query?.runId ?? "").trim(), 10);
    if (!Number.isFinite(runId) || runId <= 0) {
      return res.status(400).json({ success: false, message: "Invalid runId" });
    }

    const githubConfig = getGitHubConfig();
    if (githubConfig.missing.length > 0) {
      return res.status(500).json({
        success: false,
        message: "Missing GitHub environment variables",
        missing: githubConfig.missing,
      });
    }

    const [runResponse, jobsResponse] = await Promise.all([
      fetch(`https://api.github.com/repos/${githubConfig.owner}/${githubConfig.repo}/actions/runs/${runId}`, {
        headers: getGitHubHeaders(githubConfig.token),
      }),
      fetch(`https://api.github.com/repos/${githubConfig.owner}/${githubConfig.repo}/actions/runs/${runId}/jobs?per_page=100`, {
        headers: getGitHubHeaders(githubConfig.token),
      }),
    ]);

    if (!runResponse.ok) {
      const text = await runResponse.text();
      throw new Error(`Failed to load workflow run (${runResponse.status}): ${text}`);
    }

    if (!jobsResponse.ok) {
      const text = await jobsResponse.text();
      throw new Error(`Failed to load workflow jobs (${jobsResponse.status}): ${text}`);
    }

    const run = (await runResponse.json()) as any;
    const jobsData = (await jobsResponse.json()) as { jobs?: any[] };

    return res.status(200).json({
      success: true,
      run: {
        runId: run.id,
        runNumber: run.run_number,
        runUrl: run.html_url,
        workflowName: run.name ?? "Playwright Admin Run",
        title: run.display_title ?? `Run #${run.run_number}`,
        status: run.status ?? "unknown",
        conclusion: run.conclusion ?? null,
        createdAt: run.created_at ?? null,
        updatedAt: run.updated_at ?? null,
        actor: run.triggering_actor?.login?.trim() || run.actor?.login?.trim() || null,
        event: run.event ?? null,
        headBranch: run.head_branch ?? null,
      },
      jobs: (jobsData.jobs ?? []).map((job) => ({
        id: job.id,
        name: job.name ?? "Job",
        status: job.status ?? "unknown",
        conclusion: job.conclusion ?? null,
        startedAt: job.started_at ?? null,
        completedAt: job.completed_at ?? null,
        url: job.html_url ?? null,
      })),
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
