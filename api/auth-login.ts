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

function normalizeActive(value: unknown): "0" | "1" {
  const v = String(value ?? "")
    .trim()
    .toLowerCase();
  if (v === "1" || v === "true" || v === "yes" || v === "on") return "1";
  return "0";
}

function normalizeCourseSlug(value: unknown): string {
  return String(value ?? "").trim();
}

function uniqueNonEmptyStrings(values: unknown): string[] {
  const arr = Array.isArray(values) ? values : [];
  const set = new Set<string>();

  for (const item of arr) {
    const v = String(item ?? "").trim();
    if (v) set.add(v);
  }

  return [...set];
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
  // AuthSessions columns:
  // A=sessionId, B=username, C=role, D=firstAuthenticatedAt, E=expiresAt, F=revokedAt, G=lastSeenAt
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
  if (expiresAtMs === null || expiresAtMs <= Date.now())
    return { ok: false as const, status: 401, message: "Session expired" };

  if (role !== "admin") return { ok: false as const, status: 403, message: "Forbidden" };

  return { ok: true as const };
}

export default async function handler(req: any, res: any) {
  // CORS (needed for GitHub Pages) + preflight support
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");

  if (req.method === "OPTIONS") {
    return res.status(204).end();
  }

  if (req.method !== "GET" && req.method !== "POST") {
    return res.status(405).json({ message: "Method not allowed" });
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

    // Expect: Authorization: Bearer sess_xxx
    const authHeader = String(req.headers?.authorization ?? "");
    const sessionId = authHeader.startsWith("Bearer ") ? authHeader.slice(7).trim() : "";

    if (!sessionId) {
      return res.status(401).json({ success: false, message: "Missing session" });
    }

    const adminCheck = await requireAdminSession(sheets, spreadsheetId, sessionId);
    if (!adminCheck.ok) {
      return res.status(adminCheck.status).json({ success: false, message: adminCheck.message });
    }

    // -----------------------
    // GET: list users
    // -----------------------
    if (req.method === "GET") {
      // AuthUsers columns:
      // A=username, B=password, C=active, D=notes, E=role
      const usersResp = await sheets.spreadsheets.values.get({
        spreadsheetId,
        range: "AuthUsers!A2:E",
      });

      const rows: any[] = usersResp.data.values ?? [];

      const users = rows
        .map((r) => {
          const username = String(r?.[0] ?? "").trim();
          if (!username) return null;

          return {
            username,
            active: String(r?.[2] ?? "").trim() === "1",
            role: normalizeRole(r?.[4]),
            notes: String(r?.[3] ?? "").trim(),
          };
        })
        .filter(Boolean);

      return res.status(200).json({ success: true, users });
    }

    // -----------------------
    // POST: create user
    // -----------------------
    const body = req.body ?? {};
    const username = String(body.username ?? "").trim();
    const password = String(body.password ?? "");
    const role = normalizeRole(body.role);
    const active = normalizeActive(body.active);
    const notes = String(body.notes ?? "").trim();

    const courseSlugs: string[] = uniqueNonEmptyStrings(body.courses).map((v: string) => normalizeCourseSlug(v));

    if (!username || !password) {
      return res.status(400).json({
        success: false,
        message: "Missing username or password",
      });
    }

    if (courseSlugs.length === 0) {
      return res.status(400).json({
        success: false,
        message: "Select at least one course",
      });
    }

    // 1) Ensure user does not already exist in AuthUsers
    const existingUsersResp = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: "AuthUsers!A2:E",
    });

    const existingUserRows: any[] = existingUsersResp.data.values ?? [];
    const userExists = existingUserRows.some((r) => String(r?.[0] ?? "").trim() === username);

    if (userExists) {
      return res.status(409).json({
        success: false,
        message: "Username already exists",
      });
    }

    // 2) Append new user to AuthUsers
    await sheets.spreadsheets.values.append({
      spreadsheetId,
      range: "AuthUsers",
      valueInputOption: "USER_ENTERED",
      requestBody: {
        // A=username, B=password, C=active, D=notes, E=role
        values: [[username, password, active, notes, role]],
      },
    });

    // 3) Add access rows to AuthAccess
    // AuthAccess columns:
    // A=username, B=courseSlug, C=active
    const accessResp = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: "AuthAccess!A2:C",
    });

    const accessRows: any[] = accessResp.data.values ?? [];
    const existingPairs = new Set<string>(
      accessRows.map((r) => `${String(r?.[0] ?? "").trim()}|${String(r?.[1] ?? "").trim()}`),
    );

    const newAccessRows: string[][] = courseSlugs
      .map((slug: string) => slug.trim())
      .filter((slug: string) => Boolean(slug))
      .filter((slug: string) => !existingPairs.has(`${username}|${slug}`))
      .map((slug: string) => [username, slug, "1"]);

    if (newAccessRows.length > 0) {
      await sheets.spreadsheets.values.append({
        spreadsheetId,
        range: "AuthAccess",
        valueInputOption: "USER_ENTERED",
        requestBody: {
          values: newAccessRows,
        },
      });
    }

    return res.status(200).json({
      success: true,
      user: {
        username,
        role,
        active: active === "1",
        notes,
        coursesAdded: courseSlugs,
      },
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
