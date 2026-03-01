import { google } from "googleapis";

function nowIso() {
  return new Date().toISOString();
}

function addMonths(date: Date, months: number) {
  const d = new Date(date);
  const day = d.getDate();
  d.setMonth(d.getMonth() + months);

  // Handle month rollover (e.g. Jan 31 + 1 month)
  if (d.getDate() < day) {
    d.setDate(0);
  }

  return d;
}

function randomSessionId() {
  return `sess_${Math.random().toString(36).slice(2)}${Math.random().toString(36).slice(2)}`;
}

type Role = "student" | "teacher" | "admin";

function normalizeRole(value: unknown): Role {
  const v = String(value ?? "")
    .trim()
    .toLowerCase();
  if (v === "admin" || v === "teacher" || v === "student") return v;
  return "student";
}

export default async function handler(req: any, res: any) {
  // CORS (needed for GitHub Pages) + preflight support
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");

  if (req.method === "OPTIONS") {
    return res.status(204).end();
  }

  if (req.method !== "POST") {
    return res.status(405).json({ message: "Method not allowed" });
  }

  try {
    const missing = ["GOOGLE_CLIENT_EMAIL", "GOOGLE_PRIVATE_KEY", "GOOGLE_SHEET_ID"].filter((k) => !process.env[k]);
    if (missing.length) {
      return res.status(500).json({ success: false, message: "Missing environment variables", missing });
    }

    const { courseSlug, username, password } = req.body ?? {};
    if (!courseSlug || !username || !password) {
      return res.status(400).json({
        success: false,
        message: "Missing courseSlug, username or password",
      });
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

    const normalizedUsername = String(username).trim();
    const normalizedCourseSlug = String(courseSlug).trim();
    const normalizedPassword = String(password);

    // 1) Validate user credentials against AuthUsers
    // A=username, B=passwordHash (plaintext for now), C=active, D=notes, E=role
    const usersResp = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: "AuthUsers!A2:E",
    });

    const userRows = usersResp.data.values ?? [];
    const userMatch = userRows.find((r) => {
      const u = String(r?.[0] ?? "").trim();
      const p = String(r?.[1] ?? "");
      const active = String(r?.[2] ?? "").trim();
      return u === normalizedUsername && active === "1" && p === normalizedPassword;
    });

    if (!userMatch) {
      return res.status(401).json({ success: false, message: "Invalid credentials" });
    }

    const role = normalizeRole(userMatch?.[4]);

    // 2) Validate per-course access against AuthAccess
    // A=username, B=courseSlug, C=active
    const accessResp = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: "AuthAccess!A2:C",
    });

    const accessRows = accessResp.data.values ?? [];
    const hasAccess = accessRows.some((r) => {
      const u = String(r?.[0] ?? "").trim();
      const slug = String(r?.[1] ?? "").trim();
      const active = String(r?.[2] ?? "").trim();
      return u === normalizedUsername && slug === normalizedCourseSlug && active === "1";
    });

    if (!hasAccess) {
      return res.status(401).json({ success: false, message: "Invalid credentials" });
    }

    // 3) Create session (3 months)
    const now = new Date();
    const firstAuthenticatedAt = nowIso();
    const expiresAt = addMonths(now, 3).toISOString();
    const sessionId = randomSessionId();
    const lastSeenAt = firstAuthenticatedAt;

    // Append session to AuthSessions
    // sessionId, username, role, firstAuthenticatedAt, expiresAt, revokedAt, lastSeenAt
    await sheets.spreadsheets.values.append({
      spreadsheetId,
      range: "AuthSessions",
      valueInputOption: "USER_ENTERED",
      requestBody: {
        values: [[sessionId, normalizedUsername, role, firstAuthenticatedAt, expiresAt, "", lastSeenAt]],
      },
    });

    return res.status(200).json({
      success: true,
      sessionId,
      expiresAt,
      role,
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
