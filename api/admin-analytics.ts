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

function requireBearerSessionId(req: any) {
  const authHeader = String(req.headers?.authorization ?? "");
  const sessionId = authHeader.startsWith("Bearer ") ? authHeader.slice(7).trim() : "";
  return sessionId;
}

type UserAgg = {
  username: string;
  totalSubmissions: number;
  correct: number;
  lastActivityAt: string | null;
  courses: Set<string>;
};

type CourseAgg = {
  course: string;
  totalSubmissions: number;
  correct: number;
  users: Set<string>;
};

function toAccuracy(correct: number, total: number): number {
  if (!total) return 0;
  return correct / total;
}

export default async function handler(req: any, res: any) {
  // CORS (needed for GitHub Pages) + preflight support
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");

  if (req.method === "OPTIONS") {
    return res.status(204).end();
  }

  if (req.method !== "GET") {
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

    const sessionId = requireBearerSessionId(req);
    if (!sessionId) {
      return res.status(401).json({ success: false, message: "Missing session" });
    }

    const adminCheck = await requireAdminSession(sheets, spreadsheetId, sessionId);
    if (!adminCheck.ok) {
      return res.status(adminCheck.status).json({ success: false, message: adminCheck.message });
    }

    // 1) Load existing users (so deleted users are excluded)
    const usersResp = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: "AuthUsers!A2:E",
    });

    const userRows: any[] = usersResp.data.values ?? [];
    const existingUsers = new Set<string>();
    for (const r of userRows) {
      const username = String(r?.[0] ?? "").trim();
      if (username) existingUsers.add(username);
    }

    // 2) Load submissions
    // A=timestamp, B=course, C=quizId, D=questionId, E=selectedOption, F=isCorrect, G=attempts, H=userId
    const quizResp = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: "QuizSubmissions!A2:H",
    });

    const quizRows: any[] = quizResp.data.values ?? [];

    const userAgg = new Map<string, UserAgg>();
    const courseAgg = new Map<string, CourseAgg>();

    for (const r of quizRows) {
      const ts = String(r?.[0] ?? "").trim();
      const course = String(r?.[1] ?? "").trim();
      const isCorrectRaw = r?.[5];
      const userId = String(r?.[7] ?? "").trim();

      if (!userId || userId === "anonymous") continue;
      if (!existingUsers.has(userId)) continue;

      const correct =
        String(isCorrectRaw ?? "")
          .trim()
          .toLowerCase() === "true" || String(isCorrectRaw ?? "").trim() === "1";

      // --- user agg ---
      const u = userAgg.get(userId) ?? {
        username: userId,
        totalSubmissions: 0,
        correct: 0,
        lastActivityAt: null,
        courses: new Set<string>(),
      };

      u.totalSubmissions += 1;
      if (correct) u.correct += 1;
      if (course) u.courses.add(course);

      const tsMs = parseIsoMs(ts);
      const lastMs = parseIsoMs(u.lastActivityAt ?? "");
      if (tsMs !== null && (lastMs === null || tsMs > lastMs)) {
        u.lastActivityAt = new Date(tsMs).toISOString();
      }

      userAgg.set(userId, u);

      // --- course agg ---
      if (course) {
        const c = courseAgg.get(course) ?? {
          course,
          totalSubmissions: 0,
          correct: 0,
          users: new Set<string>(),
        };

        c.totalSubmissions += 1;
        if (correct) c.correct += 1;
        c.users.add(userId);

        courseAgg.set(course, c);
      }
    }

    const users = [...userAgg.values()]
      .map((u) => ({
        username: u.username,
        totalSubmissions: u.totalSubmissions,
        correct: u.correct,
        accuracy: toAccuracy(u.correct, u.totalSubmissions),
        lastActivityAt: u.lastActivityAt,
        courses: [...u.courses].sort((a, b) => a.localeCompare(b)),
      }))
      .sort((a, b) => {
        // Highest total activity first, then accuracy
        if (b.totalSubmissions !== a.totalSubmissions) return b.totalSubmissions - a.totalSubmissions;
        if (b.accuracy !== a.accuracy) return b.accuracy - a.accuracy;
        return a.username.localeCompare(b.username);
      });

    const courses = [...courseAgg.values()]
      .map((c) => ({
        course: c.course,
        totalSubmissions: c.totalSubmissions,
        uniqueUsers: c.users.size,
        accuracy: toAccuracy(c.correct, c.totalSubmissions),
      }))
      .sort((a, b) => b.totalSubmissions - a.totalSubmissions);

    return res.status(200).json({
      success: true,
      users,
      courses,
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
