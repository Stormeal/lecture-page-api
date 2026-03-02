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
  if (missing.length) return { ok: false as const, missing };

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

function requireBearerSessionId(req: any) {
  const authHeader = String(req.headers?.authorization ?? "");
  const sessionId = authHeader.startsWith("Bearer ") ? authHeader.slice(7).trim() : "";
  return sessionId;
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

function parseBool(value: unknown): boolean {
  const s = String(value ?? "")
    .trim()
    .toLowerCase();
  return s === "true" || s === "1" || s === "yes";
}

type CourseUserAgg = {
  username: string;
  totalSubmissions: number;
  correct: number;
  accuracy: number; // 0..1
  lastActivityAt: string | null; // ISO
};

type CourseQuestionAgg = {
  quizId: string;
  questionId: string;
  totalSubmissions: number;
  correct: number;
  accuracy: number; // 0..1
  uniqueUsers: number;
  lastAnsweredAt: string | null; // ISO
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

    const course = String(req.query?.course ?? "").trim();
    if (!course) {
      return res.status(400).json({ success: false, message: "Missing course" });
    }

    // 1) Load existing users (so deleted users are excluded)
    const usersResp = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: "AuthUsers!A2:E",
    });

    const userRows: any[] = usersResp.data.values ?? [];
    const existingUsers = new Set<string>();
    for (const r of userRows) {
      const u = String(r?.[0] ?? "").trim();
      if (u) existingUsers.add(u);
    }

    // 2) Load submissions
    // A=timestamp, B=course, C=quizId, D=questionId, E=selectedOption, F=isCorrect, G=attempts, H=userId
    const quizResp = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: "QuizSubmissions!A2:H",
    });

    const quizRows: any[] = quizResp.data.values ?? [];

    // Aggregate per-user (within this course)
    const userAgg = new Map<string, { total: number; correct: number; lastAt: string | null }>();

    // Aggregate per-question (within this course), include unique users
    const questionAgg = new Map<
      string,
      { quizId: string; questionId: string; total: number; correct: number; users: Set<string>; lastAt: string | null }
    >();

    for (const r of quizRows) {
      const ts = String(r?.[0] ?? "").trim();
      const rowCourse = String(r?.[1] ?? "").trim();
      const quizId = String(r?.[2] ?? "").trim();
      const questionId = String(r?.[3] ?? "").trim();
      const isCorrectRaw = r?.[5];
      const userId = String(r?.[7] ?? "").trim();

      if (rowCourse !== course) continue;
      if (!userId || userId === "anonymous") continue;
      if (!existingUsers.has(userId)) continue;
      if (!quizId || !questionId) continue;

      const tsMs = parseIsoMs(ts);
      if (tsMs === null) continue;

      const isCorrect = parseBool(isCorrectRaw);

      // user agg
      const u = userAgg.get(userId) ?? { total: 0, correct: 0, lastAt: null };
      u.total += 1;
      if (isCorrect) u.correct += 1;

      const uLastMs = parseIsoMs(u.lastAt ?? "");
      if (uLastMs === null || tsMs > uLastMs) u.lastAt = new Date(tsMs).toISOString();
      userAgg.set(userId, u);

      // question agg
      const qKey = `${quizId}||${questionId}`;
      const q = questionAgg.get(qKey) ?? {
        quizId,
        questionId,
        total: 0,
        correct: 0,
        users: new Set<string>(),
        lastAt: null,
      };

      q.total += 1;
      if (isCorrect) q.correct += 1;
      q.users.add(userId);

      const qLastMs = parseIsoMs(q.lastAt ?? "");
      if (qLastMs === null || tsMs > qLastMs) q.lastAt = new Date(tsMs).toISOString();

      questionAgg.set(qKey, q);
    }

    const users: CourseUserAgg[] = [...userAgg.entries()]
      .map(([username, u]) => ({
        username,
        totalSubmissions: u.total,
        correct: u.correct,
        accuracy: toAccuracy(u.correct, u.total),
        lastActivityAt: u.lastAt,
      }))
      .sort((a, b) => {
        if (b.totalSubmissions !== a.totalSubmissions) return b.totalSubmissions - a.totalSubmissions;
        if (b.accuracy !== a.accuracy) return b.accuracy - a.accuracy;
        return a.username.localeCompare(b.username);
      });

    const questions: CourseQuestionAgg[] = [...questionAgg.values()]
      .map((q) => ({
        quizId: q.quizId,
        questionId: q.questionId,
        totalSubmissions: q.total,
        correct: q.correct,
        accuracy: toAccuracy(q.correct, q.total),
        uniqueUsers: q.users.size,
        lastAnsweredAt: q.lastAt,
      }))
      .sort((a, b) => {
        if (a.quizId !== b.quizId) return a.quizId.localeCompare(b.quizId);
        return a.questionId.localeCompare(b.questionId);
      });

    const courseTotals = {
      totalSubmissions: users.reduce((sum, u) => sum + u.totalSubmissions, 0),
      uniqueUsers: users.length,
      accuracy: toAccuracy(
        users.reduce((sum, u) => sum + u.correct, 0),
        users.reduce((sum, u) => sum + u.totalSubmissions, 0),
      ),
    };

    return res.status(200).json({
      success: true,
      course,
      totals: courseTotals,
      users,
      questions,
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
