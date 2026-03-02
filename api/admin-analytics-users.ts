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

type LatestQuestionAttempt = {
  course: string;
  quizId: string;
  questionId: string;
  lastAnsweredAt: string; // ISO
  selectedOption: string;
  attempts: number;
  isCorrect: boolean;
};

type GroupedQuiz = {
  quizId: string;
  questions: LatestQuestionAttempt[];
};

type GroupedCourse = {
  course: string;
  quizzes: GroupedQuiz[];
};

function parseBool(value: unknown): boolean {
  const s = String(value ?? "")
    .trim()
    .toLowerCase();
  return s === "true" || s === "1" || s === "yes";
}

function parseAttempts(value: unknown): number {
  const n = Number(String(value ?? "").trim());
  return Number.isFinite(n) && n >= 0 ? n : 0;
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

    const username = String(req.query?.username ?? "").trim();
    const courseFilter = String(req.query?.course ?? "").trim();

    if (!username) {
      return res.status(400).json({ success: false, message: "Missing username" });
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

    if (!existingUsers.has(username)) {
      return res.status(404).json({ success: false, message: "User not found" });
    }

    // 2) Load submissions
    // A=timestamp, B=course, C=quizId, D=questionId, E=selectedOption, F=isCorrect, G=attempts, H=userId
    const quizResp = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: "QuizSubmissions!A2:H",
    });

    const quizRows: any[] = quizResp.data.values ?? [];

    // Latest-only per (course, quizId, questionId) for this user
    const latestByKey = new Map<string, LatestQuestionAttempt>();

    for (const r of quizRows) {
      const ts = String(r?.[0] ?? "").trim();
      const course = String(r?.[1] ?? "").trim();
      const quizId = String(r?.[2] ?? "").trim();
      const questionId = String(r?.[3] ?? "").trim();
      const selectedOption = String(r?.[4] ?? "").trim();
      const isCorrectRaw = r?.[5];
      const attemptsRaw = r?.[6];
      const userId = String(r?.[7] ?? "").trim();

      if (userId !== username) continue;
      if (!course || !quizId || !questionId) continue;
      if (courseFilter && course !== courseFilter) continue;

      const tsMs = parseIsoMs(ts);
      if (tsMs === null) continue;

      const isCorrect = parseBool(isCorrectRaw);
      const attempts = parseAttempts(attemptsRaw);

      const key = `${course}||${quizId}||${questionId}`;
      const prev = latestByKey.get(key);

      const prevMs = prev ? parseIsoMs(prev.lastAnsweredAt) : null;
      if (!prev || prevMs === null || tsMs > prevMs) {
        latestByKey.set(key, {
          course,
          quizId,
          questionId,
          lastAnsweredAt: new Date(tsMs).toISOString(),
          selectedOption,
          attempts,
          isCorrect,
        });
      }
    }

    const attemptsArr = [...latestByKey.values()].sort((a, b) => {
      // Newest first
      const am = parseIsoMs(a.lastAnsweredAt) ?? 0;
      const bm = parseIsoMs(b.lastAnsweredAt) ?? 0;
      if (bm !== am) return bm - am;
      // Stable sorting
      if (a.course !== b.course) return a.course.localeCompare(b.course);
      if (a.quizId !== b.quizId) return a.quizId.localeCompare(b.quizId);
      return a.questionId.localeCompare(b.questionId);
    });

    // Group: course -> quiz -> questions
    const byCourse = new Map<string, Map<string, LatestQuestionAttempt[]>>();
    for (const a of attemptsArr) {
      const quizMap = byCourse.get(a.course) ?? new Map<string, LatestQuestionAttempt[]>();
      const list = quizMap.get(a.quizId) ?? [];
      list.push(a);
      quizMap.set(a.quizId, list);
      byCourse.set(a.course, quizMap);
    }

    const grouped: GroupedCourse[] = [...byCourse.entries()]
      .map(([course, quizMap]) => ({
        course,
        quizzes: [...quizMap.entries()]
          .map(([quizId, questions]) => ({
            quizId,
            questions: questions.sort((x, y) => x.questionId.localeCompare(y.questionId)),
          }))
          .sort((a, b) => a.quizId.localeCompare(b.quizId)),
      }))
      .sort((a, b) => a.course.localeCompare(b.course));

    return res.status(200).json({
      success: true,
      username,
      course: courseFilter || null,
      latest: grouped,
      totalLatestQuestions: attemptsArr.length,
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
