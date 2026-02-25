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
  // Not cryptographically perfect, but fine for this first iteration.
  // We can upgrade to crypto.randomUUID() in a follow-up if your runtime supports it.
  return `sess_${Math.random().toString(36).slice(2)}${Math.random().toString(36).slice(2)}`;
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

    const { username, password } = req.body ?? {};
    if (!username || !password) {
      return res.status(400).json({ success: false, message: "Missing username or password" });
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

    // Read users from AuthUsers!A2:D
    const usersResp = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: "AuthUsers!A2:D",
    });

    const rows = usersResp.data.values ?? [];
    const match = rows.find((r) => {
      const u = String(r?.[0] ?? "").trim();
      const p = String(r?.[1] ?? "");
      const active = String(r?.[2] ?? "").trim();
      return u === String(username).trim() && active === "1" && p === String(password);
    });

    if (!match) {
      return res.status(401).json({ success: false, message: "Invalid credentials" });
    }

    const now = new Date();
    const firstAuthenticatedAt = nowIso();
    const expiresAt = addMonths(now, 3).toISOString();
    const sessionId = randomSessionId();
    const lastSeenAt = firstAuthenticatedAt;

    // Append session to AuthSessions
    await sheets.spreadsheets.values.append({
      spreadsheetId,
      range: "AuthSessions",
      valueInputOption: "USER_ENTERED",
      requestBody: {
        values: [[sessionId, String(username).trim(), firstAuthenticatedAt, expiresAt, "", lastSeenAt]],
      },
    });

    return res.status(200).json({
      success: true,
      sessionId,
      expiresAt,
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
