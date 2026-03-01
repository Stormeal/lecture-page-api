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

function requireBearerSessionId(req: any) {
  const authHeader = String(req.headers?.authorization ?? "");
  const sessionId = authHeader.startsWith("Bearer ") ? authHeader.slice(7).trim() : "";
  return sessionId;
}

async function getSheetIdByTitle(sheets: any, spreadsheetId: string, title: string): Promise<number | null> {
  const meta = await sheets.spreadsheets.get({ spreadsheetId });
  const sheet = (meta.data.sheets ?? []).find((s: any) => s?.properties?.title === title);
  const sheetId = sheet?.properties?.sheetId;
  return typeof sheetId === "number" ? sheetId : null;
}

async function deleteRowsByIndices(sheets: any, spreadsheetId: string, sheetId: number, rowIndicesZeroBased: number[]) {
  // Must delete from bottom to top to avoid shifting indices
  const sorted = [...rowIndicesZeroBased].sort((a, b) => b - a);

  for (const idx of sorted) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [
          {
            deleteDimension: {
              range: {
                sheetId,
                dimension: "ROWS",
                startIndex: idx,
                endIndex: idx + 1,
              },
            },
          },
        ],
      },
    });
  }
}

export default async function handler(req: any, res: any) {
  // CORS (needed for GitHub Pages) + preflight support
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET, POST, PUT, DELETE, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");

  if (req.method === "OPTIONS") {
    return res.status(204).end();
  }

  if (req.method !== "GET" && req.method !== "POST" && req.method !== "PUT" && req.method !== "DELETE") {
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
    const sessionId = requireBearerSessionId(req);
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
    if (req.method === "POST") {
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

      // Ensure user does not already exist in AuthUsers
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

      // Append new user to AuthUsers
      await sheets.spreadsheets.values.append({
        spreadsheetId,
        range: "AuthUsers",
        valueInputOption: "USER_ENTERED",
        requestBody: {
          values: [[username, password, active, notes, role]],
        },
      });

      // Add access rows to AuthAccess
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
          requestBody: { values: newAccessRows },
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
    }

    // -----------------------
    // PUT: edit user (password/role/active/notes + replace access)
    // -----------------------
    if (req.method === "PUT") {
      const body = req.body ?? {};
      const username = String(body.username ?? "").trim();

      const passwordMaybe = body.password;
      const roleMaybe = body.role;
      const activeMaybe = body.active;
      const notesMaybe = body.notes;

      const coursesProvided = Object.prototype.hasOwnProperty.call(body, "courses");
      const desiredCourseSlugs: string[] = coursesProvided
        ? uniqueNonEmptyStrings(body.courses).map((v: string) => normalizeCourseSlug(v))
        : [];

      if (!username) {
        return res.status(400).json({ success: false, message: "Missing username" });
      }

      // Find user row in AuthUsers (A2:E)
      const usersResp = await sheets.spreadsheets.values.get({
        spreadsheetId,
        range: "AuthUsers!A2:E",
      });

      const userRows: any[] = usersResp.data.values ?? [];
      const userIndex = userRows.findIndex((r) => String(r?.[0] ?? "").trim() === username);

      if (userIndex === -1) {
        return res.status(404).json({ success: false, message: "User not found" });
      }

      // Sheet row number (1-based) for update range:
      // A2 is index 0, so rowNumber = 2 + userIndex
      const userRowNumber = 2 + userIndex;
      const existing = userRows[userIndex] ?? [];

      const existingPassword = String(existing?.[1] ?? "");
      const existingActive = String(existing?.[2] ?? "").trim() === "1" ? "1" : "0";
      const existingNotes = String(existing?.[3] ?? "").trim();
      const existingRole = normalizeRole(existing?.[4]);

      const nextPassword =
        typeof passwordMaybe === "string" && passwordMaybe.length > 0 ? String(passwordMaybe) : existingPassword;

      const nextRole = Object.prototype.hasOwnProperty.call(body, "role") ? normalizeRole(roleMaybe) : existingRole;

      const nextActive = Object.prototype.hasOwnProperty.call(body, "active")
        ? normalizeActive(activeMaybe)
        : (existingActive as "0" | "1");

      const nextNotes = Object.prototype.hasOwnProperty.call(body, "notes")
        ? String(notesMaybe ?? "").trim()
        : existingNotes;

      // Update AuthUsers row (A:E)
      await sheets.spreadsheets.values.update({
        spreadsheetId,
        range: `AuthUsers!A${userRowNumber}:E${userRowNumber}`,
        valueInputOption: "USER_ENTERED",
        requestBody: {
          values: [[username, nextPassword, nextActive, nextNotes, nextRole]],
        },
      });

      // Replace access only if "courses" was provided
      if (coursesProvided) {
        if (desiredCourseSlugs.length === 0) {
          return res.status(400).json({ success: false, message: "Select at least one course" });
        }

        // Read AuthAccess rows (A2:C)
        const accessResp = await sheets.spreadsheets.values.get({
          spreadsheetId,
          range: "AuthAccess!A2:C",
        });

        const accessRows: any[] = accessResp.data.values ?? [];

        // Collect existing access rows for this user with their sheet row number
        // rowNumber = 2 + index
        const existingForUser = accessRows
          .map((r, idx) => ({
            idx,
            rowNumber: 2 + idx,
            username: String(r?.[0] ?? "").trim(),
            slug: String(r?.[1] ?? "").trim(),
            active: String(r?.[2] ?? "").trim(),
          }))
          .filter((x) => x.username === username && x.slug);

        const desired = new Set<string>(desiredCourseSlugs.map((s) => s.trim()).filter(Boolean));

        // Build batch updates:
        // - Set active=1 for slugs in desired (if row exists)
        // - Set active=0 for slugs NOT in desired (if row exists)
        const updates: Array<{ range: string; values: string[][] }> = [];

        for (const row of existingForUser) {
          const shouldBeActive = desired.has(row.slug) ? "1" : "0";
          if (row.active !== shouldBeActive) {
            updates.push({
              range: `AuthAccess!C${row.rowNumber}`,
              values: [[shouldBeActive]],
            });
          }
        }

        if (updates.length > 0) {
          await sheets.spreadsheets.values.batchUpdate({
            spreadsheetId,
            requestBody: {
              valueInputOption: "USER_ENTERED",
              data: updates,
            },
          });
        }

        // Add rows for desired slugs that do NOT exist at all for this user
        const existingSlugSet = new Set(existingForUser.map((x) => x.slug));
        const missingSlugs = [...desired].filter((slug) => !existingSlugSet.has(slug));

        if (missingSlugs.length > 0) {
          await sheets.spreadsheets.values.append({
            spreadsheetId,
            range: "AuthAccess",
            valueInputOption: "USER_ENTERED",
            requestBody: {
              values: missingSlugs.map((slug) => [username, slug, "1"]),
            },
          });
        }
      }

      return res.status(200).json({
        success: true,
        user: {
          username,
          role: nextRole,
          active: nextActive === "1",
          notes: nextNotes,
        },
      });
    }

    // -----------------------
    // DELETE: hard delete user (AuthUsers row + all AuthAccess rows)
    // -----------------------
    if (req.method === "DELETE") {
      const body = req.body ?? {};
      const username = String(body.username ?? "").trim();

      if (!username) {
        return res.status(400).json({ success: false, message: "Missing username" });
      }

      // Need sheetIds for deleteDimension
      const authUsersSheetId = await getSheetIdByTitle(sheets, spreadsheetId, "AuthUsers");
      const authAccessSheetId = await getSheetIdByTitle(sheets, spreadsheetId, "AuthAccess");

      if (authUsersSheetId === null) {
        return res.status(500).json({ success: false, message: "Sheet AuthUsers not found" });
      }

      if (authAccessSheetId === null) {
        return res.status(500).json({ success: false, message: "Sheet AuthAccess not found" });
      }

      // Find AuthUsers row index (0-based within the sheet) to delete
      const usersResp = await sheets.spreadsheets.values.get({
        spreadsheetId,
        range: "AuthUsers!A2:E",
      });

      const userRows: any[] = usersResp.data.values ?? [];
      const userIndex = userRows.findIndex((r) => String(r?.[0] ?? "").trim() === username);

      if (userIndex === -1) {
        return res.status(404).json({ success: false, message: "User not found" });
      }

      // Convert to sheet row index (0-based):
      // A2 is sheet row index 1 (because row 1 is headers)
      const authUsersRowIndexZeroBased = 1 + userIndex;

      // Find ALL AuthAccess row indices to delete for this username
      const accessResp = await sheets.spreadsheets.values.get({
        spreadsheetId,
        range: "AuthAccess!A2:C",
      });

      const accessRows: any[] = accessResp.data.values ?? [];
      const accessRowIndicesZeroBased: number[] = [];

      for (let i = 0; i < accessRows.length; i++) {
        const r = accessRows[i];
        const u = String(r?.[0] ?? "").trim();
        if (u === username) {
          // A2 is sheet row index 1
          accessRowIndicesZeroBased.push(1 + i);
        }
      }

      // Delete AuthAccess rows first (bottom-up)
      if (accessRowIndicesZeroBased.length > 0) {
        await deleteRowsByIndices(sheets, spreadsheetId, authAccessSheetId, accessRowIndicesZeroBased);
      }

      // Delete AuthUsers row
      await deleteRowsByIndices(sheets, spreadsheetId, authUsersSheetId, [authUsersRowIndexZeroBased]);

      return res.status(200).json({
        success: true,
        deleted: {
          username,
          authAccessRowsDeleted: accessRowIndicesZeroBased.length,
        },
      });
    }

    return res.status(405).json({ message: "Method not allowed" });
  } catch (error: any) {
    console.error(error);
    return res.status(500).json({
      success: false,
      message: "Server error",
      error: error?.message ?? String(error),
    });
  }
}
