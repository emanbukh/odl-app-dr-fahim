import express from "express";
import bcrypt from "bcryptjs";
import crypto from "crypto";
import jwt from "jsonwebtoken";
import Database from "better-sqlite3";
import path from "path";
import { fileURLToPath } from "url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const PORT = process.env.PORT || 5179;
const JWT_SECRET = process.env.JWT_SECRET || "odl-portal-dev-secret";

app.use(express.json({ limit: "50mb" }));
app.use((req, res, next) => {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");
  res.setHeader("Access-Control-Allow-Methods", "GET,POST,OPTIONS");
  if (req.method === "OPTIONS") return res.sendStatus(204);
  next();
});
app.use(express.static(path.join(__dirname, "public")));

const db = new Database(path.join(__dirname, "portal.db"));
let lastUpdateId = 0;

db.exec(`
  PRAGMA journal_mode = WAL;
  CREATE TABLE IF NOT EXISTS students (
    student_id TEXT PRIMARY KEY,
    name TEXT,
    ic_hash TEXT,
    intake TEXT,
    intake_year TEXT,
    intake_month TEXT,
    program TEXT,
    status TEXT,
    cgpa REAL
  );
  CREATE TABLE IF NOT EXISTS courses (
    course_code TEXT PRIMARY KEY,
    title TEXT,
    credits REAL,
    is_mpu INTEGER
  );
  CREATE TABLE IF NOT EXISTS results (
    result_id TEXT PRIMARY KEY,
    student_id TEXT,
    course_code TEXT,
    session TEXT,
    semester INTEGER,
    mark REAL,
    letter TEXT,
    point REAL,
    grade_points REAL,
    passed_course INTEGER,
    credits_earned REAL
  );
  CREATE TABLE IF NOT EXISTS meta (
    key TEXT PRIMARY KEY,
    value TEXT
  );
  CREATE TABLE IF NOT EXISTS otp_requests (
    student_id TEXT,
    otp_hash TEXT,
    expires_at INTEGER,
    attempts INTEGER DEFAULT 0
  );
`);

const storedUpdate = db.prepare("SELECT value FROM meta WHERE key = 'last_update_id'").get();
if (storedUpdate?.value) lastUpdateId = Number(storedUpdate.value) || 0;

function toNumber(value) {
  const n = Number(value);
  return Number.isFinite(n) ? n : null;
}

function hashIc(value) {
  const raw = String(value ?? "").trim();
  if (!raw) return "";
  return crypto.createHash("sha256").update(raw).digest("hex");
}

const defaultGradeScale = [
  { min: 90, letter: "A+", point: 4.0, passed: true },
  { min: 80, letter: "A", point: 4.0, passed: true },
  { min: 75, letter: "A-", point: 3.67, passed: true },
  { min: 70, letter: "B+", point: 3.33, passed: true },
  { min: 65, letter: "B", point: 3.0, passed: true },
  { min: 60, letter: "B-", point: 2.67, passed: true },
  { min: 55, letter: "C+", point: 2.33, passed: true },
  { min: 50, letter: "C", point: 2.0, passed: true },
  { min: 45, letter: "D+", point: 1.67, passed: true },
  { min: 40, letter: "D", point: 1.33, passed: true },
  { min: 0, letter: "F", point: 0.0, passed: false },
];

const mpuGradeScale = [
  { min: 75, letter: "A", point: 4.0, passed: true },
  { min: 65, letter: "B", point: 3.0, passed: true },
  { min: 50, letter: "C", point: 2.0, passed: true },
  { min: 0, letter: "F", point: 0.0, passed: false },
];

function gradeFromMark(markValue, isMPU) {
  const mark = toNumber(markValue);
  if (mark === null) return { letter: "", point: null, passed: false };
  const scale = isMPU ? mpuGradeScale : defaultGradeScale;
  for (const grade of scale) {
    if (mark >= grade.min) return grade;
  }
  return { letter: "F", point: 0.0, passed: false };
}

function computeGpa(rows) {
  let credits = 0;
  let gradePoints = 0;
  for (const row of rows) {
    const rowCredits = toNumber(row.credits);
    if (rowCredits === null) continue;
    const point = toNumber(row.point);
    const gradePts =
      toNumber(row.grade_points) ??
      (point !== null ? point * rowCredits : null);
    if (gradePts === null) continue;
    credits += rowCredits;
    gradePoints += gradePts;
  }
  return credits > 0 ? gradePoints / credits : null;
}

function requireAuth(req, res, next) {
  const header = req.headers.authorization || "";
  const token = header.startsWith("Bearer ") ? header.slice(7) : "";
  if (!token) return res.status(401).json({ error: "Missing token" });
  try {
    const payload = jwt.verify(token, JWT_SECRET);
    req.studentId = payload.studentId;
    return next();
  } catch (e) {
    return res.status(401).json({ error: "Invalid token" });
  }
}

app.post("/api/admin/import-portal-export", async (req, res) => {
  try {
    const payload = req.body;
    if (!payload || !Array.isArray(payload.students) || !Array.isArray(payload.results)) {
      return res.status(400).json({ error: "Invalid export format." });
    }

    const insertStudent = db.prepare(`
      INSERT OR REPLACE INTO students
      (student_id, name, ic_hash, intake, intake_year, intake_month, program, status, cgpa)
      VALUES (@student_id, @name, @ic_hash, @intake, @intake_year, @intake_month, @program, @status, @cgpa)
    `);
    const insertCourse = db.prepare(`
      INSERT OR REPLACE INTO courses
      (course_code, title, credits, is_mpu)
      VALUES (@course_code, @title, @credits, @is_mpu)
    `);
    const insertResult = db.prepare(`
      INSERT OR REPLACE INTO results
      (result_id, student_id, course_code, session, semester, mark, letter, point, grade_points, passed_course, credits_earned)
      VALUES (@result_id, @student_id, @course_code, @session, @semester, @mark, @letter, @point, @grade_points, @passed_course, @credits_earned)
    `);

    db.exec("DELETE FROM students; DELETE FROM courses; DELETE FROM results; DELETE FROM otp_requests;");

    const insertTx = db.transaction(() => {
      for (const student of payload.students) {
        insertStudent.run({
          student_id: String(student.studentId ?? "").trim(),
          name: String(student.name ?? "").trim(),
          ic_hash: null,
          intake: String(student.intake ?? "").trim(),
          intake_year: String(student.intakeYear ?? "").trim(),
          intake_month: String(student.intakeMonth ?? "").trim(),
          program: String(student.program ?? "").trim(),
          status: String(student.status ?? "").trim(),
          cgpa: toNumber(student.cgpa),
        });
      }

      for (const course of payload.courses ?? []) {
        insertCourse.run({
          course_code: String(course.courseCode ?? "").trim(),
          title: String(course.title ?? "").trim(),
          credits: toNumber(course.credits),
          is_mpu: course.isMPUCourse === true ? 1 : course.isMPUCourse === false ? 0 : null,
        });
      }

      for (const result of payload.results) {
        const resultId =
          String(result.resultId ?? "").trim() ||
          `${result.studentId}|${result.courseCode}|${result.session}|${result.semester ?? ""}`;
        insertResult.run({
          result_id: resultId,
          student_id: String(result.studentId ?? "").trim(),
          course_code: String(result.courseCode ?? "").trim(),
          session: String(result.session ?? "").trim(),
          semester: toNumber(result.semester),
          mark: toNumber(result.mark),
          letter: String(result.letter ?? "").trim(),
          point: toNumber(result.point),
          grade_points: toNumber(result.gradePoints ?? result.grade_points),
          passed_course: result.passedCourse === true ? 1 : result.passedCourse === false ? 0 : null,
          credits_earned: toNumber(result.creditsEarned ?? result.credits_earned),
        });
      }
    });

    insertTx();
    const updateId = Date.now();
    lastUpdateId = updateId;
    db.prepare("INSERT OR REPLACE INTO meta (key, value) VALUES ('last_update_id', ?)").run(
      String(updateId)
    );
    const offered = Array.isArray(payload.offeredCourseCodes)
      ? payload.offeredCourseCodes.map((code) => String(code ?? "").trim()).filter(Boolean)
      : [];
    db.prepare("INSERT OR REPLACE INTO meta (key, value) VALUES ('offered_course_codes', ?)").run(
      JSON.stringify(offered)
    );
    res.json({ ok: true, updateId });
  } catch (e) {
    res.status(500).json({ error: e.message ?? "Import failed" });
  }
});

app.get("/api/admin/update-status", (_req, res) => {
  res.json({ updateId: lastUpdateId, serverTime: Date.now() });
});

app.post("/api/auth/request-otp", (req, res) => {
  const { studentId, name } = req.body ?? {};
  const id = String(studentId ?? "").trim();
  const nameValue = String(name ?? "").trim().toLowerCase();
  if (!id) return res.status(400).json({ error: "Student ID is required." });

  const student = db.prepare("SELECT * FROM students WHERE student_id = ?").get(id);
  if (!student) return res.status(404).json({ error: "Student not found." });
  if (student.name && nameValue && String(student.name).trim().toLowerCase() !== nameValue) {
    return res.status(401).json({ error: "Name does not match." });
  }

  const otp = String(Math.floor(100000 + Math.random() * 900000));
  const otpHash = bcrypt.hashSync(otp, 10);
  const expiresAt = Date.now() + 5 * 60 * 1000;

  db.prepare("DELETE FROM otp_requests WHERE student_id = ?").run(id);
  db.prepare(
    "INSERT INTO otp_requests (student_id, otp_hash, expires_at, attempts) VALUES (?, ?, ?, 0)"
  ).run(id, otpHash, expiresAt);

  console.log(`[DEV OTP] Student ${id} OTP: ${otp}`);
  res.json({ ok: true, demoOtp: otp });
});

app.post("/api/auth/verify-otp", (req, res) => {
  const { studentId, otp } = req.body ?? {};
  const id = String(studentId ?? "").trim();
  const otpValue = String(otp ?? "").trim();
  if (!id || !otpValue) return res.status(400).json({ error: "Student ID and OTP are required." });

  const record = db.prepare("SELECT * FROM otp_requests WHERE student_id = ?").get(id);
  if (!record) return res.status(401).json({ error: "OTP not requested." });
  if (Date.now() > record.expires_at) {
    db.prepare("DELETE FROM otp_requests WHERE student_id = ?").run(id);
    return res.status(401).json({ error: "OTP expired." });
  }
  if (!bcrypt.compareSync(otpValue, record.otp_hash)) {
    db.prepare("UPDATE otp_requests SET attempts = attempts + 1 WHERE student_id = ?").run(id);
    return res.status(401).json({ error: "OTP invalid." });
  }

  db.prepare("DELETE FROM otp_requests WHERE student_id = ?").run(id);
  const token = jwt.sign({ studentId: id }, JWT_SECRET, { expiresIn: "2h" });
  res.json({ ok: true, token });
});

app.get("/api/me/profile", requireAuth, (req, res) => {
  const student = db.prepare("SELECT * FROM students WHERE student_id = ?").get(req.studentId);
  if (!student) return res.status(404).json({ error: "Student not found." });
  res.json({
    studentId: student.student_id,
    name: student.name,
    intake: student.intake,
    intakeYear: student.intake_year,
    intakeMonth: student.intake_month,
    program: student.program,
    status: student.status,
    cgpa: student.cgpa,
  });
});

app.get("/api/me/enrollment", requireAuth, (req, res) => {
  const student = db.prepare("SELECT * FROM students WHERE student_id = ?").get(req.studentId);
  if (!student) return res.status(404).json({ error: "Student not found." });
  const maxSemester = db
    .prepare("SELECT MAX(semester) AS sem FROM results WHERE student_id = ?")
    .get(req.studentId)?.sem;

  const offeredRow = db.prepare("SELECT value FROM meta WHERE key = 'offered_course_codes'").get();
  let offeredCodes = [];
  try {
    offeredCodes = JSON.parse(offeredRow?.value ?? "[]");
  } catch (_) {
    offeredCodes = [];
  }
  const offeredSet = new Set(
    (Array.isArray(offeredCodes) ? offeredCodes : [])
      .map((code) => String(code ?? "").trim())
      .filter(Boolean)
  );

  const allCourses = db
    .prepare("SELECT course_code, title FROM courses ORDER BY course_code")
    .all();

  const offeredCourses = offeredSet.size
    ? allCourses.filter((course) => offeredSet.has(course.course_code))
    : allCourses;

  const studentResults = db
    .prepare(
      `SELECT r.course_code, r.passed_course
       FROM results r
       WHERE r.student_id = ?`
    )
    .all(req.studentId);

  const takenSet = new Set(
    studentResults
      .filter((row) => row.course_code)
      .map((row) => String(row.course_code).trim())
  );

  const failedSet = new Set(
    studentResults
      .filter((row) => row.course_code && row.passed_course === 0)
      .map((row) => String(row.course_code).trim())
  );

  const notTaken = offeredCourses.filter((course) => !takenSet.has(course.course_code));
  const failed = offeredCourses.filter((course) => failedSet.has(course.course_code));

  res.json({
    studentId: student.student_id,
    program: student.program || "Programme",
    intake: student.intake || "",
    currentSemester: maxSemester ?? 1,
    status: student.status || "Active",
    notTaken,
    failed,
    offeredCount: offeredCourses.length,
  });
});

app.get("/api/me/results", requireAuth, (req, res) => {
  const results = db
    .prepare(
      `SELECT r.*, c.title, c.credits, c.is_mpu
       FROM results r
       LEFT JOIN courses c ON r.course_code = c.course_code
       WHERE r.student_id = ?
       ORDER BY r.session, r.semester`
    )
    .all(req.studentId);

  const byTerm = new Map();
  for (const row of results) {
    const key = `${row.session}||${row.semester ?? ""}`;
    if (!byTerm.has(key)) {
      byTerm.set(key, { session: row.session, semester: row.semester, rows: [] });
    }
    const isMPU = row.is_mpu === 1;
    let point = toNumber(row.point);
    let letter = row.letter;
    if (point === null && row.mark !== null) {
      const grade = gradeFromMark(row.mark, isMPU);
      point = grade.point;
      letter = grade.letter;
    }
    byTerm.get(key).rows.push({
      courseCode: row.course_code,
      title: row.title,
      credits: row.credits,
      mark: row.mark,
      letter,
      point,
      gradePoints: row.grade_points,
      passedCourse: row.passed_course === 1,
    });
  }

  const terms = [...byTerm.values()].map((term) => {
    const gpa = computeGpa(
      term.rows.map((row) => ({
        credits: row.credits,
        point: row.point,
        grade_points: row.gradePoints,
      }))
    );
    return { ...term, gpa };
  });

  const cgpa = computeGpa(
    results.map((row) => ({
      credits: row.credits,
      point: row.point,
      grade_points: row.grade_points,
    }))
  );

  res.json({ terms, cgpa });
});

app.listen(PORT, () => {
  console.log(`ODL portal sim running at http://localhost:${PORT}`);
});
