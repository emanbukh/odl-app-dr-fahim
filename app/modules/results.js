// app/modules/results.js
import { getAllRecords, replaceMany, upsertMany, STORES } from "../db/db.js";
import { computeStudentCgpa } from "./stats.js";
import { normalizeCourseCode } from "./course.js";

const ENROLLMENT_SLIP_LOCKS_KEY = "odlEnrollmentSlipLocks";
const ENROLLMENT_LOCKED_SLIPS_KEY = "odlEnrollmentLockedSlips";

/**
 * Normalizes truthy values from CSV:
 * e.g. "TRUE", "true", "1", "Yes" -> true
 */
function toBool(v) {
  if (v === true) return true;
  if (v === false) return false;
  const s = String(v ?? "").trim().toLowerCase();
  if (!s) return null;
  if (["true", "1", "yes", "y"].includes(s)) return true;
  if (["false", "0", "no", "n"].includes(s)) return false;
  return null;
}

function toNumberOrNull(v) {
  if (v === null || v === undefined) return null;
  const s = String(v).trim();
  if (s === "") return null;
  const n = Number(s);
  return Number.isFinite(n) ? n : null;
}

function parseSessionForOrder(value) {
  const raw = String(value ?? "").trim();
  if (!raw) return null;
  const match = raw.match(/^(\d{4})[./-](\d{1,2})$/);
  if (!match) return null;
  const year = match[1];
  const month = match[2].padStart(2, "0");
  const monthNum = Number(month);
  if (!Number.isFinite(monthNum) || monthNum < 1 || monthNum > 12) return null;
  return { intake: `${year}/${month}`, year, month, value: Number(`${year}${month}`) };
}

function normalizeSessionKey(value) {
  const parsed = parseSessionForOrder(value);
  if (parsed) return parsed.intake;
  return String(value ?? "").trim();
}

function normalizeKey(key) {
  return String(key ?? "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "");
}

function normalizePersonName(value) {
  const normalized = String(value ?? "")
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  if (!normalized) return "";
  const skipTokens = new Set(["bin", "binti", "bt", "bte", "binti", "binte"]);
  const filtered = normalized
    .split(" ")
    .filter((token) => token && !skipTokens.has(token));
  return filtered.join(" ").trim();
}

function parseSlipSnapshot(value) {
  const raw = String(value ?? "").trim();
  if (!raw) return null;
  try {
    const parsed = JSON.parse(raw);
    if (!parsed || typeof parsed !== "object") return null;
    return parsed;
  } catch (_) {
    return null;
  }
}

function getNameValue(row) {
  const preferred = [
    "fullname",
    "full name",
    "name",
    "studentname",
    "student name",
    "studentfullname",
    "student full name",
  ];
  const direct = String(getRowValue(row, preferred)).trim();
  if (direct) return direct;
  for (const key of Object.keys(row ?? {})) {
    const normalized = normalizeKey(key);
    if (!normalized.includes("name")) continue;
    if (normalized.includes("course") || normalized.includes("subject")) continue;
    const value = String(row[key] ?? "").trim();
    if (value) return value;
  }
  return "";
}

function getRowValue(row, keys) {
  for (const key of keys) {
    if (key in row) return row[key];
    const normalized = normalizeKey(key);
    if (normalized in row) return row[normalized];
  }
  return "";
}

export const defaultGradeScale = [
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

export const mpuGradeScale = [
  { min: 75, letter: "A", point: 4.0, passed: true },
  { min: 65, letter: "B", point: 3.0, passed: true },
  { min: 50, letter: "C", point: 2.0, passed: true },
  { min: 0, letter: "F", point: 0.0, passed: false },
];

function gradeFromMark(markValue, gradeScale = defaultGradeScale) {
  const mark = toNumberOrNull(markValue);
  if (mark === null) return { letter: "", point: null, passed: false };
  for (const grade of gradeScale) {
    if (mark >= grade.min) {
      return { letter: grade.letter, point: grade.point, passed: grade.passed };
    }
  }
  return { letter: "F", point: 0.0, passed: false };
}

export function getGradeForMark(markValue, isMPU) {
  const scale = isMPU ? mpuGradeScale : defaultGradeScale;
  return gradeFromMark(markValue, scale);
}

export function getSemesterForSession(studentId, session, allResults = null) {
  const id = String(studentId ?? "").trim();
  const parsedTarget = parseSessionForOrder(session);
  if (!id || !parsedTarget) return null;
  const results = allResults ?? [];
  const sessionValueMap = new Map();

  for (const row of results) {
    if (String(row?.studentId ?? "").trim() !== id) continue;
    const parsed = parseSessionForOrder(row?.session);
    if (!parsed) continue;
    sessionValueMap.set(parsed.intake, parsed.value);
  }

  sessionValueMap.set(parsedTarget.intake, parsedTarget.value);
  const ordered = [...sessionValueMap.entries()].sort((a, b) => a[1] - b[1]);
  if (!ordered.length) return 1;
  const index = ordered.findIndex(([key]) => key === parsedTarget.intake);
  if (index < 0) return null;
  return index + 1;
}

export async function resequenceSemestersForStudents(studentIds, allResults = null) {
  const normalizedIds = [...new Set(
    (studentIds ?? []).map((id) => String(id ?? "").trim()).filter(Boolean)
  )];
  if (!normalizedIds.length) {
    return { updatedCount: 0, conflictCount: 0, students: [] };
  }

  const results = allResults ?? (await getAllRecords(STORES.results));
  const resultsByStudent = new Map();
  for (const row of results) {
    const studentId = String(row?.studentId ?? "").trim();
    if (!studentId) continue;
    const entry = resultsByStudent.get(studentId) ?? [];
    entry.push(row);
    resultsByStudent.set(studentId, entry);
  }

  const candidates = [];
  const updatedStudents = new Set();

  for (const studentId of normalizedIds) {
    const rows = resultsByStudent.get(studentId) ?? [];
    if (!rows.length) continue;

    const sessionValueMap = new Map();
    for (const row of rows) {
      const sessionKey = normalizeSessionKey(row?.session);
      if (!sessionKey) continue;
      const parsed = parseSessionForOrder(row?.session);
      if (parsed) sessionValueMap.set(sessionKey, parsed.value);
    }

    const orderedSessions = [...sessionValueMap.entries()]
      .filter(([, value]) => Number.isFinite(value))
      .sort((a, b) => a[1] - b[1]);
    if (!orderedSessions.length) continue;

    const sessionToSemester = new Map();
    orderedSessions.forEach(([key], index) => {
      sessionToSemester.set(key, index + 1);
    });

    for (const row of rows) {
      const sessionKey = normalizeSessionKey(row?.session);
      if (!sessionKey || !sessionToSemester.has(sessionKey)) continue;
      const nextSemester = sessionToSemester.get(sessionKey);
      const currentSemester = toNumberOrNull(row?.semester);
      if (currentSemester === nextSemester) continue;

      const sessionRaw = String(row?.session ?? "").trim();
      const courseCode = String(row?.courseCode ?? "").trim();
      const studentIdRaw = String(row?.studentId ?? "").trim();
      if (!sessionRaw || !courseCode || !studentIdRaw) continue;

      const next = {
        ...row,
        semester: nextSemester,
        resultId: `${studentIdRaw}|${courseCode}|${sessionRaw}|${nextSemester}`,
      };
      candidates.push({ oldId: row?.resultId, next });
      updatedStudents.add(studentIdRaw);
    }
  }

  if (!candidates.length) {
    return { updatedCount: 0, conflictCount: 0, students: [] };
  }

  const existingIds = new Set(
    results.map((row) => String(row?.resultId ?? "").trim()).filter(Boolean)
  );
  const deleteSet = new Set(
    candidates.map((c) => String(c.oldId ?? "").trim()).filter(Boolean)
  );
  const finalDeletes = [];
  const finalRecords = [];
  const newIdSet = new Set();
  let conflictCount = 0;

  for (const candidate of candidates) {
    const oldId = String(candidate?.oldId ?? "").trim();
    const nextId = String(candidate?.next?.resultId ?? "").trim();
    if (!oldId || !nextId) continue;
    if (newIdSet.has(nextId)) {
      conflictCount += 1;
      continue;
    }
    if (existingIds.has(nextId) && !deleteSet.has(nextId)) {
      conflictCount += 1;
      continue;
    }
    newIdSet.add(nextId);
    finalDeletes.push(oldId);
    finalRecords.push(candidate.next);
  }

  if (!finalRecords.length) {
    return { updatedCount: 0, conflictCount, students: [] };
  }

  await replaceMany(STORES.results, finalDeletes, finalRecords);
  return { updatedCount: finalRecords.length, conflictCount, students: [...updatedStudents] };
}

/**
 * Map a CSV row (Joined.csv) to our stores.
 * Adjust column names here if your CSV headers differ.
 */
function mapRow(rawRow) {
  const row = {};
  for (const key of Object.keys(rawRow)) {
    row[normalizeKey(key)] = rawRow[key];
  }

  const studentId = String(
    getRowValue(row, ["id", "studentid", "studentno", "matric", "matricno"])
  ).trim();
  const courseCode = normalizeCourseCode(getRowValue(row, ["coursecode", "course", "subjectcode"]));
  const sessionRaw = String(getRowValue(row, ["session", "academicsession", "term"])).trim();
  const intakeRaw = String(getRowValue(row, ["intake", "intakecode", "intakesession"])).trim();
  const session = sessionRaw || intakeRaw;
  const semester = toNumberOrNull(getRowValue(row, ["semester", "sem"]));

  if (!studentId || !courseCode || !session) return null;

  const student = {
    studentId,
    name: getNameValue(row),
    intake: intakeRaw,
    intakeYear: String(getRowValue(row, ["intakeyear", "intakeyr"])).trim(),
    intakeMonth: String(getRowValue(row, ["intakemonth", "intakemo"])).trim(),
  };

  const course = {
    courseCode,
    title: String(getRowValue(row, ["title", "coursetitle", "subjecttitle"])).trim(),
    credits: toNumberOrNull(getRowValue(row, ["credits", "credit", "credithours"])) ?? 0,
    isMPUCourse: toBool(getRowValue(row, ["ismpucourse", "ismpu", "mpu"])),
  };

  const resultId = `${studentId}|${courseCode}|${session}|${semester ?? ""}`;

  const result = {
    resultId,
    studentId,
    courseCode,
    session,
    semester,
    mark: toNumberOrNull(getRowValue(row, ["mark", "score"])),
    letter: String(getRowValue(row, ["letter", "grade", "lettergrade"])).trim(),
    point: toNumberOrNull(getRowValue(row, ["point", "gradepoint"])),
    gradePoints: toNumberOrNull(getRowValue(row, ["gradepoints", "qualitypoints"])),
    passedCourse: toBool(getRowValue(row, ["passedcourse", "passed", "pass"])),
    creditsEarned: toNumberOrNull(getRowValue(row, ["creditsearned", "earnedcredits"])),
    studentName: student.name ?? "",
  };

  return { student, course, result };
}

function mapStudentOnlyRow(rawRow) {
  const row = {};
  for (const key of Object.keys(rawRow)) {
    row[normalizeKey(key)] = rawRow[key];
  }

  const studentId = String(
    getRowValue(row, ["id", "studentid", "studentno", "matric", "matricno"])
  ).trim();
  if (!studentId) return null;

  const courseCode = normalizeCourseCode(getRowValue(row, ["coursecode", "course", "subjectcode"]));
  const sessionRaw = String(getRowValue(row, ["session", "academicsession", "term"])).trim();
  if (courseCode || sessionRaw) return null;

  const intakeRaw = String(getRowValue(row, ["intake", "intakecode", "intakesession"])).trim();

  return {
    studentId,
    name: getNameValue(row),
    intake: intakeRaw,
    intakeYear: String(getRowValue(row, ["intakeyear", "intakeyr"])).trim(),
    intakeMonth: String(getRowValue(row, ["intakemonth", "intakemo"])).trim(),
  };
}

export async function importCSVFile(file, logFn) {
  if (!file) throw new Error("No file selected.");

  logFn(`Parsing: ${file.name} ...`);

  const parsed = await new Promise((resolve, reject) => {
    Papa.parse(file, {
      header: true,
      skipEmptyLines: true,
      dynamicTyping: false,
      complete: (res) => resolve(res),
      error: (err) => reject(err),
    });
  });

  if (parsed.errors?.length) {
    logFn(`CSV parse warnings/errors:`);
    for (const e of parsed.errors.slice(0, 10)) logFn(`- ${e.message}`);
  }

  const existingStudents = await getAllRecords(STORES.students);
  const existingById = new Map();
  for (const student of existingStudents) {
    const id = String(student?.studentId ?? "").trim();
    if (id) existingById.set(id, student);
  }

  const studentsMap = new Map();
  const coursesMap = new Map();
  const results = [];

  let skipped = 0;
  let studentOnlyCount = 0;
  let skippedMismatch = 0;
  const mismatches = [];
  const lockMetaById = new Map();
  const lockSlipById = new Map();
  let lockParseErrors = 0;

  const captureLockData = (studentId, row) => {
    const id = String(studentId ?? "").trim();
    if (!id) return;
    const lockedValue = getRowValue(row, ["sliplocked", "enrollmentlocked", "locked"]);
    const locked = toBool(lockedValue);
    const lockedAt = String(getRowValue(row, ["sliplockedat", "lockedat"])).trim();
    const lockSource = String(getRowValue(row, ["sliplocksource", "locksource", "lockedsource"])).trim();
    const slipJson = String(getRowValue(row, ["slipsnapshot", "lockedsnapshot", "enrollmentslip"])).trim();
    if (!locked && !lockedAt && !lockSource && !slipJson) return;

    const meta = lockMetaById.get(id) ?? {
      locked: false,
      lockedAt: "",
      source: "",
    };
    if (locked === true || slipJson) meta.locked = true;
    if (!meta.lockedAt && lockedAt) meta.lockedAt = lockedAt;
    if (!meta.source && lockSource) meta.source = lockSource;
    lockMetaById.set(id, meta);

    if (slipJson && !lockSlipById.has(id)) {
      const parsed = parseSlipSnapshot(slipJson);
      if (parsed) {
        if (!parsed.studentId) parsed.studentId = id;
        lockSlipById.set(id, parsed);
      } else {
        lockParseErrors += 1;
      }
    }
  };

  for (let index = 0; index < parsed.data.length; index += 1) {
    const row = parsed.data[index];
    const normalized = {};
    for (const key of Object.keys(row ?? {})) {
      normalized[normalizeKey(key)] = row[key];
    }
    const mapped = mapRow(row);
    if (mapped) {
      const studentId = String(mapped.student.studentId ?? "").trim();
      const existing = existingById.get(studentId);
      if (existing) {
        const existingName = normalizePersonName(existing.name);
        const incomingName = normalizePersonName(mapped.student.name);
        if (existingName && (!incomingName || incomingName !== existingName)) {
          skippedMismatch += 1;
          mismatches.push({
            rowNo: index + 2,
            studentId,
            existingName: existing.name ?? "",
            incomingName: mapped.student.name ?? "",
          });
          continue;
        }
      }

      captureLockData(studentId, normalized);
      studentsMap.set(mapped.student.studentId, mapped.student);
      coursesMap.set(mapped.course.courseCode, mapped.course);
      results.push(mapped.result);
      continue;
    }

    const studentOnly = mapStudentOnlyRow(row);
    if (!studentOnly) {
      skipped++;
      continue;
    }

    const studentId = String(studentOnly.studentId ?? "").trim();
    const existing = existingById.get(studentId);
    if (existing) {
      const existingName = normalizePersonName(existing.name);
      const incomingName = normalizePersonName(studentOnly.name);
      if (existingName && (!incomingName || incomingName !== existingName)) {
        skippedMismatch += 1;
        mismatches.push({
          rowNo: index + 2,
          studentId,
          existingName: existing.name ?? "",
          incomingName: studentOnly.name ?? "",
        });
        continue;
      }
    }

    captureLockData(studentId, normalized);
    studentOnlyCount += 1;
    studentsMap.set(studentOnly.studentId, studentOnly);
  }

  logFn(`Rows read: ${parsed.data.length}`);
  logFn(`Rows skipped (missing ID/CourseCode/Session): ${skipped}`);
  if (studentOnlyCount) {
    logFn(`Student-only rows imported: ${studentOnlyCount}`);
  }
  if (skippedMismatch) {
    logFn(`Rows skipped (ID+Name mismatch with existing): ${skippedMismatch}`);
    logFn(`Mismatched rows (first 10):`);
    for (const row of mismatches.slice(0, 10)) {
      logFn(
        `- Row ${row.rowNo}: ${row.studentId} | existing="${row.existingName}" vs uploaded="${row.incomingName}"`
      );
    }
  }
  logFn(`Upserting: ${studentsMap.size} students, ${coursesMap.size} courses, ${results.length} results ...`);

  await upsertMany(STORES.students, [...studentsMap.values()]);
  await upsertMany(STORES.courses, [...coursesMap.values()]);
  await upsertMany(STORES.results, results);

  if (lockMetaById.size || lockSlipById.size) {
    const lockPayload = {};
    for (const [studentId, meta] of lockMetaById.entries()) {
      if (!meta.locked) continue;
      lockPayload[studentId] = {
        locked: true,
        lockedAt: String(meta.lockedAt ?? ""),
        source: String(meta.source ?? ""),
      };
    }
    const slipPayload = {};
    for (const [studentId, slip] of lockSlipById.entries()) {
      slipPayload[studentId] = slip;
    }
    try {
      localStorage.setItem(ENROLLMENT_SLIP_LOCKS_KEY, JSON.stringify(lockPayload));
      localStorage.setItem(ENROLLMENT_LOCKED_SLIPS_KEY, JSON.stringify(slipPayload));
      logFn(`Enrollment slips restored: ${Object.keys(lockPayload).length} locked, ${Object.keys(slipPayload).length} snapshots.`);
      if (lockParseErrors) {
        logFn(`WARNING: ${lockParseErrors} slip snapshot(s) could not be parsed.`);
      }
    } catch (e) {
      logFn(`WARNING: Failed to store enrollment slip locks (${e.message ?? e}).`);
    }
  }

  logFn(`Import complete.`);
}

function validateSessionAndCourse(session, courseCode) {
  if (!session) throw new Error("Session is required.");
  if (!/^\d{4}\/(02|09)$/.test(session)) {
    throw new Error("Session must be in YYYY/MM format (month 02 or 09).");
  }
  const fixedCourseCode = normalizeCourseCode(courseCode ?? "");
  if (!fixedCourseCode) throw new Error("Course is required.");
  return fixedCourseCode;
}

async function parseNewResults({
  file,
  session,
  courseCode,
  gradeScale = defaultGradeScale,
}) {
  if (!file) throw new Error("No CSV file selected.");
  const fixedCourseCode = validateSessionAndCourse(session, courseCode);

  const parsed = await new Promise((resolve, reject) => {
    Papa.parse(file, {
      header: true,
      skipEmptyLines: true,
      dynamicTyping: false,
      complete: (res) => resolve(res),
      error: (err) => reject(err),
    });
  });

  const [students, courses, results] = await Promise.all([
    getAllRecords(STORES.students),
    getAllRecords(STORES.courses),
    getAllRecords(STORES.results),
  ]);

  const studentMap = new Map(students.map((s) => [String(s.studentId ?? "").trim(), s]));
  const studentsByName = new Map();
  for (const student of students) {
    const nameKey = normalizePersonName(student?.name);
    if (!nameKey) continue;
    const entry = studentsByName.get(nameKey) ?? [];
    entry.push(student);
    studentsByName.set(nameKey, entry);
  }
  const courseMap = new Map(
    courses.map((c) => [normalizeCourseCode(c.courseCode ?? ""), c])
  );
  const existingResultIds = new Set(results.map((r) => String(r.resultId ?? "").trim()));

  const maxSemesterByStudent = new Map();
  const sessionSemesterByStudent = new Map();
  for (const row of results) {
    const studentId = String(row.studentId ?? "").trim();
    if (!studentId) continue;
    const sessionKey = String(row.session ?? "").trim();
    const sem = toNumberOrNull(row.semester);
    if (sem === null) continue;
    const currentMax = maxSemesterByStudent.get(studentId) ?? 0;
    if (sem > currentMax) maxSemesterByStudent.set(studentId, sem);
    if (sessionKey) {
      sessionSemesterByStudent.set(`${studentId}|${sessionKey}`, sem);
    }
  }

  const newStudents = [];
  const updatedStudents = [];
  const newCourses = [];
  const newResults = [];
  const skippedDuplicates = [];
  let missingNameNewStudents = 0;
  let skippedMissing = 0;
  let skippedMismatch = 0;
  const mismatchedExisting = [];
  const issueRows = [];
  const issueSummary = {
    missing_mark: 0,
    missing_identifier: 0,
    name_not_found: 0,
    name_ambiguous: 0,
    name_mismatch: 0,
  };

  const addIssueRow = ({ rowNo, studentId, name, mark, reason, reasonCode }) => {
    issueRows.push({
      rowNo,
      studentId: studentId ?? "",
      name: name ?? "",
      mark: mark ?? "",
      reason,
    });
    if (reasonCode && reasonCode in issueSummary) {
      issueSummary[reasonCode] += 1;
    }
  };

  for (let index = 0; index < parsed.data.length; index += 1) {
    const rawRow = parsed.data[index];
    const row = {};
    for (const key of Object.keys(rawRow)) {
      row[normalizeKey(key)] = rawRow[key];
    }

    const studentIdRaw = String(getRowValue(row, ["id", "studentid", "studentno"])).trim();
    const name = getNameValue(row);
    const nameKey = normalizePersonName(name);
    const courseCodeFinal = fixedCourseCode;
    const markValue = getRowValue(row, ["mark", "score"]);
    const mark = toNumberOrNull(markValue);

    if (mark === null || !courseCodeFinal) {
      addIssueRow({
        rowNo: index + 2,
        studentId: studentIdRaw,
        name,
        mark: markValue,
        reason: "Missing mark",
        reasonCode: "missing_mark",
      });
      skippedMissing += 1;
      continue;
    }

    let studentId = studentIdRaw;
    let matchedByName = false;
    if (!studentId) {
      if (!nameKey) {
        addIssueRow({
          rowNo: index + 2,
          studentId: "",
          name,
          mark: markValue,
          reason: "Missing student ID and name",
          reasonCode: "missing_identifier",
        });
        skippedMissing += 1;
        continue;
      }
      const matches = studentsByName.get(nameKey) ?? [];
      if (matches.length === 1) {
        studentId = String(matches[0].studentId ?? "").trim();
        matchedByName = true;
      } else if (matches.length === 0) {
        addIssueRow({
          rowNo: index + 2,
          studentId: "",
          name,
          mark: markValue,
          reason: "Name not found in existing students",
          reasonCode: "name_not_found",
        });
        skippedMissing += 1;
        continue;
      } else {
        const ids = matches
          .map((student) => String(student.studentId ?? "").trim())
          .filter(Boolean)
          .join(", ");
        addIssueRow({
          rowNo: index + 2,
          studentId: "",
          name,
          mark: markValue,
          reason: ids ? `Name matches multiple students (${ids})` : "Name matches multiple students",
          reasonCode: "name_ambiguous",
        });
        skippedMissing += 1;
        continue;
      }
    }

    let student = studentMap.get(studentId);
    const isNewStudent = !student;
    if (student) {
      const existingName = normalizePersonName(student.name);
      const incomingName = normalizePersonName(name);
      if (existingName && (!incomingName || incomingName !== existingName)) {
        skippedMismatch += 1;
        mismatchedExisting.push({
          rowNo: index + 2,
          studentId,
          existingName: student.name ?? "",
          incomingName: name ?? "",
          reason: incomingName ? "name_mismatch" : "missing_name",
        });
        addIssueRow({
          rowNo: index + 2,
          studentId,
          name,
          mark: markValue,
          reason: "Name does not match existing student ID",
          reasonCode: "name_mismatch",
        });
        continue;
      }
    }

    if (!student) {
      if (matchedByName) {
        addIssueRow({
          rowNo: index + 2,
          studentId,
          name,
          mark: markValue,
          reason: "Name matched but student ID not found",
          reasonCode: "name_not_found",
        });
        skippedMissing += 1;
        continue;
      }
      student = {
        studentId,
        name,
        intake: "",
        intakeYear: "",
        intakeMonth: "",
      };
      studentMap.set(studentId, student);
      newStudents.push(student);
      if (!name) missingNameNewStudents += 1;
    } else if (name) {
      const currentName = String(student.name ?? "").trim();
      if (!currentName) {
        student = { ...student, name };
        studentMap.set(studentId, student);
        updatedStudents.push(student);
      }
    }

    if (!courseMap.has(courseCodeFinal)) {
      const course = {
        courseCode: courseCodeFinal,
        title: "",
        credits: 0,
        isMPUCourse: null,
      };
      courseMap.set(courseCodeFinal, course);
      newCourses.push(course);
    }

    const sessionKey = `${studentId}|${session}`;
    let semester = sessionSemesterByStudent.get(sessionKey);
    if (!semester) {
      const currentMax = maxSemesterByStudent.get(studentId) ?? 0;
      semester = currentMax + 1;
      sessionSemesterByStudent.set(sessionKey, semester);
      maxSemesterByStudent.set(studentId, semester);
    }

    const resultId = `${studentId}|${courseCodeFinal}|${session}|${semester}`;
    if (existingResultIds.has(resultId)) {
      skippedDuplicates.push(resultId);
      continue;
    }
    existingResultIds.add(resultId);

    const isMPU = courseMap.get(courseCodeFinal)?.isMPUCourse === true;
    const scale = isMPU ? mpuGradeScale : gradeScale;
    const grade = gradeFromMark(mark, scale);
    newResults.push({
      resultId,
      studentId,
      courseCode: courseCodeFinal,
      session,
      semester,
      mark,
      letter: grade.letter,
      point: grade.point,
      gradePoints: null,
      passedCourse: grade.passed,
      creditsEarned: null,
      isNewStudent,
      studentName: name || student.name || "",
    });
  }

  return {
    parsed,
    newStudents,
    updatedStudents,
    newCourses,
    newResults,
    skippedMissing,
    skippedDuplicates,
    missingNameNewStudents,
    skippedMismatch,
      mismatchedExisting,
      issueRows,
      issueSummary,
    };
}

export async function previewNewResults({ file, session, courseCode, gradeScale }) {
  const {
    parsed,
    newStudents,
    updatedStudents,
    newCourses,
    newResults,
    skippedMissing,
    skippedDuplicates,
    missingNameNewStudents,
    skippedMismatch,
    mismatchedExisting,
    issueRows,
    issueSummary,
  } = await parseNewResults({ file, session, courseCode, gradeScale });

  const affectedMap = new Map();
  for (const result of newResults) {
    const entry = affectedMap.get(result.studentId) ?? {
      studentId: result.studentId,
      name: result.studentName ?? "",
      isNew: result.isNewStudent,
      count: 0,
    };
    entry.count += 1;
    if (result.isNewStudent) entry.isNew = true;
    affectedMap.set(result.studentId, entry);
  }

  return {
    parsedRows: parsed.data.length,
    skippedMissing,
    skippedDuplicates,
    newStudents,
    updatedStudents,
    newCourses,
    newResults,
    missingNameNewStudents,
    skippedMismatch,
    mismatchedExisting,
    issueRows,
    issueSummary,
    affectedStudents: [...affectedMap.values()].sort((a, b) => a.studentId.localeCompare(b.studentId)),
  };
}

export async function uploadNewResults({
  file,
  session,
  courseCode,
  logFn,
  gradeScale = defaultGradeScale,
  selectedStudentIds,
}) {
  const fixedCourseCode = validateSessionAndCourse(session, courseCode);
  logFn(`Parsing: ${file.name} ...`);

  const {
    parsed,
    newStudents,
    updatedStudents,
    newCourses,
    newResults,
    skippedMissing,
    skippedDuplicates,
    missingNameNewStudents,
    skippedMismatch,
    mismatchedExisting,
    issueRows,
    issueSummary,
  } = await parseNewResults({ file, session, courseCode: fixedCourseCode, gradeScale });

  const selectedSet = selectedStudentIds ? new Set(selectedStudentIds) : null;
  const filteredResults = selectedSet
    ? newResults.filter((row) => selectedSet.has(row.studentId))
    : newResults;
  const filteredStudentIds = new Set(filteredResults.map((row) => row.studentId));
  const filteredStudents = selectedSet
    ? newStudents.filter((student) => selectedSet.has(student.studentId))
    : newStudents;
  const filteredUpdatedStudents = selectedSet
    ? updatedStudents.filter((student) => selectedSet.has(student.studentId))
    : updatedStudents;

  logFn(`Rows read: ${parsed.data.length}`);
  logFn(`Rows with issues: ${issueRows.length}`);
  if (issueRows.length && issueSummary) {
    if (issueSummary.missing_mark) logFn(`- Missing mark: ${issueSummary.missing_mark}`);
    if (issueSummary.missing_identifier) logFn(`- Missing ID + name: ${issueSummary.missing_identifier}`);
    if (issueSummary.name_not_found) logFn(`- Name not found: ${issueSummary.name_not_found}`);
    if (issueSummary.name_ambiguous) logFn(`- Name matches multiple students: ${issueSummary.name_ambiguous}`);
    if (issueSummary.name_mismatch) logFn(`- Name mismatch with ID: ${issueSummary.name_mismatch}`);
  } else if (skippedMissing) {
    logFn(`Rows skipped (missing ID/Name/Mark): ${skippedMissing}`);
  }
  if (skippedMismatch) {
    logFn(`Rows skipped (ID+Name mismatch with existing): ${skippedMismatch}`);
    logFn(`Mismatched rows (first 10):`);
    for (const row of mismatchedExisting.slice(0, 10)) {
      logFn(
        `- Row ${row.rowNo}: ${row.studentId} | existing="${row.existingName}" vs uploaded="${row.incomingName}"`
      );
    }
  }
  logFn(`New students: ${filteredStudents.length}`);
  if (missingNameNewStudents) {
    logFn(`New students missing names: ${missingNameNewStudents}`);
  }
  logFn(`New courses: ${newCourses.length}`);
  logFn(`New results: ${filteredResults.length}`);
  logFn(`Duplicates skipped: ${skippedDuplicates.length}`);
  if (skippedDuplicates.length) {
    logFn(`Duplicate IDs (first 10):`);
    for (const id of skippedDuplicates.slice(0, 10)) logFn(`- ${id}`);
  }

  const sanitizedResults = filteredResults.map(({ isNewStudent, ...rest }) => rest);

  if (filteredStudents.length) await upsertMany(STORES.students, filteredStudents);
  if (filteredUpdatedStudents.length) await upsertMany(STORES.students, filteredUpdatedStudents);
  if (newCourses.length) await upsertMany(STORES.courses, newCourses);
  if (sanitizedResults.length) await upsertMany(STORES.results, sanitizedResults);

  const [allResults, allCourses, allStudents] = await Promise.all([
    getAllRecords(STORES.results),
    getAllRecords(STORES.courses),
    getAllRecords(STORES.students),
  ]);
  const cgpaByStudent = computeStudentCgpa(allResults, allCourses);
  const studentMap = new Map(
    allStudents.map((student) => [String(student.studentId ?? "").trim(), student])
  );
  const cgpaUpdates = [];
  for (const [studentId, cgpa] of cgpaByStudent.entries()) {
    const existing = studentMap.get(String(studentId ?? "").trim());
    if (existing) {
      cgpaUpdates.push({ ...existing, cgpa });
    } else if (studentId) {
      cgpaUpdates.push({
        studentId,
        name: "",
        intake: "",
        intakeYear: "",
        intakeMonth: "",
        cgpa,
      });
    }
  }
  if (cgpaUpdates.length) await upsertMany(STORES.students, cgpaUpdates);

  if (filteredStudentIds.size) {
    await resequenceSemestersForStudents([...filteredStudentIds]);
  }

  return { selectedCount: selectedSet ? filteredStudentIds.size : filteredResults.length };
}
