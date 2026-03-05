// app/modules/backup.js
import { getAllRecords, STORES } from "../db/db.js";
import { importCSVFile } from "./results.js";

const ENROLLMENT_SLIP_LOCKS_KEY = "odlEnrollmentSlipLocks";
const ENROLLMENT_LOCKED_SLIPS_KEY = "odlEnrollmentLockedSlips";

function escapeCsv(value) {
  const s = String(value ?? "");
  if (/[",\n]/.test(s)) return `"${s.replace(/"/g, '""')}"`;
  return s;
}

function appendLog(el, msg) {
  if (!el) return;
  el.textContent += `${msg}\n`;
}

function readLockPayloads() {
  let locks = {};
  let slips = {};
  try {
    locks = JSON.parse(localStorage.getItem(ENROLLMENT_SLIP_LOCKS_KEY) || "{}");
  } catch (_) {
    locks = {};
  }
  try {
    slips = JSON.parse(localStorage.getItem(ENROLLMENT_LOCKED_SLIPS_KEY) || "{}");
  } catch (_) {
    slips = {};
  }

  const lockMetaById = new Map();
  for (const [studentId, meta] of Object.entries(locks ?? {})) {
    const id = String(studentId ?? "").trim();
    if (!id || !meta || typeof meta !== "object") continue;
    if (meta.locked !== true) continue;
    lockMetaById.set(id, {
      locked: true,
      lockedAt: String(meta.lockedAt ?? ""),
      source: String(meta.source ?? ""),
    });
  }

  const lockSlipById = new Map();
  for (const [studentId, slip] of Object.entries(slips ?? {})) {
    const id = String(studentId ?? "").trim();
    if (!id || !slip || typeof slip !== "object") continue;
    lockSlipById.set(id, slip);
  }

  return { lockMetaById, lockSlipById };
}

export function buildBackupCsv({ students, courses, results }) {
  const studentMap = new Map(
    students.map((s) => [String(s.studentId ?? "").trim(), s])
  );
  const courseMap = new Map(
    courses.map((c) => [String(c.courseCode ?? "").trim(), c])
  );
  const { lockMetaById, lockSlipById } = readLockPayloads();
  const lockWritten = new Set();

  const header = [
    "ID",
    "Name",
    "Intake",
    "IntakeYear",
    "IntakeMonth",
    "CourseCode",
    "Title",
    "Credits",
    "IsMPUCourse",
    "Session",
    "Semester",
    "Mark",
    "Letter",
    "Point",
    "GradePoints",
    "PassedCourse",
    "CreditsEarned",
    "SlipLocked",
    "SlipLockedAt",
    "SlipLockSource",
    "SlipSnapshot",
  ];

  const resultStudentIds = new Set(
    results.map((result) => String(result.studentId ?? "").trim()).filter(Boolean)
  );

  const buildLockColumns = (studentId) => {
    const id = String(studentId ?? "").trim();
    if (!id || lockWritten.has(id)) {
      return ["", "", "", ""];
    }
    const meta = lockMetaById.get(id);
    const slip = lockSlipById.get(id);
    if (!meta && !slip) return ["", "", "", ""];
    lockWritten.add(id);
    const locked = meta?.locked === true || Boolean(slip);
    const lockedAt = meta?.lockedAt ?? "";
    const source = meta?.source ?? "";
    const slipJson = slip ? JSON.stringify(slip) : "";
    return [locked ? "TRUE" : "", lockedAt, source, slipJson];
  };

  const rows = results.map((result) => {
    const studentId = String(result.studentId ?? "").trim();
    const courseCode = String(result.courseCode ?? "").trim();
    const student = studentMap.get(studentId) ?? {};
    const course = courseMap.get(courseCode) ?? {};
    const lockColumns = buildLockColumns(studentId);

    return [
      studentId,
      student.name ?? "",
      student.intake ?? "",
      student.intakeYear ?? "",
      student.intakeMonth ?? "",
      courseCode,
      course.title ?? "",
      course.credits ?? "",
      course.isMPUCourse ?? "",
      result.session ?? "",
      result.semester ?? "",
      result.mark ?? "",
      result.letter ?? "",
      result.point ?? "",
      result.gradePoints ?? "",
      result.passedCourse ?? "",
      result.creditsEarned ?? "",
      ...lockColumns,
    ].map(escapeCsv);
  });

  const studentOnlyRows = students
    .filter((student) => {
      const id = String(student?.studentId ?? "").trim();
      return id && !resultStudentIds.has(id);
    })
    .map((student) => {
      const lockColumns = buildLockColumns(student.studentId);
      return [
        String(student.studentId ?? "").trim(),
        student.name ?? "",
        student.intake ?? "",
        student.intakeYear ?? "",
        student.intakeMonth ?? "",
        "",
        "",
        "",
        "",
        "",
        "",
        "",
        "",
        "",
        "",
        "",
        "",
        ...lockColumns,
      ].map(escapeCsv);
    });

  return [
    header.map(escapeCsv).join(","),
    ...rows.map((r) => r.join(",")),
    ...studentOnlyRows.map((r) => r.join(",")),
  ].join("\n");
}

export function initBackup({ onDataChanged } = {}) {
  const btnBackupDownload = document.getElementById("btnBackupDownload");
  const btnBackupImport = document.getElementById("btnBackupImport");
  const elBackupLog = document.getElementById("backupLog");
  const elBackupFile = document.getElementById("backupFile");

  if (btnBackupDownload) {
    btnBackupDownload.addEventListener("click", async () => {
      if (elBackupLog) elBackupLog.textContent = "";
      try {
        const [students, courses, results] = await Promise.all([
          getAllRecords(STORES.students),
          getAllRecords(STORES.courses),
          getAllRecords(STORES.results),
        ]);
        const csv = buildBackupCsv({ students, courses, results });
        const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
        const url = URL.createObjectURL(blob);
        const link = document.createElement("a");
        link.href = url;
        link.download = "odl-backup.csv";
        document.body.appendChild(link);
        link.click();
        link.remove();
        URL.revokeObjectURL(url);
        appendLog(elBackupLog, "Backup downloaded.");
      } catch (e) {
        appendLog(elBackupLog, `ERROR: ${e.message ?? e}`);
      }
    });
  }

  if (btnBackupImport) {
    btnBackupImport.addEventListener("click", async () => {
      if (elBackupLog) elBackupLog.textContent = "";
      try {
        const file = elBackupFile?.files?.[0];
        await importCSVFile(file, (msg) => appendLog(elBackupLog, msg));
        if (onDataChanged) await onDataChanged();
      } catch (e) {
        appendLog(elBackupLog, `ERROR: ${e.message ?? e}`);
      }
    });
  }
}
