// app/main.js
import { deleteCourse, listCourses, normalizeCourseCode, upsertCourse } from "./modules/course.js";
import { initImports } from "./modules/imports-ui.js";
import { initStudentImport } from "./modules/student-import.js";
import { computeStudentCgpa } from "./modules/stats.js";
import { getGradeForMark } from "./modules/results.js";
import { initBackup } from "./modules/backup.js";
import { initPortalExport } from "./modules/portal-export.js";
import { assignBalancedCourses } from "./modules/enrollment-optimizer.js";
import { deleteRecord, getAllRecords, STORES, upsertMany } from "./db/db.js";

const btnRefresh = document.getElementById("btnRefresh");

const elTotalStudents = document.getElementById("totalStudents");
const elTotalCourses = document.getElementById("totalCourses");
const elActiveCourses = document.getElementById("activeCourses");
const elTotalResults = document.getElementById("totalResults");
const elActiveStudents = document.getElementById("activeStudents");
const elStudentsWithFailures = document.getElementById("studentsWithFailures");
const elStudentsIncomplete = document.getElementById("studentsIncomplete");
const elStudentsLowCgpa = document.getElementById("studentsLowCgpa");
const elStudentsResultIssues = document.getElementById("studentsResultIssues");
const elAvgResultsPerStudent = document.getElementById("avgResultsPerStudent");
const btnViewFailures = document.getElementById("btnViewFailures");
const btnViewIncomplete = document.getElementById("btnViewIncomplete");
const btnViewLowCgpa = document.getElementById("btnViewLowCgpa");
const btnViewResultIssues = document.getElementById("btnViewResultIssues");
const elIssueTypeSelect = document.getElementById("issueTypeSelect");
const elIssueSearch = document.getElementById("issueSearch");
const elIssueListTitle = document.getElementById("issueListTitle");
const elIssueListCount = document.getElementById("issueListCount");
const elIssueListEmpty = document.getElementById("issueListEmpty");
const elIssueListTable = document.getElementById("issueListTable");
const elIssueListBody = document.getElementById("issueListBody");
const elIssuePager = document.getElementById("issuePager");
const btnIssuePrev = document.getElementById("btnIssuePrev");
const btnIssueNext = document.getElementById("btnIssueNext");
const elIssuePageMeta = document.getElementById("issuePageMeta");
const elLastUpdated = document.getElementById("lastUpdated");
const elCourseCoverageModal = document.getElementById("courseCoverageModal");
const btnCloseCoverage = document.getElementById("btnCloseCoverage");
const btnExportCoverage = document.getElementById("btnExportCoverage");
const elCoverageCourseMeta = document.getElementById("coverageCourseMeta");
const elCoverageTakenCount = document.getElementById("coverageTakenCount");
const elCoverageNotTakenCount = document.getElementById("coverageNotTakenCount");
const elCoverageTakenList = document.getElementById("coverageTakenList");
const elCoverageNotTakenList = document.getElementById("coverageNotTakenList");
const elCourseForm = document.getElementById("courseForm");
const elCourseCode = document.getElementById("courseCode");
const elCourseTitle = document.getElementById("courseTitle");
const elCourseCredits = document.getElementById("courseCredits");
const elCourseMPU = document.getElementById("courseMPU");
const btnCourseReset = document.getElementById("btnCourseReset");
const elCourseFormHint = document.getElementById("courseFormHint");
const elCourseListBody = document.getElementById("courseListBody");
const elCourseListTable = document.getElementById("courseListTable");
const elCourseListEmpty = document.getElementById("courseListEmpty");
const elStudentListTable = document.getElementById("studentListTable");
const elStudentListBody = document.getElementById("studentListBody");
const elStudentListEmpty = document.getElementById("studentListEmpty");
const elStudentListSearch = document.getElementById("studentListSearch");
const elStudentListIntake = document.getElementById("studentListIntake");
const elStudentListStatus = document.getElementById("studentListStatus");
const elStudentListPager = document.getElementById("studentListPager");
const btnStudentListPrev = document.getElementById("btnStudentListPrev");
const btnStudentListNext = document.getElementById("btnStudentListNext");
const elStudentListPageMeta = document.getElementById("studentListPageMeta");
const elStudentProfileDrawer = document.getElementById("studentProfileDrawer");
const btnCloseStudentProfile = document.getElementById("btnCloseStudentProfile");
const elStudentProfileMeta = document.getElementById("studentProfileMeta");
const elStudentProfileInfo = document.getElementById("studentProfileInfo");
const elStudentProfileSummary = document.getElementById("studentProfileSummary");
const elStudentProfileStatus = document.getElementById("studentProfileStatus");
const btnProfileResultSlip = document.getElementById("btnProfileResultSlip");
const btnProfileEnrollmentSlip = document.getElementById("btnProfileEnrollmentSlip");
const btnProfileEdit = document.getElementById("btnProfileEdit");
const btnProfileDelete = document.getElementById("btnProfileDelete");
const elDeletedStudentEmpty = document.getElementById("deletedStudentEmpty");
const elDeletedStudentTable = document.getElementById("deletedStudentTable");
const elDeletedStudentBody = document.getElementById("deletedStudentBody");
const btnConfirmResultSlips = document.getElementById("btnConfirmResultSlips");
const btnConfirmEnrollmentSlips = document.getElementById("btnConfirmEnrollmentSlips");
const elResultSlipConfirmStatus = document.getElementById("resultSlipConfirmStatus");
const elEnrollmentSlipConfirmStatus = document.getElementById("enrollmentSlipConfirmStatus");
const elCourseTotalCount = document.getElementById("courseTotalCount");
const elCourseTotalCredits = document.getElementById("courseTotalCredits");
const tabButtons = document.querySelectorAll("[data-tab]");
const tabPanels = document.querySelectorAll(".tab-panel");
const resultSubtabButtons = document.querySelectorAll(".result-subtab-btn");
const resultSubtabPanels = document.querySelectorAll(".result-subtab-panel");
const enrollmentSubtabButtons = document.querySelectorAll(".enrollment-subtab-btn");
const enrollmentSubtabPanels = document.querySelectorAll(".enrollment-subtab-panel");
const elResultTotal = document.getElementById("resultTotal");
const elResultPassRate = document.getElementById("resultPassRate");
const elResultAvgMark = document.getElementById("resultAvgMark");
const elResultStudents = document.getElementById("resultStudents");
const elResultSessionSelect = document.getElementById("resultSessionSelect");
const elResultSearch = document.getElementById("resultSearch");
const elResultViewMode = document.getElementById("resultViewMode");
const elResultTable = document.getElementById("resultTable");
const elResultTableBody = document.getElementById("resultTableBody");
const elResultTableEmpty = document.getElementById("resultTableEmpty");
const elCourseDetailStats = document.getElementById("courseDetailStats");
const elCourseDetailRecords = document.getElementById("courseDetailRecords");
const elCourseDetailPassRate = document.getElementById("courseDetailPassRate");
const elCourseDetailAvgMark = document.getElementById("courseDetailAvgMark");
const elCourseDetailStudents = document.getElementById("courseDetailStudents");
const elCourseDetailHeader = document.getElementById("courseDetailHeader");
const elCourseDetailTitle = document.getElementById("courseDetailTitle");
const btnBackToCourseList = document.getElementById("btnBackToCourseList");
const elStudentEditModal = document.getElementById("studentEditModal");
const btnCloseStudentEdit = document.getElementById("btnCloseStudentEdit");
const btnSaveStudentEdit = document.getElementById("btnSaveStudentEdit");
const elEditStudentId = document.getElementById("editStudentId");
const elEditStudentName = document.getElementById("editStudentName");
const elEditStudentIc = document.getElementById("editStudentIc");
const elEditStudentIntake = document.getElementById("editStudentIntake");
const btnPrintSlip = document.getElementById("btnPrintSlip");
const elResultSlipModal = document.getElementById("resultSlipModal");
const btnCloseSlip = document.getElementById("btnCloseSlip");
const elSlipStudentMeta = document.getElementById("slipStudentMeta");
const elSlipCGPA = document.getElementById("slipCGPA");
const elSlipTerms = document.getElementById("slipTerms");
const elResultPager = document.getElementById("resultPager");
const btnPrevPage = document.getElementById("btnPrevPage");
const btnNextPage = document.getElementById("btnNextPage");
const elResultPageMeta = document.getElementById("resultPageMeta");
const elMarkEditModal = document.getElementById("markEditModal");
const btnCloseMarkEdit = document.getElementById("btnCloseMarkEdit");
const btnConfirmMarkEdit = document.getElementById("btnConfirmMarkEdit");
const elMarkEditStudent = document.getElementById("markEditStudent");
const elMarkEditCourse = document.getElementById("markEditCourse");
const elMarkEditSession = document.getElementById("markEditSession");
const elMarkEditSemester = document.getElementById("markEditSemester");
const elMarkEditCurrent = document.getElementById("markEditCurrent");
const elMarkEditInput = document.getElementById("markEditInput");
const elMarkEditLetter = document.getElementById("markEditLetter");
const elMarkEditPoint = document.getElementById("markEditPoint");
const elMarkEditStatus = document.getElementById("markEditStatus");
const elMarkEditWarning = document.getElementById("markEditWarning");
const elMarkEditVerify = document.getElementById("markEditVerify");
const elAddResultModal = document.getElementById("addResultModal");
const btnCloseAddResult = document.getElementById("btnCloseAddResult");
const btnConfirmAddResult = document.getElementById("btnConfirmAddResult");
const elAddResultStudent = document.getElementById("addResultStudent");
const elAddResultCourse = document.getElementById("addResultCourse");
const elAddResultSession = document.getElementById("addResultSession");
const elAddResultSemester = document.getElementById("addResultSemester");
const elAddResultMark = document.getElementById("addResultMark");
const elAddResultLetter = document.getElementById("addResultLetter");
const elAddResultPoint = document.getElementById("addResultPoint");
const elAddResultStatus = document.getElementById("addResultStatus");
const elAddResultWarning = document.getElementById("addResultWarning");
const elAddResultVerify = document.getElementById("addResultVerify");
const elEnrollmentCourseGrid = document.getElementById("enrollmentCourseGrid");
const elEnrollmentMaxPerStudent = document.getElementById("enrollmentMaxPerStudent");
const btnEnrollmentSelectAll = document.getElementById("btnEnrollmentSelectAll");
const btnEnrollmentClearAll = document.getElementById("btnEnrollmentClearAll");
const btnEnrollmentGenerateAllKeys = document.getElementById("btnEnrollmentGenerateAllKeys");
const btnEnrollmentResetAllKeys = document.getElementById("btnEnrollmentResetAllKeys");
const btnEnrollmentBuild = document.getElementById("btnEnrollmentBuild");
const elEnrollmentTargetAll = document.getElementById("enrollmentTargetAll");
const elEnrollmentTargetIntakesMode = document.getElementById("enrollmentTargetIntakesMode");
const elEnrollmentTargetIntakes = document.getElementById("enrollmentTargetIntakes");
const elEnrollmentTargetHint = document.getElementById("enrollmentTargetHint");
const btnEnrollmentLockGroups = document.getElementById("btnEnrollmentLockGroups");
const btnEnrollmentUnlockGroups = document.getElementById("btnEnrollmentUnlockGroups");
const elEnrollmentForceRebuildLocked = document.getElementById("enrollmentForceRebuildLocked");
const elEnrollmentLockHint = document.getElementById("enrollmentLockHint");
const btnEnrollmentBulkPrint = document.getElementById("btnEnrollmentBulkPrint");
const elEnrollmentSelectedCount = document.getElementById("enrollmentSelectedCount");
const elEnrollmentSearch = document.getElementById("enrollmentSearch");
const elEnrollmentCourseFilterButtons = document.getElementById("enrollmentCourseFilterButtons");
const elEnrollmentCourseFilterHint = document.getElementById("enrollmentCourseFilterHint");
const elEnrollmentCourseFilterEmpty = document.getElementById("enrollmentCourseFilterEmpty");
const elEnrollmentListIntake = document.getElementById("enrollmentListIntake");
const elEnrollmentStudentCount = document.getElementById("enrollmentStudentCount");
const elEnrollmentListEmpty = document.getElementById("enrollmentListEmpty");
const elEnrollmentListTable = document.getElementById("enrollmentListTable");
const elEnrollmentListBody = document.getElementById("enrollmentListBody");
const elEnrollmentPager = document.getElementById("enrollmentPager");
const btnEnrollmentPrev = document.getElementById("btnEnrollmentPrev");
const btnEnrollmentNext = document.getElementById("btnEnrollmentNext");
const elEnrollmentPageMeta = document.getElementById("enrollmentPageMeta");
const elEnrollmentStatsTable = document.getElementById("enrollmentStatsTable");
const elEnrollmentStatsBody = document.getElementById("enrollmentStatsBody");
const elEnrollmentStatsEmpty = document.getElementById("enrollmentStatsEmpty");
const elEnrollmentStatsCourseCount = document.getElementById("enrollmentStatsCourseCount");
const elEnrollmentStatsStudentCount = document.getElementById("enrollmentStatsStudentCount");
const elEnrollmentRedFlagTable = document.getElementById("enrollmentRedFlagTable");
const elEnrollmentRedFlagBody = document.getElementById("enrollmentRedFlagBody");
const elEnrollmentRedFlagEmpty = document.getElementById("enrollmentRedFlagEmpty");
const elEnrollmentRedFlagCount = document.getElementById("enrollmentRedFlagCount");
const elEnrollmentSlipModal = document.getElementById("enrollmentSlipModal");
const btnCloseEnrollmentSlip = document.getElementById("btnCloseEnrollmentSlip");
const btnPrintEnrollmentSlip = document.getElementById("btnPrintEnrollmentSlip");
const btnEnrollmentLockSlip = document.getElementById("btnEnrollmentLockSlip");
const elEnrollmentSlipCount = document.getElementById("enrollmentSlipCount");
const elEnrollmentSlipStudentMeta = document.getElementById("enrollmentSlipStudentMeta");
const elEnrollmentSlipBody = document.getElementById("enrollmentSlipBody");
const elEnrollmentBulkLog = document.getElementById("enrollmentBulkLog");
const elReportType = document.getElementById("reportType");
const elReportIntake = document.getElementById("reportIntake");
const elReportCourse = document.getElementById("reportCourse");
const elReportIncludeCurrent = document.getElementById("reportIncludeCurrent");
const elReportIncludePast = document.getElementById("reportIncludePast");
const btnReportGenerate = document.getElementById("btnReportGenerate");
const btnReportDownload = document.getElementById("btnReportDownload");
const elReportRowCount = document.getElementById("reportRowCount");
const elReportPreviewEmpty = document.getElementById("reportPreviewEmpty");
const elReportPreviewTable = document.getElementById("reportPreviewTable");
const elReportPreviewHead = document.getElementById("reportPreviewHead");
const elReportPreviewBody = document.getElementById("reportPreviewBody");
const elReportPager = document.getElementById("reportPager");
const elReportPageMeta = document.getElementById("reportPageMeta");
const btnReportPrev = document.getElementById("btnReportPrev");
const btnReportNext = document.getElementById("btnReportNext");
const numberFmt = new Intl.NumberFormat();
const percentFmt = new Intl.NumberFormat(undefined, { style: "percent", maximumFractionDigits: 1 });
const resultState = {
  students: [],
  courses: [],
  results: [],
};
const resultPaging = {
  page: 1,
  pageSize: 20,
};
const studentListPaging = {
  page: 1,
  pageSize: 20,
};
const studentInlineStatusMessages = new Map();
let activeCourseDetailCode = "";
let activeStudentEditId = "";
let activeMarkEditId = "";
let activeMarkEditValue = "";
let activeMarkEditRow = null;
let activeAddResultStudentId = "";
let activeEnrollmentStudentId = "";
let activeEnrollmentInlineId = "";
let activeProfileStudentId = "";
let activeStudentInlineId = "";
let activeResultInlineId = "";
const resultSortState = {
  key: "",
  dir: "asc",
};
const courseCoverageDetails = new Map();
let activeCoverageCourse = "";
let activeCourseEditCode = "";
const studentIssueLists = {
  failures: [],
  incomplete: [],
  lowCgpa: [],
  resultIssues: [],
};
const CONFIRM_RESULT_KEY = "odlConfirmResultSlipsAt";
const CONFIRM_ENROLL_KEY = "odlConfirmEnrollmentSlipsAt";
const ENROLLMENT_CONFIRM_TOKEN_KEY = "odlConfirmEnrollmentSlipsToken";
const ENROLLMENT_DATA_REVISION_KEY = "odlEnrollmentDataRevision";
const ENROLLMENT_SELECTION_KEY = "odlEnrollmentOfferedCourses";
const ENROLLMENT_KEYS_KEY = "odlEnrollmentCourseKeys";

const issueTitles = {
  failures: "Students with failures",
  incomplete: "Incomplete student info",
  lowCgpa: "CGPA below 2.00",
  resultIssues: "Result issues",
};
const issuePaging = {
  page: 1,
  pageSize: 20,
};
const reportPaging = {
  page: 1,
  pageSize: 20,
};
const enrollmentState = {
  slips: [],
  offeredCourseCodes: [],
};
const enrollmentPaging = {
  page: 1,
  pageSize: 20,
};
const ENROLLMENT_MAX_OFFERED = 15;
const ENROLLMENT_TARGET_CREDITS = 90;
const ENROLLMENT_MAX_COURSES_KEY = "odlEnrollmentMaxCoursesPerStudent";
const ENROLLMENT_PRIORITY_COURSES_KEY = "odlEnrollmentPriorityCourses";
const ENROLLMENT_CUSTOM_SELECTIONS_KEY = "odlEnrollmentCustomSelections";
const ENROLLMENT_SLIP_LOCKS_KEY = "odlEnrollmentSlipLocks";
const ENROLLMENT_LOCKED_SLIPS_KEY = "odlEnrollmentLockedSlips";
const ENROLLMENT_FORCE_REBUILD_LOCKED_KEY = "odlEnrollmentForceRebuildLocked";
const ENROLLMENT_PRIORITY_WEIGHT = 50;
const ENROLLMENT_REPEAT_PRIORITY_WEIGHT = 200;
const ENROLLMENT_PRIORITY_FOR_ALL = true;
const ENROLLMENT_DECLARATION_TEXT =
  "I confirm that I have reviewed the courses listed above. Enrollment is subject to academic regulations, prerequisite checks, and final approval by the institution.";
const enrollmentKeys = new Map();
let enrollmentDataRevision = "";
const ENROLLMENT_OPTIMIZER_MODE = "auto";
const ENROLLMENT_OPTIMIZER_MAX_SLOTS = 1200;
let enrollmentMaxPerStudent = 6;
let enrollmentPrioritySet = new Set();
let enrollmentCustomSelections = new Map();
let enrollmentSlipLocks = new Map();
let enrollmentLockedSlipSnapshots = new Map();
let enrollmentForceRebuildLocked = false;
let enrollmentCourseFilterSet = new Set();
let enrollmentRedFlagStudents = [];
const ENROLLMENT_TARGET_REDFLAG_KEY = "__redflag__";
let enrollmentTargetMode = "all";
let enrollmentTargetGroupSet = new Set();
let enrollmentSlipViewMode = "live";
const REPORT_PREVIEW_LIMIT = 100;
let reportRows = [];
let reportColumns = [];
let reportAutoGenerated = false;

function setPrintMode(mode) {
  if (!document?.body) return;
  document.body.dataset.printMode = mode;
}

function clearPrintMode() {
  if (!document?.body) return;
  delete document.body.dataset.printMode;
}

window.addEventListener("afterprint", clearPrintMode);

function loadEnrollmentKeys() {
  try {
    const stored = JSON.parse(localStorage.getItem(ENROLLMENT_KEYS_KEY) || "{}");
    if (stored && typeof stored === "object" && !Array.isArray(stored)) {
      for (const [code, key] of Object.entries(stored)) {
        const normalized = normalizeCourseCode(code);
        if (!normalized) continue;
        if (key === null) {
          enrollmentKeys.set(normalized, null);
          continue;
        }
        const value = String(key ?? "").trim();
        if (value) enrollmentKeys.set(normalized, value);
      }
    }
  } catch (e) {
    console.warn("Failed to load enrollment keys", e);
  }
}

function loadEnrollmentOptimizerSettings() {
  try {
    const stored = Number(localStorage.getItem(ENROLLMENT_MAX_COURSES_KEY));
    if (Number.isFinite(stored) && stored >= 1) {
      enrollmentMaxPerStudent = Math.min(15, Math.floor(stored));
    }
  } catch (_) {
    // ignore
  }

  try {
    const stored = JSON.parse(localStorage.getItem(ENROLLMENT_PRIORITY_COURSES_KEY) || "[]");
    if (Array.isArray(stored)) {
      enrollmentPrioritySet = new Set(stored.map((code) => normalizeCourseCode(code)).filter(Boolean));
    }
  } catch (_) {
    enrollmentPrioritySet = new Set();
  }

  if (elEnrollmentMaxPerStudent) {
    elEnrollmentMaxPerStudent.value = String(enrollmentMaxPerStudent);
  }
}

function saveEnrollmentOptimizerSettings() {
  try {
    localStorage.setItem(ENROLLMENT_MAX_COURSES_KEY, String(enrollmentMaxPerStudent));
  } catch (_) {
    // ignore
  }
  try {
    localStorage.setItem(
      ENROLLMENT_PRIORITY_COURSES_KEY,
      JSON.stringify([...enrollmentPrioritySet.values()])
    );
  } catch (_) {
    // ignore
  }
}

function loadEnrollmentCustomSelections() {
  try {
    const stored = JSON.parse(localStorage.getItem(ENROLLMENT_CUSTOM_SELECTIONS_KEY) || "{}");
    if (stored && typeof stored === "object" && !Array.isArray(stored)) {
      const next = new Map();
      for (const [studentId, list] of Object.entries(stored)) {
        if (!Array.isArray(list)) continue;
        const codes = list.map((code) => normalizeCourseCode(code)).filter(Boolean);
        next.set(String(studentId ?? "").trim(), codes);
      }
      enrollmentCustomSelections = next;
      return;
    }
  } catch (_) {
    // ignore
  }
  enrollmentCustomSelections = new Map();
}

function saveEnrollmentCustomSelections() {
  try {
    const payload = {};
    for (const [studentId, list] of enrollmentCustomSelections.entries()) {
      payload[studentId] = list;
    }
    localStorage.setItem(ENROLLMENT_CUSTOM_SELECTIONS_KEY, JSON.stringify(payload));
  } catch (_) {
    // ignore
  }
}

function loadEnrollmentSlipLocks() {
  try {
    const stored = JSON.parse(localStorage.getItem(ENROLLMENT_SLIP_LOCKS_KEY) || "{}");
    if (stored && typeof stored === "object" && !Array.isArray(stored)) {
      const next = new Map();
      for (const [studentId, meta] of Object.entries(stored)) {
        if (!meta || typeof meta !== "object") continue;
        if (!meta.locked) continue;
        next.set(String(studentId ?? "").trim(), {
          locked: true,
          lockedAt: String(meta.lockedAt ?? ""),
          source: String(meta.source ?? ""),
        });
      }
      enrollmentSlipLocks = next;
      return;
    }
  } catch (_) {
    // ignore
  }
  enrollmentSlipLocks = new Map();
}

function saveEnrollmentSlipLocks() {
  try {
    const payload = {};
    for (const [studentId, meta] of enrollmentSlipLocks.entries()) {
      payload[studentId] = {
        locked: true,
        lockedAt: String(meta?.lockedAt ?? ""),
        source: String(meta?.source ?? ""),
      };
    }
    localStorage.setItem(ENROLLMENT_SLIP_LOCKS_KEY, JSON.stringify(payload));
  } catch (_) {
    // ignore
  }
}

function loadEnrollmentLockedSlips() {
  try {
    const stored = JSON.parse(localStorage.getItem(ENROLLMENT_LOCKED_SLIPS_KEY) || "{}");
    if (stored && typeof stored === "object" && !Array.isArray(stored)) {
      const next = new Map();
      for (const [studentId, slip] of Object.entries(stored)) {
        if (!slip || typeof slip !== "object") continue;
        const id = String(studentId ?? "").trim();
        if (!id) continue;
        const courses = Array.isArray(slip.courses) ? slip.courses : [];
        next.set(id, {
          studentId: id,
          name: String(slip.name ?? "").trim(),
          intake: String(slip.intake ?? "").trim(),
          currentSemester: slip.currentSemester ?? null,
          completedCredits: slip.completedCredits ?? 0,
          isCustom: slip.isCustom === true,
          redFlag: slip.redFlag === true,
          redFlagReason: String(slip.redFlagReason ?? ""),
          courses: courses.map((course) => ({
            courseCode: normalizeCourseCode(course?.courseCode ?? ""),
            title: String(course?.title ?? "").trim(),
            enrollKey: course?.enrollKey ?? "",
            remark: String(course?.remark ?? ""),
          })),
        });
      }
      enrollmentLockedSlipSnapshots = next;
      return;
    }
  } catch (_) {
    // ignore
  }
  enrollmentLockedSlipSnapshots = new Map();
}

function saveEnrollmentLockedSlips() {
  try {
    const payload = {};
    for (const [studentId, slip] of enrollmentLockedSlipSnapshots.entries()) {
      payload[studentId] = slip;
    }
    localStorage.setItem(ENROLLMENT_LOCKED_SLIPS_KEY, JSON.stringify(payload));
  } catch (_) {
    // ignore
  }
}

function cloneSlipSnapshot(slip) {
  return {
    studentId: String(slip?.studentId ?? "").trim(),
    name: String(slip?.name ?? "").trim(),
    intake: String(slip?.intake ?? "").trim(),
    currentSemester: slip?.currentSemester ?? null,
    completedCredits: slip?.completedCredits ?? 0,
    isCustom: slip?.isCustom === true,
    redFlag: slip?.redFlag === true,
    redFlagReason: String(slip?.redFlagReason ?? ""),
    courses: Array.isArray(slip?.courses)
      ? slip.courses.map((course) => ({
          courseCode: normalizeCourseCode(course?.courseCode ?? ""),
          title: String(course?.title ?? "").trim(),
          enrollKey: course?.enrollKey ?? "",
          remark: String(course?.remark ?? ""),
        }))
      : [],
  };
}

function getLockedSlipSnapshot(studentId) {
  const id = normalizeStudentId(studentId);
  if (!id) return null;
  return enrollmentLockedSlipSnapshots.get(id) ?? null;
}

function setLockedSlipSnapshot(studentId, slip, options = {}) {
  const id = normalizeStudentId(studentId);
  if (!id || !slip) return false;
  enrollmentLockedSlipSnapshots.set(id, cloneSlipSnapshot(slip));
  if (options.save !== false) saveEnrollmentLockedSlips();
  return true;
}

function removeLockedSlipSnapshot(studentId, options = {}) {
  const id = normalizeStudentId(studentId);
  if (!id) return false;
  const removed = enrollmentLockedSlipSnapshots.delete(id);
  if (removed && options.save !== false) saveEnrollmentLockedSlips();
  return removed;
}

function captureLockedSlipSnapshots(studentIds = []) {
  if (!Array.isArray(studentIds) || !studentIds.length) return 0;
  const slipMap = new Map(
    (enrollmentState.slips ?? [])
      .map((slip) => [String(slip?.studentId ?? "").trim(), slip])
      .filter(([id]) => id)
  );
  let count = 0;
  for (const rawId of studentIds) {
    const id = normalizeStudentId(rawId);
    if (!id) continue;
    const slip = slipMap.get(id);
    if (!slip) continue;
    setLockedSlipSnapshot(id, slip, { save: false });
    count += 1;
  }
  if (count) saveEnrollmentLockedSlips();
  return count;
}

function loadEnrollmentLockPreferences() {
  try {
    const stored = localStorage.getItem(ENROLLMENT_FORCE_REBUILD_LOCKED_KEY);
    enrollmentForceRebuildLocked = stored === "1";
  } catch (_) {
    enrollmentForceRebuildLocked = false;
  }
  if (elEnrollmentForceRebuildLocked) {
    elEnrollmentForceRebuildLocked.checked = enrollmentForceRebuildLocked;
  }
}

function saveEnrollmentLockPreferences() {
  try {
    localStorage.setItem(
      ENROLLMENT_FORCE_REBUILD_LOCKED_KEY,
      enrollmentForceRebuildLocked ? "1" : "0"
    );
  } catch (_) {
    // ignore
  }
}

function getEnrollmentPrioritySet(offeredCourseCodes = []) {
  const offered = new Set(offeredCourseCodes.map((code) => normalizeCourseCode(code)).filter(Boolean));
  const filtered = new Set([...enrollmentPrioritySet].filter((code) => offered.has(code)));
  enrollmentPrioritySet = filtered;
  return filtered;
}

function saveEnrollmentKeys() {
  try {
    const payload = {};
    for (const [code, key] of enrollmentKeys.entries()) {
      payload[code] = key === null ? null : String(key ?? "");
    }
    localStorage.setItem(ENROLLMENT_KEYS_KEY, JSON.stringify(payload));
  } catch (e) {
    console.warn("Failed to save enrollment keys", e);
  }
}

function getEnrollmentKey(courseCode) {
  const code = normalizeCourseCode(courseCode);
  if (!code) return "";
  if (!enrollmentKeys.size) loadEnrollmentKeys();
  return enrollmentKeys.has(code) ? enrollmentKeys.get(code) : undefined;
}

function generateEnrollmentKey(length = 10) {
  const size = Math.max(1, Number(length) || 10);
  let output = "";
  if (window.crypto?.getRandomValues) {
    const bytes = new Uint8Array(32);
    while (output.length < size) {
      window.crypto.getRandomValues(bytes);
      for (const value of bytes) {
        if (value >= 250) continue;
        output += String(value % 10);
        if (output.length >= size) break;
      }
    }
    return output;
  }
  while (output.length < size) {
    output += String(Math.floor(Math.random() * 10));
  }
  return output;
}

function applyEnrollmentKeyValue(courseCode, nextValue, selectedSet = null) {
  const code = normalizeCourseCode(courseCode);
  if (!code) return { code: "", changed: false, selectedChanged: false };
  const hasPrevious = enrollmentKeys.has(code);
  const previousValue = hasPrevious ? enrollmentKeys.get(code) : undefined;
  const normalizedNextValue = nextValue === undefined
    ? undefined
    : nextValue === null
      ? null
      : String(nextValue).trim() || null;

  let changed = false;
  if (normalizedNextValue === undefined) {
    changed = enrollmentKeys.delete(code);
  } else if (!hasPrevious || previousValue !== normalizedNextValue) {
    enrollmentKeys.set(code, normalizedNextValue);
    changed = true;
  }

  const selectedCodes = selectedSet ?? new Set(getEnrollmentSelectionSnapshot());
  const selectedChanged = changed && selectedCodes.has(code);
  return { code, changed, selectedChanged };
}

function setEnrollmentKeyValue(courseCode, nextValue) {
  const selectedCodes = new Set(getEnrollmentSelectionSnapshot());
  const { changed, selectedChanged } = applyEnrollmentKeyValue(courseCode, nextValue, selectedCodes);
  if (!changed) return false;
  saveEnrollmentKeys();
  if (selectedChanged) invalidateEnrollmentConfirmation();
  return true;
}

function generateEnrollmentKeysForAllCourses() {
  const courseCodes = [...buildCourseMap().keys()].sort();
  if (!courseCodes.length) {
    alert("No courses available to generate keys.");
    return;
  }
  if (!enrollmentKeys.size) loadEnrollmentKeys();
  const selectedCodes = new Set(getEnrollmentSelectionSnapshot());
  let changedCount = 0;
  let selectedChanged = false;
  for (const code of courseCodes) {
    const nextKey = generateEnrollmentKey(10);
    const update = applyEnrollmentKeyValue(code, nextKey, selectedCodes);
    if (!update.changed) continue;
    changedCount += 1;
    selectedChanged = selectedChanged || update.selectedChanged;
  }
  if (!changedCount) return;
  saveEnrollmentKeys();
  if (selectedChanged) invalidateEnrollmentConfirmation();
  populateEnrollmentCourseOptions();
  alert(`Generated 10-digit enrollment keys for ${numberFmt.format(changedCount)} course(s).`);
}

function resetEnrollmentKeysForAllCourses() {
  if (!enrollmentKeys.size) loadEnrollmentKeys();
  if (!enrollmentKeys.size) {
    alert("No enrollment keys to reset.");
    return;
  }
  const ok = window.confirm("Reset enrollment keys for all courses?");
  if (!ok) return;
  const selectedCodes = new Set(getEnrollmentSelectionSnapshot());
  const selectedChanged = [...selectedCodes].some((code) => enrollmentKeys.has(code));
  enrollmentKeys.clear();
  saveEnrollmentKeys();
  if (selectedChanged) invalidateEnrollmentConfirmation();
  populateEnrollmentCourseOptions();
}

function getStoredEnrollmentSelection() {
  try {
    const stored = JSON.parse(localStorage.getItem(ENROLLMENT_SELECTION_KEY) || "[]");
    if (!Array.isArray(stored)) return [];
    return stored
      .map((code) => normalizeCourseCode(code))
      .filter(Boolean)
      .slice(0, ENROLLMENT_MAX_OFFERED);
  } catch (e) {
    return [];
  }
}

function getEnrollmentSelectionSnapshot() {
  const raw = (enrollmentState.offeredCourseCodes ?? []).length
    ? (enrollmentState.offeredCourseCodes ?? [])
    : getStoredEnrollmentSelection();
  return [...new Set(raw.map((code) => normalizeCourseCode(code)).filter(Boolean))]
    .sort()
    .slice(0, ENROLLMENT_MAX_OFFERED);
}

function hashTextFNV1a(value, seed = 0x811c9dc5) {
  let hash = seed >>> 0;
  for (let i = 0; i < value.length; i += 1) {
    hash ^= value.charCodeAt(i);
    hash = Math.imul(hash, 0x01000193) >>> 0;
  }
  return hash >>> 0;
}

function toHashHex(value) {
  return (value >>> 0).toString(16).padStart(8, "0");
}

function computeEnrollmentDataRevision(students = [], courses = [], results = []) {
  const studentRows = students
    .map((student) => {
      const studentId = String(student?.studentId ?? "").trim();
      const name = String(student?.name ?? "").trim();
      const intake = String(student?.intake ?? "").trim();
      const deleted = student?.isDeleted ? "1" : "0";
      return `${studentId}|${name}|${intake}|${deleted}`;
    })
    .sort();
  const courseRows = courses
    .map((course) => {
      const code = normalizeCourseCode(course?.courseCode ?? "");
      const title = String(course?.title ?? "").trim();
      const credits = String(course?.credits ?? "").trim();
      return `${code}|${title}|${credits}`;
    })
    .sort();
  const resultRows = results
    .map((row) => {
      const studentId = String(row?.studentId ?? "").trim();
      const courseCode = normalizeCourseCode(row?.courseCode ?? "");
      const semester = String(row?.semester ?? "").trim();
      const mark = String(row?.mark ?? "").trim();
      const letter = String(row?.letter ?? "").trim().toUpperCase();
      const passed = row?.passedCourse === true ? "1" : row?.passedCourse === false ? "0" : "";
      const creditsEarned = String(row?.creditsEarned ?? "").trim();
      return `${studentId}|${courseCode}|${semester}|${mark}|${letter}|${passed}|${creditsEarned}`;
    })
    .sort();

  let hash = hashTextFNV1a("enrollment-data-revision:v1\n");
  for (const row of studentRows) hash = hashTextFNV1a(`S|${row}\n`, hash);
  for (const row of courseRows) hash = hashTextFNV1a(`C|${row}\n`, hash);
  for (const row of resultRows) hash = hashTextFNV1a(`R|${row}\n`, hash);
  return toHashHex(hash);
}

function computeEnrollmentConfirmToken() {
  const selectedCodes = getEnrollmentSelectionSnapshot();
  if (!enrollmentKeys.size) loadEnrollmentKeys();
  const revision = localStorage.getItem(ENROLLMENT_DATA_REVISION_KEY) || enrollmentDataRevision || "";

  let hash = hashTextFNV1a("enrollment-confirm-token:v1\n");
  hash = hashTextFNV1a(`V|${revision}\n`, hash);
  for (const code of selectedCodes) {
    const key = getEnrollmentKey(code);
    const keyValue = key === null ? "~" : key === undefined ? "" : String(key).trim();
    hash = hashTextFNV1a(`C|${code}\n`, hash);
    hash = hashTextFNV1a(`K|${code}|${keyValue}\n`, hash);
  }
  return toHashHex(hash);
}

function invalidateEnrollmentConfirmation() {
  let changed = false;
  if (localStorage.getItem(CONFIRM_ENROLL_KEY) !== null) {
    localStorage.removeItem(CONFIRM_ENROLL_KEY);
    changed = true;
  }
  if (localStorage.getItem(ENROLLMENT_CONFIRM_TOKEN_KEY) !== null) {
    localStorage.removeItem(ENROLLMENT_CONFIRM_TOKEN_KEY);
    changed = true;
  }
  if (changed) updateSlipConfirmStatus();
}

function setActiveResultSubtab(subtabId) {
  resultSubtabButtons.forEach((button) => {
    const isActive = button.dataset.subtab === subtabId;
    button.classList.toggle("active", isActive);
    button.setAttribute("aria-selected", isActive ? "true" : "false");
  });
  resultSubtabPanels.forEach((panel) => {
    const isActive = panel.id === `result-subtab-${subtabId}`;
    panel.classList.toggle("active", isActive);
  });
}

function setActiveEnrollmentSubtab(subtabId) {
  enrollmentSubtabButtons.forEach((button) => {
    const isActive = button.dataset.enrollmentSubtab === subtabId;
    button.classList.toggle("active", isActive);
    button.setAttribute("aria-selected", isActive ? "true" : "false");
  });
  enrollmentSubtabPanels.forEach((panel) => {
    const isActive = panel.id === `enrollment-subtab-${subtabId}`;
    panel.classList.toggle("active", isActive);
  });
}


function setText(el, value) {
  if (!el) return;
  el.textContent = value;
}

function setLastUpdated() {
  if (!elLastUpdated) return;
  const now = new Date();
  elLastUpdated.textContent = `Last updated: ${now.toLocaleString()}`;
}

function toNumber(value) {
  const n = Number(value);
  return Number.isFinite(n) ? n : null;
}

function formatNumber(value, digits = 1) {
  const n = toNumber(value);
  if (n === null) return "--";
  return n.toFixed(digits);
}

function formatYearMonth(value, fallback = "-") {
  const raw = String(value ?? "").trim();
  if (!raw) return fallback;
  const match = raw.match(/^(\d{4})[./-](\d{1,2})$/);
  if (match) {
    const month = match[2].padStart(2, "0");
    return `${match[1]}/${month}`;
  }
  return raw;
}

function parseYearMonth(value) {
  const raw = String(value ?? "").trim();
  if (!raw) return null;
  const match = raw.match(/^(\d{4})[./-](\d{1,2})$/);
  if (!match) return null;
  const year = match[1];
  const month = match[2].padStart(2, "0");
  if (!/^(02|09)$/.test(month)) return null;
  return { intake: `${year}/${month}`, year, month };
}

function isFailedResult(result) {
  if (result.passedCourse === false) return true;
  const letter = String(result.letter ?? "").trim().toUpperCase();
  if (letter === "F" || letter === "FAIL") return true;
  const mark = toNumber(result.mark);
  if (mark !== null && mark < 40) return true;
  return false;
}

function isPassedResult(result) {
  if (result.passedCourse === true) return true;
  if (result.passedCourse === false) return false;
  const letter = String(result.letter ?? "").trim().toUpperCase();
  if (letter) return letter !== "F" && letter !== "FAIL";
  const mark = toNumber(result.mark);
  if (mark !== null) return mark >= 40;
  return false;
}

function isMissingMark(result) {
  if (result.mark === null || result.mark === undefined) return true;
  const s = String(result.mark).trim();
  return s === "";
}

function getGradePoints(result, courseCredits) {
  const direct = toNumber(result.gradePoints);
  const point = toNumber(result.point);
  const credits = toNumber(courseCredits);
  if (direct !== null) {
    if (direct === 0 && point !== null && point > 0 && credits !== null && credits > 0) {
      return point * credits;
    }
    return direct;
  }
  if (point !== null && credits !== null) return point * credits;
  return null;
}

function buildStudentMap() {
  const map = new Map();
  for (const student of resultState.students) {
    const studentId = String(student.studentId ?? "").trim();
    if (!studentId) continue;
    map.set(studentId, student);
  }
  return map;
}

async function reconcileStudentsFromResults(students, results) {
  const existing = new Map();
  for (const student of students) {
    const id = String(student?.studentId ?? "").trim();
    if (!id) continue;
    existing.set(id, student);
  }

  const toInsert = new Map();
  const toUpdate = new Map();
  for (const row of results) {
    const studentId = String(row?.studentId ?? "").trim();
    if (!studentId) continue;
    const nameFromResult = String(row?.studentName ?? row?.name ?? "").trim();
    if (existing.has(studentId)) {
      const current = existing.get(studentId);
      if ((!current?.name || !String(current.name).trim()) && nameFromResult) {
        toUpdate.set(studentId, { ...current, name: nameFromResult });
      }
      continue;
    }
    if (toInsert.has(studentId)) continue;
    toInsert.set(studentId, {
      studentId,
      name: nameFromResult,
      ic: "",
      intake: "",
      intakeYear: "",
      intakeMonth: "",
    });
  }

  const inserts = [...toInsert.values()];
  const updates = [...toUpdate.values()];
  if (inserts.length) await upsertMany(STORES.students, inserts);
  if (updates.length) await upsertMany(STORES.students, updates);

  if (!inserts.length && !updates.length) return students;
  return [...students, ...inserts].map((student) => {
    const id = String(student?.studentId ?? "").trim();
    return updates.find((item) => String(item.studentId ?? "").trim() === id) ?? student;
  });
}

async function backfillResultNamesFromStudents(students, results) {
  const nameById = new Map();
  for (const student of students) {
    const id = String(student?.studentId ?? "").trim();
    const name = String(student?.name ?? "").trim();
    if (id && name) nameById.set(id, name);
  }

  const updates = [];
  const nextResults = results.map((row) => {
    const studentId = String(row?.studentId ?? "").trim();
    if (!studentId) return row;
    const existingName = String(row?.studentName ?? "").trim();
    if (existingName) return row;
    const name = nameById.get(studentId);
    if (!name) return row;
    const updated = { ...row, studentName: name };
    updates.push(updated);
    return updated;
  });

  if (updates.length) {
    await upsertMany(STORES.results, updates);
  }
  return nextResults;
}

function buildCourseMap() {
  const map = new Map();
  for (const course of resultState.courses) {
    const code = normalizeCourseCode(course.courseCode ?? "");
    if (!code) continue;
    map.set(code, course);
  }
  return map;
}

function renderResultStats(results, studentIds) {
  const total = results.length;
  const passed = results.filter((row) => row.passedCourse === true).length;
  const failed = results.filter((row) => isFailedResult(row)).length;
  const missingMark = results.filter((row) => isMissingMark(row)).length;
  const passRate = total ? passed / total : 0;
  const marks = results.map((row) => toNumber(row.mark)).filter((mark) => mark !== null);
  const avgMark = marks.length ? marks.reduce((a, b) => a + b, 0) / marks.length : null;

  setText(elResultTotal, numberFmt.format(total));
  setText(elResultPassRate, total ? percentFmt.format(passRate) : "--");
  if (elResultPassRate) {
    elResultPassRate.title = `Failed: ${numberFmt.format(failed)} | No mark: ${numberFmt.format(missingMark)}`;
  }
  setText(elResultAvgMark, avgMark === null ? "--" : avgMark.toFixed(1));
  setText(elResultStudents, numberFmt.format(studentIds.size));
}

function renderResultTableHeader(mode) {
  const thead = elResultTable.querySelector("thead");
  if (!thead) return;
  if (mode === "students") {
    thead.innerHTML = `
      <tr>
        <th class="sortable" data-sort="studentId">Student ID</th>
        <th class="sortable" data-sort="studentName">Student Name</th>
        <th class="sortable" data-sort="currentSemester">Current Semester</th>
        <th class="sortable" data-sort="cgpa">CGPA</th>
        <th class="sortable" data-sort="records">Records</th>
        <th>Action</th>
      </tr>
    `;
    return;
  }

  if (mode === "courses") {
    thead.innerHTML = `
      <tr>
        <th class="sortable" data-sort="courseLabel">Course</th>
        <th class="sortable" data-sort="records">Records</th>
        <th class="sortable" data-sort="passRate">Pass Rate</th>
        <th class="sortable" data-sort="avgMark">Avg Mark</th>
        <th>Action</th>
      </tr>
    `;
    return;
  }

  const actionHeader = mode === "detail" ? "<th>Action</th>" : "";
  thead.innerHTML = `
    <tr>
      <th class="sortable" data-sort="studentId">Student ID</th>
      <th class="sortable" data-sort="studentName">Student Name</th>
      <th class="sortable" data-sort="session">Session</th>
      <th class="sortable" data-sort="semester">Semester</th>
      <th class="sortable" data-sort="mark">Mark</th>
      <th class="sortable" data-sort="letter">Letter</th>
      <th class="sortable" data-sort="point">Point</th>
      <th class="sortable" data-sort="status">Status</th>
      ${actionHeader}
    </tr>
  `;
}

  function renderResultTable(rows, mode) {
    if (!elResultTable || !elResultTableBody || !elResultTableEmpty || !elResultPager) return;
    if (!rows.length) {
      elResultTable.style.display = "none";
      elResultTableEmpty.style.display = "block";
      elResultPager.style.display = "none";
      return;
    }
    const redFlagSet = getRedFlagStudentSet();

  elResultTableEmpty.style.display = "none";
  elResultTable.style.display = "table";
  elResultTableBody.textContent = "";
  renderResultTableHeader(mode);
  const thead = elResultTable.querySelector("thead");
  if (thead) {
    thead.querySelectorAll("th.sortable").forEach((th) => {
      th.addEventListener("click", () => {
        const key = th.dataset.sort || "";
        if (!key) return;
        if (resultSortState.key === key) {
          resultSortState.dir = resultSortState.dir === "asc" ? "desc" : "asc";
        } else {
          resultSortState.key = key;
          resultSortState.dir = "asc";
        }
        renderResultsView();
      });
    });
  }


  if (mode === "students") {
    elResultPager.style.display = "flex";
    for (const row of rows) {
      const tr = document.createElement("tr");

      const tdStudentId = document.createElement("td");
      tdStudentId.textContent = row.studentId ?? "-";
      tr.appendChild(tdStudentId);

      const tdStudentName = document.createElement("td");
      tdStudentName.textContent = row.studentName || "-";
      if (row.studentId && redFlagSet.has(String(row.studentId ?? "").trim())) {
        appendRedFlagTag(tdStudentName, row.studentId);
      }
      tr.appendChild(tdStudentName);

      const tdSemester = document.createElement("td");
      tdSemester.textContent = row.currentSemester ?? "-";
      tr.appendChild(tdSemester);

      const tdCgpa = document.createElement("td");
      tdCgpa.textContent = row.cgpa ?? "--";
      tr.appendChild(tdCgpa);

      const tdCount = document.createElement("td");
      tdCount.textContent = numberFmt.format(row.records);
      tr.appendChild(tdCount);

      const tdAction = document.createElement("td");
      const btnView = document.createElement("button");
      btnView.className = "btn secondary small";
      btnView.type = "button";
      const isExpanded = activeResultInlineId === row.studentId;
      btnView.textContent = isExpanded ? "Hide preview" : "Preview";
      btnView.addEventListener("click", () => {
        if (!row.studentId) return;
        activeResultInlineId = activeResultInlineId === row.studentId ? "" : row.studentId;
        renderResultsView();
      });
      tdAction.appendChild(btnView);
      tr.appendChild(tdAction);

      elResultTableBody.appendChild(tr);

      if (row.studentId && row.studentId === activeResultInlineId) {
        const inlineRow = document.createElement("tr");
        inlineRow.className = "result-inline-row";
        const inlineCell = document.createElement("td");
        inlineCell.colSpan = 6;

        const panel = document.createElement("div");
        panel.className = "result-inline-panel";

        const header = document.createElement("div");
        header.className = "result-inline-header";
        const title = document.createElement("strong");
        title.textContent = "Result Preview";
        header.appendChild(title);

        const actions = document.createElement("div");
        actions.className = "form-actions";
        const btnAddResult = document.createElement("button");
        btnAddResult.className = "btn secondary small";
        btnAddResult.type = "button";
        btnAddResult.textContent = "Add Result";
        btnAddResult.addEventListener("click", () => openAddResultModal(row.studentId));
        actions.appendChild(btnAddResult);
        const btnSlip = document.createElement("button");
        btnSlip.className = "btn secondary small";
        btnSlip.type = "button";
        btnSlip.textContent = "Open Slip";
        btnSlip.addEventListener("click", () => openSlip(row.studentId));
        actions.appendChild(btnSlip);
        header.appendChild(actions);
        panel.appendChild(header);

        const studentResults = resultState.results.filter(
          (result) => String(result.studentId ?? "").trim() === row.studentId
        );
        const cgpaMap = computeStudentCgpa(studentResults, resultState.courses);
        const cgpa = cgpaMap.get(row.studentId) ?? null;
        const latestSemester = studentResults.reduce((max, result) => {
          const sem = toNumber(result.semester);
          if (sem === null) return max;
          return Math.max(max, sem);
        }, 0);

        const summary = document.createElement("div");
        summary.className = "result-inline-summary";
        summary.innerHTML = `
          <span>Records: ${numberFmt.format(studentResults.length)}</span>
          <span>Latest Semester: ${latestSemester || "-"}</span>
          <span>CGPA: ${cgpa === null ? "--" : cgpa.toFixed(2)}</span>
        `;
        panel.appendChild(summary);

        if (!studentResults.length) {
          const empty = document.createElement("div");
          empty.className = "muted";
          empty.textContent = "No results recorded for this student yet.";
          panel.appendChild(empty);
        } else {
          const courseMap = buildCourseMap();
          const termGroups = new Map();
          for (const result of studentResults) {
            const session = String(result.session ?? "").trim();
            const semester = result.semester ?? "";
            const key = `${session}||${semester}`;
            if (!termGroups.has(key)) {
              termGroups.set(key, { session, semester, rows: [] });
            }
            termGroups.get(key).rows.push(result);
          }

          const termList = [...termGroups.values()].sort((a, b) => {
            if (a.session === b.session) {
              return String(a.semester ?? "").localeCompare(String(b.semester ?? ""));
            }
            return String(a.session ?? "").localeCompare(String(b.session ?? ""));
          });

          for (const term of termList) {
            const section = document.createElement("div");
            section.className = "result-inline-section";
            const heading = document.createElement("h4");
            heading.textContent = `Session ${formatYearMonth(term.session)} - Semester ${term.semester ?? "-"}`;
            section.appendChild(heading);

            const table = document.createElement("table");
            table.className = "result-inline-table";
            const colgroup = document.createElement("colgroup");
            const colWidths = ["14%", "40%", "10%", "10%", "16%", "10%"];
            for (const width of colWidths) {
              const col = document.createElement("col");
              col.style.width = width;
              colgroup.appendChild(col);
            }
            table.appendChild(colgroup);
            const thead = document.createElement("thead");
            thead.innerHTML = `
              <tr>
                <th>Course</th>
                <th>Title</th>
                <th>Mark</th>
                <th>Letter</th>
                <th>Status</th>
                <th>Action</th>
              </tr>
            `;
            table.appendChild(thead);
            const tbody = document.createElement("tbody");

            const sorted = [...term.rows].sort((a, b) => {
              const aCode = normalizeCourseCode(a.courseCode ?? "");
              const bCode = normalizeCourseCode(b.courseCode ?? "");
              return aCode.localeCompare(bCode);
            });

            for (const result of sorted) {
              const trResult = document.createElement("tr");
              const courseCode = normalizeCourseCode(result.courseCode ?? "");
              const course = courseMap.get(courseCode);
              const tdCourse = document.createElement("td");
              tdCourse.textContent = courseCode || "-";
              trResult.appendChild(tdCourse);
              const tdTitle = document.createElement("td");
              tdTitle.textContent = String(course?.title ?? "").trim() || "-";
              trResult.appendChild(tdTitle);
              const tdMark = document.createElement("td");
              tdMark.textContent = result.mark ?? "-";
              if (isFailedResult(result)) tdMark.classList.add("fail");
              trResult.appendChild(tdMark);
              const tdLetter = document.createElement("td");
              tdLetter.textContent = result.letter ?? "-";
              if (isFailedResult(result)) tdLetter.classList.add("fail");
              trResult.appendChild(tdLetter);
              const tdStatus = document.createElement("td");
              if (isMissingMark(result)) {
                tdStatus.textContent = "No mark";
                tdStatus.classList.add("remark");
              } else if (isFailedResult(result)) {
                tdStatus.textContent = "Failed";
                tdStatus.classList.add("remark");
              } else if (isPassedResult(result)) {
                tdStatus.textContent = "Passed";
              } else {
                tdStatus.textContent = "-";
              }
              trResult.appendChild(tdStatus);

              const tdAction = document.createElement("td");
              tdAction.className = "table-actions";
              const btnEdit = document.createElement("button");
              btnEdit.className = "btn secondary small";
              btnEdit.type = "button";
              btnEdit.textContent = "Edit";
              btnEdit.addEventListener("click", () => openMarkEdit(result));
              tdAction.appendChild(btnEdit);
              const btnDelete = document.createElement("button");
              btnDelete.className = "btn secondary small";
              btnDelete.type = "button";
              btnDelete.textContent = "Delete";
              btnDelete.addEventListener("click", async () => {
                const courseCode = normalizeCourseCode(result.courseCode ?? "");
                const session = formatYearMonth(result.session);
                const semester = result.semester ?? "-";
                const label = `${courseCode} (${session} - Sem ${semester})`;
                const ok = window.confirm(`Delete result ${label}?`);
                if (!ok) return;
                try {
                  await deleteRecord(STORES.results, result.resultId);
                  activeResultInlineId = row.studentId;
                  await refreshResults();
                  await refreshStats();
                } catch (e) {
                  alert(`Failed to delete result: ${e.message ?? e}`);
                }
              });
              tdAction.appendChild(btnDelete);
              trResult.appendChild(tdAction);

              tbody.appendChild(trResult);
            }

            table.appendChild(tbody);
            section.appendChild(table);
            panel.appendChild(section);
          }
        }

        inlineCell.appendChild(panel);
        inlineRow.appendChild(inlineCell);
        elResultTableBody.appendChild(inlineRow);
      }
    }
    return;
  }

  if (mode === "courses") {
    elResultPager.style.display = "flex";
    for (const row of rows) {
      const tr = document.createElement("tr");

      const tdCourse = document.createElement("td");
      tdCourse.textContent = row.courseLabel;
      tr.appendChild(tdCourse);

      const tdCount = document.createElement("td");
      tdCount.textContent = numberFmt.format(row.records);
      tr.appendChild(tdCount);

      const tdPass = document.createElement("td");
      tdPass.textContent = row.passRate === null ? "--" : percentFmt.format(row.passRate);
      tr.appendChild(tdPass);

      const tdAvg = document.createElement("td");
      tdAvg.textContent = row.avgMark === null ? "--" : row.avgMark.toFixed(1);
      tr.appendChild(tdAvg);

      const tdAction = document.createElement("td");
      const btnView = document.createElement("button");
      btnView.className = "btn secondary small";
      btnView.type = "button";
      btnView.textContent = "View results";
      btnView.addEventListener("click", () => {
        setActiveTab("results");
        renderCourseResults(row.courseCode);
      });
      tdAction.appendChild(btnView);
      tr.appendChild(tdAction);

      elResultTableBody.appendChild(tr);
    }
    return;
  }

  elResultPager.style.display = "none";
  for (const row of rows) {
    const tr = document.createElement("tr");

    const tdStudentId = document.createElement("td");
    tdStudentId.textContent = row.studentId ?? "-";
    tr.appendChild(tdStudentId);

      const tdStudentName = document.createElement("td");
      tdStudentName.textContent = row.studentName || "-";
      if (row.studentId && redFlagSet.has(String(row.studentId ?? "").trim())) {
        appendRedFlagTag(tdStudentName, row.studentId);
      }
      tr.appendChild(tdStudentName);

    const tdSession = document.createElement("td");
    tdSession.textContent = formatYearMonth(row.session);
    tr.appendChild(tdSession);

    const tdSemester = document.createElement("td");
    tdSemester.textContent = row.semester ?? "-";
    tr.appendChild(tdSemester);

    const tdMark = document.createElement("td");
    tdMark.textContent = row.mark ?? "-";
    if (row.failed) tdMark.classList.add("fail");
    tr.appendChild(tdMark);

    const tdLetter = document.createElement("td");
    tdLetter.textContent = row.letter || "-";
    if (row.failed) tdLetter.classList.add("fail");
    tr.appendChild(tdLetter);

    const tdPoint = document.createElement("td");
    tdPoint.textContent = row.point ?? "-";
    tr.appendChild(tdPoint);

    const tdStatus = document.createElement("td");
    if (row.failed) {
      tdStatus.textContent = "Failed";
      tdStatus.classList.add("remark");
    } else {
      tdStatus.textContent = row.passedCourse === true ? "Passed" : "-";
    }
    tr.appendChild(tdStatus);

    if (mode === "detail") {
      const tdAction = document.createElement("td");
      tdAction.className = "table-actions";
      const btnEdit = document.createElement("button");
      btnEdit.className = "btn secondary small";
      btnEdit.type = "button";
      btnEdit.textContent = "Edit";
      btnEdit.addEventListener("click", () => openMarkEdit(row));
      tdAction.appendChild(btnEdit);
      tr.appendChild(tdAction);
    }

    elResultTableBody.appendChild(tr);
  }
}

function buildSlip(studentId) {
  const studentMap = buildStudentMap();
  const courseMap = buildCourseMap();
  const student = studentMap.get(studentId);

  if (!student) {
    elSlipStudentMeta.textContent = "Select a student to generate a slip.";
    elSlipCGPA.textContent = "--";
    elSlipTerms.textContent = "";
    return;
  }

  const studentResults = resultState.results.filter(
    (row) => String(row.studentId ?? "").trim() === studentId
  );

  if (!studentResults.length) {
    elSlipStudentMeta.textContent = "No results found for this student.";
    elSlipCGPA.textContent = "--";
    elSlipTerms.textContent = "";
    return;
  }

  elSlipStudentMeta.textContent = "";
  const table = document.createElement("table");
  const tbody = document.createElement("tbody");

  const makeCell = (label, value) => {
    const td = document.createElement("td");
    const labelSpan = document.createElement("span");
    labelSpan.textContent = label;
    td.appendChild(labelSpan);
    td.appendChild(document.createElement("br"));
    td.appendChild(document.createTextNode(String(value)));
    return td;
  };

  const row1 = document.createElement("tr");
  row1.appendChild(makeCell("Student Name", student.name ?? "-"));
  row1.appendChild(makeCell("Student ID", student.studentId ?? studentId));
  tbody.appendChild(row1);

  const row2 = document.createElement("tr");
  row2.appendChild(makeCell("IC", student.ic ?? "-"));
  row2.appendChild(makeCell("Intake", formatYearMonth(student.intake)));
  tbody.appendChild(row2);

  table.appendChild(tbody);
  elSlipStudentMeta.appendChild(table);

  const termGroups = new Map();
  for (const row of studentResults) {
    const session = String(row.session ?? "").trim();
    const semester = row.semester ?? "";
    const key = `${session}||${semester}`;
    if (!termGroups.has(key)) {
      termGroups.set(key, { session, semester, rows: [] });
    }
    termGroups.get(key).rows.push(row);
  }

  const termList = [...termGroups.values()].sort((a, b) => {
    if (a.session === b.session) {
      return String(a.semester ?? "").localeCompare(String(b.semester ?? ""));
    }
    return a.session.localeCompare(b.session);
  });

  let totalCredits = 0;
  let totalGradePoints = 0;

  elSlipTerms.textContent = "";

  const failedBuckets = [];

  for (const term of termList) {
    const section = document.createElement("div");
    section.className = "slip-section";

    const heading = document.createElement("h4");
    heading.textContent = `Session ${formatYearMonth(term.session)} - Semester ${term.semester ?? "-"}`;
    section.appendChild(heading);

    const table = document.createElement("table");
    table.className = "slip-table";
    const colgroup = document.createElement("colgroup");
    const colWidths = ["14%", "32%", "8%", "8%", "8%", "8%", "22%"];
    for (const width of colWidths) {
      const col = document.createElement("col");
      col.style.width = width;
      colgroup.appendChild(col);
    }
    table.appendChild(colgroup);
    const thead = document.createElement("thead");
    thead.innerHTML = `
      <tr>
        <th>Course</th>
        <th>Title</th>
        <th>Credits</th>
        <th>Mark</th>
        <th>Letter</th>
        <th>Point</th>
        <th>Grade Pts</th>
      </tr>
    `;
    table.appendChild(thead);
    const tbody = document.createElement("tbody");

    let termCredits = 0;
    let termCreditsEarned = 0;
    let termGradePoints = 0;

    for (const row of term.rows) {
      const tr = document.createElement("tr");
      const courseCode = normalizeCourseCode(row.courseCode ?? "");
      const course = courseMap.get(courseCode);
      const credits = toNumber(course?.credits);
      const gradePoints = getGradePoints(row, credits);
      const failed = isFailedResult(row);

      const tdCourse = document.createElement("td");
      tdCourse.textContent = courseCode || "-";
      tr.appendChild(tdCourse);

      const tdTitle = document.createElement("td");
      tdTitle.textContent = course?.title || "-";
      tr.appendChild(tdTitle);

      const tdCredits = document.createElement("td");
      tdCredits.textContent = credits === null ? "-" : numberFmt.format(credits);
      tr.appendChild(tdCredits);

      const tdMark = document.createElement("td");
      tdMark.textContent = row.mark ?? "-";
      if (failed) tdMark.classList.add("fail");
      tr.appendChild(tdMark);

      const tdLetter = document.createElement("td");
      tdLetter.textContent = row.letter ?? "-";
      if (failed) tdLetter.classList.add("fail");
      tr.appendChild(tdLetter);

      const tdPoint = document.createElement("td");
      tdPoint.textContent = row.point ?? "-";
      tr.appendChild(tdPoint);

      const tdGradePoints = document.createElement("td");
      tdGradePoints.textContent = gradePoints === null ? "-" : gradePoints.toFixed(2);
      tr.appendChild(tdGradePoints);

      if (failed) {
        failedBuckets.push({
          session: term.session,
          semester: term.semester,
          courseCode,
          title: course?.title || "-",
          mark: row.mark ?? "-",
          letter: row.letter ?? "-",
          point: row.point ?? "-",
        });
      }

      if (credits !== null) {
        termCredits += credits;
        const earned = toNumber(row.creditsEarned);
        if (earned !== null) {
          termCreditsEarned += earned;
        } else if (!failed) {
          termCreditsEarned += credits;
        }
      }
      if (credits !== null && gradePoints !== null) {
        termGradePoints += gradePoints;
      }

      tbody.appendChild(tr);
    }

    table.appendChild(tbody);

    const tfoot = document.createElement("tfoot");
    const cumulativeGradePoints = totalGradePoints + termGradePoints;
    const cumulativeCredits = totalCredits + termCredits;
    const termGpa = termCredits ? termGradePoints / termCredits : null;
    const cumulativeCgpa = cumulativeCredits ? cumulativeGradePoints / cumulativeCredits : null;

    const tr = document.createElement("tr");
    const summaryCells = [
      ["Credits Attempted", termCredits ? numberFmt.format(termCredits) : "-"],
      ["Credits Earned", termCreditsEarned ? numberFmt.format(termCreditsEarned) : "-"],
      ["Grade Pts (Sem)", termGradePoints ? termGradePoints.toFixed(2) : "-"],
      ["Grade Pts (Cum)", cumulativeGradePoints ? cumulativeGradePoints.toFixed(2) : "-"],
      ["GPA", termGpa === null ? "--" : termGpa.toFixed(2)],
      ["CGPA", cumulativeCgpa === null ? "--" : cumulativeCgpa.toFixed(2)],
    ];

    summaryCells.forEach(([label, value], index) => {
      const td = document.createElement("td");
      const strong = document.createElement("strong");
      strong.textContent = `${label}: `;
      td.appendChild(strong);
      td.appendChild(document.createTextNode(String(value)));
      if (index === summaryCells.length - 1) td.colSpan = 2;
      tr.appendChild(td);
    });

    tfoot.appendChild(tr);

    table.appendChild(tfoot);
    section.appendChild(table);

    const termSummary = document.createElement("div");
    termSummary.className = "slip-meta";
    const creditsLine = document.createElement("div");
    const creditsLabel = document.createElement("strong");
    creditsLabel.textContent = "Credits: ";
    creditsLine.appendChild(creditsLabel);
    creditsLine.appendChild(
      document.createTextNode(termCredits ? numberFmt.format(termCredits) : "-")
    );
    termSummary.appendChild(creditsLine);

    const gpaLine = document.createElement("div");
    const gpaLabel = document.createElement("strong");
    gpaLabel.textContent = "GPA: ";
    gpaLine.appendChild(gpaLabel);
    gpaLine.appendChild(
      document.createTextNode(termGpa === null ? "--" : termGpa.toFixed(2))
    );
    termSummary.appendChild(gpaLine);
    section.appendChild(termSummary);

    totalCredits += termCredits;
    totalGradePoints += termGradePoints;

    elSlipTerms.appendChild(section);
  }

  const cgpa = totalCredits ? totalGradePoints / totalCredits : null;
  elSlipCGPA.textContent = cgpa === null ? "--" : cgpa.toFixed(2);

  if (failedBuckets.length) {
    const failedSection = document.createElement("div");
    failedSection.className = "slip-section";
    const heading = document.createElement("h4");
    heading.textContent = "Failed Courses";
    failedSection.appendChild(heading);

    const table = document.createElement("table");
    table.className = "slip-table";
    const colgroup = document.createElement("colgroup");
    const colWidths = ["16%", "10%", "14%", "30%", "10%", "10%"];
    for (const width of colWidths) {
      const col = document.createElement("col");
      col.style.width = width;
      colgroup.appendChild(col);
    }
    table.appendChild(colgroup);

    const thead = document.createElement("thead");
    thead.innerHTML = `
      <tr>
        <th>Session</th>
        <th>Sem</th>
        <th>Course</th>
        <th>Title</th>
        <th>Mark</th>
        <th>Letter</th>
      </tr>
    `;
    table.appendChild(thead);
    const tbody = document.createElement("tbody");

    for (const row of failedBuckets) {
      const tr = document.createElement("tr");
        const cells = [
          formatYearMonth(row.session),
          row.semester ?? "-",
          row.courseCode || "-",
          row.title || "-",
          row.mark ?? "-",
          row.letter ?? "-",
        ];
      for (const value of cells) {
        const td = document.createElement("td");
        td.textContent = value;
        td.classList.add("fail");
        tr.appendChild(td);
      }
      tbody.appendChild(tr);
    }

    table.appendChild(tbody);
    failedSection.appendChild(table);
    elSlipTerms.appendChild(failedSection);
  }
}

function openSlip(studentId) {
  buildSlip(studentId);
  if (!elResultSlipModal) return;
  elResultSlipModal.classList.add("active");
  elResultSlipModal.setAttribute("aria-hidden", "false");
}

function closeSlip() {
  if (!elResultSlipModal) return;
  elResultSlipModal.classList.remove("active");
  elResultSlipModal.setAttribute("aria-hidden", "true");
}

function populateResultFilters() {
  if (!elResultSessionSelect) return;
  const studentMap = buildStudentMap();

  const currentSession = formatYearMonth(elResultSessionSelect.value, "");

  const sessions = [...new Set(
    [...studentMap.values()]
      .map((student) => formatYearMonth(student.intake, ""))
      .filter(Boolean)
  )].sort();
  elResultSessionSelect.textContent = "";
  const sessionDefault = document.createElement("option");
  sessionDefault.value = "";
  sessionDefault.textContent = "All intakes";
  elResultSessionSelect.appendChild(sessionDefault);
  for (const session of sessions) {
    const option = document.createElement("option");
    option.value = session;
    option.textContent = session;
    elResultSessionSelect.appendChild(option);
  }

  if (currentSession) elResultSessionSelect.value = currentSession;
}

function renderResultsView() {
  const studentMap = buildStudentMap();
  const courseMap = buildCourseMap();
  const resultNameMap = new Map();
  for (const row of resultState.results) {
    const studentId = String(row.studentId ?? "").trim();
    const name = String(row.studentName ?? "").trim();
    if (studentId && name && !resultNameMap.has(studentId)) {
      resultNameMap.set(studentId, name);
    }
  }

  const search = elResultSearch?.value?.trim().toLowerCase() ?? "";
  const intakeFilter = formatYearMonth(elResultSessionSelect?.value ?? "", "");
  const viewMode = elResultViewMode?.value ?? "students";

  const filtered = resultState.results.filter((row) => {
    const studentId = String(row.studentId ?? "").trim();
    if (intakeFilter) {
      const student = studentMap.get(studentId);
      const intake = formatYearMonth(student?.intake ?? "", "");
      if (intake !== intakeFilter) return false;
    }
    if (activeCourseDetailCode) {
      const courseCode = normalizeCourseCode(row.courseCode ?? "");
      if (courseCode !== activeCourseDetailCode) return false;
    }
    if (search) {
      const student = studentMap.get(studentId);
      const courseCode = normalizeCourseCode(row.courseCode ?? "");
      const course = courseMap.get(courseCode);
      const label = `${student?.studentId ?? ""} ${student?.name ?? ""} ${courseCode} ${course?.title ?? ""}`
        .toLowerCase();
      if (!label.includes(search)) return false;
    }
    return true;
  });

  const studentIds = new Set(filtered.map((row) => String(row.studentId ?? "").trim()).filter(Boolean));

  renderResultStats(filtered, studentIds);
  if (activeCourseDetailCode) {
    const stats = filtered.reduce(
      (acc, row) => {
        acc.records += 1;
        if (row.passedCourse === true) acc.passed += 1;
        const mark = toNumber(row.mark);
        if (mark !== null) acc.marks.push(mark);
        if (row.studentId) acc.students.add(String(row.studentId).trim());
        return acc;
      },
      { records: 0, passed: 0, marks: [], students: new Set() }
    );
    const avg = stats.marks.length ? stats.marks.reduce((a, b) => a + b, 0) / stats.marks.length : null;
    const passRate = stats.records ? stats.passed / stats.records : null;
    if (elCourseDetailStats) elCourseDetailStats.style.display = "grid";
    if (elCourseDetailRecords) setText(elCourseDetailRecords, numberFmt.format(stats.records));
    if (elCourseDetailPassRate) setText(elCourseDetailPassRate, passRate === null ? "--" : percentFmt.format(passRate));
    if (elCourseDetailAvgMark) setText(elCourseDetailAvgMark, avg === null ? "--" : avg.toFixed(1));
    if (elCourseDetailStudents) setText(elCourseDetailStudents, numberFmt.format(stats.students.size));
    const courseMap = buildCourseMap();
    const course = courseMap.get(activeCourseDetailCode);
    const courseLabel = course
      ? `${activeCourseDetailCode} - ${course.title ?? ""}`.trim()
      : activeCourseDetailCode;
    if (elCourseDetailHeader) elCourseDetailHeader.style.display = "flex";
    if (elCourseDetailTitle) elCourseDetailTitle.textContent = courseLabel;
    const rows = filtered.map((row) => {
      const studentId = String(row.studentId ?? "").trim();
      const student = studentMap.get(studentId);
      return {
        resultId: row.resultId,
        courseCode: activeCourseDetailCode,
        courseLabel,
        studentId,
        studentName: student?.name ?? "",
        session: row.session ?? "",
        semester: row.semester ?? "",
        mark: row.mark ?? "",
        letter: row.letter ?? "",
        point: row.point ?? "",
        passedCourse: row.passedCourse,
        failed: isFailedResult(row),
        creditsEarned: row.creditsEarned ?? null,
      };
    });
    const sortedRows = sortResultRows(rows);
    const totalPages = Math.max(1, Math.ceil(rows.length / resultPaging.pageSize));
    if (resultPaging.page > totalPages) resultPaging.page = totalPages;
    const startIndex = (resultPaging.page - 1) * resultPaging.pageSize;
    const pageRows = sortedRows.slice(startIndex, startIndex + resultPaging.pageSize);
    if (elResultPageMeta) elResultPageMeta.textContent = `Page ${resultPaging.page} of ${totalPages}`;
    if (btnPrevPage) btnPrevPage.disabled = resultPaging.page <= 1;
    if (btnNextPage) btnNextPage.disabled = resultPaging.page >= totalPages;
    renderResultTable(pageRows, "detail");
    if (elResultPager) elResultPager.style.display = "flex";
    return;
  } else if (elCourseDetailStats) {
    elCourseDetailStats.style.display = "none";
    if (elCourseDetailHeader) elCourseDetailHeader.style.display = "none";
  }
  if (viewMode === "courses" && !activeCourseDetailCode) {
    const courseCounts = new Map();
    for (const row of filtered) {
      const courseCode = normalizeCourseCode(row.courseCode ?? "");
      if (!courseCode) continue;
      const entry = courseCounts.get(courseCode) ?? {
        records: 0,
        passed: 0,
        marks: [],
      };
      entry.records += 1;
      if (row.passedCourse === true) entry.passed += 1;
      const mark = toNumber(row.mark);
      if (mark !== null) entry.marks.push(mark);
      courseCounts.set(courseCode, entry);
    }

    const listRows = [...courseCounts.entries()]
      .map(([courseCode, entry]) => {
        const course = courseMap.get(courseCode);
        const courseLabel = course ? `${courseCode} - ${course.title ?? ""}`.trim() : courseCode;
        const avgMark = entry.marks.length
          ? entry.marks.reduce((a, b) => a + b, 0) / entry.marks.length
          : null;
        const passRate = entry.records ? entry.passed / entry.records : null;
        return {
          courseCode,
          courseLabel,
          records: entry.records,
          avgMark,
          passRate,
        };
      })
      .sort((a, b) => a.courseLabel.localeCompare(b.courseLabel));

    const totalPages = Math.max(1, Math.ceil(listRows.length / resultPaging.pageSize));
    if (resultPaging.page > totalPages) resultPaging.page = totalPages;
    const startIndex = (resultPaging.page - 1) * resultPaging.pageSize;
    const tableRows = sortResultRows(listRows).slice(startIndex, startIndex + resultPaging.pageSize);

    if (elResultPageMeta) elResultPageMeta.textContent = `Page ${resultPaging.page} of ${totalPages}`;
    if (btnPrevPage) btnPrevPage.disabled = resultPaging.page <= 1;
    if (btnNextPage) btnNextPage.disabled = resultPaging.page >= totalPages;

    renderResultTable(tableRows, "courses");
  } else {
    const studentCounts = new Map();
    const studentSemesterMax = new Map();
    for (const row of filtered) {
      const studentId = String(row.studentId ?? "").trim();
      if (!studentId) continue;
      studentCounts.set(studentId, (studentCounts.get(studentId) ?? 0) + 1);
      const sem = toNumber(row.semester);
      if (sem !== null) {
        const current = studentSemesterMax.get(studentId) ?? 0;
        if (sem > current) studentSemesterMax.set(studentId, sem);
      }
    }

    const cgpaByStudent = computeStudentCgpa(filtered, resultState.courses);
    const resultStudentIds = new Set(
      filtered.map((row) => String(row.studentId ?? "").trim()).filter(Boolean)
    );

    const listRows = [...studentMap.values()]
      .filter((student) => !student?.isDeleted)
      .map((student) => {
        const studentId = String(student?.studentId ?? "").trim();
        const name =
          String(student?.name ?? "").trim() || resultNameMap.get(studentId) || "";
        const intake = formatYearMonth(student?.intake ?? "", "");
        const label = `${studentId} ${name}`.toLowerCase();
        const matchesSearch = !search || label.includes(search) || resultStudentIds.has(studentId);
        const matchesIntake = !intakeFilter || intake === intakeFilter;
        if (!studentId || !matchesSearch || !matchesIntake) return null;
        const records = studentCounts.get(studentId) ?? 0;
        const cgpa = cgpaByStudent.get(studentId) ?? null;
        return {
          studentId,
          studentName: name,
          records,
          currentSemester: studentSemesterMax.get(studentId) ?? "-",
          cgpa: cgpa === null ? "--" : cgpa.toFixed(2),
        };
      })
      .filter(Boolean)
      .sort((a, b) => {
        const aName = String(a.studentName ?? "").toLowerCase();
        const bName = String(b.studentName ?? "").toLowerCase();
        if (aName && bName) return aName.localeCompare(bName);
        return String(a.studentId ?? "").localeCompare(String(b.studentId ?? ""));
      });

    const totalPages = Math.max(1, Math.ceil(listRows.length / resultPaging.pageSize));
    if (resultPaging.page > totalPages) resultPaging.page = totalPages;
    const startIndex = (resultPaging.page - 1) * resultPaging.pageSize;
    const tableRows = sortResultRows(listRows).slice(startIndex, startIndex + resultPaging.pageSize);

    if (elResultPageMeta) elResultPageMeta.textContent = `Page ${resultPaging.page} of ${totalPages}`;
    if (btnPrevPage) btnPrevPage.disabled = resultPaging.page <= 1;
    if (btnNextPage) btnNextPage.disabled = resultPaging.page >= totalPages;

    renderResultTable(tableRows, "students");
  }
}

function sortResultRows(rows) {
  const sortKey = resultSortState.key;
  if (!sortKey) return rows;
  const getSortValue = (row, key) => {
    if (key === "status") {
      if (row.failed) return "failed";
      if (row.passedCourse === true) return "passed";
      return "unknown";
    }
    return row[key];
  };
  return rows.slice().sort((a, b) => {
    const av = getSortValue(a, sortKey);
    const bv = getSortValue(b, sortKey);
    if (av === null || av === undefined) return 1;
    if (bv === null || bv === undefined) return -1;
    if (typeof av === "number" && typeof bv === "number") {
      return resultSortState.dir === "asc" ? av - bv : bv - av;
    }
    const as = String(av).toLowerCase();
    const bs = String(bv).toLowerCase();
    return resultSortState.dir === "asc" ? as.localeCompare(bs) : bs.localeCompare(as);
  });
}

function openStudentEdit(studentId) {
  const student = resultState.students.find(
    (s) => String(s.studentId ?? "").trim() === studentId
  );
  if (!student || !elStudentEditModal) return;
  activeStudentEditId = studentId;
  elEditStudentId.value = studentId;
  elEditStudentName.value = student.name ?? "";
  elEditStudentIc.value = student.ic ?? "";
  elEditStudentIntake.value = formatYearMonth(student.intake, "");
  elStudentEditModal.classList.add("active");
  elStudentEditModal.setAttribute("aria-hidden", "false");
}

function closeStudentEdit() {
  activeStudentEditId = "";
  if (!elStudentEditModal) return;
  elStudentEditModal.classList.remove("active");
  elStudentEditModal.setAttribute("aria-hidden", "true");
}

function openMarkEdit(row) {
  if (!row?.resultId || !elMarkEditModal) return;
  activeMarkEditId = row.resultId;
  activeMarkEditRow = row;
  activeMarkEditValue = row.mark ?? "";
  if (elMarkEditVerify) elMarkEditVerify.checked = false;
  updateMarkEditPreview();
  elMarkEditModal.classList.add("active");
  elMarkEditModal.setAttribute("aria-hidden", "false");
  queueMicrotask(() => elMarkEditInput?.focus());
}

function closeMarkEdit() {
  activeMarkEditId = "";
  activeMarkEditValue = "";
  activeMarkEditRow = null;
  if (!elMarkEditModal) return;
  elMarkEditModal.classList.remove("active");
  elMarkEditModal.setAttribute("aria-hidden", "true");
}

function updateMarkEditPreview() {
  if (!activeMarkEditRow) return;
  const row = activeMarkEditRow;
  const studentMap = buildStudentMap();
  const courseMap = buildCourseMap();
  const student = studentMap.get(String(row.studentId ?? "").trim());
  const course = courseMap.get(normalizeCourseCode(row.courseCode ?? ""));
  const studentLabel = student ? `${student.studentId ?? ""} - ${student.name ?? ""}`.trim() : row.studentId ?? "-";
  const courseLabel = course
    ? `${normalizeCourseCode(row.courseCode ?? "")} - ${course.title ?? ""}`.trim()
    : normalizeCourseCode(row.courseCode ?? "");

  if (elMarkEditStudent) elMarkEditStudent.textContent = studentLabel || "-";
  if (elMarkEditCourse) elMarkEditCourse.textContent = courseLabel || "-";
  if (elMarkEditSession) elMarkEditSession.textContent = formatYearMonth(row.session);
  if (elMarkEditSemester) elMarkEditSemester.textContent = row.semester ?? "-";
  if (elMarkEditCurrent) elMarkEditCurrent.textContent = row.mark ?? "-";
  if (elMarkEditInput) elMarkEditInput.value = activeMarkEditValue ?? "";

  const raw = String(activeMarkEditValue ?? "").trim();
  let mark = null;
  let letter = "";
  let point = null;
  let passedCourse = null;
  let warning = "";
  let isValid = true;

  if (raw) {
    const n = Number(raw);
    if (!Number.isFinite(n)) {
      isValid = false;
      warning = "Mark must be a number.";
    } else if (n < 0 || n > 100) {
      isValid = false;
      warning = "Mark must be between 0 and 100.";
    } else {
      mark = n;
      const grade = getGradeForMark(mark, course?.isMPUCourse === true);
      letter = grade.letter;
      point = grade.point;
      passedCourse = grade.passed;
    }
  } else {
    warning = "Blank mark will clear letter/point/status.";
  }

  if (elMarkEditLetter) elMarkEditLetter.textContent = letter || "-";
  if (elMarkEditPoint) elMarkEditPoint.textContent = point === null ? "-" : String(point);
  if (elMarkEditStatus) {
    elMarkEditStatus.textContent = passedCourse === null ? "-" : (passedCourse ? "Passed" : "Failed");
    elMarkEditStatus.classList.toggle("remark", passedCourse === false);
  }
  if (elMarkEditWarning) {
    elMarkEditWarning.textContent = warning;
    elMarkEditWarning.style.display = warning ? "block" : "none";
  }
  if (btnConfirmMarkEdit) {
    const verified = elMarkEditVerify?.checked;
    btnConfirmMarkEdit.disabled = !isValid || !verified;
  }
}

async function confirmMarkEdit() {
  if (!activeMarkEditRow) return;
  const current = resultState.results.find((r) => r.resultId === activeMarkEditRow.resultId);
  if (!current) return;

  const raw = String(activeMarkEditValue ?? "").trim();
  let mark = null;
  let letter = "";
  let point = null;
  let passedCourse = null;

  if (raw) {
    const n = Number(raw);
    if (!Number.isFinite(n) || n < 0 || n > 100) {
      alert("Mark must be a number between 0 and 100.");
      return;
    }
    mark = n;
    const courseCode = normalizeCourseCode(current.courseCode ?? "");
    const course = buildCourseMap().get(courseCode);
    const grade = getGradeForMark(mark, course?.isMPUCourse === true);
    letter = grade.letter;
    point = grade.point;
    passedCourse = grade.passed;
  }

  const updated = {
    ...current,
    mark,
    letter,
    point,
    gradePoints: null,
    passedCourse,
  };

  try {
    await upsertMany(STORES.results, [updated]);
    const index = resultState.results.findIndex((r) => r.resultId === current.resultId);
    if (index >= 0) resultState.results[index] = updated;
    closeMarkEdit();
    await refreshStats();
    renderResultsView();
  } catch (e) {
    console.error(e);
    alert(`Failed to save mark: ${e.message ?? e}`);
  }
}

function populateAddResultCourseOptions(selectedCode = "") {
  if (!elAddResultCourse) return;
  const courses = [...buildCourseMap().values()].sort((a, b) =>
    normalizeCourseCode(a.courseCode ?? "").localeCompare(normalizeCourseCode(b.courseCode ?? ""))
  );
  elAddResultCourse.textContent = "";
  const defaultOption = document.createElement("option");
  defaultOption.value = "";
  defaultOption.textContent = "Select course";
  elAddResultCourse.appendChild(defaultOption);
  for (const course of courses) {
    const code = normalizeCourseCode(course.courseCode ?? "");
    if (!code) continue;
    const option = document.createElement("option");
    option.value = code;
    option.textContent = course.title ? `${code} - ${course.title}` : code;
    if (code === selectedCode) option.selected = true;
    elAddResultCourse.appendChild(option);
  }
}

function getLatestSessionForStudent(studentId) {
  const id = String(studentId ?? "").trim();
  if (!id) return "";
  let best = "";
  let bestValue = -1;
  for (const row of resultState.results) {
    if (String(row.studentId ?? "").trim() !== id) continue;
    const parsed = parseYearMonth(row.session);
    if (!parsed) continue;
    const value = Number(`${parsed.year}${parsed.month}`);
    if (value > bestValue) {
      bestValue = value;
      best = parsed.intake;
    }
  }
  return best;
}

function getNextSemesterForStudent(studentId) {
  const id = String(studentId ?? "").trim();
  if (!id) return 1;
  let maxSem = 0;
  for (const row of resultState.results) {
    if (String(row.studentId ?? "").trim() !== id) continue;
    const sem = toNumber(row.semester);
    if (sem !== null && sem > maxSem) maxSem = sem;
  }
  return maxSem + 1;
}

function updateAddResultComputed() {
  if (!elAddResultCourse || !elAddResultMark) return;
  const courseCode = normalizeCourseCode(elAddResultCourse.value ?? "");
  const course = buildCourseMap().get(courseCode);
  const raw = String(elAddResultMark.value ?? "").trim();
  let warning = "";
  let letter = "-";
  let point = "-";
  let status = "--";

  if (raw) {
    const n = Number(raw);
    if (!Number.isFinite(n) || n < 0 || n > 100) {
      warning = "Mark must be between 0 and 100.";
    } else {
      const grade = getGradeForMark(n, course?.isMPUCourse === true);
      letter = grade.letter;
      point = grade.point === null ? "-" : String(grade.point);
      status = grade.passed ? "Passed" : "Failed";
    }
  }

  if (elAddResultLetter) elAddResultLetter.textContent = letter;
  if (elAddResultPoint) elAddResultPoint.textContent = point;
  if (elAddResultStatus) {
    elAddResultStatus.textContent = status;
    elAddResultStatus.classList.toggle("remark", status === "Failed");
  }
  if (elAddResultWarning) {
    elAddResultWarning.textContent = warning;
    elAddResultWarning.style.display = warning ? "block" : "none";
  }

  validateAddResultInputs();
}

function validateAddResultInputs() {
  if (!btnConfirmAddResult) return false;
  const studentId = String(activeAddResultStudentId ?? "").trim();
  const courseCode = normalizeCourseCode(elAddResultCourse?.value ?? "");
  const sessionRaw = String(elAddResultSession?.value ?? "").trim();
  const parsedSession = parseYearMonth(sessionRaw);
  const semesterRaw = String(elAddResultSemester?.value ?? "").trim();
  const semesterValue = Number(semesterRaw);
  const markRaw = String(elAddResultMark?.value ?? "").trim();
  const verifyOk = elAddResultVerify ? elAddResultVerify.checked : true;

  let message = "";
  let valid = true;
  if (!studentId) {
    message = "Student is required.";
    valid = false;
  } else if (!courseCode) {
    message = "Course is required.";
    valid = false;
  } else if (!parsedSession) {
    message = "Session must be in YYYY/MM format (month 02 or 09).";
    valid = false;
  } else if (!Number.isFinite(semesterValue) || semesterValue < 1) {
    message = "Semester must be a number greater than 0.";
    valid = false;
  } else if (markRaw) {
    const n = Number(markRaw);
    if (!Number.isFinite(n) || n < 0 || n > 100) {
      message = "Mark must be between 0 and 100.";
      valid = false;
    }
  }

  if (valid && parsedSession) {
    const resultId = `${studentId}|${courseCode}|${parsedSession.intake}|${Math.floor(semesterValue)}`;
    if (resultState.results.some((row) => String(row.resultId ?? "").trim() === resultId)) {
      message = "Result already exists for this student/course/session/semester.";
      valid = false;
    }
  }

  if (elAddResultWarning) {
    if (message) {
      elAddResultWarning.textContent = message;
      elAddResultWarning.style.display = "block";
    } else {
      elAddResultWarning.textContent = "";
      elAddResultWarning.style.display = "none";
    }
  }

  btnConfirmAddResult.disabled = !valid || !verifyOk;
  return valid;
}

function openAddResultModal(studentId) {
  const id = String(studentId ?? "").trim();
  if (!id || !elAddResultModal) return;
  activeAddResultStudentId = id;
  const student = buildStudentMap().get(id);
  if (elAddResultStudent) {
    const label = student?.name ? `${student.name} (${id})` : id;
    elAddResultStudent.textContent = label;
  }
  if (elAddResultWarning) {
    elAddResultWarning.textContent = "";
    elAddResultWarning.style.display = "none";
  }
  populateAddResultCourseOptions("");
  if (elAddResultSession) {
    elAddResultSession.value = getLatestSessionForStudent(id) || "";
  }
  if (elAddResultSemester) {
    elAddResultSemester.value = String(getNextSemesterForStudent(id));
  }
  if (elAddResultMark) elAddResultMark.value = "";
  if (elAddResultVerify) elAddResultVerify.checked = false;
  updateAddResultComputed();
  elAddResultModal.classList.add("active");
  elAddResultModal.setAttribute("aria-hidden", "false");
}

function closeAddResultModal() {
  activeAddResultStudentId = "";
  if (!elAddResultModal) return;
  elAddResultModal.classList.remove("active");
  elAddResultModal.setAttribute("aria-hidden", "true");
}


function renderCourseResults(courseCode) {
  activeCourseDetailCode = normalizeCourseCode(courseCode);
  if (elResultViewMode) elResultViewMode.value = "courses";
  resultPaging.page = 1;
  renderResultsView();
}

async function refreshResults() {
  let [students, courses, results] = await Promise.all([
    getAllRecords(STORES.students),
    getAllRecords(STORES.courses),
    getAllRecords(STORES.results),
  ]);
  students = await reconcileStudentsFromResults(students, results);
  results = await backfillResultNamesFromStudents(students, results);
  const previousRevision = enrollmentDataRevision || localStorage.getItem(ENROLLMENT_DATA_REVISION_KEY) || "";

  resultState.students = students;
  resultState.courses = courses;
  resultState.results = results;
  resultPaging.page = 1;
  studentListPaging.page = 1;
  const nextRevision = computeEnrollmentDataRevision(students, courses, results);
  enrollmentDataRevision = nextRevision;
  try {
    localStorage.setItem(ENROLLMENT_DATA_REVISION_KEY, nextRevision);
  } catch (e) {
    console.warn("Failed to persist enrollment data revision", e);
  }
  if (previousRevision && previousRevision !== nextRevision) {
    invalidateEnrollmentConfirmation();
  }

  populateResultFilters();
  renderResultsView();
  refreshEnrollmentModule();
  renderStudentList();
}

function updateEnrollmentSelectedCount() {
  if (!elEnrollmentSelectedCount) return;
  const selectedCount = enrollmentState.offeredCourseCodes.length;
  elEnrollmentSelectedCount.textContent = `${numberFmt.format(selectedCount)} selected (max ${ENROLLMENT_MAX_OFFERED})`;
}

function persistEnrollmentSelection() {
  try {
    const trimmed = [...new Set((enrollmentState.offeredCourseCodes ?? []).map((code) => normalizeCourseCode(code)).filter(Boolean))]
      .sort()
      .slice(0, ENROLLMENT_MAX_OFFERED);
    const payload = JSON.stringify(trimmed);
    const previous = localStorage.getItem(ENROLLMENT_SELECTION_KEY);
    if (payload !== previous) {
      localStorage.setItem(ENROLLMENT_SELECTION_KEY, payload);
      invalidateEnrollmentConfirmation();
    }
  } catch (e) {
    console.warn("Failed to persist enrollment selection", e);
  }
}

function loadEnrollmentSelection() {
  try {
    const stored = JSON.parse(localStorage.getItem(ENROLLMENT_SELECTION_KEY) || "[]");
    if (Array.isArray(stored)) {
      enrollmentState.offeredCourseCodes = stored
        .map((code) => normalizeCourseCode(code))
        .filter(Boolean)
        .slice(0, ENROLLMENT_MAX_OFFERED);
    }
  } catch (e) {
    console.warn("Failed to load enrollment selection", e);
  }
}

function getSelectedEnrollmentCourseCodes() {
  if (!elEnrollmentCourseGrid) return [];
  return [...elEnrollmentCourseGrid.querySelectorAll(".enrollment-course-item.selected")]
    .map((item) => normalizeCourseCode(item.dataset.courseCode))
    .filter(Boolean);
}

function populateEnrollmentCourseOptions() {
  if (!elEnrollmentCourseGrid) return;
  if (!enrollmentState.offeredCourseCodes.length) {
    loadEnrollmentSelection();
  }
  if (!enrollmentKeys.size) {
    loadEnrollmentKeys();
  }
  const previous = new Set((enrollmentState.offeredCourseCodes ?? []).slice(0, ENROLLMENT_MAX_OFFERED));
  const courses = [...buildCourseMap().values()].sort((a, b) =>
    normalizeCourseCode(a.courseCode ?? "").localeCompare(normalizeCourseCode(b.courseCode ?? ""))
  );

  elEnrollmentCourseGrid.textContent = "";
  for (const course of courses) {
    const code = normalizeCourseCode(course.courseCode ?? "");
    if (!code) continue;
    const item = document.createElement("div");
    item.className = "enrollment-course-item";
    item.dataset.courseCode = code;
    item.setAttribute("role", "button");
    item.setAttribute("tabindex", "0");
    const setSelected = (selected) => {
      item.classList.toggle("selected", selected);
      item.setAttribute("aria-pressed", selected ? "true" : "false");
    };
    setSelected(previous.has(code));
    const keyLabel = document.createElement("span");
    keyLabel.className = "enrollment-key-label";
    const keyActions = document.createElement("div");
    keyActions.className = "enrollment-key-actions";
    const keyEditButton = document.createElement("button");
    keyEditButton.type = "button";
    keyEditButton.className = "enrollment-key-btn";
    const keyGenerateButton = document.createElement("button");
    keyGenerateButton.type = "button";
    keyGenerateButton.className = "enrollment-key-btn";
    keyGenerateButton.textContent = "Generate";
    const keyResetButton = document.createElement("button");
    keyResetButton.type = "button";
    keyResetButton.className = "enrollment-key-btn";
    keyResetButton.textContent = "Reset";
    const priorityButton = document.createElement("button");
    priorityButton.type = "button";
    priorityButton.className = "enrollment-key-btn";
    const stopToggleKeyBubble = (event) => {
      if (event.key === "Enter" || event.key === " ") {
        event.stopPropagation();
      }
    };
    keyEditButton.addEventListener("keydown", stopToggleKeyBubble);
    keyGenerateButton.addEventListener("keydown", stopToggleKeyBubble);
    keyResetButton.addEventListener("keydown", stopToggleKeyBubble);
    priorityButton.addEventListener("keydown", stopToggleKeyBubble);
    const updateKeyUi = () => {
      const currentKey = enrollmentKeys.has(code) ? enrollmentKeys.get(code) : undefined;
      const keyText = formatEnrollmentKey(currentKey);
      keyLabel.textContent = `Key: ${keyText === "-" ? "Not set" : keyText}`;
      keyEditButton.textContent = currentKey === undefined ? "Set key" : "Edit key";
      keyResetButton.disabled = currentKey === undefined;
      const isPriority = enrollmentPrioritySet.has(code);
      item.classList.toggle("priority", isPriority);
      priorityButton.textContent = isPriority ? "Priority ✓" : "Priority";
    };
    updateKeyUi();
    keyEditButton.addEventListener("click", (event) => {
      event.stopPropagation();
      const currentKey = enrollmentKeys.has(code) ? enrollmentKeys.get(code) : undefined;
      if (!promptEnrollmentKey(code, currentKey ?? "")) return;
      updateKeyUi();
    });
    keyGenerateButton.addEventListener("click", (event) => {
      event.stopPropagation();
      if (!setEnrollmentKeyValue(code, generateEnrollmentKey(10))) return;
      updateKeyUi();
    });
    keyResetButton.addEventListener("click", (event) => {
      event.stopPropagation();
      if (!setEnrollmentKeyValue(code, undefined)) return;
      updateKeyUi();
    });
    priorityButton.addEventListener("click", (event) => {
      event.stopPropagation();
      const isPriority = enrollmentPrioritySet.has(code);
      if (isPriority) {
        enrollmentPrioritySet.delete(code);
      } else {
        enrollmentPrioritySet.add(code);
        if (!item.classList.contains("selected")) {
          setSelected(true);
        }
      }
      enrollmentState.offeredCourseCodes = getSelectedEnrollmentCourseCodes();
      updateEnrollmentSelectedCount();
      persistEnrollmentSelection();
      saveEnrollmentOptimizerSettings();
      updateKeyUi();
    });
    const toggleSelection = () => {
      const isSelected = item.classList.contains("selected");
      if (!isSelected) {
        const selectedCodes = getSelectedEnrollmentCourseCodes();
        if (selectedCodes.length >= ENROLLMENT_MAX_OFFERED) {
          alert(`Maximum ${ENROLLMENT_MAX_OFFERED} offered courses allowed.`);
          return;
        }
        setSelected(true);
      } else {
        setSelected(false);
        if (enrollmentPrioritySet.has(code)) {
          enrollmentPrioritySet.delete(code);
          saveEnrollmentOptimizerSettings();
        }
      }
      enrollmentState.offeredCourseCodes = getSelectedEnrollmentCourseCodes();
      updateEnrollmentSelectedCount();
      persistEnrollmentSelection();
      updateKeyUi();
    };
    item.addEventListener("click", toggleSelection);
    item.addEventListener("keydown", (event) => {
      if (event.key === "Enter" || event.key === " ") {
        event.preventDefault();
        toggleSelection();
      }
    });
    const content = document.createElement("div");
    const strong = document.createElement("strong");
    strong.textContent = code;
    const span = document.createElement("span");
    span.textContent = String(course.title ?? "").trim() || "Untitled course";
    content.appendChild(strong);
    content.appendChild(span);
    content.appendChild(keyLabel);
    keyActions.appendChild(keyEditButton);
    keyActions.appendChild(keyGenerateButton);
    keyActions.appendChild(keyResetButton);
    keyActions.appendChild(priorityButton);
    content.appendChild(keyActions);
    item.appendChild(content);
    elEnrollmentCourseGrid.appendChild(item);
  }

  enrollmentState.offeredCourseCodes = getSelectedEnrollmentCourseCodes();
  updateEnrollmentSelectedCount();
  persistEnrollmentSelection();
  populateEnrollmentCourseFilterOptions();
}

function getEnrollmentCourseFilterSelection() {
  return [...enrollmentCourseFilterSet.values()];
}

function populateEnrollmentCourseFilterOptions() {
  if (!elEnrollmentCourseFilterButtons) return;
  const courseMap = buildCourseMap();
  const offered = (enrollmentState.offeredCourseCodes.length
    ? enrollmentState.offeredCourseCodes
    : getSelectedEnrollmentCourseCodes()
  )
    .map((code) => normalizeCourseCode(code))
    .filter(Boolean);
  const offeredSet = new Set(offered);
  if (enrollmentCourseFilterSet.size) {
    enrollmentCourseFilterSet = new Set(
      [...enrollmentCourseFilterSet].filter((code) => offeredSet.has(code))
    );
  }

  elEnrollmentCourseFilterButtons.textContent = "";
  if (elEnrollmentCourseFilterEmpty) {
    elEnrollmentCourseFilterEmpty.style.display = offered.length ? "none" : "block";
  }
  if (elEnrollmentCourseFilterHint) {
    elEnrollmentCourseFilterHint.style.display = offered.length ? "block" : "none";
  }
  if (!offered.length) {
    if (enrollmentCourseFilterSet.size) {
      enrollmentCourseFilterSet = new Set();
    }
    return;
  }

  const options = [...new Set(offered)].sort((a, b) => a.localeCompare(b));
  for (const code of options) {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "filter-chip";
    if (enrollmentCourseFilterSet.has(code)) {
      button.classList.add("active");
    }
    const title = String(courseMap.get(code)?.title ?? "").trim();
    button.textContent = code;
    if (title) button.title = title;
    button.addEventListener("click", () => {
      if (enrollmentCourseFilterSet.has(code)) {
        enrollmentCourseFilterSet.delete(code);
      } else {
        enrollmentCourseFilterSet.add(code);
      }
      enrollmentPaging.page = 1;
      populateEnrollmentCourseFilterOptions();
      renderEnrollmentList();
    });
    elEnrollmentCourseFilterButtons.appendChild(button);
  }
}

function populateEnrollmentIntakeSelectors() {
  const students = getEnrollmentStudents();
  const intakeSet = new Set();
  for (const student of students) {
    const intake = formatYearMonth(student.intake, "");
    if (intake) intakeSet.add(intake);
  }
  const intakes = [...intakeSet].sort((a, b) => a.localeCompare(b));

  const fillSelect = (el, includeAll = true, allLabel = "All intakes") => {
    if (!el) return;
    const current = formatYearMonth(el.value ?? "", "");
    el.textContent = "";
    if (includeAll) {
      const option = document.createElement("option");
      option.value = "";
      option.textContent = allLabel;
      el.appendChild(option);
    }
    for (const intake of intakes) {
      const option = document.createElement("option");
      option.value = intake;
      option.textContent = intake;
      if (intake === current) option.selected = true;
      el.appendChild(option);
    }
  };

  fillSelect(elEnrollmentListIntake, true, "All intakes");

  if (elEnrollmentTargetIntakes) {
    const current = enrollmentTargetGroupSet.size
      ? new Set(enrollmentTargetGroupSet)
      : new Set([...elEnrollmentTargetIntakes.selectedOptions].map((opt) => opt.value));
    elEnrollmentTargetIntakes.textContent = "";

    const redFlagOption = document.createElement("option");
    redFlagOption.value = ENROLLMENT_TARGET_REDFLAG_KEY;
    redFlagOption.textContent = "Red Flag Group";
    if (current.has(ENROLLMENT_TARGET_REDFLAG_KEY)) redFlagOption.selected = true;
    elEnrollmentTargetIntakes.appendChild(redFlagOption);

    for (const intake of intakes) {
      const option = document.createElement("option");
      option.value = intake;
      option.textContent = intake;
      if (current.has(intake)) option.selected = true;
      elEnrollmentTargetIntakes.appendChild(option);
    }
  }
}

function updateEnrollmentTargetUi() {
  const useGroups = enrollmentTargetMode === "groups";
  const { intakeSet, includeRedFlags } = getEnrollmentTargetGroups();
  const hasGroups = intakeSet.size > 0 || includeRedFlags;
  const redFlagCount = getRedFlagStudentSet().size;
  const hasRedFlags = redFlagCount > 0;
  const enabled = !useGroups || (hasGroups && (!includeRedFlags || hasRedFlags || intakeSet.size > 0));

  if (elEnrollmentTargetAll) {
    elEnrollmentTargetAll.classList.toggle("active", enrollmentTargetMode === "all");
  }
  if (elEnrollmentTargetIntakesMode) {
    elEnrollmentTargetIntakesMode.classList.toggle("active", useGroups);
  }
  if (elEnrollmentTargetHint) {
    if (useGroups) {
      elEnrollmentTargetHint.style.display = "block";
      if (!hasGroups) {
        elEnrollmentTargetHint.textContent = "Select at least one group to enable course selection.";
      } else if (includeRedFlags && !hasRedFlags && intakeSet.size === 0) {
        elEnrollmentTargetHint.textContent = "No red flag students found.";
      } else if (includeRedFlags) {
        elEnrollmentTargetHint.textContent = `Red flag group selected (${numberFmt.format(redFlagCount)} students).`;
      } else {
        elEnrollmentTargetHint.textContent = "Group selection active.";
      }
    } else {
      elEnrollmentTargetHint.style.display = "none";
    }
  }
  if (elEnrollmentTargetIntakes) {
    elEnrollmentTargetIntakes.disabled = !useGroups;
  }
  if (elEnrollmentCourseGrid) {
    elEnrollmentCourseGrid.classList.toggle("disabled", !enabled);
  }
  if (btnEnrollmentSelectAll) btnEnrollmentSelectAll.disabled = !enabled;
  if (btnEnrollmentClearAll) btnEnrollmentClearAll.disabled = !enabled;
  if (btnEnrollmentBuild) btnEnrollmentBuild.disabled = !enabled;
  if (btnEnrollmentBulkPrint) btnEnrollmentBulkPrint.disabled = !enabled;
  updateEnrollmentLockUi();
}

function updateEnrollmentLockUi() {
  const selected = elEnrollmentTargetIntakes
    ? [...elEnrollmentTargetIntakes.selectedOptions].map((opt) => opt.value).filter(Boolean)
    : [];
  const hasSelection = selected.length > 0;
  if (btnEnrollmentLockGroups) btnEnrollmentLockGroups.disabled = !hasSelection;
  if (btnEnrollmentUnlockGroups) btnEnrollmentUnlockGroups.disabled = !hasSelection;
  if (elEnrollmentLockHint) {
    elEnrollmentLockHint.textContent = hasSelection
      ? "Locking applies to all students in the selected group(s)."
      : "Select group(s) above to lock or unlock slips in bulk.";
  }
}

function getEnrollmentTargetIntakes() {
  if (enrollmentTargetMode !== "groups") return [];
  return [...enrollmentTargetGroupSet.values()];
}

function getEnrollmentTargetMode() {
  return enrollmentTargetMode;
}

function getEnrollmentTargetGroups() {
  const groups = enrollmentTargetMode === "groups" ? [...enrollmentTargetGroupSet.values()] : [];
  const intakeSet = new Set(groups.filter((item) => item && item !== ENROLLMENT_TARGET_REDFLAG_KEY));
  const includeRedFlags = groups.includes(ENROLLMENT_TARGET_REDFLAG_KEY);
  return { intakeSet, includeRedFlags };
}

function getEnrollmentStudents() {
  const studentMap = buildStudentMap();
  for (const [id, student] of studentMap.entries()) {
    if (student?.isDeleted) studentMap.delete(id);
  }
  for (const row of resultState.results) {
    const studentId = String(row.studentId ?? "").trim();
    if (!studentId || studentMap.has(studentId)) continue;
    studentMap.set(studentId, {
      studentId,
      name: "",
      intake: "",
      ic: "",
    });
  }
  return [...studentMap.values()].sort((a, b) => {
    const aName = String(a.name ?? "").trim().toLowerCase();
    const bName = String(b.name ?? "").trim().toLowerCase();
    if (aName && bName) return aName.localeCompare(bName);
    if (aName) return -1;
    if (bName) return 1;
    return String(a.studentId ?? "").localeCompare(String(b.studentId ?? ""));
  });
}

function computeEnrollmentRedFlags() {
  const studentMap = buildStudentMap();
  for (const row of resultState.results) {
    const studentId = String(row.studentId ?? "").trim();
    if (!studentId) continue;
    if (!studentMap.has(studentId)) {
      studentMap.set(studentId, {
        studentId,
        name: "",
        intake: "",
      });
    }
  }

  const progressByStudent = new Map();
  for (const row of resultState.results) {
    const studentId = String(row.studentId ?? "").trim();
    const courseCode = normalizeCourseCode(row.courseCode ?? "");
    if (!studentId || !courseCode) continue;
    let progress = progressByStudent.get(studentId);
    if (!progress) {
      progress = {
        attempted: new Set(),
        failed: new Set(),
        passed: new Set(),
      };
      progressByStudent.set(studentId, progress);
    }
    progress.attempted.add(courseCode);
    if (isPassedResult(row)) progress.passed.add(courseCode);
    if (isFailedResult(row) || isMissingMark(row)) progress.failed.add(courseCode);
  }

  const flags = [];
  for (const [studentId, progress] of progressByStudent.entries()) {
    if (!progress.attempted.size) continue;
    const student = studentMap.get(studentId);
    if (student?.isDeleted) continue;
    const pendingFails = [...progress.failed].filter((code) => !progress.passed.has(code));
    if (pendingFails.length < 3) continue;
    flags.push({
      studentId,
      name: String(student?.name ?? "").trim(),
      intake: String(student?.intake ?? "").trim(),
      failedCount: pendingFails.length,
      reason: `Failed/no mark in ${pendingFails.length} course(s)`,
    });
  }

  return flags.sort((a, b) => {
    if (b.failedCount !== a.failedCount) return b.failedCount - a.failedCount;
    return String(a.studentId ?? "").localeCompare(String(b.studentId ?? ""));
  });
}

function getRedFlagStudentSet() {
  const flags = computeEnrollmentRedFlags();
  enrollmentRedFlagStudents = flags;
  return new Set(flags.map((item) => item.studentId));
}

function appendRedFlagTag(container, studentId) {
  if (!container || !studentId) return;
  const tag = document.createElement("span");
  tag.className = "chip flag-tag";
  tag.textContent = "RED FLAG";
  container.appendChild(tag);
}

function appendLockTag(container, studentId) {
  if (!container || !studentId) return;
  const tag = document.createElement("span");
  tag.className = "chip lock-tag";
  tag.textContent = "LOCKED";
  container.appendChild(tag);
}

function buildSlipIndicator({ icon, label, className }) {
  const wrap = document.createElement("span");
  wrap.className = className ? `slip-indicator ${className}` : "slip-indicator";
  wrap.title = label;
  const iconSpan = document.createElement("span");
  iconSpan.className = "slip-indicator-icon";
  iconSpan.innerHTML = icon;
  const labelSpan = document.createElement("span");
  labelSpan.className = "label";
  labelSpan.textContent = label;
  wrap.appendChild(iconSpan);
  wrap.appendChild(labelSpan);
  return wrap;
}

function normalizeStudentId(value) {
  return String(value ?? "").trim();
}

function isSlipLocked(studentId) {
  const id = normalizeStudentId(studentId);
  if (!id) return false;
  return enrollmentSlipLocks.get(id)?.locked === true;
}

function setSlipLock(studentId, locked, meta = {}, options = {}) {
  const id = normalizeStudentId(studentId);
  if (!id) return false;
  if (locked) {
    const existing = enrollmentSlipLocks.get(id);
    enrollmentSlipLocks.set(id, {
      locked: true,
      lockedAt: existing?.lockedAt || new Date().toISOString(),
      source: meta?.source ?? existing?.source ?? "",
    });
  } else {
    enrollmentSlipLocks.delete(id);
    removeLockedSlipSnapshot(id, { save: false });
  }
  if (options.save !== false) {
    saveEnrollmentSlipLocks();
    saveEnrollmentLockedSlips();
  }
  return true;
}

function getGroupStudentIds(groupKey) {
  const key = String(groupKey ?? "").trim();
  if (!key) return [];
  if (key === ENROLLMENT_TARGET_REDFLAG_KEY) {
    return [...getRedFlagStudentSet().values()].map((id) => String(id ?? "").trim()).filter(Boolean);
  }
  const intakeKey = formatYearMonth(key, "");
  if (!intakeKey) return [];
  const ids = [];
  for (const student of getEnrollmentStudents()) {
    const studentId = String(student?.studentId ?? "").trim();
    if (!studentId) continue;
    const intake = formatYearMonth(student?.intake ?? "", "");
    if (intake === intakeKey) ids.push(studentId);
  }
  return ids;
}

function applyGroupSlipLock({ locked, groups }) {
  const uniqueGroups = Array.isArray(groups)
    ? groups.map((group) => String(group ?? "").trim()).filter(Boolean)
    : [];
  if (!uniqueGroups.length) return 0;
  const ids = new Set();
  for (const group of uniqueGroups) {
    for (const studentId of getGroupStudentIds(group)) {
      if (studentId) ids.add(studentId);
    }
  }
  if (!ids.size) return 0;
  for (const studentId of ids) {
    setSlipLock(studentId, locked, { source: "group" }, { save: false });
  }
  saveEnrollmentSlipLocks();
  if (locked) {
    captureLockedSlipSnapshots([...ids]);
  } else {
    for (const studentId of ids) {
      removeLockedSlipSnapshot(studentId, { save: false });
    }
    saveEnrollmentLockedSlips();
  }
  return ids.size;
}

function buildCsvContent(rows) {
  return rows
    .map((row) =>
      row
        .map((value) => {
          const s = String(value ?? "");
          if (/[",\n]/.test(s)) return `"${s.replace(/"/g, '""')}"`;
          return s;
        })
        .join(",")
    )
    .join("\n");
}

function downloadCsvFile(fileName, rows) {
  if (!rows.length) return;
  const csv = buildCsvContent(rows);
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function getReportCourseOptions() {
  const courseMap = buildCourseMap();
  const codes = new Set();
  for (const code of courseMap.keys()) codes.add(code);
  for (const row of resultState.results) {
    const code = normalizeCourseCode(row.courseCode ?? "");
    if (code) codes.add(code);
  }
  for (const slip of enrollmentState.slips ?? []) {
    for (const course of slip.courses ?? []) {
      const code = normalizeCourseCode(course.courseCode ?? "");
      if (code) codes.add(code);
    }
  }
  return [...codes].sort((a, b) => a.localeCompare(b));
}

function populateReportFilters() {
  if (!elReportIntake || !elReportCourse) return;
  const students = getEnrollmentStudents();
  const intakeSet = new Set();
  for (const student of students) {
    const intake = formatYearMonth(student.intake, "");
    if (intake) intakeSet.add(intake);
  }
  const intakes = [...intakeSet].sort((a, b) => a.localeCompare(b));
  const intakeValue = formatYearMonth(elReportIntake.value, "");
  elReportIntake.textContent = "";
  const allIntake = document.createElement("option");
  allIntake.value = "";
  allIntake.textContent = "All intakes";
  elReportIntake.appendChild(allIntake);
  for (const intake of intakes) {
    const option = document.createElement("option");
    option.value = intake;
    option.textContent = intake;
    if (intake === intakeValue) option.selected = true;
    elReportIntake.appendChild(option);
  }

  const courseValue = normalizeCourseCode(elReportCourse.value ?? "");
  const courses = getReportCourseOptions();
  elReportCourse.textContent = "";
  const allCourse = document.createElement("option");
  allCourse.value = "";
  allCourse.textContent = "All courses";
  elReportCourse.appendChild(allCourse);
  for (const code of courses) {
    const option = document.createElement("option");
    option.value = code;
    option.textContent = code;
    if (code === courseValue) option.selected = true;
    elReportCourse.appendChild(option);
  }
}

function buildActiveReport(options = {}) {
  const type = elReportType?.value ?? "students";
  if (type === "students") {
    buildStudentListReport(options);
    return;
  }
  reportRows = [];
  reportColumns = [];
  if (options?.resetPage !== false) reportPaging.page = 1;
  renderReportPreview();
}

function buildStudentListReport({ resetPage = true } = {}) {
  const students = getEnrollmentStudents();
  const redFlagSet = getRedFlagStudentSet();
  const courseMap = buildCourseMap();
  const creditsByStudent = new Map();
  for (const row of resultState.results) {
    const studentId = String(row.studentId ?? "").trim();
    if (!studentId) continue;
    const courseCode = normalizeCourseCode(row.courseCode ?? "");
    const credits = toNumber(courseMap.get(courseCode)?.credits);
    if (credits === null) continue;
    creditsByStudent.set(studentId, (creditsByStudent.get(studentId) ?? 0) + credits);
  }

  const intakeFilter = formatYearMonth(elReportIntake?.value ?? "", "");
  const courseFilter = normalizeCourseCode(elReportCourse?.value ?? "");
  const includeCurrent = elReportIncludeCurrent?.checked !== false;
  const includePast = elReportIncludePast?.checked !== false;

  const hasCourseFilter = Boolean(courseFilter);
  const slipsByStudent = new Map();
  if (hasCourseFilter && includeCurrent) {
    for (const slip of enrollmentState.slips ?? []) {
      const studentId = String(slip.studentId ?? "").trim();
      if (!studentId) continue;
      if (!slipsByStudent.has(studentId)) slipsByStudent.set(studentId, new Set());
      const set = slipsByStudent.get(studentId);
      for (const course of slip.courses ?? []) {
        const code = normalizeCourseCode(course.courseCode ?? "");
        if (code) set.add(code);
      }
    }
  }

  const resultsByStudent = new Map();
  if (hasCourseFilter && includePast) {
    for (const row of resultState.results) {
      const studentId = String(row.studentId ?? "").trim();
      const code = normalizeCourseCode(row.courseCode ?? "");
      if (!studentId || !code) continue;
      if (!resultsByStudent.has(studentId)) resultsByStudent.set(studentId, new Set());
      resultsByStudent.get(studentId).add(code);
    }
  }

  const filtered = students.filter((student) => {
    const studentId = String(student.studentId ?? "").trim();
    const intake = formatYearMonth(student.intake, "");
    if (intakeFilter && intake !== intakeFilter) return false;
    if (!hasCourseFilter) return true;
    if (!includeCurrent && !includePast) return false;
    const inCurrent = includeCurrent && slipsByStudent.get(studentId)?.has(courseFilter);
    const inPast = includePast && resultsByStudent.get(studentId)?.has(courseFilter);
    return Boolean(inCurrent || inPast);
  });

  const rows = filtered.map((student) => {
    const studentId = String(student.studentId ?? "").trim();
    const takenCredits = creditsByStudent.get(studentId) ?? 0;
    return {
      StudentID: studentId,
      Name: String(student.name ?? "").trim(),
      IC: String(student.ic ?? "").trim(),
      Intake: formatYearMonth(student.intake, ""),
      "Taken Credit": `(${numberFmt.format(takenCredits)}/${ENROLLMENT_TARGET_CREDITS})`,
      RedFlag: redFlagSet.has(studentId) ? "Yes" : "No",
    };
  });

  reportColumns = ["StudentID", "Name", "IC", "Intake", "Taken Credit", "RedFlag"];
  reportRows = rows;
  if (resetPage) reportPaging.page = 1;
  renderReportPreview();
}

function renderReportPreview() {
  if (!elReportPreviewTable || !elReportPreviewBody || !elReportPreviewHead || !elReportPreviewEmpty) return;
  const rows = reportRows ?? [];
  const columns = reportColumns ?? [];
  if (elReportRowCount) elReportRowCount.textContent = `${numberFmt.format(rows.length)} rows`;

  if (!rows.length) {
    elReportPreviewBody.textContent = "";
    elReportPreviewHead.textContent = "";
    elReportPreviewTable.style.display = "none";
    elReportPreviewEmpty.style.display = "block";
    if (elReportPager) elReportPager.style.display = "none";
    if (elReportPageMeta) elReportPageMeta.textContent = "Page 1 of 1";
    if (btnReportDownload) btnReportDownload.disabled = true;
    return;
  }

  elReportPreviewEmpty.style.display = "none";
  elReportPreviewTable.style.display = "table";
  elReportPreviewHead.textContent = "";
  const headRow = document.createElement("tr");
  for (const col of columns) {
    const th = document.createElement("th");
    th.textContent = col;
    headRow.appendChild(th);
  }
  elReportPreviewHead.appendChild(headRow);

  elReportPreviewBody.textContent = "";
  const totalPages = Math.max(1, Math.ceil(rows.length / reportPaging.pageSize));
  if (reportPaging.page > totalPages) reportPaging.page = totalPages;
  const startIndex = (reportPaging.page - 1) * reportPaging.pageSize;
  const pageRows = rows.slice(startIndex, startIndex + reportPaging.pageSize);
  for (const row of pageRows) {
    const tr = document.createElement("tr");
    for (const col of columns) {
      const td = document.createElement("td");
      td.textContent = row[col] ?? "-";
      tr.appendChild(td);
    }
    elReportPreviewBody.appendChild(tr);
  }
  if (elReportPager) elReportPager.style.display = totalPages > 1 ? "flex" : "none";
  if (elReportPageMeta) elReportPageMeta.textContent = `Page ${reportPaging.page} of ${totalPages}`;
  if (btnReportPrev) btnReportPrev.disabled = reportPaging.page <= 1;
  if (btnReportNext) btnReportNext.disabled = reportPaging.page >= totalPages;
  if (btnReportDownload) btnReportDownload.disabled = false;
}

function renderEnrollmentRedFlags() {
  if (!elEnrollmentRedFlagTable || !elEnrollmentRedFlagBody || !elEnrollmentRedFlagEmpty) return;
  const flags = computeEnrollmentRedFlags();
  enrollmentRedFlagStudents = flags;

  if (elEnrollmentRedFlagCount) {
    elEnrollmentRedFlagCount.textContent = numberFmt.format(flags.length);
  }

  if (!flags.length) {
    elEnrollmentRedFlagBody.textContent = "";
    elEnrollmentRedFlagTable.style.display = "none";
    elEnrollmentRedFlagEmpty.style.display = "block";
    return;
  }

  elEnrollmentRedFlagEmpty.style.display = "none";
  elEnrollmentRedFlagTable.style.display = "table";
  elEnrollmentRedFlagBody.textContent = "";

  for (const item of flags) {
    const tr = document.createElement("tr");
    const tdId = document.createElement("td");
    tdId.textContent = item.studentId || "-";
    tr.appendChild(tdId);
    const tdName = document.createElement("td");
    tdName.textContent = item.name || "-";
    tr.appendChild(tdName);
    const tdIntake = document.createElement("td");
    tdIntake.textContent = formatYearMonth(item.intake);
    tr.appendChild(tdIntake);
    const tdCount = document.createElement("td");
    tdCount.textContent = numberFmt.format(item.failedCount ?? 0);
    tr.appendChild(tdCount);
    const tdReason = document.createElement("td");
    tdReason.textContent = item.reason;
    tr.appendChild(tdReason);
    elEnrollmentRedFlagBody.appendChild(tr);
  }
}

function buildEnrollmentSlips(options = {}) {
  const suppressMinAlert = options.suppressMinAlert === true;
  const targetMode = options.targetMode ?? enrollmentTargetMode;
  const targetGroups = Array.isArray(options.targetGroups) ? options.targetGroups : [];
  const forceRebuildLocked =
    options.forceRebuildLocked === true ? true : enrollmentForceRebuildLocked;
  const overrideLockedIds = new Set(
    (options.overrideLockedIds ?? [])
      .map((studentId) => String(studentId ?? "").trim())
      .filter(Boolean)
  );
  const shouldBypassLock = (studentId) =>
    forceRebuildLocked || overrideLockedIds.has(String(studentId ?? "").trim());
  const normalizedGroups = targetGroups.map((group) => String(group ?? "").trim()).filter(Boolean);
  const targetIntakeSet = new Set(
    normalizedGroups
      .filter((group) => group !== ENROLLMENT_TARGET_REDFLAG_KEY)
      .map((intake) => formatYearMonth(intake, ""))
      .filter(Boolean)
  );
  const includeRedFlags = normalizedGroups.includes(ENROLLMENT_TARGET_REDFLAG_KEY);
  const selectedCodes = getSelectedEnrollmentCourseCodes();
  const prioritySet = getEnrollmentPrioritySet(selectedCodes);
  saveEnrollmentOptimizerSettings();
  enrollmentState.offeredCourseCodes = selectedCodes;
  updateEnrollmentSelectedCount();
  persistEnrollmentSelection();

  if (!selectedCodes.length) {
    if (!suppressMinAlert) {
      alert("Please select at least 1 offered course.");
    }
    return;
  }

  const courseMap = buildCourseMap();
  const progressByStudent = new Map();
  for (const row of resultState.results) {
    const studentId = String(row.studentId ?? "").trim();
    const courseCode = normalizeCourseCode(row.courseCode ?? "");
    if (!studentId || !courseCode) continue;
    let progress = progressByStudent.get(studentId);
    if (!progress) {
      progress = {
        attempted: new Set(),
        passed: new Set(),
        failed: new Set(),
        completedCredits: 0,
        currentSemester: null,
      };
      progressByStudent.set(studentId, progress);
    }
    progress.attempted.add(courseCode);
    const semester = toNumber(row.semester);
    if (semester !== null) {
      progress.currentSemester =
        progress.currentSemester === null ? semester : Math.max(progress.currentSemester, semester);
    }
    if (isFailedResult(row)) {
      progress.failed.add(courseCode);
    }
    if (isPassedResult(row)) {
      if (!progress.passed.has(courseCode)) {
        progress.passed.add(courseCode);
        const earned = toNumber(row.creditsEarned);
        const fallbackCredits = toNumber(courseMap.get(courseCode)?.credits);
        progress.completedCredits += earned ?? fallbackCredits ?? 0;
      }
    }
  }

  const redFlags = computeEnrollmentRedFlags();
  enrollmentRedFlagStudents = redFlags;
  const redFlagSet = new Set(redFlags.map((item) => item.studentId));
  const redFlagReasonById = new Map(redFlags.map((item) => [item.studentId, item.reason]));
  let enrollmentStudents = getEnrollmentStudents();
  if (targetMode === "groups" && (targetIntakeSet.size || includeRedFlags)) {
    enrollmentStudents = enrollmentStudents.filter(
      (student) => {
        const intake = formatYearMonth(student.intake, "");
        const studentId = String(student?.studentId ?? "").trim();
        if (includeRedFlags && redFlagSet.has(studentId)) return true;
        if (targetIntakeSet.size && targetIntakeSet.has(intake)) return true;
        return false;
      }
    );
  }
  const repeatByStudent = new Map();
  for (const student of enrollmentStudents) {
    const studentId = String(student?.studentId ?? "").trim();
    if (!studentId) continue;
    const progress = progressByStudent.get(studentId);
    if (!progress) {
      repeatByStudent.set(studentId, new Set());
      continue;
    }
    const repeats = new Set();
    for (const courseCode of progress.failed ?? []) {
      if (!progress.passed?.has(courseCode)) repeats.add(courseCode);
    }
    repeatByStudent.set(studentId, repeats);
  }

    const priorityCodes = new Set(
      [...(prioritySet ?? [])].map((code) => normalizeCourseCode(code)).filter(Boolean)
    );
    const nonPriorityCodes = selectedCodes.filter((code) => !priorityCodes.has(code));
    const requiredPriorityByStudent = new Map();
    const remainingSlotsByStudent = new Map();
    for (const student of enrollmentStudents) {
      const studentId = String(student?.studentId ?? "").trim();
      if (!studentId) continue;
      const progress = progressByStudent.get(studentId) ?? { passed: new Set() };
      const requiredPriority = selectedCodes.filter(
        (code) => priorityCodes.has(code) && !progress.passed.has(code)
      );
      const cappedPriority = requiredPriority.slice(0, enrollmentMaxPerStudent);
      requiredPriorityByStudent.set(studentId, cappedPriority);
      remainingSlotsByStudent.set(
        studentId,
        Math.max(0, enrollmentMaxPerStudent - cappedPriority.length)
      );
    }

    let assignedCourses = null;
    try {
        assignedCourses = assignBalancedCourses({
        students: enrollmentStudents,
        offeredCourseCodes: nonPriorityCodes,
        progressByStudent,
        maxPerStudent: enrollmentMaxPerStudent,
        maxPerStudentById: remainingSlotsByStudent,
        mode: ENROLLMENT_OPTIMIZER_MODE,
        maxSlotsForMcmf: ENROLLMENT_OPTIMIZER_MAX_SLOTS,
        priorityCourseCodes: [...prioritySet],
        priorityWeight: ENROLLMENT_PRIORITY_WEIGHT,
        priorityForAll: ENROLLMENT_PRIORITY_FOR_ALL,
        studentRepeatCourses: repeatByStudent,
        repeatWeight: ENROLLMENT_REPEAT_PRIORITY_WEIGHT,
      });
    } catch (e) {
      console.warn("Enrollment optimizer failed, falling back to first eligible courses.", e);
      assignedCourses = null;
    }

  const slips = [];
  const existingSlipsById = new Map(
    (enrollmentState.slips ?? [])
      .map((slip) => [String(slip?.studentId ?? "").trim(), slip])
      .filter(([studentId]) => studentId)
  );
    for (const student of enrollmentStudents) {
      const studentId = String(student.studentId ?? "").trim();
      if (!studentId) continue;
      const progress = progressByStudent.get(studentId) ?? {
        attempted: new Set(),
        passed: new Set(),
        completedCredits: 0,
        currentSemester: null,
      };
      const isLocked = isSlipLocked(studentId);
      const existingSlip = existingSlipsById.get(studentId);
      const lockedSnapshot = getLockedSlipSnapshot(studentId);
      if (isLocked && !shouldBypassLock(studentId)) {
        if (existingSlip) {
          slips.push(existingSlip);
          continue;
        }
        if (lockedSnapshot) {
          slips.push(lockedSnapshot);
          continue;
        }
      }
      const eligibleCodes = selectedCodes.filter((courseCode) => !progress.passed.has(courseCode));
      const assigned = assignedCourses?.get(studentId);
      const customSelection = enrollmentCustomSelections.get(studentId);
      const cappedPriority = requiredPriorityByStudent.get(studentId) ?? [];
      const eligibleNonPriority = nonPriorityCodes.filter((code) => !progress.passed.has(code));
      let finalCodes = assigned && assigned.length
        ? buildEnrollmentFinalCodes({
            priorityCodes: cappedPriority,
            baseCodes: assigned,
            max: enrollmentMaxPerStudent,
          })
        : buildEnrollmentFinalCodes({
            priorityCodes: cappedPriority,
            baseCodes: eligibleNonPriority,
            max: enrollmentMaxPerStudent,
          });
      let isCustom = false;

      if (Array.isArray(customSelection)) {
        const filteredCustom = customSelection
          .map((code) => normalizeCourseCode(code))
          .filter(Boolean)
          .filter((code) => selectedCodes.includes(code))
          .filter((code) => !progress.passed.has(code));
        finalCodes = buildEnrollmentFinalCodes({
          priorityCodes: cappedPriority,
          baseCodes: filteredCustom,
          max: enrollmentMaxPerStudent,
        });
        isCustom = true;
        enrollmentCustomSelections.set(studentId, filteredCustom);
      }

      const courses = finalCodes.map((courseCode) => {
        const course = courseMap.get(courseCode);
        return {
          courseCode,
          title: String(course?.title ?? "").trim(),
          enrollKey: getEnrollmentKey(courseCode),
          remark: progress.attempted.has(courseCode) ? "Repeat" : "New",
        };
      });
      slips.push({
        studentId,
        name: String(student.name ?? "").trim(),
        intake: String(student.intake ?? "").trim(),
        currentSemester: progress.currentSemester ?? null,
        completedCredits: progress.completedCredits ?? 0,
        isCustom,
        redFlag: redFlagSet.has(studentId),
        redFlagReason: redFlagReasonById.get(studentId) ?? "",
        courses,
      });
    }
    saveEnrollmentCustomSelections();

  let nextSlips = slips;
  if (targetMode === "groups" && (targetIntakeSet.size || includeRedFlags)) {
    const preserved = enrollmentState.slips.filter((slip) => {
      const intake = formatYearMonth(slip.intake, "");
      const studentId = String(slip.studentId ?? "").trim();
      const isRedFlag = includeRedFlags && redFlagSet.has(studentId);
      const isIntake = targetIntakeSet.size && targetIntakeSet.has(intake);
      return !(isRedFlag || isIntake);
    });
    nextSlips = [...preserved, ...slips];
  }
  nextSlips.sort((a, b) => {
    const aName = String(a.name ?? "").trim().toLowerCase();
    const bName = String(b.name ?? "").trim().toLowerCase();
    if (aName && bName) return aName.localeCompare(bName);
    if (aName) return -1;
    if (bName) return 1;
    return String(a.studentId ?? "").localeCompare(String(b.studentId ?? ""));
  });

  enrollmentState.slips = nextSlips;
  enrollmentPaging.page = 1;
  populateEnrollmentCourseFilterOptions();
  renderEnrollmentList();
  renderEnrollmentStats();
  renderEnrollmentRedFlags();
  if (activeEnrollmentStudentId && !slips.some((slip) => slip.studentId === activeEnrollmentStudentId)) {
    closeEnrollmentSlip();
  }
}

  function renderEnrollmentList() {
    if (!elEnrollmentListTable || !elEnrollmentListEmpty || !elEnrollmentListBody) return;
    const search = String(elEnrollmentSearch?.value ?? "").trim().toLowerCase();
    let filtered = search
      ? enrollmentState.slips.filter((slip) =>
          `${slip.studentId} ${slip.name}`.toLowerCase().includes(search)
        )
      : enrollmentState.slips;
    const intakeFilter = formatYearMonth(elEnrollmentListIntake?.value ?? "", "");
    if (intakeFilter) {
      filtered = filtered.filter(
        (slip) => formatYearMonth(slip.intake, "") === intakeFilter
      );
    }
    const redFlagSet = getRedFlagStudentSet();
    const courseFilters = new Set(getEnrollmentCourseFilterSelection());
    if (courseFilters.size) {
      filtered = filtered.filter((slip) => {
        const enrolled = new Set(
          (slip.courses ?? []).map((course) => normalizeCourseCode(course.courseCode)).filter(Boolean)
        );
        for (const code of courseFilters) {
          if (enrolled.has(code)) return true;
        }
        return false;
      });
    }

  if (elEnrollmentStudentCount) {
    elEnrollmentStudentCount.textContent = numberFmt.format(filtered.length);
  }

  if (!filtered.length) {
    elEnrollmentListBody.textContent = "";
    elEnrollmentListTable.style.display = "none";
    elEnrollmentListEmpty.style.display = "block";
    if (elEnrollmentPager) elEnrollmentPager.style.display = "none";
    return;
  }

    elEnrollmentListEmpty.style.display = "none";
    elEnrollmentListTable.style.display = "table";
    elEnrollmentListBody.textContent = "";

    const courseMap = buildCourseMap();
    const allCourseCodes = resultState.courses
      .map((course) => normalizeCourseCode(course.courseCode))
      .filter(Boolean);
    const allCourseSet = new Set(allCourseCodes);

    const totalPages = Math.max(1, Math.ceil(filtered.length / enrollmentPaging.pageSize));
    if (enrollmentPaging.page > totalPages) enrollmentPaging.page = totalPages;
    const startIndex = (enrollmentPaging.page - 1) * enrollmentPaging.pageSize;
    const pageRows = filtered.slice(startIndex, startIndex + enrollmentPaging.pageSize);

  if (elEnrollmentPager) elEnrollmentPager.style.display = "flex";
    if (elEnrollmentPageMeta) {
      elEnrollmentPageMeta.textContent = `Page ${enrollmentPaging.page} of ${totalPages}`;
    }
    if (btnEnrollmentPrev) btnEnrollmentPrev.disabled = enrollmentPaging.page <= 1;
    if (btnEnrollmentNext) btnEnrollmentNext.disabled = enrollmentPaging.page >= totalPages;

    const buildCourseColumn = (title, codes, emptyText, className) => {
      const col = document.createElement("div");
      col.className = className ? `enrollment-grid-col ${className}` : "enrollment-grid-col";
      const header = document.createElement("div");
      header.className = "enrollment-grid-header";
      header.textContent = `${title} (${codes.length})`;
      col.appendChild(header);
      const list = document.createElement("div");
      list.className = "enrollment-chip-list";
      if (!codes.length) {
        const empty = document.createElement("div");
        empty.className = "muted";
        empty.textContent = emptyText;
        list.appendChild(empty);
      } else {
        for (const code of codes) {
          const chip = document.createElement("div");
          chip.className = "enrollment-chip compact";
          const strong = document.createElement("strong");
          strong.textContent = code;
          const span = document.createElement("span");
          span.className = "muted";
          span.textContent = String(courseMap.get(code)?.title ?? "").trim() || "Untitled course";
          chip.appendChild(strong);
          chip.appendChild(span);
          list.appendChild(chip);
        }
      }
      col.appendChild(list);
      return col;
    };

    for (const slip of pageRows) {
      const tr = document.createElement("tr");
      const isLocked = isSlipLocked(slip.studentId);

      const tdId = document.createElement("td");
      tdId.textContent = slip.studentId || "-";
      tr.appendChild(tdId);

      const tdName = document.createElement("td");
      const nameButton = document.createElement("button");
      nameButton.type = "button";
      nameButton.className = "link-button";
      nameButton.textContent = slip.name || "-";
      nameButton.setAttribute(
        "aria-expanded",
        activeEnrollmentInlineId === slip.studentId ? "true" : "false"
      );
      nameButton.addEventListener("click", () => {
        activeEnrollmentInlineId =
          activeEnrollmentInlineId === slip.studentId ? "" : slip.studentId;
        renderEnrollmentList();
      });
      tdName.appendChild(nameButton);
      if (redFlagSet.has(String(slip.studentId ?? "").trim())) {
        appendRedFlagTag(tdName, slip.studentId);
      }
      if (isLocked) {
        appendLockTag(tdName, slip.studentId);
      }
      tr.appendChild(tdName);

    const tdIntake = document.createElement("td");
    tdIntake.textContent = formatYearMonth(slip.intake);
    tr.appendChild(tdIntake);

    const tdSemester = document.createElement("td");
    tdSemester.textContent = slip.currentSemester === null ? "-" : String(slip.currentSemester);
    tr.appendChild(tdSemester);

    const tdProgress = document.createElement("td");
    const completedCredits = toNumber(slip.completedCredits) ?? 0;
    const progressRatio = ENROLLMENT_TARGET_CREDITS
      ? Math.min(1, Math.max(0, completedCredits / ENROLLMENT_TARGET_CREDITS))
      : 0;
    tdProgress.textContent = `${numberFmt.format(completedCredits)} / ${ENROLLMENT_TARGET_CREDITS}`;
    tdProgress.title = percentFmt.format(progressRatio);
    tr.appendChild(tdProgress);

    const tdCount = document.createElement("td");
    tdCount.textContent = numberFmt.format(slip.courses.length);
    tr.appendChild(tdCount);

    const tdStatus = document.createElement("td");
    if (slip.isCustom) {
      tdStatus.textContent = slip.courses.length ? "Custom" : "Custom (empty)";
    } else {
      tdStatus.textContent = slip.courses.length ? "Enrollment Needed" : "No Enrollment";
    }
    if (isLocked) {
      tdStatus.textContent = `${tdStatus.textContent} • Locked`;
    }
    tr.appendChild(tdStatus);

    const tdAction = document.createElement("td");
    const btnView = document.createElement("button");
    btnView.className = "btn secondary small";
    btnView.type = "button";
    btnView.textContent = "View slip";
    btnView.addEventListener("click", () => {
      openEnrollmentSlipWithMode(slip.studentId, { mode: "live" });
    });
      tdAction.appendChild(btnView);

    const btnLock = document.createElement("button");
    btnLock.className = "btn secondary small";
    btnLock.type = "button";
    btnLock.textContent = isLocked ? "Unlock" : "Lock";
    btnLock.addEventListener("click", () => {
      const nextLocked = !isLocked;
      setSlipLock(slip.studentId, nextLocked, { source: "manual" });
      if (nextLocked) {
        setLockedSlipSnapshot(slip.studentId, slip);
      } else {
        removeLockedSlipSnapshot(slip.studentId);
      }
      renderEnrollmentList();
      renderEnrollmentSlip();
    });
      tdAction.appendChild(btnLock);
      tr.appendChild(tdAction);

      elEnrollmentListBody.appendChild(tr);

      if (slip.studentId && slip.studentId === activeEnrollmentInlineId) {
        const passedSet = new Set();
        for (const row of resultState.results) {
          const rowStudentId = String(row.studentId ?? "").trim();
          if (rowStudentId !== slip.studentId) continue;
          const code = normalizeCourseCode(row.courseCode ?? "");
          if (!code) continue;
          if (isPassedResult(row)) passedSet.add(code);
        }

        const enrolledSet = new Set(
          (slip.courses ?? []).map((course) => normalizeCourseCode(course.courseCode)).filter(Boolean)
        );

        const passedCodes = [...passedSet]
          .filter((code) => allCourseSet.has(code))
          .sort((a, b) => a.localeCompare(b));
        const enrolledCodes = [...enrolledSet]
          .filter((code) => allCourseSet.has(code))
          .sort((a, b) => a.localeCompare(b));
        const notYetCodes = allCourseCodes
          .filter((code) => !passedSet.has(code) && !enrolledSet.has(code))
          .sort((a, b) => a.localeCompare(b));

        const inlineRow = document.createElement("tr");
        inlineRow.className = "enrollment-inline-row";
        const inlineCell = document.createElement("td");
        inlineCell.colSpan = 8;

        const wrapper = document.createElement("div");
        wrapper.className = "enrollment-inline-panel";
        const grid = document.createElement("div");
        grid.className = "enrollment-grid";

        grid.appendChild(buildCourseColumn("Passed", passedCodes, "No passed courses.", "passed"));
        grid.appendChild(buildCourseColumn("Enrolled", enrolledCodes, "No enrolled courses.", "enrolled"));
        grid.appendChild(
          buildCourseColumn(
            "Not Yet Taken",
            notYetCodes,
            allCourseCodes.length ? "All courses passed/enrolled." : "No courses found.",
            "pending"
          )
        );

        wrapper.appendChild(grid);
        inlineCell.appendChild(wrapper);
        inlineRow.appendChild(inlineCell);
        elEnrollmentListBody.appendChild(inlineRow);
      }
    }
  }

function renderEnrollmentStats() {
  if (!elEnrollmentStatsTable || !elEnrollmentStatsBody || !elEnrollmentStatsEmpty) return;
  const offered = enrollmentState.offeredCourseCodes.length
    ? enrollmentState.offeredCourseCodes
    : getSelectedEnrollmentCourseCodes();
  const slips = enrollmentState.slips ?? [];

  if (elEnrollmentStatsCourseCount) {
    elEnrollmentStatsCourseCount.textContent = `Courses: ${numberFmt.format(offered.length)}`;
  }
  if (elEnrollmentStatsStudentCount) {
    elEnrollmentStatsStudentCount.textContent = `Students: ${numberFmt.format(slips.length)}`;
  }

  if (!offered.length) {
    elEnrollmentStatsEmpty.textContent = "Select offered courses to see enrollment counts.";
    elEnrollmentStatsEmpty.style.display = "block";
    elEnrollmentStatsTable.style.display = "none";
    return;
  }

  if (!slips.length) {
    elEnrollmentStatsEmpty.textContent = "Build slips to see enrollment counts.";
    elEnrollmentStatsEmpty.style.display = "block";
    elEnrollmentStatsTable.style.display = "none";
    return;
  }

  const counts = new Map();
  const repeatCounts = new Map();
  for (const code of offered) counts.set(code, 0);
  for (const code of offered) repeatCounts.set(code, 0);
  for (const slip of slips) {
    for (const course of slip.courses ?? []) {
      const code = normalizeCourseCode(course.courseCode ?? "");
      if (counts.has(code)) counts.set(code, counts.get(code) + 1);
      if (repeatCounts.has(code) && course.remark === "Repeat") {
        repeatCounts.set(code, repeatCounts.get(code) + 1);
      }
    }
  }

  const courseMap = buildCourseMap();
  const rows = [...counts.entries()]
    .map(([code, count]) => ({
      code,
      title: String(courseMap.get(code)?.title ?? "").trim(),
      count,
      repeatCount: repeatCounts.get(code) ?? 0,
    }))
    .sort((a, b) => {
      if (b.count !== a.count) return b.count - a.count;
      return a.code.localeCompare(b.code);
    });

  elEnrollmentStatsBody.textContent = "";
  for (const row of rows) {
    const tr = document.createElement("tr");
    const tdCode = document.createElement("td");
    tdCode.textContent = row.code || "-";
    tr.appendChild(tdCode);
    const tdTitle = document.createElement("td");
    tdTitle.textContent = row.title || "-";
    tr.appendChild(tdTitle);
    const tdCount = document.createElement("td");
    tdCount.textContent = numberFmt.format(row.count);
    tr.appendChild(tdCount);
    const tdRepeat = document.createElement("td");
    tdRepeat.textContent = numberFmt.format(row.repeatCount ?? 0);
    tr.appendChild(tdRepeat);
    elEnrollmentStatsBody.appendChild(tr);
  }

  elEnrollmentStatsEmpty.style.display = "none";
  elEnrollmentStatsTable.style.display = "table";
}

function openEnrollmentSlip(studentId) {
  openEnrollmentSlipWithMode(studentId, { mode: "live" });
}

function openEnrollmentSlipWithMode(studentId, options = {}) {
  const mode = options.mode === "locked" ? "locked" : "live";
  enrollmentSlipViewMode = mode;
  if (mode === "live") {
    if (!enrollmentState.slips.length) {
      buildEnrollmentSlips({ suppressMinAlert: true });
    }
    if (!enrollmentState.slips.length) {
      alert("Select offered courses in the Enrollment tab to build slips.");
      return;
    }
  }
  activeEnrollmentStudentId = String(studentId ?? "").trim();
  if (!activeEnrollmentStudentId || !elEnrollmentSlipModal) return;
  renderEnrollmentSlip();
  elEnrollmentSlipModal.classList.add("active");
  elEnrollmentSlipModal.setAttribute("aria-hidden", "false");
}

function closeEnrollmentSlip() {
  activeEnrollmentStudentId = "";
  enrollmentSlipViewMode = "live";
  if (!elEnrollmentSlipModal) return;
  elEnrollmentSlipModal.classList.remove("active");
  elEnrollmentSlipModal.setAttribute("aria-hidden", "true");
}

function renderEnrollmentSlip() {
  if (!elEnrollmentSlipBody) return;
  const useLockedView = enrollmentSlipViewMode === "locked";
  const slip = useLockedView
    ? getLockedSlipSnapshot(activeEnrollmentStudentId)
    : enrollmentState.slips.find((row) => row.studentId === activeEnrollmentStudentId);
  if (!slip) {
    if (elEnrollmentSlipCount) elEnrollmentSlipCount.textContent = "0 course(s)";
    if (elEnrollmentSlipStudentMeta) {
      elEnrollmentSlipStudentMeta.textContent = useLockedView
        ? "No locked slip available for this student."
        : "--";
    }
    elEnrollmentSlipBody.textContent = "";
    const elCustomSection = document.getElementById("enrollmentCustomSection");
    if (elCustomSection) elCustomSection.style.display = useLockedView ? "none" : "";
    if (btnEnrollmentLockSlip) {
      btnEnrollmentLockSlip.disabled = true;
    }
    return;
  }

  const isLocked = isSlipLocked(slip.studentId) || useLockedView;
  if (btnEnrollmentLockSlip) {
    btnEnrollmentLockSlip.disabled = false;
    btnEnrollmentLockSlip.textContent = isLocked ? "Unlock Slip" : "Lock Slip";
    btnEnrollmentLockSlip.onclick = () => {
      const nextLocked = !isSlipLocked(slip.studentId);
      setSlipLock(slip.studentId, nextLocked, { source: "manual" });
      if (nextLocked) {
        setLockedSlipSnapshot(slip.studentId, slip);
      } else {
        removeLockedSlipSnapshot(slip.studentId);
      }
      renderEnrollmentList();
      renderEnrollmentSlip();
    };
  }

  const elCustomSection = document.getElementById("enrollmentCustomSection");
  if (elCustomSection) elCustomSection.style.display = useLockedView ? "none" : "";

  if (elEnrollmentSlipStudentMeta) {
    elEnrollmentSlipStudentMeta.textContent = "";
    elEnrollmentSlipStudentMeta.appendChild(buildEnrollmentMetaTable(slip));
  }
    const elCustomList = document.getElementById("enrollmentCustomList");
  const elCustomCount = document.getElementById("enrollmentCustomCount");
  const elCustomMeta = document.getElementById("enrollmentCustomMeta");
  const elCustomStatus = document.getElementById("enrollmentCustomStatus");
  const btnCustomSelectAll = document.getElementById("btnEnrollmentCustomSelectAll");
  const btnCustomClear = document.getElementById("btnEnrollmentCustomClear");
  const btnCustomSave = document.getElementById("btnEnrollmentCustomSave");
  const btnCustomReset = document.getElementById("btnEnrollmentCustomReset");
  const lockEditing = isLocked;
  if (btnCustomSelectAll) btnCustomSelectAll.disabled = lockEditing;
  if (btnCustomClear) btnCustomClear.disabled = lockEditing;
  if (btnCustomSave) btnCustomSave.disabled = lockEditing;
  if (btnCustomReset) btnCustomReset.disabled = lockEditing;
  if (elCustomStatus) {
    elCustomStatus.textContent = lockEditing ? "Slip locked." : " ";
  }

    const courseMap = buildCourseMap();
    const offered = enrollmentState.offeredCourseCodes.length
      ? enrollmentState.offeredCourseCodes
      : getSelectedEnrollmentCourseCodes();
    const prioritySet = new Set(
      [...(getEnrollmentPrioritySet(offered) ?? [])]
        .map((code) => normalizeCourseCode(code))
        .filter(Boolean)
    );

  const progress = { attempted: new Set(), passed: new Set() };
  for (const row of resultState.results) {
    const studentId = String(row.studentId ?? "").trim();
    if (studentId !== slip.studentId) continue;
    const courseCode = normalizeCourseCode(row.courseCode ?? "");
    if (!courseCode) continue;
    progress.attempted.add(courseCode);
    if (isPassedResult(row)) progress.passed.add(courseCode);
  }

  const eligibleCodes = offered.filter((code) => !progress.passed.has(code));
    const requiredPriority = offered.filter(
      (code) => prioritySet.has(code) && !progress.passed.has(code)
    );
    const cappedPriority = requiredPriority.slice(0, enrollmentMaxPerStudent);
    const cappedPrioritySet = new Set(cappedPriority);
    const customSelection = enrollmentCustomSelections.get(slip.studentId);
    const selectedSet = new Set(
      Array.isArray(customSelection) && customSelection.length
        ? customSelection.map((code) => normalizeCourseCode(code)).filter(Boolean)
        : slip.courses.map((course) => normalizeCourseCode(course.courseCode))
    );
    for (const code of cappedPrioritySet) selectedSet.add(code);

    if (elCustomMeta) {
      let meta = "Select any offered course(s) for this student.";
      if (requiredPriority.length > enrollmentMaxPerStudent) {
        meta = `Priority courses exceed max ${enrollmentMaxPerStudent}. Only the first ${enrollmentMaxPerStudent} priority courses (by offered order) will be included.`;
      }
      if (lockEditing) {
        meta = "Slip is locked. Unlock to edit custom enrollment.";
      }
      elCustomMeta.textContent = meta;
    }

    const buildCourseRowsFromCodes = (codes) => {
      const rows = [];
      for (const rawCode of codes) {
        const courseCode = normalizeCourseCode(rawCode);
        if (!courseCode) continue;
        const course = courseMap.get(courseCode);
        rows.push({
          courseCode,
          title: String(course?.title ?? "").trim(),
          enrollKey: getEnrollmentKey(courseCode),
          remark: progress.attempted.has(courseCode) ? "Repeat" : "New",
        });
      }
      return rows;
    };

    const renderSlipCourses = (courses) => {
      if (!elEnrollmentSlipBody) return;
      if (elEnrollmentSlipCount) {
        elEnrollmentSlipCount.textContent = `${numberFmt.format(courses.length)} course(s)`;
      }
      elEnrollmentSlipBody.textContent = "";
      if (!courses.length) {
        const tr = document.createElement("tr");
        const td = document.createElement("td");
        td.colSpan = 4;
        td.textContent = "No enrollment needed.";
        tr.appendChild(td);
        elEnrollmentSlipBody.appendChild(tr);
        return;
      }
      for (const course of courses) {
        const tr = document.createElement("tr");
        const tdCode = document.createElement("td");
        tdCode.textContent = course.courseCode;
        tr.appendChild(tdCode);
        const tdTitle = document.createElement("td");
        tdTitle.textContent = course.title || "-";
        tr.appendChild(tdTitle);
        const tdKey = document.createElement("td");
        tdKey.textContent = formatEnrollmentKey(course.enrollKey);
        tr.appendChild(tdKey);
        const tdRemark = document.createElement("td");
        tdRemark.textContent = course.remark;
        tr.appendChild(tdRemark);
        elEnrollmentSlipBody.appendChild(tr);
      }
    };

    const readSelected = () => {
      const selected = [];
      elCustomList?.querySelectorAll("input[type='checkbox']").forEach((cb) => {
        if (cb.checked) selected.push(normalizeCourseCode(cb.value));
      });
      return selected.filter(Boolean);
    };

    const updateCustomCount = () => {
      const selected = readSelected();
      if (elCustomCount) elCustomCount.textContent = `${selected.length} selected`;
      return selected;
    };

    const refreshSlipPreview = () => {
      if (!elCustomList) {
        renderSlipCourses(slip.courses);
        return;
      }
      const selected = readSelected();
      const finalCodes = buildEnrollmentFinalCodes({
        priorityCodes: cappedPriority,
        baseCodes: selected,
        max: enrollmentMaxPerStudent,
      });
      renderSlipCourses(buildCourseRowsFromCodes(finalCodes));
    };

    if (elCustomList) {
      elCustomList.textContent = "";
    if (!eligibleCodes.length) {
      const empty = document.createElement("div");
      empty.className = "muted";
      empty.textContent = "No eligible courses for this student.";
      elCustomList.appendChild(empty);
    } else {
      for (const code of eligibleCodes) {
        const item = document.createElement("div");
        item.className = "enrollment-custom-item";
        const label = document.createElement("label");
        const checkbox = document.createElement("input");
        checkbox.type = "checkbox";
        checkbox.value = code;
          const isPriority = prioritySet.has(code);
          const isForcedPriority = cappedPrioritySet.has(code);
          checkbox.checked = selectedSet.has(code) || isForcedPriority;
          if (isPriority || lockEditing) checkbox.disabled = true;
          const title = courseMap.get(code)?.title ?? "";
          const remarkBase = progress.attempted.has(code) ? "Repeat" : "New";
          let remark = remarkBase;
          if (isPriority && isForcedPriority) {
            remark = `${remarkBase} • Priority`;
          } else if (isPriority) {
            remark = `${remarkBase} • Priority (max cap)`;
          }
        const strong = document.createElement("strong");
        strong.textContent = code;
        const spanTitle = document.createElement("span");
        spanTitle.textContent = title || "Untitled course";
        const spanRemark = document.createElement("span");
        spanRemark.className = "muted";
        spanRemark.textContent = `Remark: ${remark}`;

        checkbox.addEventListener("click", (event) => {
          event.stopPropagation();
        });
          checkbox.addEventListener("change", () => {
            updateCustomCount();
            refreshSlipPreview();
          });

          item.addEventListener("click", (event) => {
            if (event.target.closest("label")) return;
            if (checkbox.disabled) return;
            checkbox.checked = !checkbox.checked;
            updateCustomCount();
            refreshSlipPreview();
          });

        label.appendChild(checkbox);
        label.appendChild(strong);
        label.appendChild(spanTitle);
        label.appendChild(spanRemark);
        item.appendChild(label);
        elCustomList.appendChild(item);
      }
    }
  }

    updateCustomCount();
    refreshSlipPreview();

    if (btnCustomSelectAll) {
      btnCustomSelectAll.onclick = () => {
        if (lockEditing) return;
        const boxes = [...(elCustomList?.querySelectorAll("input[type='checkbox']") ?? [])];
        boxes.forEach((cb) => {
          if (!cb.disabled) cb.checked = false;
        });
        if (elCustomStatus) elCustomStatus.textContent = " ";
        boxes.forEach((cb) => {
          cb.checked = true;
        });
        updateCustomCount();
        refreshSlipPreview();
      };
    }

    if (btnCustomClear) {
      btnCustomClear.onclick = () => {
        if (lockEditing) return;
        elCustomList?.querySelectorAll("input[type='checkbox']").forEach((cb) => {
          if (!cb.disabled) cb.checked = false;
        });
        if (elCustomStatus) elCustomStatus.textContent = " ";
        updateCustomCount();
        refreshSlipPreview();
      };
    }

    if (btnCustomSave) {
      btnCustomSave.onclick = () => {
        if (lockEditing) return;
        const selected = readSelected();
        enrollmentCustomSelections.set(slip.studentId, selected);
        saveEnrollmentCustomSelections();
        setSlipLock(slip.studentId, true, { source: "custom" });
        buildEnrollmentSlips({ suppressMinAlert: true, overrideLockedIds: [slip.studentId] });
        captureLockedSlipSnapshots([slip.studentId]);
        renderEnrollmentSlip();
        if (elCustomStatus) elCustomStatus.textContent = "Custom enrollment saved.";
      };
    }

    if (btnCustomReset) {
      btnCustomReset.onclick = () => {
        if (lockEditing) return;
        enrollmentCustomSelections.delete(slip.studentId);
        saveEnrollmentCustomSelections();
        buildEnrollmentSlips({ suppressMinAlert: true, overrideLockedIds: [slip.studentId] });
        renderEnrollmentSlip();
        if (elCustomStatus) elCustomStatus.textContent = "Reverted to auto selection.";
      };
    }
  }

function buildEnrollmentMetaTable(slip) {
  const makeCell = (label, value) => {
    const td = document.createElement("td");
    const labelSpan = document.createElement("span");
    labelSpan.textContent = label;
    td.appendChild(labelSpan);
    td.appendChild(document.createElement("br"));
    td.appendChild(document.createTextNode(String(value)));
    return td;
  };
  const table = document.createElement("table");
  const tbody = document.createElement("tbody");
  const row1 = document.createElement("tr");
  row1.appendChild(makeCell("Student Name", slip.name || "-"));
  row1.appendChild(makeCell("Student ID", slip.studentId || "-"));
  tbody.appendChild(row1);
  const row2 = document.createElement("tr");
  row2.appendChild(makeCell("Intake", formatYearMonth(slip.intake)));
  row2.appendChild(
    makeCell(
      "Current Semester",
      slip.currentSemester === null ? "-" : String(slip.currentSemester)
    )
  );
  tbody.appendChild(row2);
  const completedCredits = toNumber(slip.completedCredits) ?? 0;
  const row3 = document.createElement("tr");
  row3.appendChild(
    makeCell(
      "Credits Completed",
      `${numberFmt.format(completedCredits)} / ${ENROLLMENT_TARGET_CREDITS}`
    )
  );
  const statusLabel = slip.courses.length ? "Enrollment Needed" : "No Enrollment";
  const isLocked = isSlipLocked(slip.studentId);
  const remarkRow = slip.redFlag
    ? `
                  <tr>
                    <td colspan="2"><span>Remark</span><br />${escapeHtml(
                      slip.redFlagReason || "Red flag student"
                    )}</td>
                  </tr>
                `
    : "";
  row3.appendChild(
    makeCell("Status", statusLabel)
  );
  tbody.appendChild(row3);
  if (isLocked || slip.redFlag) {
    const rowIndicators = document.createElement("tr");
    const td = document.createElement("td");
    td.colSpan = 2;
    const labelSpan = document.createElement("span");
    labelSpan.textContent = "Indicators";
    td.appendChild(labelSpan);
    td.appendChild(document.createElement("br"));
    const indicators = document.createElement("div");
    indicators.className = "slip-indicators";
    if (isLocked) {
      indicators.appendChild(
        buildSlipIndicator({
          icon: "&#128274;",
          label: "Locked slip",
          className: "lock",
        })
      );
    }
    if (slip.redFlag) {
      indicators.appendChild(
        buildSlipIndicator({
          icon: "&#9888;",
          label: "Red flag student",
          className: "flag",
        })
      );
    }
    td.appendChild(indicators);
    rowIndicators.appendChild(td);
    tbody.appendChild(rowIndicators);
  }
  if (slip.redFlag) {
    const row4 = document.createElement("tr");
    const td = document.createElement("td");
    td.colSpan = 2;
    const labelSpan = document.createElement("span");
    labelSpan.textContent = "Remark";
    td.appendChild(labelSpan);
    td.appendChild(document.createElement("br"));
    td.appendChild(document.createTextNode(slip.redFlagReason || "Red flag student"));
    row4.appendChild(td);
    tbody.appendChild(row4);
  }
  table.appendChild(tbody);
  return table;
}

function formatEnrollmentKey(value) {
  if (value === null) return "No key";
  if (value === undefined) return "-";
  const s = String(value ?? "").trim();
  return s || "-";
}

function buildEnrollmentFinalCodes({ priorityCodes = [], baseCodes = [], max = 0 }) {
  const result = [];
  const seen = new Set();
  const limit = Number.isFinite(max) && max > 0 ? max : Number.MAX_SAFE_INTEGER;

  for (const code of priorityCodes) {
    const normalized = normalizeCourseCode(code);
    if (!normalized || seen.has(normalized)) continue;
    result.push(normalized);
    seen.add(normalized);
    if (result.length >= limit) return result;
  }

  for (const code of baseCodes) {
    const normalized = normalizeCourseCode(code);
    if (!normalized || seen.has(normalized)) continue;
    result.push(normalized);
    seen.add(normalized);
    if (result.length >= limit) break;
  }

  return result;
}

function promptEnrollmentKey(courseCode, existingKey) {
  const normalizedCode = normalizeCourseCode(courseCode);
  if (!normalizedCode) return false;
  const input = window.prompt(
    `Set enrollment key for ${normalizedCode}.\nLeave blank to reset.`,
    existingKey ?? ""
  );
  if (input === null) return false;
  const trimmed = String(input).trim();
  return setEnrollmentKeyValue(normalizedCode, trimmed || undefined);
}

function appendBulkLog(message) {
  if (!elEnrollmentBulkLog) return;
  elEnrollmentBulkLog.textContent += `${message}\n`;
}

function resetBulkLog(message = "") {
  if (!elEnrollmentBulkLog) return;
  elEnrollmentBulkLog.textContent = message;
}

function setBulkLogStatus(message) {
  if (!elEnrollmentBulkLog) return;
  const lines = elEnrollmentBulkLog.textContent.split("\n");
  lines[0] = message;
  elEnrollmentBulkLog.textContent = lines.filter(Boolean).join("\n");
}

function sanitizeFileName(name) {
  const trimmed = String(name ?? "").trim();
  const safe = trimmed.replace(/[<>:"/\\|?*\x00-\x1F]/g, "").replace(/\s+/g, " ").trim();
  return safe || "student";
}

function buildEnrollmentSlipExportNode(slip) {
  const wrapper = document.createElement("div");
  wrapper.className = "slip";
  const safeCourses = Array.isArray(slip.courses) ? slip.courses : [];

  const header = document.createElement("div");
  header.className = "slip-header";
  const headerLeft = document.createElement("div");
  const title = document.createElement("h3");
  title.className = "slip-title";
  title.textContent = "Enrollment Slip";
  headerLeft.appendChild(title);
  const meta = document.createElement("div");
  meta.className = "slip-meta";
  meta.appendChild(buildEnrollmentMetaTable(slip));
  headerLeft.appendChild(meta);
  header.appendChild(headerLeft);
  wrapper.appendChild(header);

  const section = document.createElement("div");
  section.className = "slip-section";
  const heading = document.createElement("h4");
  heading.textContent = "Courses To Enroll";
  section.appendChild(heading);

  const table = document.createElement("table");
  const thead = document.createElement("thead");
  thead.innerHTML = `
    <tr>
      <th>Course Code</th>
      <th>Title</th>
      <th>Enrollment Key</th>
      <th>Remark</th>
    </tr>
  `;
  table.appendChild(thead);
  const tbody = document.createElement("tbody");
  if (!safeCourses.length) {
    const tr = document.createElement("tr");
    const td = document.createElement("td");
    td.colSpan = 4;
    td.textContent = "No enrollment needed.";
    tr.appendChild(td);
    tbody.appendChild(tr);
  } else {
    for (const course of safeCourses) {
      const tr = document.createElement("tr");
      const tdCode = document.createElement("td");
      tdCode.textContent = course.courseCode;
      tr.appendChild(tdCode);
      const tdTitle = document.createElement("td");
      tdTitle.textContent = course.title || "-";
      tr.appendChild(tdTitle);
      const tdKey = document.createElement("td");
      tdKey.textContent = formatEnrollmentKey(course.enrollKey);
      tr.appendChild(tdKey);
      const tdRemark = document.createElement("td");
      tdRemark.textContent = course.remark;
      tr.appendChild(tdRemark);
      tbody.appendChild(tr);
    }
  }
  table.appendChild(tbody);
  section.appendChild(table);
  wrapper.appendChild(section);

  const declaration = document.createElement("div");
  declaration.className = "slip-disclaimer";
  const label = document.createElement("strong");
  label.textContent = "Declaration:";
  declaration.appendChild(label);
  declaration.appendChild(document.createTextNode(` ${ENROLLMENT_DECLARATION_TEXT}`));
  wrapper.appendChild(declaration);

  return wrapper;
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function buildEnrollmentSlipHtml(slip) {
  const safeCourses = Array.isArray(slip.courses) ? slip.courses : [];
  const coursesRows = safeCourses.length
    ? safeCourses
        .map(
          (course) => `
            <tr>
              <td>${escapeHtml(course.courseCode)}</td>
              <td>${escapeHtml(course.title || "-")}</td>
              <td>${escapeHtml(formatEnrollmentKey(course.enrollKey))}</td>
              <td>${escapeHtml(course.remark)}</td>
            </tr>
          `
        )
        .join("")
    : `
      <tr>
        <td colspan="4">No enrollment needed.</td>
      </tr>
    `;

  const completedCredits = toNumber(slip.completedCredits) ?? 0;
  const statusLabel = safeCourses.length ? "Enrollment Needed" : "No Enrollment";
  const isLocked = isSlipLocked(slip.studentId);
  const indicatorsHtml =
    isLocked || slip.redFlag
      ? `
                  <tr>
                    <td colspan="2">
                      <span>Indicators</span><br />
                      <div class="slip-indicators">
                        ${
                          isLocked
                            ? `<span class="slip-indicator lock" title="Locked slip"><span class="slip-indicator-icon">&#128274;</span><span class="label">Locked slip</span></span>`
                            : ""
                        }
                        ${
                          slip.redFlag
                            ? `<span class="slip-indicator flag" title="Red flag student"><span class="slip-indicator-icon">&#9888;</span><span class="label">Red flag student</span></span>`
                            : ""
                        }
                      </div>
                    </td>
                  </tr>
                `
      : "";
  const remarkRow = slip.redFlag
    ? `
                  <tr>
                    <td colspan="2"><span>Remark</span><br />${escapeHtml(
                      slip.redFlagReason || "Red flag student"
                    )}</td>
                  </tr>
                `
    : "";

  return `<!doctype html>
<html lang="en">
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width,initial-scale=1" />
    <title>Enrollment Slip - ${escapeHtml(slip.name || slip.studentId || "Student")}</title>
    <style>
      :root {
        color-scheme: light;
        --ink: #101418;
        --muted: #5f6b7a;
        --border: rgba(16, 20, 24, 0.12);
        --surface: #ffffff;
        --surface-2: #f6f8fb;
        --radius: 16px;
      }
      * { box-sizing: border-box; }
      body { margin: 0; font-family: "Segoe UI", sans-serif; color: var(--ink); background: #fff; }
      .shell { max-width: 900px; margin: 0 auto; padding: 24px; }
      .slip { background: var(--surface); border: 1px solid var(--border); border-radius: var(--radius); padding: 20px; }
      .slip-header { display: flex; justify-content: space-between; align-items: flex-start; gap: 16px; margin-bottom: 16px; }
      .slip-title { margin: 0; font-size: 22px; }
      .slip-meta { font-size: 13px; color: var(--muted); }
      .slip-meta table { width: auto; border-collapse: collapse; }
      .slip-meta td { padding: 4px 12px 4px 0; color: var(--ink); }
      .slip-meta td span { color: var(--muted); font-size: 12px; text-transform: uppercase; letter-spacing: 0.08em; }
      .slip-indicators { display: inline-flex; flex-wrap: wrap; gap: 8px; align-items: center; }
      .slip-indicator { display: inline-flex; align-items: center; gap: 6px; font-size: 12px; color: var(--muted); }
      .slip-indicator .label { display: none; }
      .slip-indicator:hover .label { display: inline; }
      .slip-indicator.lock { color: #1d4ed8; }
      .slip-indicator.flag { color: #b42318; }
      .slip-section h4 { margin: 0 0 8px; font-size: 16px; }
      .slip-disclaimer {
        margin-top: 16px;
        font-size: 12px;
        color: var(--muted);
        border-top: 1px solid var(--border);
        padding-top: 10px;
      }
      .slip-disclaimer strong { color: var(--ink); }
      table { width: 100%; border-collapse: collapse; font-size: 14px; }
      th, td { text-align: left; padding: 8px; border-bottom: 1px solid rgba(16, 20, 24, 0.08); }
      th { font-size: 12px; text-transform: uppercase; letter-spacing: 0.08em; color: var(--muted); }
      @media print {
        @page { size: A4 portrait; margin: 12mm; }
        .shell { padding: 0; }
        .slip { border: none; border-radius: 0; padding: 0; }
        .slip-indicator .label { display: inline; }
      }
    </style>
  </head>
  <body>
    <div class="shell">
      <div class="slip">
        <div class="slip-header">
          <div>
            <h3 class="slip-title">Enrollment Slip</h3>
            <div class="slip-meta">
              <table>
                <tbody>
                  <tr>
                    <td><span>Student Name</span><br />${escapeHtml(slip.name || "-")}</td>
                    <td><span>Student ID</span><br />${escapeHtml(slip.studentId || "-")}</td>
                  </tr>
                  <tr>
                    <td><span>Intake</span><br />${escapeHtml(formatYearMonth(slip.intake))}</td>
                    <td><span>Current Semester</span><br />${escapeHtml(
                      slip.currentSemester === null ? "-" : String(slip.currentSemester)
                    )}</td>
                  </tr>
                  <tr>
                    <td><span>Credits Completed</span><br />${escapeHtml(
                      `${numberFmt.format(completedCredits)} / ${ENROLLMENT_TARGET_CREDITS}`
                    )}</td>
                    <td><span>Status</span><br />${escapeHtml(statusLabel)}</td>
                  </tr>
                  ${indicatorsHtml}
                  ${remarkRow}
                </tbody>
              </table>
            </div>
          </div>
        </div>
        <div class="slip-section">
          <h4>Courses To Enroll</h4>
          <table>
            <thead>
              <tr>
                <th>Course Code</th>
                <th>Title</th>
                <th>Enrollment Key</th>
                <th>Remark</th>
              </tr>
            </thead>
            <tbody>
              ${coursesRows}
            </tbody>
          </table>
        </div>
        <div class="slip-disclaimer">
          <strong>Declaration:</strong> ${escapeHtml(ENROLLMENT_DECLARATION_TEXT)}
        </div>
      </div>
    </div>
  </body>
</html>`;
}

async function bulkExportEnrollmentSlips() {
  if (!btnEnrollmentBulkPrint) return;
  if (!window.JSZip || !window.saveAs) {
    alert("Bulk export libraries not loaded.");
    return;
  }

  resetBulkLog("");

  if (!enrollmentState.offeredCourseCodes.length) {
    buildEnrollmentSlips({ suppressMinAlert: true });
  }
  if (!enrollmentState.offeredCourseCodes.length) {
    alert("Select offered courses in the Enrollment tab before bulk export.");
    return;
  }

  const targetGroups = getEnrollmentTargetIntakes();
  const targetMode = getEnrollmentTargetMode();
  const redFlagSet = getRedFlagStudentSet();
  const { intakeSet, includeRedFlags } = getEnrollmentTargetGroups();
  if (targetMode === "groups" && !(intakeSet.size || includeRedFlags)) {
    alert("Please select at least one group.");
    return;
  }
  if (targetMode === "groups" && includeRedFlags && redFlagSet.size === 0 && intakeSet.size === 0) {
    alert("No red flag students found.");
    return;
  }
  buildEnrollmentSlips({ suppressMinAlert: true, targetGroups, targetMode });

  let slips = enrollmentState.slips.slice();
  if (targetMode === "groups" && (intakeSet.size || includeRedFlags)) {
    slips = slips.filter((slip) => {
      const intake = formatYearMonth(slip.intake, "");
      const studentId = String(slip.studentId ?? "").trim();
      const isIntake = intakeSet.size && intakeSet.has(intake);
      const isRedFlag = includeRedFlags && redFlagSet.has(studentId);
      return isIntake || isRedFlag;
    });
  }
  if (!slips.length) {
    resetBulkLog("No students available for bulk export.");
    return;
  }

  btnEnrollmentBulkPrint.disabled = true;
  const originalLabel = btnEnrollmentBulkPrint.textContent;
  btnEnrollmentBulkPrint.textContent = "Exporting...";
  resetBulkLog(`Written 0 of ${slips.length}`);

  try {
    const zip = new window.JSZip();
    const nameCounts = new Map();
    const failedNames = [];
    let written = 0;

    for (let i = 0; i < slips.length; i += 1) {
      const slip = slips[i];
      const baseName = sanitizeFileName(slip.name || slip.studentId || "student");
      const count = (nameCounts.get(baseName) ?? 0) + 1;
      nameCounts.set(baseName, count);
      const fileName = count > 1 ? `${baseName} (${count}).html` : `${baseName}.html`;

      try {
        const html = buildEnrollmentSlipHtml(slip);
        zip.file(fileName, html);
        written += 1;
        setBulkLogStatus(`Written ${written} of ${slips.length}`);
      } catch (e) {
        const label = slip.name || slip.studentId || "Unknown";
        failedNames.push(`${label} (${e.message ?? e})`);
      }
    }

    setBulkLogStatus(`Written ${written} of ${slips.length}`);
    if (failedNames.length) {
      appendBulkLog(`Failed ${failedNames.length} slip(s):`);
      failedNames.forEach((name) => appendBulkLog(`- ${name}`));
    }
    if (written === 0) {
      appendBulkLog("No slips were exported.");
      alert("No slips were exported. Please check the bulk log for details.");
      return;
    }
      const zipBlob = await zip.generateAsync({ type: "blob" });
    const stamp = new Date().toISOString().slice(0, 10);
    window.saveAs(zipBlob, `enrollment-slips-${stamp}.zip`);
    setBulkLogStatus(`Written ${written} of ${slips.length} (ZIP ready)`);
  } catch (e) {
    console.error(e);
    appendBulkLog(`ERROR: ${e.message ?? e}`);
    alert(`Bulk export failed: ${e.message ?? e}`);
  } finally {
    btnEnrollmentBulkPrint.textContent = originalLabel;
    btnEnrollmentBulkPrint.disabled = false;
  }
}

function computeStudentSummary(studentId) {
  const studentResults = resultState.results.filter(
    (row) => String(row.studentId ?? "").trim() === studentId
  );
  const courseMap = buildCourseMap();
  let creditsAttempted = 0;
  let creditsEarned = 0;
  let failedCount = 0;

  for (const row of studentResults) {
    const courseCode = normalizeCourseCode(row.courseCode ?? "");
    const course = courseMap.get(courseCode);
    const credits = toNumber(course?.credits);
    if (credits !== null) {
      creditsAttempted += credits;
      const earned = toNumber(row.creditsEarned);
      if (earned !== null) {
        creditsEarned += earned;
      } else if (!isFailedResult(row)) {
        creditsEarned += credits;
      }
    }
    if (isFailedResult(row)) failedCount += 1;
  }

  const cgpaByStudent = computeStudentCgpa(studentResults, resultState.courses);
  const cgpa = cgpaByStudent.get(studentId) ?? null;

  return {
    cgpa,
    creditsAttempted,
    creditsEarned,
    failedCount,
    totalResults: studentResults.length,
  };
}

function populateStudentListFilters(students) {
  if (!elStudentListIntake) return;
  const current = formatYearMonth(elStudentListIntake.value, "");
  const intakes = [...new Set(
    students
      .map((student) => formatYearMonth(student.intake, ""))
      .filter(Boolean)
  )].sort();
  elStudentListIntake.textContent = "";
  const optionAll = document.createElement("option");
  optionAll.value = "";
  optionAll.textContent = "All intakes";
  elStudentListIntake.appendChild(optionAll);
  for (const intake of intakes) {
    const option = document.createElement("option");
    option.value = intake;
    option.textContent = intake;
    elStudentListIntake.appendChild(option);
  }
  if (current) elStudentListIntake.value = current;
}

async function saveStudentEdits({
  originalId,
  nextId,
  nextName,
  nextIc,
  intakeRaw,
  statusEl,
}) {
  const setStatus = (msg) => {
    if (statusEl) statusEl.textContent = msg || " ";
  };

  if (!nextId) {
    setStatus("Student ID is required.");
    return false;
  }
  if (!nextName) {
    setStatus("Student name is required.");
    return false;
  }

  const parsedIntake = intakeRaw ? parseYearMonth(intakeRaw) : null;
  if (intakeRaw && !parsedIntake) {
    setStatus("Intake must be in YYYY/MM format (month 02 or 09).");
    return false;
  }

  const normalizedOriginal = String(originalId ?? "").trim();
  const normalizedNext = String(nextId ?? "").trim();
  if (
    normalizedNext !== normalizedOriginal
    && resultState.students.some((s) => String(s.studentId ?? "").trim() === normalizedNext)
  ) {
    setStatus("Student ID already exists. Please use a unique ID.");
    return false;
  }
  if (
    normalizedNext !== normalizedOriginal
    && resultState.results.some((row) => String(row.studentId ?? "").trim() === normalizedNext)
  ) {
    setStatus("Student ID already exists in results. Please use a unique ID.");
    return false;
  }

  const existing = resultState.students.find(
    (s) => String(s.studentId ?? "").trim() === normalizedOriginal
  );
  if (!existing) {
    setStatus("Student record not found.");
    return false;
  }

  const updated = {
    ...existing,
    studentId: normalizedNext,
    name: nextName,
    ic: nextIc,
    intake: parsedIntake ? parsedIntake.intake : "",
    intakeYear: parsedIntake ? parsedIntake.year : "",
    intakeMonth: parsedIntake ? parsedIntake.month : "",
  };

  try {
    if (normalizedNext === normalizedOriginal) {
      await upsertMany(STORES.students, [updated]);
    } else {
      await upsertMany(STORES.students, [updated]);
      await deleteRecord(STORES.students, normalizedOriginal);

      const results = await getAllRecords(STORES.results);
      const affected = results.filter(
        (row) => String(row.studentId ?? "").trim() === normalizedOriginal
      );
      if (affected.length) {
        const updatedResults = affected.map((row) => {
          const courseCode = normalizeCourseCode(row.courseCode ?? "");
          const session = String(row.session ?? "").trim();
          const semester = row.semester ?? "";
          return {
            ...row,
            studentId: normalizedNext,
            resultId: `${normalizedNext}|${courseCode}|${session}|${semester}`,
          };
        });
        await upsertMany(STORES.results, updatedResults);
        await Promise.all(affected.map((row) => deleteRecord(STORES.results, row.resultId)));
      }
    }

    if (activeProfileStudentId === normalizedOriginal) activeProfileStudentId = normalizedNext;
    if (activeStudentInlineId === normalizedOriginal) activeStudentInlineId = normalizedNext;
    studentInlineStatusMessages.set(normalizedNext, "Student profile updated.");

    await refreshResults();
    await refreshStats();
    return true;
  } catch (e) {
    console.error(e);
    setStatus(`Failed to save: ${e.message ?? e}`);
    return false;
  }
}

async function deleteStudentById(studentId, statusEl) {
  const setStatus = (msg) => {
    if (statusEl) statusEl.textContent = msg || " ";
  };
  const normalizedId = String(studentId ?? "").trim();
  if (!normalizedId) return false;
  const summary = computeStudentSummary(normalizedId);
  if (summary.totalResults > 0) {
    setStatus("Cannot delete this student because results exist. Delete results first.");
    return false;
  }
  const ok = window.confirm(
    `Delete student ${normalizedId}? This will hide the student and can be restored later.`
  );
  if (!ok) return false;
  const student = resultState.students.find(
    (s) => String(s.studentId ?? "").trim() === normalizedId
  );
  if (!student) return false;
  try {
    await upsertMany(STORES.students, [{ ...student, isDeleted: true }]);
    if (activeStudentInlineId === normalizedId) activeStudentInlineId = "";
    if (activeProfileStudentId === normalizedId) activeProfileStudentId = "";
    await refreshResults();
    await refreshStats();
    renderDeletedStudentsList();
    return true;
  } catch (e) {
    console.error(e);
    setStatus(`Failed to delete student: ${e.message ?? e}`);
    return false;
  }
}

function renderDeletedStudentsList() {
  if (!elDeletedStudentBody || !elDeletedStudentTable || !elDeletedStudentEmpty) return;
  const deleted = resultState.students
    .filter((student) => student?.isDeleted)
    .sort((a, b) => String(a.studentId ?? "").localeCompare(String(b.studentId ?? "")));

  if (!deleted.length) {
    elDeletedStudentBody.textContent = "";
    elDeletedStudentTable.style.display = "none";
    elDeletedStudentEmpty.style.display = "block";
    return;
  }

  elDeletedStudentEmpty.style.display = "none";
  elDeletedStudentTable.style.display = "table";
  elDeletedStudentBody.textContent = "";
  const redFlagSet = getRedFlagStudentSet();

  for (const student of deleted) {
    const tr = document.createElement("tr");
    const tdId = document.createElement("td");
    tdId.textContent = student.studentId || "-";
    tr.appendChild(tdId);
    const tdName = document.createElement("td");
    tdName.textContent = student.name || "-";
    if (redFlagSet.has(String(student.studentId ?? "").trim())) {
      appendRedFlagTag(tdName, student.studentId);
    }
    tr.appendChild(tdName);
    const tdAction = document.createElement("td");
    const btnRestore = document.createElement("button");
    btnRestore.className = "btn secondary small";
    btnRestore.type = "button";
    btnRestore.textContent = "Restore";
    btnRestore.addEventListener("click", async () => {
      try {
        await upsertMany(STORES.students, [
          { ...student, isDeleted: false },
        ]);
        await refreshResults();
        await refreshStats();
        renderDeletedStudentsList();
      } catch (e) {
        console.error(e);
        alert(`Failed to restore student: ${e.message ?? e}`);
      }
    });
    tdAction.appendChild(btnRestore);
    tr.appendChild(tdAction);
    elDeletedStudentBody.appendChild(tr);
  }
}

function renderStudentProfile() {
  if (!elStudentProfileInfo || !elStudentProfileSummary || !elStudentProfileMeta) return;
  if (!activeProfileStudentId) {
    elStudentProfileMeta.textContent = "--";
    elStudentProfileInfo.textContent = "Select a student to view details.";
    elStudentProfileSummary.textContent = "--";
    if (elStudentProfileStatus) elStudentProfileStatus.textContent = "--";
    if (btnProfileResultSlip) btnProfileResultSlip.disabled = true;
    if (btnProfileEnrollmentSlip) btnProfileEnrollmentSlip.disabled = true;
    if (btnProfileEdit) btnProfileEdit.disabled = true;
    if (btnProfileDelete) btnProfileDelete.disabled = true;
    renderDeletedStudentsList();
    return;
  }

  const student = resultState.students.find(
    (s) => String(s.studentId ?? "").trim() === activeProfileStudentId
  );

  if (!student) {
    elStudentProfileMeta.textContent = "Student not found.";
    elStudentProfileInfo.textContent = "--";
    elStudentProfileSummary.textContent = "--";
    if (elStudentProfileStatus) elStudentProfileStatus.textContent = "--";
    return;
  }

  elStudentProfileMeta.textContent = `${student.name ?? "-"} (${student.studentId ?? "-"})`;

  const form = document.createElement("div");
  form.className = "form-grid";

  const buildField = (label, id, value, wide = false) => {
    const field = document.createElement("div");
    field.className = wide ? "field field-wide" : "field";
    const labelEl = document.createElement("label");
    labelEl.setAttribute("for", id);
    labelEl.textContent = label;
    const input = document.createElement("input");
    input.id = id;
    input.type = "text";
    input.value = value ?? "";
    field.appendChild(labelEl);
    field.appendChild(input);
    return field;
  };

  form.appendChild(buildField("Student ID", "profileStudentId", student.studentId ?? ""));
  form.appendChild(buildField("Name", "profileStudentName", student.name ?? "", true));
  form.appendChild(buildField("IC", "profileStudentIc", student.ic ?? ""));
  form.appendChild(
    buildField("Intake (YYYY/MM)", "profileStudentIntake", formatYearMonth(student.intake, ""))
  );

  const editStatus = document.createElement("div");
  editStatus.className = "muted";
  editStatus.id = "studentProfileEditStatus";
  editStatus.style.marginTop = "8px";

  elStudentProfileInfo.textContent = "";
  elStudentProfileInfo.appendChild(form);
  elStudentProfileInfo.appendChild(editStatus);

  const elId = document.getElementById("profileStudentId");
  const elName = document.getElementById("profileStudentName");
  const elIc = document.getElementById("profileStudentIc");
  const elIntake = document.getElementById("profileStudentIntake");

  const validateProfileInputs = () => {
    if (!editStatus) return true;
    const nextId = String(elId?.value ?? "").trim();
    const nextName = String(elName?.value ?? "").trim();
    const nextIntake = String(elIntake?.value ?? "").trim();
    let message = "";
    let valid = true;

    if (!nextId) {
      message = "Student ID is required.";
      valid = false;
    } else if (
      nextId !== String(activeProfileStudentId ?? "").trim()
      && resultState.students.some((s) => String(s.studentId ?? "").trim() === nextId)
    ) {
      message = "Student ID already exists. Please use a unique ID.";
      valid = false;
    } else if (!nextName) {
      message = "Student name is required.";
      valid = false;
    } else if (nextIntake && !parseYearMonth(nextIntake)) {
      message = "Intake must be in YYYY/MM format (month 02 or 09).";
      valid = false;
    }

    editStatus.textContent = message || " ";
    if (btnProfileEdit) btnProfileEdit.disabled = !valid;
    return valid;
  };

  if (elId) elId.addEventListener("input", validateProfileInputs);
  if (elName) elName.addEventListener("input", validateProfileInputs);
  if (elIntake) elIntake.addEventListener("input", validateProfileInputs);

  validateProfileInputs();

  const summary = computeStudentSummary(activeProfileStudentId);
  const summaryTable = document.createElement("table");
  const summaryBody = document.createElement("tbody");
  summaryBody.innerHTML = `
    <tr>
      <td><span>CGPA</span><br />${summary.cgpa === null ? "--" : formatNumber(summary.cgpa, 2)}</td>
      <td><span>Total Credits Attempted</span><br />${numberFmt.format(summary.creditsAttempted)}</td>
    </tr>
    <tr>
      <td><span>Total Credits Earned</span><br />${numberFmt.format(summary.creditsEarned)}</td>
      <td><span>Failed Courses</span><br />${numberFmt.format(summary.failedCount)}</td>
    </tr>
  `;
  summaryTable.appendChild(summaryBody);
  elStudentProfileSummary.textContent = "";
  elStudentProfileSummary.appendChild(summaryTable);

  const hasResults = summary.totalResults > 0;
  if (elStudentProfileStatus) {
    elStudentProfileStatus.textContent = hasResults
      ? `Results on file: ${numberFmt.format(summary.totalResults)}`
      : "No results on file.";
  }
  if (btnProfileResultSlip) btnProfileResultSlip.disabled = !activeProfileStudentId;
  if (btnProfileEnrollmentSlip) btnProfileEnrollmentSlip.disabled = !activeProfileStudentId;
  if (btnProfileEdit) btnProfileEdit.disabled = !activeProfileStudentId;
  if (btnProfileDelete) btnProfileDelete.disabled = hasResults;

  renderDeletedStudentsList();
}

function openStudentProfile(studentId) {
  const nextId = String(studentId ?? "").trim();
  if (!nextId) return;
  activeStudentInlineId = activeStudentInlineId === nextId ? "" : nextId;
  renderStudentList();
}

function closeStudentProfile() {
  activeStudentInlineId = "";
  renderStudentList();
}

function refreshEnrollmentModule() {
  loadEnrollmentOptimizerSettings();
  loadEnrollmentCustomSelections();
  loadEnrollmentSlipLocks();
  loadEnrollmentLockPreferences();
  loadEnrollmentLockedSlips();
  populateEnrollmentCourseOptions();
  populateEnrollmentCourseFilterOptions();
  populateEnrollmentIntakeSelectors();
  renderEnrollmentRedFlags();
  updateEnrollmentTargetUi();
  if (enrollmentState.offeredCourseCodes.length) {
    buildEnrollmentSlips();
    return;
  }
  enrollmentState.slips = [];
  enrollmentPaging.page = 1;
  renderEnrollmentList();
  renderEnrollmentStats();
  closeEnrollmentSlip();
}

function setActiveTab(tabId) {
  for (const button of tabButtons) {
    const isActive = button.dataset.tab === tabId;
    button.classList.toggle("active", isActive);
    button.setAttribute("aria-selected", isActive ? "true" : "false");
  }

  for (const panel of tabPanels) {
    const isActive = panel.id === `panel-${tabId}`;
    panel.classList.toggle("active", isActive);
  }

  if (tabId === "reporting") {
    populateReportFilters();
    if (!reportAutoGenerated) {
      buildActiveReport({ resetPage: true });
      reportAutoGenerated = true;
    } else {
      renderReportPreview();
    }
  }
}

function renderCourseList(courses) {
  if (!elCourseListTable || !elCourseListEmpty || !elCourseListBody) return;
  if (!courses.length) {
    elCourseListTable.style.display = "none";
    elCourseListEmpty.style.display = "block";
    return;
  }

  elCourseListEmpty.style.display = "none";
  elCourseListTable.style.display = "table";
  elCourseListBody.textContent = "";

  for (const course of courses) {
    const detail = courseCoverageDetails.get(course.courseCode);
    const taken = detail ? detail.taken.length : 0;
    const total = detail ? detail.totalStudents : 0;
    const coverageLabel = total ? `${taken} / ${total}` : "-";
    const tr = document.createElement("tr");

    const tdCode = document.createElement("td");
    tdCode.textContent = course.courseCode;
    tr.appendChild(tdCode);

    const tdTitle = document.createElement("td");
    tdTitle.textContent = course.title || "-";
    tr.appendChild(tdTitle);

    const tdCoverage = document.createElement("td");
    tdCoverage.textContent = coverageLabel;
    tr.appendChild(tdCoverage);

    const tdCredits = document.createElement("td");
    tdCredits.textContent = course.credits === null ? "-" : numberFmt.format(course.credits);
    tr.appendChild(tdCredits);

    const tdMPU = document.createElement("td");
    if (course.isMPUCourse === null) {
      tdMPU.textContent = "-";
    } else {
      tdMPU.textContent = course.isMPUCourse ? "Yes" : "No";
    }
    tr.appendChild(tdMPU);

    const tdAction = document.createElement("td");
    tdAction.className = "table-actions";
    const btnEdit = document.createElement("button");
    btnEdit.className = "btn secondary small";
    btnEdit.type = "button";
    btnEdit.textContent = "Edit";
    btnEdit.addEventListener("click", () => {
      fillCourseForm(course);
      elCourseCode.focus();
    });
    tdAction.appendChild(btnEdit);

    const btnDelete = document.createElement("button");
    btnDelete.className = "btn secondary small";
    btnDelete.type = "button";
    btnDelete.textContent = "Delete";
    btnDelete.addEventListener("click", async () => {
      const courseCode = normalizeCourseCode(course.courseCode ?? "");
      const detail = courseCoverageDetails.get(courseCode);
      if (detail && detail.taken.length > 0) {
        alert("Cannot delete this course because results already exist for it.");
        return;
      }
      const ok = window.confirm(`Delete course ${courseCode}?`);
      if (!ok) return;
      try {
        await deleteCourse(courseCode);
        if (activeCourseEditCode === courseCode) resetCourseForm();
        await refreshCourses();
        await refreshStats();
        await refreshResults();
      } catch (e) {
        console.error(e);
        alert(`Failed to delete course: ${e.message ?? e}`);
      }
    });
    tdAction.appendChild(btnDelete);
    tr.appendChild(tdAction);

    const tdView = document.createElement("td");
    tdView.className = "table-actions";
    const btnView = document.createElement("button");
    btnView.className = "btn secondary small";
    btnView.type = "button";
    btnView.textContent = "View list";
    btnView.addEventListener("click", () => {
      openCourseCoverageDetail(course.courseCode);
    });
    tdView.appendChild(btnView);
    tr.appendChild(tdView);

    elCourseListBody.appendChild(tr);
  }
}

  function renderStudentList() {
    if (!elStudentListTable || !elStudentListBody || !elStudentListEmpty) return;
    const students = getEnrollmentStudents();
    const redFlagSet = getRedFlagStudentSet();
    const studentsWithResults = new Set();
  for (const row of resultState.results) {
    const studentId = String(row.studentId ?? "").trim();
    if (!studentId) continue;
    studentsWithResults.add(studentId);
  }
  populateStudentListFilters(students);

  const search = String(elStudentListSearch?.value ?? "").trim().toLowerCase();
  const intakeFilter = formatYearMonth(elStudentListIntake?.value ?? "", "");
  const statusFilter = String(elStudentListStatus?.value ?? "").trim();
  const hasFilters = Boolean(search || intakeFilter || statusFilter);

  const filtered = students.filter((student) => {
    const studentId = String(student.studentId ?? "").trim();
    const name = String(student.name ?? "").trim();
    const intake = formatYearMonth(student.intake, "");
    if (intakeFilter && intake !== intakeFilter) return false;
    if (statusFilter === "new" && studentsWithResults.has(studentId)) return false;
    if (statusFilter === "has-results" && !studentsWithResults.has(studentId)) return false;
    if (search) {
      const label = `${studentId} ${name}`.toLowerCase();
      if (!label.includes(search)) return false;
    }
    return true;
  });

  if (!students.length) {
    elStudentListBody.textContent = "";
    elStudentListTable.style.display = "none";
    elStudentListEmpty.style.display = "block";
    if (elStudentListPager) elStudentListPager.style.display = "none";
    return;
  }

  if (!filtered.length) {
    elStudentListBody.textContent = "";
    elStudentListTable.style.display = "none";
    elStudentListEmpty.textContent = hasFilters
      ? "No students match the current filters."
      : "No students yet. Import results or students to begin.";
    elStudentListEmpty.style.display = "block";
    if (elStudentListPager) elStudentListPager.style.display = "none";
    return;
  }

  elStudentListEmpty.style.display = "none";
  elStudentListTable.style.display = "table";
  elStudentListBody.textContent = "";

  const totalPages = Math.max(1, Math.ceil(filtered.length / studentListPaging.pageSize));
  if (studentListPaging.page > totalPages) studentListPaging.page = totalPages;
  const startIndex = (studentListPaging.page - 1) * studentListPaging.pageSize;
  const pageRows = filtered.slice(startIndex, startIndex + studentListPaging.pageSize);

  if (elStudentListPager) elStudentListPager.style.display = "flex";
  if (elStudentListPageMeta) {
    elStudentListPageMeta.textContent = `Page ${studentListPaging.page} of ${totalPages}`;
  }
  if (btnStudentListPrev) btnStudentListPrev.disabled = studentListPaging.page <= 1;
  if (btnStudentListNext) btnStudentListNext.disabled = studentListPaging.page >= totalPages;

  for (const student of pageRows) {
    const studentId = String(student.studentId ?? "").trim();
    const tr = document.createElement("tr");

    const tdId = document.createElement("td");
    tdId.textContent = studentId || "-";
    tr.appendChild(tdId);

      const tdName = document.createElement("td");
      if (studentId) {
        const btn = document.createElement("button");
        btn.className = "link-button";
        btn.type = "button";
        btn.textContent = String(student.name ?? "").trim() || studentId;
        btn.addEventListener("click", () => openStudentProfile(studentId));
      tdName.appendChild(btn);
        if (!studentsWithResults.has(studentId)) {
          const tag = document.createElement("span");
          tag.className = "chip new-tag";
          tag.textContent = "NEW";
          tdName.appendChild(tag);
        }
        if (redFlagSet.has(studentId)) {
          appendRedFlagTag(tdName, studentId);
        }
      } else {
        tdName.textContent = String(student.name ?? "").trim() || "-";
      }
    tr.appendChild(tdName);

    const tdAction = document.createElement("td");
    const btnResult = document.createElement("button");
    btnResult.className = "btn secondary small";
    btnResult.type = "button";
    btnResult.textContent = "Result Slip";
    btnResult.addEventListener("click", () => {
      if (!studentId) return;
      openSlip(studentId);
    });
    tdAction.appendChild(btnResult);

    const btnEnroll = document.createElement("button");
    btnEnroll.className = "btn secondary small";
    btnEnroll.type = "button";
    btnEnroll.textContent = "Enrollment Slip";
    btnEnroll.style.marginLeft = "8px";
    btnEnroll.addEventListener("click", () => {
      if (!studentId) return;
      openEnrollmentSlipWithMode(studentId, { mode: "live" });
    });
    tdAction.appendChild(btnEnroll);
    tr.appendChild(tdAction);

    elStudentListBody.appendChild(tr);

    if (studentId && studentId === activeStudentInlineId) {
      const inlineRow = document.createElement("tr");
      inlineRow.className = "student-inline-row";
      const inlineCell = document.createElement("td");
      inlineCell.colSpan = 3;

      const wrapper = document.createElement("div");
      wrapper.className = "inline-editor";

      const form = document.createElement("div");
      form.className = "form-grid";

      const buildField = (label, value, wide = false) => {
        const field = document.createElement("div");
        field.className = wide ? "field field-wide" : "field";
        const labelEl = document.createElement("label");
        labelEl.textContent = label;
        const input = document.createElement("input");
        input.type = "text";
        input.value = value ?? "";
        field.appendChild(labelEl);
        field.appendChild(input);
        return { field, input };
      };

      const idField = buildField("Student ID", student.studentId ?? "");
      const nameField = buildField("Name", student.name ?? "", true);
      const icField = buildField("IC", student.ic ?? "");
      const intakeField = buildField("Intake (YYYY/MM)", formatYearMonth(student.intake, ""));

      form.appendChild(idField.field);
      form.appendChild(nameField.field);
      form.appendChild(icField.field);
      form.appendChild(intakeField.field);
      wrapper.appendChild(form);

      const status = document.createElement("div");
      status.className = "muted";
      status.style.marginTop = "8px";
      status.textContent = studentInlineStatusMessages.get(studentId) ?? " ";
      studentInlineStatusMessages.delete(studentId);
      wrapper.appendChild(status);

      const actions = document.createElement("div");
      actions.className = "form-actions";
      actions.style.marginTop = "10px";

      const btnSave = document.createElement("button");
      btnSave.className = "btn small";
      btnSave.type = "button";
      btnSave.textContent = "Save";
      actions.appendChild(btnSave);

      const btnCancel = document.createElement("button");
      btnCancel.className = "btn secondary small";
      btnCancel.type = "button";
      btnCancel.textContent = "Cancel";
      actions.appendChild(btnCancel);

      const btnInlineResult = document.createElement("button");
      btnInlineResult.className = "btn secondary small";
      btnInlineResult.type = "button";
      btnInlineResult.textContent = "Result Slip";
      actions.appendChild(btnInlineResult);

      const btnInlineEnroll = document.createElement("button");
      btnInlineEnroll.className = "btn secondary small";
      btnInlineEnroll.type = "button";
      btnInlineEnroll.textContent = "Enrollment Slip";
      actions.appendChild(btnInlineEnroll);

      const btnDelete = document.createElement("button");
      btnDelete.className = "btn secondary small";
      btnDelete.type = "button";
      btnDelete.textContent = "Delete";
      btnDelete.disabled = studentsWithResults.has(studentId);
      actions.appendChild(btnDelete);

      wrapper.appendChild(actions);
      inlineCell.appendChild(wrapper);
      inlineRow.appendChild(inlineCell);
      elStudentListBody.appendChild(inlineRow);

      const validateInlineInputs = () => {
        const nextId = String(idField.input.value ?? "").trim();
        const nextName = String(nameField.input.value ?? "").trim();
        const nextIntake = String(intakeField.input.value ?? "").trim();
        let message = "";
        let valid = true;

        if (!nextId) {
          message = "Student ID is required.";
          valid = false;
        } else if (
          nextId !== studentId
          && resultState.students.some((s) => String(s.studentId ?? "").trim() === nextId)
        ) {
          message = "Student ID already exists. Please use a unique ID.";
          valid = false;
        } else if (
          nextId !== studentId
          && resultState.results.some((row) => String(row.studentId ?? "").trim() === nextId)
        ) {
          message = "Student ID already exists in results. Please use a unique ID.";
          valid = false;
        } else if (!nextName) {
          message = "Student name is required.";
          valid = false;
        } else if (nextIntake && !parseYearMonth(nextIntake)) {
          message = "Intake must be in YYYY/MM format (month 02 or 09).";
          valid = false;
        }

        status.textContent = message || " ";
        btnSave.disabled = !valid;
        return valid;
      };

      idField.input.addEventListener("input", validateInlineInputs);
      nameField.input.addEventListener("input", validateInlineInputs);
      intakeField.input.addEventListener("input", validateInlineInputs);
      validateInlineInputs();

      btnSave.addEventListener("click", async () => {
        const ok = await saveStudentEdits({
          originalId: studentId,
          nextId: String(idField.input.value ?? "").trim(),
          nextName: String(nameField.input.value ?? "").trim(),
          nextIc: String(icField.input.value ?? "").trim(),
          intakeRaw: String(intakeField.input.value ?? "").trim(),
          statusEl: status,
        });
        if (ok) activeStudentInlineId = String(idField.input.value ?? "").trim();
      });

      btnCancel.addEventListener("click", () => {
        activeStudentInlineId = "";
        renderStudentList();
      });

      btnInlineResult.addEventListener("click", () => openSlip(studentId));
      btnInlineEnroll.addEventListener("click", () =>
        openEnrollmentSlipWithMode(studentId, { mode: "live" })
      );
      btnDelete.addEventListener("click", async () => {
        await deleteStudentById(studentId, status);
      });
    }
  }
}

function updateSlipConfirmStatus() {
  if (elResultSlipConfirmStatus) {
    const stamp = localStorage.getItem(CONFIRM_RESULT_KEY);
    elResultSlipConfirmStatus.textContent = stamp
      ? `Result: Confirmed ${new Date(stamp).toLocaleString()}`
      : "Result: Not confirmed";
  }
  if (elEnrollmentSlipConfirmStatus) {
    const stamp = localStorage.getItem(CONFIRM_ENROLL_KEY);
    const savedToken = localStorage.getItem(ENROLLMENT_CONFIRM_TOKEN_KEY);
    const currentToken = computeEnrollmentConfirmToken();
    if (stamp && savedToken && savedToken === currentToken) {
      elEnrollmentSlipConfirmStatus.textContent = `Enrollment: Confirmed ${new Date(stamp).toLocaleString()}`;
    } else if (stamp) {
      elEnrollmentSlipConfirmStatus.textContent = "Enrollment: Outdated (reconfirm required)";
    } else {
      elEnrollmentSlipConfirmStatus.textContent = "Enrollment: Not confirmed";
    }
  }
}

function renderCourseStats(courses) {
  const totalCourses = courses.length;
  const totalCredits = courses.reduce((sum, course) => {
    const credits = toNumber(course?.credits);
    return sum + (credits ?? 0);
  }, 0);
  setText(elCourseTotalCount, numberFmt.format(totalCourses));
  setText(elCourseTotalCredits, numberFmt.format(totalCredits));
}

function fillCourseForm(course) {
  activeCourseEditCode = normalizeCourseCode(course.courseCode ?? "");
  elCourseCode.value = course.courseCode ?? "";
  elCourseTitle.value = course.title ?? "";
  elCourseCredits.value = course.credits ?? "";
  if (course.isMPUCourse === null || course.isMPUCourse === undefined) {
    elCourseMPU.value = "";
  } else {
    elCourseMPU.value = course.isMPUCourse ? "true" : "false";
  }

  const detail = courseCoverageDetails.get(activeCourseEditCode);
  const hasCoverage = detail && detail.taken.length > 0;
  if (hasCoverage) {
    elCourseCode.setAttribute("disabled", "disabled");
    if (elCourseFormHint) {
      elCourseFormHint.textContent = "Course code is locked because results exist for this course.";
    }
  } else {
    elCourseCode.removeAttribute("disabled");
    if (elCourseFormHint) {
      elCourseFormHint.textContent = "Use the same course code to update an existing course.";
    }
  }
}

function resetCourseForm() {
  elCourseForm.reset();
  elCourseCode.value = "";
  elCourseTitle.value = "";
  elCourseCredits.value = "";
  elCourseMPU.value = "";
  activeCourseEditCode = "";
  elCourseCode.removeAttribute("disabled");
  if (elCourseFormHint) {
    elCourseFormHint.textContent = "Use the same course code to update an existing course.";
  }
}

async function refreshCourses() {
  const courses = await listCourses();
  renderCourseStats(courses);
  renderCourseList(courses);
}

async function refreshStats() {
  try {
    const [students, courses, results] = await Promise.all([
      getAllRecords(STORES.students),
      getAllRecords(STORES.courses),
      getAllRecords(STORES.results),
    ]);

    const totalResults = results.length;
    const resultsByStudent = new Map();
    const failedStudents = new Set();
    const incompleteStudents = new Set();
    let incompleteCount = 0;
    const lowCgpaStudents = new Set();
    const resultIssueStudents = new Set();
    const seenCoursesByStudent = new Map();

    const statsByCourse = new Map();
    for (const result of results) {
      const courseCode = normalizeCourseCode(result.courseCode ?? "");
      if (!courseCode) continue;
      let entry = statsByCourse.get(courseCode);
      if (!entry) {
        entry = { takenIds: new Set(), results: 0 };
        statsByCourse.set(courseCode, entry);
      }
      entry.results += 1;
      if (result.studentId) entry.takenIds.add(result.studentId);

      const studentId = String(result.studentId ?? "").trim();
      if (studentId) {
        resultsByStudent.set(studentId, (resultsByStudent.get(studentId) ?? 0) + 1);
        if (isFailedResult(result)) failedStudents.add(studentId);
        if (isMissingMark(result)) resultIssueStudents.add(studentId);

        if (!seenCoursesByStudent.has(studentId)) {
          seenCoursesByStudent.set(studentId, new Set());
        }
        const seen = seenCoursesByStudent.get(studentId);
        if (seen.has(courseCode)) {
          resultIssueStudents.add(studentId);
        } else {
          seen.add(courseCode);
        }
      }
    }

    const courseMap = new Map();
    for (const course of courses) {
      const courseCode = normalizeCourseCode(course.courseCode ?? "");
      if (!courseCode) continue;
      courseMap.set(courseCode, String(course.title ?? "").trim());
    }
    for (const [courseCode] of statsByCourse) {
      if (!courseMap.has(courseCode)) courseMap.set(courseCode, "");
    }

    const studentMap = new Map();
    const activeStudents = students.filter((student) => !student?.isDeleted);
    const totalStudents = activeStudents.length;
    for (const student of activeStudents) {
      const id = String(student.studentId ?? "").trim();
      const ic = String(student.ic ?? "").trim();
      if (!id || !ic) {
        if (id) {
          incompleteStudents.add(id);
        } else {
          incompleteCount += 1;
        }
      }
      if (!id) continue;
      studentMap.set(id, student);
    }
    const allStudentIds = [...studentMap.keys()];

    courseCoverageDetails.clear();
    const totalCourses = courseMap.size;
    let activeCourses = 0;
    for (const [code, title] of [...courseMap.entries()].sort((a, b) => a[0].localeCompare(b[0]))) {
      const entry = statsByCourse.get(code);
      const taken = entry ? entry.takenIds.size : 0;
      if (taken > 0) activeCourses += 1;
      const coverage = totalStudents > 0 ? Math.round((taken / totalStudents) * 100) : 0;
      const takenIds = entry ? [...entry.takenIds] : [];
      const notTakenIds = allStudentIds.filter((id) => !(entry?.takenIds?.has(id)));
      const toRecord = (id) => {
        const student = studentMap.get(id);
        return { id, name: String(student?.name ?? "").trim() };
      };
      courseCoverageDetails.set(code, {
        code,
        title,
        coverage,
        totalStudents,
        taken: takenIds.map(toRecord).sort((a, b) => (a.id + a.name).localeCompare(b.id + b.name)),
        notTaken: notTakenIds
          .map(toRecord)
          .sort((a, b) => (a.id + a.name).localeCompare(b.id + b.name)),
      });
    }

    const cgpaByStudentAll = computeStudentCgpa(results, courses);
    for (const [studentId, cgpa] of cgpaByStudentAll.entries()) {
      if (cgpa < 2.0) lowCgpaStudents.add(studentId);
    }

    setText(elTotalStudents, numberFmt.format(activeStudents.length));
    setText(elTotalCourses, numberFmt.format(totalCourses));
    setText(elActiveCourses, numberFmt.format(activeCourses));
    setText(elTotalResults, numberFmt.format(totalResults));
    const getStudentRecord = (studentId) => {
      const student = studentMap.get(studentId) ?? {};
      return {
        studentId,
        name: student.name ?? "",
        ic: student.ic ?? "",
        intake: student.intake ?? "",
      };
    };

    studentIssueLists.failures = [...failedStudents].map(getStudentRecord);
    studentIssueLists.incomplete = [
      ...incompleteStudents,
      ...Array.from({ length: incompleteCount }, () => ""),
    ].map((id) => (id ? getStudentRecord(id) : { studentId: "", name: "", ic: "", intake: "" }));
    studentIssueLists.lowCgpa = [...lowCgpaStudents].map(getStudentRecord);
    studentIssueLists.resultIssues = [...resultIssueStudents].map(getStudentRecord);

    if (elIssueTypeSelect) {
      renderIssueList(elIssueTypeSelect.value);
    }

    setText(elActiveStudents, numberFmt.format(resultsByStudent.size));
    setText(elStudentsWithFailures, numberFmt.format(failedStudents.size));
    setText(elStudentsIncomplete, numberFmt.format(incompleteStudents.size + incompleteCount));
    setText(elStudentsLowCgpa, numberFmt.format(lowCgpaStudents.size));
    setText(elStudentsResultIssues, numberFmt.format(resultIssueStudents.size));
    const activeCount = activeStudents.length;
    const avgResultsPerStudent = activeCount ? totalResults / activeCount : 0;
    setText(elAvgResultsPerStudent, activeCount ? avgResultsPerStudent.toFixed(1) : "--");
    setLastUpdated();
    populateReportFilters();
  } catch (e) {
    console.error(e);
  }
}

function openCourseCoverageDetail(courseCode) {
  if (!elCourseCoverageModal || !elCoverageCourseMeta || !elCoverageTakenCount || !elCoverageNotTakenCount || !elCoverageTakenList || !elCoverageNotTakenList) {
    return;
  }
  const detail = courseCoverageDetails.get(courseCode);
  if (!detail) return;
  activeCoverageCourse = courseCode;
  elCoverageCourseMeta.textContent = `${detail.code} ${detail.title ? "- " + detail.title : ""}`;
  elCoverageTakenCount.textContent = numberFmt.format(detail.taken.length);
  elCoverageNotTakenCount.textContent = numberFmt.format(detail.notTaken.length);
  elCoverageTakenList.textContent = detail.taken.length
    ? detail.taken.map((row) => `${row.id} - ${row.name}`.trim()).join("\n")
    : "—";
  elCoverageNotTakenList.textContent = detail.notTaken.length
    ? detail.notTaken.map((row) => `${row.id} - ${row.name}`.trim()).join("\n")
    : "—";
  elCourseCoverageModal.classList.add("active");
  elCourseCoverageModal.setAttribute("aria-hidden", "false");
}

function closeCourseCoverageDetail() {
  activeCoverageCourse = "";
  elCourseCoverageModal.classList.remove("active");
  elCourseCoverageModal.setAttribute("aria-hidden", "true");
}

function renderIssueList(type) {
  if (!elIssueListTable || !elIssueListEmpty || !elIssueListBody) return;
  const list = studentIssueLists[type] ?? [];
  const redFlagSet = getRedFlagStudentSet();
  const search = (elIssueSearch?.value ?? "").trim().toLowerCase();
  const filtered = search
    ? list.filter((student) => {
        const label = `${student.studentId ?? ""} ${student.name ?? ""}`.toLowerCase();
        return label.includes(search);
      })
    : list;

  if (elIssueListTitle) elIssueListTitle.textContent = issueTitles[type] ?? "Student List";
  if (elIssueListCount) elIssueListCount.textContent = numberFmt.format(filtered.length);

  const totalPages = Math.max(1, Math.ceil(filtered.length / issuePaging.pageSize));
  if (issuePaging.page > totalPages) issuePaging.page = totalPages;
  const startIndex = (issuePaging.page - 1) * issuePaging.pageSize;
  const pageRows = filtered.slice(startIndex, startIndex + issuePaging.pageSize);

  if (!pageRows.length) {
    elIssueListTable.style.display = "none";
    elIssueListEmpty.style.display = "block";
    if (elIssuePager) elIssuePager.style.display = "none";
    return;
  }

  if (elIssuePager) elIssuePager.style.display = "flex";
  if (elIssuePageMeta) elIssuePageMeta.textContent = `Page ${issuePaging.page} of ${totalPages}`;
  if (btnIssuePrev) btnIssuePrev.disabled = issuePaging.page <= 1;
  if (btnIssueNext) btnIssueNext.disabled = issuePaging.page >= totalPages;

  elIssueListEmpty.style.display = "none";
  elIssueListTable.style.display = "table";
  elIssueListBody.textContent = "";
  for (const student of pageRows) {
    const tr = document.createElement("tr");
    const tdId = document.createElement("td");
    tdId.textContent = student.studentId || "-";
    tr.appendChild(tdId);
    const tdName = document.createElement("td");
    tdName.textContent = student.name || "-";
    if (redFlagSet.has(String(student.studentId ?? "").trim())) {
      appendRedFlagTag(tdName, student.studentId);
    }
    tr.appendChild(tdName);
    const tdIc = document.createElement("td");
    tdIc.textContent = student.ic || "-";
    tr.appendChild(tdIc);
      const tdIntake = document.createElement("td");
      tdIntake.textContent = formatYearMonth(student.intake);
      tr.appendChild(tdIntake);
    const tdAction = document.createElement("td");
    const btnView = document.createElement("button");
    btnView.className = "btn secondary small";
    btnView.type = "button";
    btnView.textContent = "View results";
    btnView.addEventListener("click", () => {
      if (student.studentId) {
        openSlip(student.studentId);
      }
    });
    tdAction.appendChild(btnView);
    tr.appendChild(tdAction);
    elIssueListBody.appendChild(tr);
  }
}

function exportCoverageCSV() {
  if (!activeCoverageCourse) return;
  const detail = courseCoverageDetails.get(activeCoverageCourse);
  if (!detail) return;

  const rows = [
    ["Course Code", "Course Title", "Status", "Student ID", "Student Name"],
  ];

  for (const student of detail.taken) {
    rows.push([detail.code, detail.title, "Taken", student.id, student.name]);
  }
  for (const student of detail.notTaken) {
    rows.push([detail.code, detail.title, "Not Taken", student.id, student.name]);
  }

  const csv = rows
    .map((row) =>
      row
        .map((value) => {
          const s = String(value ?? "");
          if (/[",\n]/.test(s)) return `"${s.replace(/"/g, '""')}"`;
          return s;
        })
        .join(",")
    )
    .join("\n");

  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  const safeCode = detail.code || "course";
  link.href = url;
  link.download = `coverage-${safeCode}.csv`;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

if (btnRefresh) {
  btnRefresh.addEventListener("click", async () => {
    await refreshStats();
    await refreshCourses();
    await refreshResults();
  });
}

loadEnrollmentOptimizerSettings();
loadEnrollmentSlipLocks();
loadEnrollmentLockPreferences();
loadEnrollmentLockedSlips();

for (const button of tabButtons) {
  button.addEventListener("click", () => {
    setActiveTab(button.dataset.tab);
  });
}

if (btnEnrollmentSelectAll) {
  btnEnrollmentSelectAll.addEventListener("click", () => {
    if (!elEnrollmentCourseGrid) return;
    const items = [...elEnrollmentCourseGrid.querySelectorAll(".enrollment-course-item")];
    items.forEach((item, index) => {
      const shouldSelect = index < ENROLLMENT_MAX_OFFERED;
      item.classList.toggle("selected", shouldSelect);
      item.setAttribute("aria-pressed", shouldSelect ? "true" : "false");
    });
    if (items.length > ENROLLMENT_MAX_OFFERED) {
      alert(`Selected first ${ENROLLMENT_MAX_OFFERED} courses (maximum allowed).`);
    }
    enrollmentState.offeredCourseCodes = getSelectedEnrollmentCourseCodes();
    updateEnrollmentSelectedCount();
    persistEnrollmentSelection();
  });
}

if (btnEnrollmentClearAll) {
  btnEnrollmentClearAll.addEventListener("click", () => {
    if (!elEnrollmentCourseGrid) return;
    for (const item of elEnrollmentCourseGrid.querySelectorAll(".enrollment-course-item")) {
      item.classList.remove("selected");
      item.setAttribute("aria-pressed", "false");
    }
    enrollmentState.offeredCourseCodes = [];
    updateEnrollmentSelectedCount();
    buildEnrollmentSlips({ suppressMinAlert: true });
  });
}

if (btnEnrollmentGenerateAllKeys) {
  btnEnrollmentGenerateAllKeys.addEventListener("click", () => {
    generateEnrollmentKeysForAllCourses();
  });
}

if (btnEnrollmentResetAllKeys) {
  btnEnrollmentResetAllKeys.addEventListener("click", () => {
    resetEnrollmentKeysForAllCourses();
  });
}

if (btnEnrollmentBuild) {
  btnEnrollmentBuild.addEventListener("click", () => {
    const targetGroups = getEnrollmentTargetIntakes();
    const targetMode = getEnrollmentTargetMode();
    const { intakeSet, includeRedFlags } = getEnrollmentTargetGroups();
    if (targetMode === "groups" && !(intakeSet.size || includeRedFlags)) {
      alert("Please select at least one group.");
      return;
    }
    if (targetMode === "groups" && includeRedFlags && getRedFlagStudentSet().size === 0 && intakeSet.size === 0) {
      alert("No red flag students found.");
      return;
    }
    buildEnrollmentSlips({ targetGroups, targetMode });
  });
}

enrollmentSubtabButtons.forEach((button) => {
  button.addEventListener("click", () => {
    const subtab = button.dataset.enrollmentSubtab || "listing";
    setActiveEnrollmentSubtab(subtab);
  });
});

if (elEnrollmentMaxPerStudent) {
  elEnrollmentMaxPerStudent.addEventListener("change", () => {
    const next = Math.max(1, Math.min(15, Number(elEnrollmentMaxPerStudent.value) || 1));
    enrollmentMaxPerStudent = next;
    elEnrollmentMaxPerStudent.value = String(next);
    saveEnrollmentOptimizerSettings();
  });
}

if (btnEnrollmentBulkPrint) {
  btnEnrollmentBulkPrint.addEventListener("click", () => {
    bulkExportEnrollmentSlips();
  });
}

if (elEnrollmentSearch) {
  elEnrollmentSearch.addEventListener("input", () => {
    enrollmentPaging.page = 1;
    renderEnrollmentList();
  });
}

if (elEnrollmentTargetAll) {
  elEnrollmentTargetAll.addEventListener("click", () => {
    enrollmentTargetMode = "all";
    updateEnrollmentTargetUi();
  });
}

if (elEnrollmentTargetIntakesMode) {
  elEnrollmentTargetIntakesMode.addEventListener("click", () => {
    enrollmentTargetMode = "groups";
    updateEnrollmentTargetUi();
  });
}

if (elEnrollmentTargetIntakes) {
  elEnrollmentTargetIntakes.addEventListener("change", () => {
    enrollmentTargetGroupSet = new Set(
      [...elEnrollmentTargetIntakes.selectedOptions].map((opt) => opt.value).filter(Boolean)
    );
    updateEnrollmentTargetUi();
  });
}

if (btnEnrollmentLockGroups) {
  btnEnrollmentLockGroups.addEventListener("click", () => {
    const groups = elEnrollmentTargetIntakes
      ? [...elEnrollmentTargetIntakes.selectedOptions].map((opt) => opt.value).filter(Boolean)
      : [];
    if (!groups.length) {
      alert("Select at least one group to lock.");
      return;
    }
    const count = applyGroupSlipLock({ locked: true, groups });
    if (!count) {
      alert("No students found for the selected groups.");
      return;
    }
    renderEnrollmentList();
    renderEnrollmentSlip();
    alert(`Locked ${numberFmt.format(count)} student(s).`);
  });
}

if (btnEnrollmentUnlockGroups) {
  btnEnrollmentUnlockGroups.addEventListener("click", () => {
    const groups = elEnrollmentTargetIntakes
      ? [...elEnrollmentTargetIntakes.selectedOptions].map((opt) => opt.value).filter(Boolean)
      : [];
    if (!groups.length) {
      alert("Select at least one group to unlock.");
      return;
    }
    const count = applyGroupSlipLock({ locked: false, groups });
    if (!count) {
      alert("No students found for the selected groups.");
      return;
    }
    renderEnrollmentList();
    renderEnrollmentSlip();
    alert(`Unlocked ${numberFmt.format(count)} student(s).`);
  });
}

if (elEnrollmentForceRebuildLocked) {
  elEnrollmentForceRebuildLocked.addEventListener("change", () => {
    enrollmentForceRebuildLocked = elEnrollmentForceRebuildLocked.checked;
    saveEnrollmentLockPreferences();
  });
}

if (elEnrollmentListIntake) {
  elEnrollmentListIntake.addEventListener("change", () => {
    enrollmentPaging.page = 1;
    renderEnrollmentList();
  });
}

if (btnEnrollmentPrev) {
  btnEnrollmentPrev.addEventListener("click", () => {
    if (enrollmentPaging.page > 1) {
      enrollmentPaging.page -= 1;
      renderEnrollmentList();
    }
  });
}

if (btnEnrollmentNext) {
  btnEnrollmentNext.addEventListener("click", () => {
    enrollmentPaging.page += 1;
    renderEnrollmentList();
  });
}

if (elReportType) {
  elReportType.addEventListener("change", () => {
    populateReportFilters();
    buildActiveReport({ resetPage: true });
    reportAutoGenerated = true;
  });
}

const handleReportFilterChange = () => {
  buildActiveReport({ resetPage: true });
  reportAutoGenerated = true;
};

if (elReportIntake) {
  elReportIntake.addEventListener("change", handleReportFilterChange);
}
if (elReportCourse) {
  elReportCourse.addEventListener("change", handleReportFilterChange);
}
if (elReportIncludeCurrent) {
  elReportIncludeCurrent.addEventListener("change", handleReportFilterChange);
}
if (elReportIncludePast) {
  elReportIncludePast.addEventListener("change", handleReportFilterChange);
}

if (btnReportGenerate) {
  btnReportGenerate.addEventListener("click", () => {
    buildActiveReport({ resetPage: true });
    reportAutoGenerated = true;
  });
}

if (btnReportDownload) {
  btnReportDownload.addEventListener("click", () => {
    if (!reportRows.length || !reportColumns.length) return;
    const rows = [reportColumns, ...reportRows.map((row) => reportColumns.map((col) => row[col]))];
    const timestamp = new Date();
    const pad = (value) => String(value).padStart(2, "0");
    const type = elReportType?.value ?? "students";
    const fileName = `report-${type}-${timestamp.getFullYear()}${pad(timestamp.getMonth() + 1)}${pad(timestamp.getDate())}-${pad(timestamp.getHours())}${pad(timestamp.getMinutes())}.csv`;
    downloadCsvFile(fileName, rows);
  });
}

if (btnReportPrev) {
  btnReportPrev.addEventListener("click", () => {
    if (reportPaging.page > 1) {
      reportPaging.page -= 1;
      renderReportPreview();
    }
  });
}

if (btnReportNext) {
  btnReportNext.addEventListener("click", () => {
    const totalPages = Math.max(1, Math.ceil((reportRows ?? []).length / reportPaging.pageSize));
    if (reportPaging.page < totalPages) {
      reportPaging.page += 1;
      renderReportPreview();
    }
  });
}

resultSubtabButtons.forEach((button) => {
  button.addEventListener("click", () => {
    setActiveResultSubtab(button.dataset.subtab);
  });
});

if (elResultSessionSelect) {
  elResultSessionSelect.addEventListener("change", () => {
    resultPaging.page = 1;
    activeCourseDetailCode = "";
    renderResultsView();
  });
}
if (elResultSearch) {
  elResultSearch.addEventListener("input", () => {
    resultPaging.page = 1;
    renderResultsView();
  });
}
if (elResultViewMode) {
  elResultViewMode.addEventListener("change", () => {
    resultPaging.page = 1;
    activeCourseDetailCode = "";
    renderResultsView();
  });
}


if (btnBackToCourseList) {
  btnBackToCourseList.addEventListener("click", () => {
    activeCourseDetailCode = "";
    resultPaging.page = 1;
    renderResultsView();
  });
}

if (btnCloseStudentEdit) {
  btnCloseStudentEdit.addEventListener("click", () => {
    closeStudentEdit();
  });
}

if (btnCloseStudentProfile) {
  btnCloseStudentProfile.addEventListener("click", () => {
    closeStudentProfile();
  });
}

if (elStudentProfileDrawer) {
  elStudentProfileDrawer.addEventListener("click", (event) => {
    if (event.target === elStudentProfileDrawer) {
      closeStudentProfile();
    }
  });
}

if (btnSaveStudentEdit) {
  btnSaveStudentEdit.addEventListener("click", async () => {
    if (!activeStudentEditId) return;
    const nextName = elEditStudentName.value.trim();
    const nextIc = elEditStudentIc.value.trim();
    const intakeRaw = String(elEditStudentIntake.value ?? "").trim();
    const parsedIntake = intakeRaw ? parseYearMonth(intakeRaw) : null;
    if (!nextName) {
      alert("Student name is required.");
      return;
    }
    if (intakeRaw && !parsedIntake) {
      alert("Intake must be in YYYY/MM format (month 02 or 09).");
      return;
    }
    const existing = resultState.students.find(
      (s) => String(s.studentId ?? "").trim() === String(activeStudentEditId ?? "").trim()
    );
    if (!existing) {
      alert("Student record not found.");
      return;
    }
    const updated = {
      ...existing,
      studentId: activeStudentEditId,
      name: nextName,
      ic: nextIc,
      intake: parsedIntake ? parsedIntake.intake : "",
      intakeYear: parsedIntake ? parsedIntake.year : "",
      intakeMonth: parsedIntake ? parsedIntake.month : "",
    };
    try {
      await upsertMany(STORES.students, [updated]);
      await refreshResults();
      await refreshStats();
      closeStudentEdit();
    } catch (e) {
      console.error(e);
      alert(`Failed to save: ${e.message ?? e}`);
    }
  });
}

if (btnProfileResultSlip) {
  btnProfileResultSlip.addEventListener("click", () => {
    if (!activeProfileStudentId) return;
    openSlip(activeProfileStudentId);
  });
}

if (btnProfileEnrollmentSlip) {
  btnProfileEnrollmentSlip.addEventListener("click", () => {
    if (!activeProfileStudentId) return;
    openEnrollmentSlipWithMode(activeProfileStudentId, { mode: "locked" });
  });
}

if (btnProfileEdit) {
  btnProfileEdit.addEventListener("click", async () => {
    if (!activeProfileStudentId) return;
    const elId = document.getElementById("profileStudentId");
    const elName = document.getElementById("profileStudentName");
    const elIc = document.getElementById("profileStudentIc");
    const elIntake = document.getElementById("profileStudentIntake");
    const elEditStatus = document.getElementById("studentProfileEditStatus");
    if (!elId || !elName || !elIc || !elIntake) return;

    const nextId = String(elId.value ?? "").trim();
    const nextName = String(elName.value ?? "").trim();
    const nextIc = String(elIc.value ?? "").trim();
    const intakeRaw = String(elIntake.value ?? "").trim();
    const parsedIntake = intakeRaw ? parseYearMonth(intakeRaw) : null;

    if (!nextId) {
      if (elEditStatus) elEditStatus.textContent = "Student ID is required.";
      return;
    }
    if (!nextName) {
      if (elEditStatus) elEditStatus.textContent = "Student name is required.";
      return;
    }
    if (intakeRaw && !parsedIntake) {
      if (elEditStatus) {
        elEditStatus.textContent = "Intake must be in YYYY/MM format (month 02 or 09).";
      }
      return;
    }
    if (
      nextId !== String(activeProfileStudentId ?? "").trim()
      && resultState.students.some((s) => String(s.studentId ?? "").trim() === nextId)
    ) {
      if (elEditStatus) {
        elEditStatus.textContent = "Student ID already exists. Please use a unique ID.";
      }
      return;
    }

    const existing = resultState.students.find(
      (s) => String(s.studentId ?? "").trim() === String(activeProfileStudentId ?? "").trim()
    );
    if (!existing) {
      if (elEditStatus) elEditStatus.textContent = "Student record not found.";
      return;
    }

    const updated = {
      ...existing,
      studentId: nextId,
      name: nextName,
      ic: nextIc,
      intake: parsedIntake ? parsedIntake.intake : "",
      intakeYear: parsedIntake ? parsedIntake.year : "",
      intakeMonth: parsedIntake ? parsedIntake.month : "",
    };

    try {
      if (nextId === activeProfileStudentId) {
        await upsertMany(STORES.students, [updated]);
      } else {
        await upsertMany(STORES.students, [updated]);
        await deleteRecord(STORES.students, activeProfileStudentId);

        const results = await getAllRecords(STORES.results);
        const affected = results.filter(
          (row) => String(row.studentId ?? "").trim() === String(activeProfileStudentId ?? "").trim()
        );
        if (affected.length) {
          const updatedResults = affected.map((row) => {
            const courseCode = normalizeCourseCode(row.courseCode ?? "");
            const session = String(row.session ?? "").trim();
            const semester = row.semester ?? "";
            return {
              ...row,
              studentId: nextId,
              resultId: `${nextId}|${courseCode}|${session}|${semester}`,
            };
          });
          await upsertMany(STORES.results, updatedResults);
          await Promise.all(affected.map((row) => deleteRecord(STORES.results, row.resultId)));
        }
      }

      if (nextId !== activeProfileStudentId) {
        if (isSlipLocked(activeProfileStudentId)) {
          const lockedMeta = enrollmentSlipLocks.get(activeProfileStudentId);
          const lockedSlip = getLockedSlipSnapshot(activeProfileStudentId);
          if (lockedMeta) {
            enrollmentSlipLocks.set(nextId, lockedMeta);
            enrollmentSlipLocks.delete(activeProfileStudentId);
            saveEnrollmentSlipLocks();
          }
          if (lockedSlip) {
            enrollmentLockedSlipSnapshots.set(nextId, lockedSlip);
            enrollmentLockedSlipSnapshots.delete(activeProfileStudentId);
            saveEnrollmentLockedSlips();
          }
        }
      }
      activeProfileStudentId = nextId;
      await refreshResults();
      await refreshStats();
      const status = document.getElementById("studentProfileEditStatus");
      if (status) status.textContent = "Student profile updated.";
    } catch (e) {
      console.error(e);
      const status = document.getElementById("studentProfileEditStatus");
      if (status) status.textContent = `Failed to save: ${e.message ?? e}`;
    }
  });
}

if (btnProfileDelete) {
  btnProfileDelete.addEventListener("click", async () => {
    if (!activeProfileStudentId) return;
    const summary = computeStudentSummary(activeProfileStudentId);
    if (summary.totalResults > 0) {
      alert("Cannot delete this student because results exist. Delete results first.");
      return;
    }
    const ok = window.confirm(
      `Delete student ${activeProfileStudentId}? This will hide the student and can be restored later.`
    );
    if (!ok) return;
    const student = resultState.students.find(
      (s) => String(s.studentId ?? "").trim() === activeProfileStudentId
    );
    if (!student) return;
    try {
      await upsertMany(STORES.students, [{ ...student, isDeleted: true }]);
      closeStudentProfile();
      await refreshResults();
      await refreshStats();
      renderDeletedStudentsList();
    } catch (e) {
      console.error(e);
      alert(`Failed to delete student: ${e.message ?? e}`);
    }
  });
}

if (elStudentListSearch) {
  elStudentListSearch.addEventListener("input", () => {
    studentListPaging.page = 1;
    activeStudentInlineId = "";
    renderStudentList();
  });
}

if (elStudentListIntake) {
  elStudentListIntake.addEventListener("change", () => {
    studentListPaging.page = 1;
    activeStudentInlineId = "";
    renderStudentList();
  });
}

if (elStudentListStatus) {
  elStudentListStatus.addEventListener("change", () => {
    studentListPaging.page = 1;
    activeStudentInlineId = "";
    renderStudentList();
  });
}

if (btnStudentListPrev) {
  btnStudentListPrev.addEventListener("click", () => {
    if (studentListPaging.page <= 1) return;
    studentListPaging.page -= 1;
    activeStudentInlineId = "";
    renderStudentList();
  });
}

if (btnStudentListNext) {
  btnStudentListNext.addEventListener("click", () => {
    studentListPaging.page += 1;
    activeStudentInlineId = "";
    renderStudentList();
  });
}

if (elStudentEditModal) {
  elStudentEditModal.addEventListener("click", (event) => {
    if (event.target === elStudentEditModal) closeStudentEdit();
  });
}


if (btnPrintSlip) {
  btnPrintSlip.addEventListener("click", () => {
    setPrintMode("result");
    window.print();
  });
}

if (btnCloseSlip) {
  btnCloseSlip.addEventListener("click", () => {
    closeSlip();
  });
}

if (btnCloseEnrollmentSlip) {
  btnCloseEnrollmentSlip.addEventListener("click", () => {
    closeEnrollmentSlip();
  });
}

if (btnPrintEnrollmentSlip) {
  btnPrintEnrollmentSlip.addEventListener("click", () => {
    setPrintMode("enrollment");
    window.print();
  });
}

if (btnCloseMarkEdit) {
  btnCloseMarkEdit.addEventListener("click", () => {
    closeMarkEdit();
  });
}

if (btnConfirmMarkEdit) {
  btnConfirmMarkEdit.addEventListener("click", () => {
    confirmMarkEdit();
  });
}

if (elMarkEditInput) {
  elMarkEditInput.addEventListener("input", () => {
    activeMarkEditValue = elMarkEditInput.value;
    updateMarkEditPreview();
  });
}

if (elMarkEditVerify) {
  elMarkEditVerify.addEventListener("change", () => {
    updateMarkEditPreview();
  });
}

if (elMarkEditModal) {
  elMarkEditModal.addEventListener("click", (event) => {
    if (event.target === elMarkEditModal) closeMarkEdit();
  });
}

if (btnCloseAddResult) {
  btnCloseAddResult.addEventListener("click", () => {
    closeAddResultModal();
  });
}

if (btnConfirmAddResult) {
  btnConfirmAddResult.addEventListener("click", async () => {
    if (!validateAddResultInputs()) return;
    const studentId = String(activeAddResultStudentId ?? "").trim();
    const courseCode = normalizeCourseCode(elAddResultCourse?.value ?? "");
    const sessionParsed = parseYearMonth(elAddResultSession?.value ?? "");
    const semesterValue = Math.floor(Number(elAddResultSemester?.value ?? ""));
    if (!studentId || !courseCode || !sessionParsed || !Number.isFinite(semesterValue)) return;

    const course = buildCourseMap().get(courseCode);
    const markRaw = String(elAddResultMark?.value ?? "").trim();
    let mark = null;
    let letter = "";
    let point = null;
    let passedCourse = null;
    if (markRaw) {
      mark = Number(markRaw);
      const grade = getGradeForMark(mark, course?.isMPUCourse === true);
      letter = grade.letter;
      point = grade.point;
      passedCourse = grade.passed;
    }

    const resultId = `${studentId}|${courseCode}|${sessionParsed.intake}|${semesterValue}`;
    const student = buildStudentMap().get(studentId);
    const result = {
      resultId,
      studentId,
      courseCode,
      session: sessionParsed.intake,
      semester: semesterValue,
      mark,
      letter,
      point,
      gradePoints: null,
      passedCourse,
      creditsEarned: null,
      studentName: String(student?.name ?? "").trim(),
    };

    try {
      await upsertMany(STORES.results, [result]);
      closeAddResultModal();
      activeResultInlineId = studentId;
      await refreshResults();
      await refreshStats();
    } catch (e) {
      if (elAddResultWarning) {
        elAddResultWarning.textContent = `Failed to save result: ${e.message ?? e}`;
        elAddResultWarning.style.display = "block";
      }
    }
  });
}

if (elAddResultCourse) {
  elAddResultCourse.addEventListener("change", () => {
    updateAddResultComputed();
  });
}

if (elAddResultSession) {
  elAddResultSession.addEventListener("input", () => {
    validateAddResultInputs();
  });
}

if (elAddResultSemester) {
  elAddResultSemester.addEventListener("input", () => {
    validateAddResultInputs();
  });
}

if (elAddResultMark) {
  elAddResultMark.addEventListener("input", () => {
    updateAddResultComputed();
  });
}

if (elAddResultVerify) {
  elAddResultVerify.addEventListener("change", () => {
    validateAddResultInputs();
  });
}

if (elAddResultModal) {
  elAddResultModal.addEventListener("click", (event) => {
    if (event.target === elAddResultModal) closeAddResultModal();
  });
}

if (elResultSlipModal) {
  elResultSlipModal.addEventListener("click", (event) => {
    if (event.target === elResultSlipModal) {
      closeSlip();
    }
  });
}

if (elEnrollmentSlipModal) {
  elEnrollmentSlipModal.addEventListener("click", (event) => {
    if (event.target === elEnrollmentSlipModal) {
      closeEnrollmentSlip();
    }
  });
}

if (btnCloseCoverage) {
  btnCloseCoverage.addEventListener("click", () => {
    closeCourseCoverageDetail();
  });
}

if (btnExportCoverage) {
  btnExportCoverage.addEventListener("click", () => {
    exportCoverageCSV();
  });
}

if (elCourseCoverageModal) {
  elCourseCoverageModal.addEventListener("click", (event) => {
    if (event.target === elCourseCoverageModal) {
      closeCourseCoverageDetail();
    }
  });
}

const openIssuesTab = (type) => {
  if (elIssueTypeSelect) elIssueTypeSelect.value = type;
  issuePaging.page = 1;
  if (elIssueSearch) elIssueSearch.value = "";
  setActiveTab("results");
  setActiveResultSubtab("issues");
  renderIssueList(type);
};

if (btnViewFailures) {
  btnViewFailures.addEventListener("click", () => openIssuesTab("failures"));
}
if (btnViewIncomplete) {
  btnViewIncomplete.addEventListener("click", () => openIssuesTab("incomplete"));
}
if (btnViewLowCgpa) {
  btnViewLowCgpa.addEventListener("click", () => openIssuesTab("lowCgpa"));
}
if (btnViewResultIssues) {
  btnViewResultIssues.addEventListener("click", () => openIssuesTab("resultIssues"));
}

if (elIssueTypeSelect) {
  elIssueTypeSelect.addEventListener("change", () => {
    issuePaging.page = 1;
    renderIssueList(elIssueTypeSelect.value);
  });
}

if (elIssueSearch) {
  elIssueSearch.addEventListener("input", () => {
    issuePaging.page = 1;
    renderIssueList(elIssueTypeSelect?.value ?? "failures");
  });
}

if (btnIssuePrev) {
  btnIssuePrev.addEventListener("click", () => {
    if (issuePaging.page > 1) {
      issuePaging.page -= 1;
      renderIssueList(elIssueTypeSelect?.value ?? "failures");
    }
  });
}

if (btnIssueNext) {
  btnIssueNext.addEventListener("click", () => {
    issuePaging.page += 1;
    renderIssueList(elIssueTypeSelect?.value ?? "failures");
  });
}

if (btnPrevPage) {
  btnPrevPage.addEventListener("click", () => {
    if (resultPaging.page > 1) {
      resultPaging.page -= 1;
      renderResultsView();
    }
  });
}

if (btnNextPage) {
  btnNextPage.addEventListener("click", () => {
    resultPaging.page += 1;
    renderResultsView();
  });
}

if (elCourseForm) {
  elCourseForm.addEventListener("submit", async (event) => {
    event.preventDefault();
    const courseCode = normalizeCourseCode(elCourseCode.value);
    try {
      if (activeCourseEditCode && activeCourseEditCode !== courseCode) {
        const detail = courseCoverageDetails.get(activeCourseEditCode);
        if (detail && detail.taken.length > 0) {
          throw new Error("Cannot change course code when results exist. Create a new course instead.");
        }
      }
      await upsertCourse({
        courseCode,
        title: elCourseTitle.value,
        credits: elCourseCredits.value,
        isMPUCourse: elCourseMPU.value,
      });
      console.log(`Course saved: ${courseCode}`);
      resetCourseForm();
      await refreshCourses();
      await refreshStats();
      await refreshResults();
    } catch (e) {
      console.error(e);
    }
  });
}

if (btnCourseReset) {
  btnCourseReset.addEventListener("click", () => {
    resetCourseForm();
  });
}

if (btnConfirmResultSlips) {
  btnConfirmResultSlips.addEventListener("click", () => {
    localStorage.setItem(CONFIRM_RESULT_KEY, new Date().toISOString());
    updateSlipConfirmStatus();
    alert("Result slips confirmed.");
  });
}

if (btnConfirmEnrollmentSlips) {
  btnConfirmEnrollmentSlips.addEventListener("click", () => {
    if (!enrollmentState.offeredCourseCodes.length) {
      buildEnrollmentSlips({ suppressMinAlert: true });
    }
    if (!enrollmentState.offeredCourseCodes.length) {
      alert("Select offered courses in the Enrollment tab before confirming.");
      return;
    }
    buildEnrollmentSlips({ suppressMinAlert: true });
    const lockedIds = new Set();
    for (const slip of enrollmentState.slips ?? []) {
      const studentId = String(slip?.studentId ?? "").trim();
      if (!studentId) continue;
      lockedIds.add(studentId);
    }
    if (lockedIds.size) {
      for (const studentId of lockedIds) {
        setSlipLock(studentId, true, { source: "confirm" }, { save: false });
      }
      saveEnrollmentSlipLocks();
      captureLockedSlipSnapshots([...lockedIds]);
    }
    localStorage.setItem(CONFIRM_ENROLL_KEY, new Date().toISOString());
    localStorage.setItem(ENROLLMENT_CONFIRM_TOKEN_KEY, computeEnrollmentConfirmToken());
    updateSlipConfirmStatus();
    renderEnrollmentList();
    renderEnrollmentSlip();
    alert("Enrollment slips confirmed and locked.");
  });
}

Promise.resolve().then(async () => {
  try {
    initImports({
      onDataChanged: async () => {
        await refreshStats();
        await refreshCourses();
        await refreshResults();
      },
    });
  } catch (e) {
    console.error(e);
  }
  try {
    initStudentImport({
      onDataChanged: async () => {
        await refreshStats();
        await refreshCourses();
        await refreshResults();
      },
    });
  } catch (e) {
    console.error(e);
  }
  try {
    initBackup({
      onDataChanged: async () => {
        await refreshStats();
        await refreshCourses();
        await refreshResults();
      },
    });
  } catch (e) {
    console.error(e);
  }
  try {
    initPortalExport();
  } catch (e) {
    console.error(e);
  }
  await refreshStats();
  await refreshCourses();
  await refreshResults();
  updateSlipConfirmStatus();
  setActiveResultSubtab("listing");
});
