// app/modules/imports-ui.js
import { importCSVFile, previewNewResults, uploadNewResults } from "./results.js";
import { getAllRecords, getCounts, STORES } from "../db/db.js";

function appendLog(el, msg) {
  if (!el) return;
  el.textContent += `${msg}\n`;
}

function escapeCsv(value) {
  const s = String(value ?? "");
  if (/[",\n]/.test(s)) return `"${s.replace(/"/g, '""')}"`;
  return s;
}

function downloadCsv(fileName, rows) {
  const csv = rows.map((row) => row.map(escapeCsv).join(",")).join("\n");
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

export function initImports({ onDataChanged }) {
  const elFileImports = document.getElementById("csvFileImports");
  const btnImportImports = document.getElementById("btnImportImports");
  const elLogImports = document.getElementById("logImports");

  const elNewResultYear = document.getElementById("newResultYear");
  const elNewResultMonth = document.getElementById("newResultMonth");
  const elNewResultCourse = document.getElementById("newResultCourse");
  const elNewResultFile = document.getElementById("newResultFile");
  const btnUploadNewResults = document.getElementById("btnUploadNewResults");
  const elNewResultLog = document.getElementById("newResultLog");
  const btnPreviewResults = document.getElementById("btnPreviewResults");
  const btnDownloadResultIssues = document.getElementById("btnDownloadResultIssues");
  const elPreviewModal = document.getElementById("resultPreviewModal");
  const elPreviewBody = document.getElementById("resultPreviewBody");
  const elPreviewTable = document.getElementById("resultPreviewTable");
  const elPreviewEmpty = document.getElementById("resultPreviewEmpty");
  const elPreviewCount = document.getElementById("resultPreviewCount");
  const elResultPreviewSimilarWrap = document.getElementById("resultPreviewSimilarWrap");
  const elResultPreviewSimilarCount = document.getElementById("resultPreviewSimilarCount");
  const elResultPreviewSimilarBody = document.getElementById("resultPreviewSimilarBody");
  const btnSelectAllPreview = document.getElementById("btnSelectAllPreview");
  const btnSelectNonePreview = document.getElementById("btnSelectNonePreview");
  const btnConfirmPreview = document.getElementById("btnConfirmPreview");
  const btnClosePreview = document.getElementById("btnClosePreview");
  let previewCache = null;
  const RESULT_NAME_SIMILARITY_THRESHOLD = 0.75;
  const RESULT_NAME_SKIP_TOKENS = new Set(["bin", "binti", "bt", "bte", "binte"]);

  const normalizeNameForCompare = (value) => {
    const normalized = String(value ?? "")
      .trim()
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, " ")
      .replace(/\s+/g, " ")
      .trim();
    if (!normalized) return "";
    return normalized
      .split(" ")
      .filter((token) => token && !RESULT_NAME_SKIP_TOKENS.has(token))
      .join(" ")
      .trim();
  };

  const levenshteinDistance = (a, b) => {
    if (a === b) return 0;
    const aLen = a.length;
    const bLen = b.length;
    if (aLen === 0) return bLen;
    if (bLen === 0) return aLen;
    const prev = Array.from({ length: bLen + 1 }, (_, i) => i);
    const curr = new Array(bLen + 1).fill(0);
    for (let i = 1; i <= aLen; i += 1) {
      curr[0] = i;
      const aChar = a.charAt(i - 1);
      for (let j = 1; j <= bLen; j += 1) {
        const cost = aChar === b.charAt(j - 1) ? 0 : 1;
        curr[j] = Math.min(
          prev[j] + 1,
          curr[j - 1] + 1,
          prev[j - 1] + cost
        );
      }
      for (let j = 0; j <= bLen; j += 1) {
        prev[j] = curr[j];
      }
    }
    return prev[bLen];
  };

  const nameSimilarity = (a, b) => {
    const left = normalizeNameForCompare(a);
    const right = normalizeNameForCompare(b);
    if (!left || !right) return 0;
    const maxLen = Math.max(left.length, right.length);
    if (!maxLen) return 0;
    const distance = levenshteinDistance(left, right);
    return 1 - distance / maxLen;
  };

  const renderSimilarMismatches = (preview) => {
    if (!elResultPreviewSimilarWrap || !elResultPreviewSimilarBody || !elResultPreviewSimilarCount) return;
    const mismatches = preview?.mismatchedExisting ?? [];
    const similarRows = [];
    for (const row of mismatches) {
      const incoming = row.incomingName ?? "";
      const existing = row.existingName ?? "";
      const score = nameSimilarity(incoming, existing);
      if (score >= RESULT_NAME_SIMILARITY_THRESHOLD) {
        similarRows.push({
          rowNo: row.rowNo ?? "",
          studentId: row.studentId ?? "",
          incoming,
          existing,
          score,
        });
      }
    }

    if (!similarRows.length) {
      elResultPreviewSimilarWrap.style.display = "none";
      elResultPreviewSimilarBody.textContent = "";
      elResultPreviewSimilarCount.textContent = "0";
      return;
    }

    elResultPreviewSimilarWrap.style.display = "block";
    elResultPreviewSimilarBody.textContent = "";
    elResultPreviewSimilarCount.textContent = String(similarRows.length);
    for (const row of similarRows) {
      const tr = document.createElement("tr");

      const tdRow = document.createElement("td");
      tdRow.textContent = String(row.rowNo ?? "-");
      tr.appendChild(tdRow);

      const tdId = document.createElement("td");
      tdId.textContent = row.studentId || "-";
      tr.appendChild(tdId);

      const tdIncoming = document.createElement("td");
      tdIncoming.textContent = row.incoming || "-";
      tr.appendChild(tdIncoming);

      const tdExisting = document.createElement("td");
      tdExisting.textContent = row.existing || "-";
      tr.appendChild(tdExisting);

      const tdScore = document.createElement("td");
      tdScore.textContent = `${Math.round(row.score * 100)}%`;
      tr.appendChild(tdScore);

      elResultPreviewSimilarBody.appendChild(tr);
    }
  };
  const updateResultIssueButton = () => {
    if (!btnDownloadResultIssues) return;
    const issueRows = previewCache?.preview?.issueRows ?? [];
    btnDownloadResultIssues.disabled = issueRows.length === 0;
  };

  const loadCourses = async () => {
    if (!elNewResultCourse) return;
    const courses = await getAllRecords(STORES.courses);
    const sorted = courses
      .filter((course) => course?.courseCode)
      .sort((a, b) => String(a.courseCode).localeCompare(String(b.courseCode)));
    elNewResultCourse.textContent = "";
    const defaultOption = document.createElement("option");
    defaultOption.value = "";
    defaultOption.textContent = "Course";
    elNewResultCourse.appendChild(defaultOption);
    for (const course of sorted) {
      const option = document.createElement("option");
      option.value = String(course.courseCode ?? "");
      option.textContent = course.title
        ? `${course.courseCode} - ${course.title}`
        : String(course.courseCode ?? "");
      elNewResultCourse.appendChild(option);
    }
  };

  if (btnImportImports) {
    btnImportImports.addEventListener("click", async () => {
      if (elLogImports) elLogImports.textContent = "";
      try {
        const file = elFileImports?.files?.[0];
        await importCSVFile(file, (msg) => appendLog(elLogImports, msg));
        const counts = await getCounts();
        appendLog(
          elLogImports,
          `DB counts -> students: ${counts.students}, courses: ${counts.courses}, results: ${counts.results}`
        );
        if (onDataChanged) await onDataChanged();
      } catch (e) {
        appendLog(elLogImports, `ERROR: ${e.message ?? e}`);
      }
    });
  }

  const openPreview = async () => {
      if (elNewResultLog) elNewResultLog.textContent = "";
      try {
        const year = String(elNewResultYear?.value ?? "").trim();
        const month = String(elNewResultMonth?.value ?? "").trim();
        const courseCode = String(elNewResultCourse?.value ?? "").trim();
        const session = year && month ? `${year}/${month}` : "";
        const file = elNewResultFile?.files?.[0];

        const preview = await previewNewResults({ file, session, courseCode });
        previewCache = { preview, session, courseCode, file };
        updateResultIssueButton();
        renderSimilarMismatches(preview);

        if (elPreviewCount) {
          elPreviewCount.textContent = String(preview.affectedStudents.length);
        }
        if (preview.missingNameNewStudents) {
          appendLog(
            elNewResultLog,
            `WARNING: ${preview.missingNameNewStudents} new student(s) have no name in this file. Upload a student list CSV to fill names.`
          );
        }
        if (preview.skippedMismatch) {
          appendLog(
            elNewResultLog,
            `WARNING: ${preview.skippedMismatch} row(s) were blocked because ID+Name did not match existing student records.`
          );
          const mismatches = preview.mismatchedExisting ?? [];
          mismatches.slice(0, 10).forEach((row) => {
            appendLog(
              elNewResultLog,
              `- Row ${row.rowNo}: ${row.studentId} | existing="${row.existingName}" vs uploaded="${row.incomingName}"`
            );
          });
        }
        if (preview.issueRows?.length) {
          appendLog(elNewResultLog, `Issue rows detected: ${preview.issueRows.length}`);
          if (preview.issueSummary?.missing_mark) {
            appendLog(elNewResultLog, `- Missing mark: ${preview.issueSummary.missing_mark}`);
          }
          if (preview.issueSummary?.missing_identifier) {
            appendLog(elNewResultLog, `- Missing ID + name: ${preview.issueSummary.missing_identifier}`);
          }
          if (preview.issueSummary?.name_not_found) {
            appendLog(elNewResultLog, `- Name not found: ${preview.issueSummary.name_not_found}`);
          }
          if (preview.issueSummary?.name_ambiguous) {
            appendLog(elNewResultLog, `- Name matches multiple students: ${preview.issueSummary.name_ambiguous}`);
          }
          if (preview.issueSummary?.name_mismatch) {
            appendLog(elNewResultLog, `- Name mismatch with ID: ${preview.issueSummary.name_mismatch}`);
          }
          appendLog(elNewResultLog, "Use Download Issues Report for details.");
        }

        if (!preview.affectedStudents.length) {
          elPreviewBody.textContent = "";
          elPreviewEmpty.style.display = "block";
          if (elPreviewTable) elPreviewTable.style.display = "none";
        } else {
          elPreviewEmpty.style.display = "none";
          if (elPreviewTable) elPreviewTable.style.display = "table";
          elPreviewBody.textContent = "";
          for (const student of preview.affectedStudents) {
            const tr = document.createElement("tr");
            const tdCheck = document.createElement("td");
            const checkbox = document.createElement("input");
            checkbox.type = "checkbox";
            checkbox.checked = true;
            checkbox.dataset.studentId = student.studentId;
            tdCheck.appendChild(checkbox);
            tr.appendChild(tdCheck);

            const tdId = document.createElement("td");
            tdId.textContent = student.studentId;
            tr.appendChild(tdId);

            const tdName = document.createElement("td");
            tdName.textContent = student.name || "-";
            tr.appendChild(tdName);

            const tdNew = document.createElement("td");
            tdNew.textContent = student.isNew ? "New" : "Existing";
            tr.appendChild(tdNew);

            const tdCount = document.createElement("td");
            tdCount.textContent = String(student.count);
            tr.appendChild(tdCount);

            elPreviewBody.appendChild(tr);
          }
        }

        if (elPreviewModal) {
          elPreviewModal.classList.add("active");
          elPreviewModal.setAttribute("aria-hidden", "false");
        }
      } catch (e) {
        appendLog(elNewResultLog, `ERROR: ${e.message ?? e}`);
        renderSimilarMismatches({ mismatchedExisting: [] });
      }
  };

  if (btnPreviewResults) {
    btnPreviewResults.addEventListener("click", async () => {
      await openPreview();
    });
  }

  if (btnDownloadResultIssues) {
    btnDownloadResultIssues.addEventListener("click", () => {
      const issueRows = previewCache?.preview?.issueRows ?? [];
      if (!issueRows.length) {
        appendLog(elNewResultLog, "No issue rows to export. Run preview first.");
        return;
      }
      const header = ["Row", "StudentID", "Name", "Mark", "Reason"];
      const rows = issueRows.map((row) => [
        row.rowNo ?? "",
        row.studentId ?? "",
        row.name ?? "",
        row.mark ?? "",
        row.reason ?? "",
      ]);
      const stamp = new Date().toISOString().slice(0, 10);
      downloadCsv(`result-issues-${stamp}.csv`, [header, ...rows]);
      appendLog(elNewResultLog, `Issues report downloaded (${rows.length} row(s)).`);
    });
  }

  if (btnUploadNewResults) {
    btnUploadNewResults.addEventListener("click", async () => {
      if (elNewResultLog) elNewResultLog.textContent = "";
      try {
        const year = String(elNewResultYear?.value ?? "").trim();
        const month = String(elNewResultMonth?.value ?? "").trim();
        const courseCode = String(elNewResultCourse?.value ?? "").trim();
        const session = year && month ? `${year}/${month}` : "";
        const file = elNewResultFile?.files?.[0];
        if (previewCache?.session !== session || previewCache?.courseCode !== courseCode || previewCache?.file !== file) {
          previewCache = null;
        }
        if (!previewCache) {
          await openPreview();
          return;
        }
        appendLog(elNewResultLog, "Use Confirm Update in the preview window.");
      } catch (e) {
        appendLog(elNewResultLog, `ERROR: ${e.message ?? e}`);
      }
    });
  }

  if (btnSelectAllPreview) {
    btnSelectAllPreview.addEventListener("click", () => {
      elPreviewBody?.querySelectorAll("input[type='checkbox']").forEach((cb) => {
        cb.checked = true;
      });
    });
  }

  if (btnSelectNonePreview) {
    btnSelectNonePreview.addEventListener("click", () => {
      elPreviewBody?.querySelectorAll("input[type='checkbox']").forEach((cb) => {
        cb.checked = false;
      });
    });
  }

  if (btnConfirmPreview) {
    btnConfirmPreview.addEventListener("click", async () => {
      if (!previewCache) return;
      const selected = [];
      elPreviewBody?.querySelectorAll("input[type='checkbox']").forEach((cb) => {
        if (cb.checked) selected.push(cb.dataset.studentId);
      });
      try {
        if (!selected.length) {
          appendLog(elNewResultLog, "No students selected.");
          return;
        }
        const result = await uploadNewResults({
          session: previewCache.session,
          file: previewCache.file,
          courseCode: previewCache.courseCode,
          selectedStudentIds: selected,
          logFn: (msg) => appendLog(elNewResultLog, msg),
        });
        if (onDataChanged) await onDataChanged();
        previewCache = null;
        updateResultIssueButton();
        if (elPreviewModal) {
          elPreviewModal.classList.remove("active");
          elPreviewModal.setAttribute("aria-hidden", "true");
        }
      } catch (e) {
        appendLog(elNewResultLog, `ERROR: ${e.message ?? e}`);
      }
    });
  }

  if (btnClosePreview) {
    btnClosePreview.addEventListener("click", () => {
      if (elPreviewModal) {
        elPreviewModal.classList.remove("active");
        elPreviewModal.setAttribute("aria-hidden", "true");
      }
      previewCache = null;
      updateResultIssueButton();
      renderSimilarMismatches({ mismatchedExisting: [] });
    });
  }

  if (elPreviewModal) {
    elPreviewModal.addEventListener("click", (event) => {
      if (event.target === elPreviewModal) {
        elPreviewModal.classList.remove("active");
        elPreviewModal.setAttribute("aria-hidden", "true");
      }
    });
  }

  loadCourses().catch((e) => appendLog(elNewResultLog, `ERROR: ${e.message ?? e}`));
}
