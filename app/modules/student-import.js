// app/modules/student-import.js
import { getAllRecords, upsertMany, STORES } from "../db/db.js";

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

function normalizeKey(key) {
  return String(key ?? "")
    .toLowerCase()
    .replace(/[\s_-]+/g, "");
}

function getRowValue(row, keys) {
  for (const key of keys) {
    if (key in row) return row[key];
  }
  return "";
}

function normalizeNameKey(value) {
  return String(value ?? "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
}

function parseIntake(value) {
  const raw = String(value ?? "").trim();
  if (!raw) return null;
  const match = raw.match(/^(\d{4})[./-](\d{1,2})$/);
  if (!match) return null;
  const year = match[1];
  const month = match[2].padStart(2, "0");
  if (!/^(02|09)$/.test(month)) return null;
  return { intake: `${year}/${month}`, year, month };
}

function normalizeStudentRow(rawRow) {
  const row = {};
  for (const key of Object.keys(rawRow)) {
    row[normalizeKey(key)] = rawRow[key];
  }

  const studentId = String(
    getRowValue(row, ["studentid", "id", "studentno", "matric", "matricno", "student_id"])
  ).trim();
  const name = String(getRowValue(row, ["name", "fullname", "studentname"])).trim();
  const intakeRaw = getRowValue(row, ["intake", "intakecode", "intakesession"]);
  const intakeParsed = parseIntake(intakeRaw);

  if (!studentId || !name || !intakeParsed) return null;

  return {
    studentId,
    name,
    intake: intakeParsed.intake,
    intakeYear: intakeParsed.year,
    intakeMonth: intakeParsed.month,
  };
}

function toDisplayStatus(row) {
  if (row.statusType === "new") return "New";
  if (row.statusType === "duplicate_id_existing") return "Duplicate ID (existing)";
  if (row.statusType === "duplicate_id_csv") return "Duplicate ID (CSV)";
  if (row.statusType === "duplicate_name_existing") {
    return `Duplicate Name (existing ID: ${row.existingStudentId})`;
  }
  if (row.statusType === "duplicate_name_csv") return "Duplicate Name (CSV)";
  if (row.statusType === "missing_required") return "Missing required fields";
  return "Unknown";
}

function buildStudentImportPreview({
  parsedRows,
  existingStudents,
}) {
  const existingById = new Map();
  const existingByName = new Map();
  for (const student of existingStudents) {
    const existingId = String(student?.studentId ?? "").trim();
    if (existingId) existingById.set(existingId, student);
    const nameKey = normalizeNameKey(student?.name);
    if (nameKey && !existingByName.has(nameKey)) {
      existingByName.set(nameKey, student);
    }
  }

  const seenIds = new Set();
  const seenNames = new Map();
  const rows = [];
  const summary = {
    totalRows: parsedRows.length,
    newCount: 0,
    duplicateIdExisting: 0,
    duplicateIdCsv: 0,
    duplicateNameExisting: 0,
    duplicateNameCsv: 0,
    missingRequired: 0,
  };

  parsedRows.forEach((rawRow, index) => {
    const rowNo = index + 2;
    const normalizedRaw = {};
    for (const key of Object.keys(rawRow ?? {})) {
      normalizedRaw[normalizeKey(key)] = rawRow[key];
    }
    const draftId = String(
      getRowValue(normalizedRaw, ["studentid", "id", "studentno", "matric", "matricno", "student_id"])
    ).trim();
    const draftName = String(getRowValue(normalizedRaw, ["name", "fullname", "studentname"])).trim();
    const draftIntake = String(
      getRowValue(normalizedRaw, ["intake", "intakecode", "intakesession"])
    ).trim();

    const candidate = normalizeStudentRow(rawRow);
    if (!candidate) {
      summary.missingRequired += 1;
      const missing = [];
      if (!draftId) missing.push("Student ID");
      if (!draftName) missing.push("Name");
      if (!draftIntake) {
        missing.push("Intake");
      } else if (!parseIntake(draftIntake)) {
        missing.push("Intake format");
      }
      rows.push({
        rowNo,
        studentId: draftId,
        name: draftName,
        intake: draftIntake,
        reason: missing.length ? missing.join(", ") : "Missing required fields",
        statusType: "missing_required",
        selectable: false,
        defaultChecked: false,
      });
      return;
    }

    const studentId = String(candidate.studentId ?? "").trim();
    const nameKey = normalizeNameKey(candidate.name);

    if (seenIds.has(studentId)) {
      summary.duplicateIdCsv += 1;
      rows.push({
        rowNo,
        studentId: candidate.studentId,
        name: candidate.name,
        intake: candidate.intake,
        reason: "Duplicate ID (CSV)",
        statusType: "duplicate_id_csv",
        selectable: false,
        defaultChecked: false,
        candidate,
      });
      return;
    }
    if (existingById.has(studentId)) {
      const existingStudent = existingById.get(studentId);
      const existingNameKey = normalizeNameKey(existingStudent?.name);
      const existingIntakeParsed = parseIntake(existingStudent?.intake);
      const existingIntake = existingIntakeParsed
        ? existingIntakeParsed.intake
        : String(existingStudent?.intake ?? "").trim();
      const intakeDifferent = Boolean(existingIntake && candidate.intake && existingIntake !== candidate.intake);
      const nameMatches = Boolean(nameKey && existingNameKey && nameKey === existingNameKey);
      const allowUpdate = nameMatches && intakeDifferent;
      summary.duplicateIdExisting += 1;
      rows.push({
        rowNo,
        studentId: candidate.studentId,
        name: candidate.name,
        intake: candidate.intake,
        reason: allowUpdate ? "Duplicate ID (existing) - intake update" : "Duplicate ID (existing)",
        statusType: "duplicate_id_existing",
        selectable: allowUpdate,
        defaultChecked: allowUpdate,
        allowUpdate,
        candidate,
      });
      seenIds.add(studentId);
      if (nameKey) seenNames.set(nameKey, studentId);
      return;
    }

    const seenNameId = seenNames.get(nameKey);
    if (seenNameId && seenNameId !== studentId) {
      summary.duplicateNameCsv += 1;
      rows.push({
        rowNo,
        studentId: candidate.studentId,
        name: candidate.name,
        intake: candidate.intake,
        reason: "Duplicate Name (CSV)",
        statusType: "duplicate_name_csv",
        selectable: false,
        defaultChecked: false,
        candidate,
      });
      return;
    }

    const matchedByName = nameKey ? existingByName.get(nameKey) : null;
    if (matchedByName) {
      const existingId = String(matchedByName.studentId ?? "").trim();
      if (existingId && existingId !== studentId) {
        summary.duplicateNameExisting += 1;
        rows.push({
          rowNo,
          studentId: candidate.studentId,
          name: candidate.name,
          intake: candidate.intake,
          reason: `Duplicate Name (existing ID: ${existingId})`,
          statusType: "duplicate_name_existing",
          selectable: true,
          defaultChecked: false,
          actionDefault: "skip",
          existingStudentId: existingId,
          candidate,
        });
        seenIds.add(studentId);
        if (nameKey) seenNames.set(nameKey, studentId);
        return;
      }
    }

    summary.newCount += 1;
    rows.push({
      rowNo,
      studentId: candidate.studentId,
      name: candidate.name,
      intake: candidate.intake,
      statusType: "new",
      selectable: true,
      defaultChecked: true,
      actionDefault: "import",
      candidate,
    });
    seenIds.add(studentId);
    if (nameKey) seenNames.set(nameKey, studentId);
  });

  return { rows, summary, existingById };
}

export function initStudentImport({ onDataChanged } = {}) {
  const elFile = document.getElementById("studentCsvFile");
  const btnImport = document.getElementById("btnImportStudents");
  const btnDownloadMissing = document.getElementById("btnDownloadMissingStudents");
  const elLog = document.getElementById("studentImportLog");
  const elPreviewModal = document.getElementById("studentImportPreviewModal");
  const elPreviewBody = document.getElementById("studentImportPreviewBody");
  const elPreviewTable = document.getElementById("studentImportPreviewTable");
  const elPreviewEmpty = document.getElementById("studentImportPreviewEmpty");
  const elPreviewCount = document.getElementById("studentImportPreviewCount");
  const btnClosePreview = document.getElementById("btnCloseStudentImportPreview");
  const btnSelectAllPreview = document.getElementById("btnSelectAllStudentImportPreview");
  const btnSelectNonePreview = document.getElementById("btnSelectNoneStudentImportPreview");
  const btnConfirmPreview = document.getElementById("btnConfirmStudentImportPreview");

  let previewCache = null;
  const reportIssueTypes = new Set([
    "missing_required",
    "duplicate_id_existing",
    "duplicate_id_csv",
    "duplicate_name_existing",
    "duplicate_name_csv",
  ]);
  const updateMissingReportButton = () => {
    if (!btnDownloadMissing) return;
    const issueRows = previewCache?.preview?.rows?.filter(
      (row) => reportIssueTypes.has(row.statusType)
    );
    btnDownloadMissing.disabled = !issueRows || issueRows.length === 0;
  };

  const closePreview = () => {
    if (!elPreviewModal) return;
    elPreviewModal.classList.remove("active");
    elPreviewModal.setAttribute("aria-hidden", "true");
  };

  const openPreview = () => {
    if (!elPreviewModal) return;
    elPreviewModal.classList.add("active");
    elPreviewModal.setAttribute("aria-hidden", "false");
  };

  const renderPreview = (preview) => {
    if (!elPreviewBody || !elPreviewTable || !elPreviewEmpty || !elPreviewCount) return;
    const rows = preview?.rows ?? [];
    const actionableCount = rows.filter((row) => row.selectable).length;
    elPreviewCount.textContent = String(actionableCount);
    elPreviewBody.textContent = "";

    if (!rows.length) {
      elPreviewEmpty.style.display = "block";
      elPreviewTable.style.display = "none";
      return;
    }

    const priorityRank = (row) => {
      if (row.statusType === "new") return 0;
      if (row.statusType === "duplicate_id_existing" && row.allowUpdate) return 1;
      if (row.statusType === "duplicate_name_existing") return 2;
      return 3;
    };
    const orderedRows = rows
      .map((row, index) => ({ row, index }))
      .sort((a, b) => {
        const rankDiff = priorityRank(a.row) - priorityRank(b.row);
        if (rankDiff !== 0) return rankDiff;
        return a.index - b.index;
      })
      ;

    elPreviewEmpty.style.display = "none";
    elPreviewTable.style.display = "table";

    orderedRows.forEach(({ row, index: originalIndex }) => {
      const tr = document.createElement("tr");
      const tdCheck = document.createElement("td");
      let rowCheckbox = null;
      if (row.selectable) {
        rowCheckbox = document.createElement("input");
        rowCheckbox.type = "checkbox";
        rowCheckbox.checked = row.defaultChecked === true;
        rowCheckbox.dataset.rowIndex = String(originalIndex);
        tdCheck.appendChild(rowCheckbox);
      } else {
        tdCheck.textContent = "-";
      }
      tr.appendChild(tdCheck);

      const tdRow = document.createElement("td");
      tdRow.textContent = String(row.rowNo ?? "-");
      tr.appendChild(tdRow);

      const tdId = document.createElement("td");
      tdId.textContent = row.studentId || "-";
      tr.appendChild(tdId);

      const tdName = document.createElement("td");
      tdName.textContent = row.name || "-";
      tr.appendChild(tdName);

      const tdStatus = document.createElement("td");
      tdStatus.textContent = toDisplayStatus(row);
      tr.appendChild(tdStatus);

      const tdAction = document.createElement("td");
      if (row.statusType === "duplicate_name_existing") {
        const select = document.createElement("select");
        select.className = "preview-action-select";
        select.dataset.rowIndex = String(originalIndex);
        const optSkip = document.createElement("option");
        optSkip.value = "skip";
        optSkip.textContent = "Skip";
        const optUpdate = document.createElement("option");
        optUpdate.value = "update";
        optUpdate.textContent = `Update existing (${row.existingStudentId})`;
        select.appendChild(optSkip);
        select.appendChild(optUpdate);
        select.value = row.actionDefault ?? "skip";
        select.addEventListener("change", () => {
          if (select.value === "update" && rowCheckbox) rowCheckbox.checked = true;
        });
        tdAction.appendChild(select);
      } else if (row.statusType === "duplicate_id_existing" && row.allowUpdate) {
        tdAction.textContent = "Update intake";
      } else if (row.statusType === "new") {
        tdAction.textContent = "Import";
      } else {
        tdAction.textContent = "Skip";
      }
      tr.appendChild(tdAction);

      elPreviewBody.appendChild(tr);
    });
  };

  if (!btnImport) return;

  btnImport.addEventListener("click", async () => {
    if (elLog) elLog.textContent = "";
    try {
      const file = elFile?.files?.[0];
      if (!file) throw new Error("No CSV file selected.");

      appendLog(elLog, `Parsing: ${file.name} ...`);

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
        appendLog(elLog, "CSV parse warnings/errors:");
        for (const e of parsed.errors.slice(0, 10)) appendLog(elLog, `- ${e.message}`);
      }

      const existing = await getAllRecords(STORES.students);
      const preview = buildStudentImportPreview({
        parsedRows: parsed.data,
        existingStudents: existing,
      });
      previewCache = {
        fileName: file.name,
        preview,
      };
      updateMissingReportButton();
      appendLog(elLog, `Rows read: ${preview.summary.totalRows}`);
      appendLog(elLog, `New rows detected: ${preview.summary.newCount}`);
      appendLog(elLog, `Duplicate name (existing ID mismatch): ${preview.summary.duplicateNameExisting}`);
      if (preview.summary.missingRequired) {
        appendLog(elLog, `Missing required fields: ${preview.summary.missingRequired}`);
      }
      appendLog(elLog, "Review the preview popup and confirm which rows to import.");
      renderPreview(preview);
      openPreview();
    } catch (e) {
      appendLog(elLog, `ERROR: ${e.message ?? e}`);
    }
  });

  if (btnDownloadMissing) {
    btnDownloadMissing.addEventListener("click", () => {
      const issueRows = previewCache?.preview?.rows?.filter(
        (row) => reportIssueTypes.has(row.statusType)
      ) ?? [];
      if (!issueRows.length) {
        appendLog(elLog, "No missing/duplicate issues to export. Run import preview first.");
        return;
      }
      const header = ["Row", "StudentID", "Name", "Intake", "Reason"];
      const rows = issueRows.map((row) => [
        row.rowNo ?? "",
        row.studentId ?? "",
        row.name ?? "",
        row.intake ?? "",
        row.reason ?? toDisplayStatus(row),
      ]);
      const stamp = new Date().toISOString().slice(0, 10);
      downloadCsv(`missing-students-${stamp}.csv`, [header, ...rows]);
      appendLog(elLog, `Missing/duplicate report downloaded (${rows.length} row(s)).`);
    });
  }

  if (btnSelectAllPreview) {
    btnSelectAllPreview.addEventListener("click", () => {
      elPreviewBody?.querySelectorAll("input[type='checkbox']").forEach((cb) => {
        if (!cb.disabled) cb.checked = true;
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
      if (!previewCache?.preview?.rows?.length) return;
      try {
        const rows = previewCache.preview.rows;
        const existingById = previewCache.preview.existingById;
        const toInsert = [];
        const toUpdateById = new Map();
        let selectedRows = 0;
        let skippedConflicts = 0;

        for (let index = 0; index < rows.length; index += 1) {
          const row = rows[index];
          if (!row.selectable) continue;
          const checkbox = elPreviewBody?.querySelector(`input[type='checkbox'][data-row-index='${index}']`);
          if (!checkbox?.checked) continue;
          selectedRows += 1;

          if (row.statusType === "new") {
            if (row.candidate) toInsert.push(row.candidate);
            continue;
          }

          if (row.statusType === "duplicate_name_existing") {
            const actionEl = elPreviewBody?.querySelector(`select[data-row-index='${index}']`);
            const action = String(actionEl?.value ?? "skip");
            if (action !== "update") {
              skippedConflicts += 1;
              continue;
            }
            const existingStudent = existingById.get(String(row.existingStudentId ?? "").trim());
            if (!existingStudent || !row.candidate) {
              skippedConflicts += 1;
              continue;
            }
            toUpdateById.set(existingStudent.studentId, {
              ...existingStudent,
              name: row.candidate.name || existingStudent.name,
              intake: row.candidate.intake,
              intakeYear: row.candidate.intakeYear,
              intakeMonth: row.candidate.intakeMonth,
            });
            continue;
          }

          if (row.statusType === "duplicate_id_existing") {
            if (!row.allowUpdate || !row.candidate) {
              skippedConflicts += 1;
              continue;
            }
            const existingStudent = existingById.get(String(row.studentId ?? "").trim());
            if (!existingStudent) {
              skippedConflicts += 1;
              continue;
            }
            const existingNameKey = normalizeNameKey(existingStudent?.name);
            const incomingNameKey = normalizeNameKey(row.candidate?.name);
            if (!incomingNameKey || !existingNameKey || incomingNameKey !== existingNameKey) {
              skippedConflicts += 1;
              continue;
            }
            toUpdateById.set(existingStudent.studentId, {
              ...existingStudent,
              intake: row.candidate.intake,
              intakeYear: row.candidate.intakeYear,
              intakeMonth: row.candidate.intakeMonth,
            });
            continue;
          }
        }

        const toUpdate = [...toUpdateById.values()];
        if (!selectedRows) {
          appendLog(elLog, "No rows selected.");
          return;
        }

        if (toUpdate.length) {
          const ok = window.confirm(
            `Update ${toUpdate.length} existing student record(s) where name matched but CSV ID was different?`
          );
          if (!ok) {
            appendLog(elLog, "Student import update step cancelled.");
            return;
          }
        }

        if (toInsert.length) await upsertMany(STORES.students, toInsert);
        if (toUpdate.length) await upsertMany(STORES.students, toUpdate);

        appendLog(elLog, `Rows selected: ${selectedRows}`);
        appendLog(elLog, `Imported new students: ${toInsert.length}`);
        appendLog(elLog, `Updated existing students: ${toUpdate.length}`);
        appendLog(elLog, `Skipped selected conflicts: ${skippedConflicts}`);
        appendLog(
          elLog,
          `Auto-skipped duplicates/missing: ${
            previewCache.preview.summary.duplicateIdExisting
            + previewCache.preview.summary.duplicateIdCsv
            + previewCache.preview.summary.duplicateNameCsv
            + previewCache.preview.summary.missingRequired
          }`
        );

        if (!toInsert.length && !toUpdate.length) {
          appendLog(elLog, "No changes were applied.");
        } else {
          appendLog(elLog, "Student import complete.");
          if (onDataChanged) await onDataChanged();
        }

        previewCache = null;
        updateMissingReportButton();
        closePreview();
      } catch (e) {
        appendLog(elLog, `ERROR: ${e.message ?? e}`);
      }
    });
  }

  if (btnClosePreview) {
    btnClosePreview.addEventListener("click", () => {
      closePreview();
    });
  }

  if (elPreviewModal) {
    elPreviewModal.addEventListener("click", (event) => {
      if (event.target === elPreviewModal) {
        closePreview();
      }
    });
  }
}
