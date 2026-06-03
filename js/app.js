const excelInput = document.getElementById("excelFile");
const dlpInput = document.getElementById("dlpDate");
const intervalSelect = document.getElementById("intervalType");
const processBtn = document.getElementById("processBtn");

const statusSection = document.getElementById("statusSection");
const statusMessage = document.getElementById("statusMessage");
const dedupSection = document.getElementById("dedupSection");
const dedupStats = document.getElementById("dedupStats");
const dedupTableBody = document.querySelector("#dedupTable tbody");
const intervalSection = document.getElementById("intervalSection");
const intervalLabels = document.getElementById("intervalLabels");
const resultsSection = document.getElementById("resultsSection");
const resultsTableHead = document.querySelector("#resultsTable thead");
const resultsTableBody = document.querySelector("#resultsTable tbody");
const intervalTotals = document.getElementById("intervalTotals");

function updateProcessButtonState() {
  processBtn.disabled = !(excelInput.files.length && dlpInput.value);
}

excelInput.addEventListener("change", updateProcessButtonState);
dlpInput.addEventListener("change", updateProcessButtonState);
processBtn.addEventListener("click", handleProcess);

function showStatus(message, type = "info") {
  statusSection.classList.remove("hidden", "success", "warning", "error");
  statusSection.classList.add(type === "info" ? "" : type);
  statusMessage.textContent = message;
}

function formatDate(date) {
  const day = String(date.getDate()).padStart(2, "0");
  const month = date.toLocaleString("en-GB", { month: "short" });
  const year = date.getFullYear();
  return `${day}-${month}-${year}`;
}

function parseExcelDate(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  }

  if (typeof value === "number") {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed) {
      return new Date(parsed.y, parsed.m - 1, parsed.d);
    }
  }

  if (typeof value === "string" && value.trim()) {
    const parsed = new Date(value.trim());
    if (!Number.isNaN(parsed.getTime())) {
      return new Date(parsed.getFullYear(), parsed.getMonth(), parsed.getDate());
    }
  }

  return null;
}

function normalizeAdr(value) {
  return String(value ?? "").trim().toLowerCase();
}

function displayAdr(value) {
  return String(value ?? "").trim();
}

function deduplicateRows(rows) {
  const grouped = new Map();
  const duplicatesRemoved = [];

  for (const row of rows) {
    const key = `${row.caseId}||${normalizeAdr(row.adr)}`;

    if (!grouped.has(key)) {
      grouped.set(key, {
        caseId: row.caseId,
        adr: displayAdr(row.adr),
        irdDate: row.irdDate,
        removedCount: 0,
      });
      continue;
    }

    const existing = grouped.get(key);
    existing.removedCount += 1;

    if (row.irdDate && (!existing.irdDate || row.irdDate < existing.irdDate)) {
      existing.irdDate = row.irdDate;
    }

    duplicatesRemoved.push(row);
  }

  return {
    uniqueRows: Array.from(grouped.values()),
    duplicatesRemoved,
  };
}

function getQuarterIndex(date) {
  return Math.floor(date.getMonth() / 3);
}

function getQuarterStart(year, quarterIndex) {
  return new Date(year, quarterIndex * 3, 1);
}

function getQuarterEnd(year, quarterIndex) {
  return new Date(year, quarterIndex * 3 + 3, 0);
}

function getHalfYearIndex(date) {
  return date.getMonth() < 6 ? 0 : 1;
}

function getHalfYearStart(year, halfIndex) {
  return new Date(year, halfIndex === 0 ? 0 : 6, 1);
}

function getHalfYearEnd(year, halfIndex) {
  return new Date(year, halfIndex === 0 ? 6 : 12, 0);
}

function buildIntervals(dlpDate, intervalType) {
  const intervals = [];
  const dlp = new Date(dlpDate.getFullYear(), dlpDate.getMonth(), dlpDate.getDate());

  if (intervalType === "yearly") {
    const endYear = dlp.getFullYear();

    for (let i = 2; i >= 0; i -= 1) {
      const year = endYear - i;
      const start = new Date(year, 0, 1);
      let end = new Date(year, 11, 31);

      if (year === endYear && dlp < end) {
        end = dlp;
      }

      intervals.push({ label: `Year ${year}`, start, end });
    }

    return intervals;
  }

  if (intervalType === "quarterly") {
    let year = dlp.getFullYear();
    let quarter = getQuarterIndex(dlp);

    for (let i = 0; i < 3; i += 1) {
      const start = getQuarterStart(year, quarter);
      let end = getQuarterEnd(year, quarter);

      if (end > dlp) {
        end = dlp;
      }

      intervals.unshift({
        label: `Q${quarter + 1} ${year}`,
        start,
        end,
      });

      quarter -= 1;
      if (quarter < 0) {
        quarter = 3;
        year -= 1;
      }
    }

    return intervals;
  }

  if (intervalType === "biannual") {
    let year = dlp.getFullYear();
    let half = getHalfYearIndex(dlp);

    for (let i = 0; i < 3; i += 1) {
      const start = getHalfYearStart(year, half);
      let end = getHalfYearEnd(year, half);

      if (end > dlp) {
        end = dlp;
      }

      intervals.unshift({
        label: half === 0 ? `H1 ${year}` : `H2 ${year}`,
        start,
        end,
      });

      half -= 1;
      if (half < 0) {
        half = 1;
        year -= 1;
      }
    }

    return intervals;
  }

  return intervals;
}

function isDateInInterval(date, interval) {
  if (!date) {
    return false;
  }

  const time = date.getTime();
  return time >= interval.start.getTime() && time <= interval.end.getTime();
}

function countByInterval(uniqueRows, intervals) {
  const adrMap = new Map();

  for (const row of uniqueRows) {
    const adr = displayAdr(row.adr);
    if (!adrMap.has(adr)) {
      adrMap.set(adr, intervals.map(() => 0));
    }

    if (!row.irdDate) {
      continue;
    }

    const counts = adrMap.get(adr);
    intervals.forEach((interval, index) => {
      if (isDateInInterval(row.irdDate, interval)) {
        counts[index] += 1;
      }
    });
  }

  return adrMap;
}

function readExcelRows(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = (event) => {
      try {
        const workbook = XLSX.read(event.target.result, { type: "array", cellDates: true });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        const rawRows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });

        const parsedRows = [];

        for (let i = 0; i < rawRows.length; i += 1) {
          const row = rawRows[i];
          const caseId = String(row[0] ?? "").trim();
          const adr = row[1];
          const irdRaw = row[2];

          if (!caseId && !adr && !irdRaw) {
            continue;
          }

          if (i === 0 && /case|adr|ird|date|pt/i.test(`${caseId} ${adr}`)) {
            continue;
          }

          if (!caseId || !adr) {
            continue;
          }

          parsedRows.push({
            caseId,
            adr,
            irdDate: parseExcelDate(irdRaw),
          });
        }

        resolve(parsedRows);
      } catch (error) {
        reject(error);
      }
    };

    reader.onerror = () => reject(new Error("Unable to read the selected file."));
    reader.readAsArrayBuffer(file);
  });
}

function renderDedupSummary(originalCount, uniqueRows, duplicatesRemoved) {
  dedupStats.innerHTML = `
    <span class="stat-pill">Original rows: ${originalCount}</span>
    <span class="stat-pill">Unique case + ADR: ${uniqueRows.length}</span>
    <span class="stat-pill">Duplicates removed: ${duplicatesRemoved.length}</span>
  `;

  dedupTableBody.innerHTML = uniqueRows
    .map(
      (row) => `
        <tr>
          <td>${row.caseId}</td>
          <td>${row.adr}</td>
          <td>${row.irdDate ? formatDate(row.irdDate) : "—"}</td>
          <td>${row.removedCount}</td>
        </tr>
      `
    )
    .join("");
}

function renderIntervals(intervals) {
  intervalLabels.innerHTML = intervals
    .map(
      (interval, index) => `
        <div class="interval-chip">
          <strong>Interval ${index + 1}: ${interval.label}</strong>
          <span>${formatDate(interval.start)} to ${formatDate(interval.end)}</span>
        </div>
      `
    )
    .join("");
}

function renderResults(adrMap, intervals) {
  const sortedAdrs = Array.from(adrMap.keys()).sort((a, b) => a.localeCompare(b));

  resultsTableHead.innerHTML = `
    <tr>
      <th>ADR PT</th>
      ${intervals.map((interval, index) => `<th>${interval.label}<br><small>${formatDate(interval.start)} – ${formatDate(interval.end)}</small></th>`).join("")}
      <th>Total</th>
    </tr>
  `;

  const columnTotals = intervals.map(() => 0);

  resultsTableBody.innerHTML = sortedAdrs
    .map((adr) => {
      const counts = adrMap.get(adr);
      const rowTotal = counts.reduce((sum, value) => sum + value, 0);

      counts.forEach((value, index) => {
        columnTotals[index] += value;
      });

      return `
        <tr>
          <td>${adr}</td>
          ${counts.map((value) => `<td>${value}</td>`).join("")}
          <td><strong>${rowTotal}</strong></td>
        </tr>
      `;
    })
    .join("");

  intervalTotals.innerHTML = columnTotals
    .map(
      (total, index) => `
        <div class="total-box">
          Interval ${index + 1} total: ${total}
        </div>
      `
    )
    .join("");
}

async function handleProcess() {
  const file = excelInput.files[0];
  const dlpDate = new Date(dlpInput.value);
  const intervalType = intervalSelect.value;

  if (!file || Number.isNaN(dlpDate.getTime())) {
    showStatus("Please upload an Excel file and select a DLP date.", "error");
    return;
  }

  try {
    const rows = await readExcelRows(file);

    if (!rows.length) {
      showStatus("No valid rows found. Ensure columns A, B, and C contain Case ID, ADR PT, and IRD date.", "error");
      return;
    }

    const { uniqueRows, duplicatesRemoved } = deduplicateRows(rows);
    const intervals = buildIntervals(dlpDate, intervalType);
    const adrMap = countByInterval(uniqueRows, intervals);

    renderDedupSummary(rows.length, uniqueRows, duplicatesRemoved);
    renderIntervals(intervals);
    renderResults(adrMap, intervals);

    dedupSection.classList.remove("hidden");
    intervalSection.classList.remove("hidden");
    resultsSection.classList.remove("hidden");

    const missingDates = uniqueRows.filter((row) => !row.irdDate).length;
    let message = `Processed ${rows.length} rows. ${duplicatesRemoved.length} duplicate case + ADR entries were removed.`;

    if (missingDates) {
      message += ` Warning: ${missingDates} unique rows have missing or invalid IRD dates and were excluded from interval counts.`;
      showStatus(message, "warning");
    } else {
      showStatus(message, "success");
    }
  } catch (error) {
    showStatus(`Error processing file: ${error.message}`, "error");
  }
}
