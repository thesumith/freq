let excelInput;
let dlpInput;
let intervalSelect;
let processBtn;
let statusSection;
let statusMessage;
let intervalSection;
let intervalLabels;
let resultsSection;
let resultsTableHead;
let resultsTableBody;
let intervalTotals;

function getRequiredElement(id) {
  const element = document.getElementById(id);
  if (!element) {
    throw new Error(`Missing required element: #${id}`);
  }
  return element;
}

function initApp() {
  excelInput = getRequiredElement("excelFile");
  dlpInput = getRequiredElement("dlpDate");
  intervalSelect = getRequiredElement("intervalType");
  processBtn = getRequiredElement("processBtn");
  statusSection = getRequiredElement("statusSection");
  statusMessage = getRequiredElement("statusMessage");
  intervalSection = getRequiredElement("intervalSection");
  intervalLabels = getRequiredElement("intervalLabels");
  resultsSection = getRequiredElement("resultsSection");
  resultsTableHead = getRequiredElement("resultsTableHead");
  resultsTableBody = getRequiredElement("resultsTableBody");
  intervalTotals = getRequiredElement("intervalTotals");

  excelInput.addEventListener("change", updateProcessButtonState);
  dlpInput.addEventListener("change", updateProcessButtonState);
  processBtn.addEventListener("click", handleProcess);
}

function updateProcessButtonState() {
  processBtn.disabled = !(excelInput.files.length && dlpInput.value);
}

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

  for (const row of rows) {
    const key = `${row.caseId}||${normalizeAdr(row.adr)}`;

    if (!grouped.has(key)) {
      grouped.set(key, {
        caseId: row.caseId,
        adr: displayAdr(row.adr),
        irdDate: row.irdDate,
      });
      continue;
    }

    const existing = grouped.get(key);

    if (row.irdDate && (!existing.irdDate || row.irdDate < existing.irdDate)) {
      existing.irdDate = row.irdDate;
    }
  }

  return Array.from(grouped.values());
}

const INTERVAL_MONTHS = {
  quarterly: 3,
  biannual: 6,
  yearly: 12,
};

const INTERVAL_LABELS = {
  quarterly: "Quarter",
  biannual: "Half-year",
  yearly: "Year",
};

function normalizeDate(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

function addDays(date, days) {
  const result = normalizeDate(date);
  result.setDate(result.getDate() + days);
  return result;
}

function subtractMonths(date, months) {
  const result = normalizeDate(date);
  const day = result.getDate();
  result.setDate(1);
  result.setMonth(result.getMonth() - months);
  const lastDay = new Date(result.getFullYear(), result.getMonth() + 1, 0).getDate();
  result.setDate(Math.min(day, lastDay));
  return result;
}

function buildIntervals(dlpDate, intervalType) {
  const dlp = normalizeDate(dlpDate);
  const periodMonths = INTERVAL_MONTHS[intervalType];
  const periodLabel = INTERVAL_LABELS[intervalType];
  const intervals = [];

  for (let i = 0; i < 3; i += 1) {
    const periodsBeforeEnd = 2 - i;
    const end = periodsBeforeEnd === 0 ? dlp : subtractMonths(dlp, periodsBeforeEnd * periodMonths);
    const start = addDays(subtractMonths(end, periodMonths), 1);

    intervals.push({
      label: `${periodLabel} ${i + 1}`,
      start,
      end,
    });
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
      ${intervals.map((interval) => `<th>${interval.label}<br><small>${formatDate(interval.start)} – ${formatDate(interval.end)}</small></th>`).join("")}
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

    const uniqueRows = deduplicateRows(rows);
    const intervals = buildIntervals(dlpDate, intervalType);
    const adrMap = countByInterval(uniqueRows, intervals);

    renderIntervals(intervals);
    renderResults(adrMap, intervals);

    intervalSection.classList.remove("hidden");
    resultsSection.classList.remove("hidden");

    const missingDates = uniqueRows.filter((row) => !row.irdDate).length;
    let message = `Processed ${uniqueRows.length} records successfully.`;

    if (missingDates) {
      message += ` Warning: ${missingDates} records have missing or invalid IRD dates and were excluded from interval counts.`;
      showStatus(message, "warning");
    } else {
      showStatus(message, "success");
    }
  } catch (error) {
    showStatus(`Error processing file: ${error.message}`, "error");
  }
}

document.addEventListener("DOMContentLoaded", initApp);
