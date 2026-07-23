let excelInput;
let dlpInput;
let authDateInput;
let intervalSelect;
let processBtn;
let feedbackMessage;
let intervalSection;
let intervalLabels;
let resultsSection;
let yearsGrossInfo;
let ptSearch;
let highlightOnlyFilter;
let filterSummary;
let resultsTableHead;
let resultsTableBody;
let downloadBtn;
let tableScroll;

let exportData = null;
let allResultsRows = [];
let resultsRows = [];
let virtualScrollBound = false;
let isProcessing = false;
let useVirtualScroll = false;

const ROW_HEIGHT = 38;
const VIRTUAL_BUFFER = 12;
const DIRECT_RENDER_LIMIT = 80;
const PROCESS_BTN_LABEL = "Process File";
const MONTH_NAMES = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
const MONTH_INDEX = Object.fromEntries(MONTH_NAMES.map((name, index) => [name.toLowerCase(), index]));

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
  authDateInput = getRequiredElement("authDate");
  intervalSelect = getRequiredElement("intervalType");
  processBtn = getRequiredElement("processBtn");
  feedbackMessage = getRequiredElement("feedbackMessage");
  intervalSection = getRequiredElement("intervalSection");
  intervalLabels = getRequiredElement("intervalLabels");
  resultsSection = getRequiredElement("resultsSection");
  yearsGrossInfo = getRequiredElement("yearsGrossInfo");
  ptSearch = getRequiredElement("ptSearch");
  highlightOnlyFilter = getRequiredElement("highlightOnlyFilter");
  filterSummary = getRequiredElement("filterSummary");
  resultsTableHead = getRequiredElement("resultsTableHead");
  resultsTableBody = getRequiredElement("resultsTableBody");
  downloadBtn = getRequiredElement("downloadBtn");
  tableScroll = getRequiredElement("tableScroll");

  excelInput.addEventListener("change", updateProcessButtonState);
  dlpInput.addEventListener("input", updateProcessButtonState);
  dlpInput.addEventListener("change", normalizeDateInputField);
  authDateInput.addEventListener("input", updateProcessButtonState);
  authDateInput.addEventListener("change", normalizeDateInputField);
  processBtn.addEventListener("click", handleProcess);
  downloadBtn.addEventListener("click", handleDownload);
  ptSearch.addEventListener("input", handleFilterChange);
  highlightOnlyFilter.addEventListener("change", handleFilterChange);
}

function updateProcessButtonState() {
  processBtn.disabled =
    isProcessing ||
    !(excelInput.files.length && parseUserDate(dlpInput.value) && parseUserDate(authDateInput.value));
}

function normalizeDateInputField(event) {
  const parsed = parseUserDate(event.target.value);
  if (parsed) {
    event.target.value = formatDate(parsed);
  }
  updateProcessButtonState();
}

function setProcessing(processing) {
  isProcessing = processing;
  processBtn.textContent = processing ? "Processing…" : PROCESS_BTN_LABEL;
  updateProcessButtonState();
}

function showFeedback(message, type = "info") {
  feedbackMessage.textContent = message;
  feedbackMessage.className = `feedback feedback-${type}`;
}

function hideFeedback() {
  feedbackMessage.textContent = "";
  feedbackMessage.className = "feedback hidden";
}

function calculateYearsGross(authDate, dlpDate) {
  const auth = normalizeDate(authDate);
  const dlp = normalizeDate(dlpDate);
  const diffMs = dlp.getTime() - auth.getTime();

  if (diffMs < 0) {
    return null;
  }

  const days = diffMs / (1000 * 60 * 60 * 24);
  return days / 365.25;
}

function calculateAdjustedCount(total, yearsGross) {
  return (total / yearsGross) * 2;
}

function formatDate(date) {
  const day = String(date.getDate()).padStart(2, "0");
  const month = MONTH_NAMES[date.getMonth()];
  const year = date.getFullYear();
  return `${day}-${month}-${year}`;
}

function getUploadedFileBaseName(fileName) {
  const trimmed = String(fileName ?? "").trim();
  if (!trimmed) {
    return "export";
  }

  return trimmed.replace(/\.(xlsx|xls|csv)$/i, "") || "export";
}

function parseUserDate(value) {
  if (!value || typeof value !== "string") {
    return null;
  }

  const trimmed = value.trim();
  const match = trimmed.match(/^(\d{1,2})[-\/\s]([A-Za-z]{3})[-\/\s](\d{4})$/);
  if (!match) {
    return null;
  }

  const day = Number(match[1]);
  const month = MONTH_INDEX[match[2].toLowerCase()];
  const year = Number(match[3]);

  if (month === undefined || day < 1 || day > 31) {
    return null;
  }

  const parsed = new Date(year, month, day);
  if (
    Number.isNaN(parsed.getTime()) ||
    parsed.getFullYear() !== year ||
    parsed.getMonth() !== month ||
    parsed.getDate() !== day
  ) {
    return null;
  }

  return parsed;
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
    const fromUserFormat = parseUserDate(value);
    if (fromUserFormat) {
      return fromUserFormat;
    }

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

function countCumulativeByAdr(uniqueRows, authDate, dlpDate) {
  const auth = normalizeDate(authDate);
  const dlp = normalizeDate(dlpDate);
  const adrMap = new Map();

  for (const row of uniqueRows) {
    if (!row.irdDate) {
      continue;
    }

    const time = row.irdDate.getTime();
    if (time < auth.getTime() || time > dlp.getTime()) {
      continue;
    }

    const adr = displayAdr(row.adr);
    adrMap.set(adr, (adrMap.get(adr) || 0) + 1);
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

function hasIncrementalTrend(counts) {
  return counts[0] > 0 && counts[1] > counts[0] && counts[2] > counts[1];
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function buildResultsRows(adrMap, cumulativeMap, yearsGross, intervalCount) {
  const allAdrs = new Set([...adrMap.keys(), ...cumulativeMap.keys()]);
  const sortedAdrs = Array.from(allAdrs).sort((a, b) => a.localeCompare(b));

  return sortedAdrs.map((adr) => {
    const counts = adrMap.get(adr) || Array(intervalCount).fill(0);
    const rowTotal = cumulativeMap.get(adr) || 0;

    const adjustedCount = calculateAdjustedCount(rowTotal, yearsGross);

    return {
      adr,
      counts,
      rowTotal,
      adjustedCount,
      highlighted: hasIncrementalTrend(counts) && counts[2] > adjustedCount,
    };
  });
}

function getFilteredRows() {
  let rows = allResultsRows;

  if (highlightOnlyFilter.checked) {
    rows = rows.filter((row) => row.highlighted);
  }

  const query = ptSearch.value.trim().toLowerCase();
  if (query) {
    rows = rows.filter((row) => row.adr.toLowerCase().includes(query));
  }

  return rows;
}

function updateFilterSummary() {
  const total = allResultsRows.length;
  const visible = resultsRows.length;
  const highlighted = allResultsRows.filter((row) => row.highlighted).length;

  if (visible === total) {
    filterSummary.textContent = `${total.toLocaleString()} PT(s) · ${highlighted.toLocaleString()} highlighted`;
    return;
  }

  filterSummary.textContent = `Showing ${visible.toLocaleString()} of ${total.toLocaleString()} PT(s)`;
}

function buildRowHtml(row) {
  return `
    <tr class="${row.highlighted ? "row-highlight" : ""}">
      <td title="${escapeHtml(row.adr)}">${escapeHtml(row.adr)}</td>
      ${row.counts.map((value) => `<td>${value}</td>`).join("")}
      <td><strong>${row.rowTotal}</strong></td>
      <td>${row.adjustedCount.toFixed(2)}</td>
    </tr>
  `;
}

function getColumnCount(intervals) {
  return intervals.length + 3;
}

function renderTableHeader(intervals) {
  resultsTableHead.innerHTML = `
    <tr>
      <th>PT(s)</th>
      ${intervals.map((interval) => `<th>${interval.label}<br><small>${formatDate(interval.start)} – ${formatDate(interval.end)}</small></th>`).join("")}
      <th>Total<br><small>Cumulative (till DLP)</small></th>
      <th>Adjusted<br><small>(Cumulative ÷ Years) × 2</small></th>
    </tr>
  `;
}

function renderAllRows() {
  if (!resultsRows.length) {
    resultsTableBody.innerHTML = `
      <tr>
        <td colspan="${getColumnCount(exportData.intervals)}" class="empty-row">No PT(s) match the current filter.</td>
      </tr>
    `;
    return;
  }

  resultsTableBody.innerHTML = resultsRows.map(buildRowHtml).join("");
}

function renderVirtualRows() {
  if (!resultsRows.length) {
    resultsTableBody.innerHTML = `
      <tr>
        <td colspan="${getColumnCount(exportData.intervals)}" class="empty-row">No PT(s) match the current filter.</td>
      </tr>
    `;
    return;
  }

  const columnCount = getColumnCount(exportData.intervals);
  const scrollTop = tableScroll.scrollTop;
  const viewportHeight = tableScroll.clientHeight;
  const headerHeight = resultsTableHead.offsetHeight || 0;
  const visibleStart = Math.max(0, scrollTop - headerHeight);
  const startIndex = Math.max(0, Math.floor(visibleStart / ROW_HEIGHT) - VIRTUAL_BUFFER);
  const visibleCount = Math.ceil(viewportHeight / ROW_HEIGHT) + VIRTUAL_BUFFER * 2;
  const endIndex = Math.min(resultsRows.length, startIndex + visibleCount);
  const topSpacerHeight = startIndex * ROW_HEIGHT;
  const bottomSpacerHeight = (resultsRows.length - endIndex) * ROW_HEIGHT;

  const parts = [];

  if (topSpacerHeight > 0) {
    parts.push(`<tr class="virtual-spacer" aria-hidden="true"><td colspan="${columnCount}" style="height:${topSpacerHeight}px;padding:0;border:none;"></td></tr>`);
  }

  for (let i = startIndex; i < endIndex; i += 1) {
    parts.push(buildRowHtml(resultsRows[i]));
  }

  if (bottomSpacerHeight > 0) {
    parts.push(`<tr class="virtual-spacer" aria-hidden="true"><td colspan="${columnCount}" style="height:${bottomSpacerHeight}px;padding:0;border:none;"></td></tr>`);
  }

  resultsTableBody.innerHTML = parts.join("");
}

function renderVisibleRows() {
  if (useVirtualScroll) {
    renderVirtualRows();
    return;
  }

  renderAllRows();
}

function bindVirtualScroll() {
  if (virtualScrollBound) {
    return;
  }

  let ticking = false;

  tableScroll.addEventListener(
    "scroll",
    () => {
      if (!useVirtualScroll || ticking) {
        return;
      }

      ticking = true;
      requestAnimationFrame(() => {
        renderVirtualRows();
        ticking = false;
      });
    },
    { passive: true }
  );

  virtualScrollBound = true;
}

function refreshTableView() {
  resultsRows = getFilteredRows();
  useVirtualScroll = resultsRows.length > DIRECT_RENDER_LIMIT;
  updateFilterSummary();
  tableScroll.scrollTop = 0;
  renderVisibleRows();
}

function handleFilterChange() {
  if (!exportData) {
    return;
  }

  refreshTableView();
}

function renderResults(adrMap, intervals, yearsGross, uniqueRows, authDate, dlpDate) {
  const cumulativeMap = countCumulativeByAdr(uniqueRows, authDate, dlpDate);
  allResultsRows = buildResultsRows(adrMap, cumulativeMap, yearsGross, intervals.length);
  ptSearch.value = "";
  highlightOnlyFilter.checked = false;

  yearsGrossInfo.textContent = `Years gross: ${yearsGross.toFixed(2)} (first authorisation to DLP). Total = cumulative count for each event till DLP. Adjusted = (Cumulative total ÷ Years gross) × 2. Highlight rule: Interval 1 > 0 and Interval 1 < Interval 2 < Interval 3, and Interval 3 > Adjusted.`;

  renderTableHeader(intervals);
  bindVirtualScroll();
  refreshTableView();
}

function buildExportRows() {
  const { intervals } = exportData;
  const intervalHeaders = intervals.map(
    (interval) => `${interval.label} (${formatDate(interval.start)} - ${formatDate(interval.end)})`
  );

  const headerRow = ["PT(s)", ...intervalHeaders, "Total (cumulative, till DLP)", "Adjusted ((Cumulative / Years) x 2)", "Incremental Trend"];

  const dataRows = allResultsRows.map((row) => [
    row.adr,
    ...row.counts,
    row.rowTotal,
    Number(row.adjustedCount.toFixed(2)),
    row.highlighted ? "Yes" : "No",
  ]);

  return { headerRow, dataRows };
}

function downloadExcel() {
  if (!exportData) {
    return;
  }

  const { yearsGross, dlpDate, authDate, intervalType } = exportData;
  const { headerRow, dataRows } = buildExportRows();

  const sheetData = [
    ["DLP Date", formatDate(dlpDate)],
    ["Date of First Authorisation", formatDate(authDate)],
    ["Interval Type", intervalType],
    ["Years Gross", Number(yearsGross.toFixed(2))],
    [],
    headerRow,
    ...dataRows,
  ];

  const worksheet = XLSX.utils.aoa_to_sheet(sheetData);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Trend Analysis");
  XLSX.writeFile(workbook, `${exportData.sourceFileBase}_Trend_output.xlsx`);
}

function handleDownload() {
  downloadExcel();
}

async function handleProcess() {
  const file = excelInput.files[0];
  const dlpDate = parseUserDate(dlpInput.value);
  const authDate = parseUserDate(authDateInput.value);
  const intervalType = intervalSelect.value;

  if (!file) {
    showFeedback("Please upload an Excel file.", "error");
    return;
  }

  if (!dlpDate) {
    showFeedback("Please enter a valid DLP date (dd-MMM-yyyy).", "error");
    return;
  }

  if (!authDate) {
    showFeedback("Please enter a valid first authorisation date (dd-MMM-yyyy).", "error");
    return;
  }

  dlpInput.value = formatDate(dlpDate);
  authDateInput.value = formatDate(authDate);

  const yearsGross = calculateYearsGross(authDate, dlpDate);

  if (yearsGross === null) {
    showFeedback("First authorisation date must be on or before the DLP date.", "error");
    return;
  }

  if (yearsGross === 0) {
    showFeedback("Years gross must be greater than zero. Check your authorisation and DLP dates.", "error");
    return;
  }

  setProcessing(true);
  hideFeedback();
  intervalSection.classList.add("hidden");
  resultsSection.classList.add("hidden");

  try {
    const rows = await readExcelRows(file);

    if (!rows.length) {
      showFeedback("No valid rows found. Check columns A (Case ID), B (PT(s)), and C (IRD date).", "error");
      return;
    }

    const uniqueRows = deduplicateRows(rows);
    const intervals = buildIntervals(dlpDate, intervalType);
    const adrMap = countByInterval(uniqueRows, intervals);

    exportData = {
      adrMap,
      intervals,
      yearsGross,
      dlpDate: normalizeDate(dlpDate),
      authDate: normalizeDate(authDate),
      intervalType,
      sourceFileBase: getUploadedFileBaseName(file.name),
    };

    renderIntervals(intervals);

    await new Promise((resolve) => requestAnimationFrame(resolve));
    renderResults(adrMap, intervals, yearsGross, uniqueRows, authDate, dlpDate);

    intervalSection.classList.remove("hidden");
    resultsSection.classList.remove("hidden");

    const highlightedCount = allResultsRows.filter((row) => row.highlighted).length;
    const missingDates = uniqueRows.filter((row) => !row.irdDate).length;
    let message = `Processed ${rows.length.toLocaleString()} rows · ${allResultsRows.length.toLocaleString()} PT(s) · ${highlightedCount.toLocaleString()} highlighted.`;

    if (missingDates) {
      message += ` ${missingDates.toLocaleString()} record(s) had missing IRD dates and were excluded from interval counts.`;
      showFeedback(message, "warning");
      return;
    }

    showFeedback(message, "success");
  } catch (error) {
    showFeedback(`Could not process file: ${error.message}`, "error");
    console.error(error);
  } finally {
    setProcessing(false);
  }
}

document.addEventListener("DOMContentLoaded", initApp);
