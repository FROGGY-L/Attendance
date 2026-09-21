// ============================================================
// KFL Attendance Calculator - Fixed Version
// ============================================================

// ---------- Compact layout toggle ----------
function toggleCompactLayout(compact) {
  const headerSection = document.querySelector(".header-section");
  const mainContent = document.getElementById("mainContent");
  if (!headerSection || !mainContent) return;

  if (compact) {
    headerSection.classList.add("compact");
    mainContent.classList.add("has-data");
  } else {
    headerSection.classList.remove("compact");
    mainContent.classList.remove("has-data");
  }
}

// ---------- File input auto-trigger ----------
document.addEventListener("DOMContentLoaded", function () {
  const fileInput = document.getElementById("fileInput");
  const fileNameSpan = document.getElementById("fileName");

  if (fileInput && fileNameSpan) {
    fileInput.addEventListener("change", function (event) {
      const file = event.target.files[0];

      if (file) {
        const fileName = file.name.toLowerCase();
        if (
          fileName.endsWith(".xlsx") ||
          fileName.endsWith(".xls") ||
          fileName.endsWith(".csv")
        ) {
          fileNameSpan.textContent = `Selected File: ${file.name}`;
          console.log("File selected:", file.name);
          showToast("📁 File uploaded successfully!");

          // Auto-calculate
          setTimeout(() => {
            toggleCompactLayout(true);
            processFile();
          }, 100);
        } else {
          fileNameSpan.textContent = "Invalid file type";
          fileInput.value = "";
          showToast(
            "❌ Please select a valid Excel (.xlsx, .xls) or CSV file!",
          );
        }
      } else {
        fileNameSpan.textContent = "No file chosen";
        console.log("No file selected");
      }
    });
  } else {
    console.error("File input or fileName span not found");
  }
});

// ---------- Clear ----------
function clearFile() {
  const fileInput = document.getElementById("fileInput");
  const fileNameSpan = document.getElementById("fileName");
  const output = document.getElementById("output");

  if (fileInput) fileInput.value = "";
  if (fileNameSpan) fileNameSpan.textContent = "No file selected";
  if (output) output.innerHTML = "";

  toggleCompactLayout(false);
  window.maxBreaksFound = 0;

  showToast("✅ Data cleared successfully!");
}

// ---------- Date helpers ----------
function parseDate(dateStr) {
  if (!dateStr) return new Date(0);
  if (typeof dateStr !== "string") return new Date(dateStr);

  if (dateStr.includes("-")) {
    const parts = dateStr.split("-");
    if (parts.length === 3) {
      // Try DD-MM-YYYY first
      let day = parseInt(parts[0]);
      let month = parseInt(parts[1]);
      let year = parseInt(parts[2]);
      // If the "day" is > 12 and "month" <= 12, it's DD-MM-YYYY
      // If the "month" is > 12, it's probably YYYY-MM-DD
      if (month > 12 && day <= 12) {
        // Swap
        [day, month] = [month, day];
      }
      if (year < 100) year += 2000;
      return new Date(year, month - 1, day);
    }
  }

  if (dateStr.includes("/")) {
    const parts = dateStr.split("/");
    if (parts.length === 3) {
      let day = parseInt(parts[0]);
      let month = parseInt(parts[1]);
      let year = parseInt(parts[2]);
      if (month > 12 && day <= 12) {
        [day, month] = [month, day];
      }
      if (year < 100) year += 2000;
      return new Date(year, month - 1, day);
    }
  }

  return new Date(dateStr);
}

function combineDateTime(dateStr, timeStr) {
  if (!dateStr || !timeStr) return null;

  try {
    let day, month, year;

    if (typeof dateStr === "string" && dateStr.includes("-")) {
      const parts = dateStr.split("-");
      if (parts.length === 3) {
        day = parseInt(parts[0]);
        month = parseInt(parts[1]);
        year = parseInt(parts[2]);
        // Detect YYYY-MM-DD
        if (month > 12 && day <= 12) {
          [day, month] = [month, day];
        }
        if (year < 100) year += 2000;
      }
    } else if (typeof dateStr === "string" && dateStr.includes("/")) {
      const parts = dateStr.split("/");
      if (parts.length === 3) {
        day = parseInt(parts[0]);
        month = parseInt(parts[1]);
        year = parseInt(parts[2]);
        if (month > 12 && day <= 12) {
          [day, month] = [month, day];
        }
        if (year < 100) year += 2000;
      }
    } else if (dateStr instanceof Date) {
      day = dateStr.getDate();
      month = dateStr.getMonth() + 1;
      year = dateStr.getFullYear();
    }

    // Excel serial date number
    if (typeof dateStr === "number") {
      const excelEpoch = new Date(Date.UTC(1899, 11, 30));
      const d = new Date(excelEpoch.getTime() + dateStr * 86400000);
      day = d.getUTCDate();
      month = d.getUTCMonth() + 1;
      year = d.getUTCFullYear();
    }

    // Numeric time (Excel decimal)
    if (typeof timeStr === "number") {
      const totalMinutes = Math.round(timeStr * 24 * 60);
      const hours = Math.floor(totalMinutes / 60) % 24;
      const minutes = totalMinutes % 60;
      timeStr = `${hours.toString().padStart(2, "0")}:${minutes
        .toString()
        .padStart(2, "0")}`;
    }

    const timeParts = String(timeStr).trim().split(/[:\s]/);
    if (timeParts.length < 2) {
      console.warn(`Invalid time format: ${timeStr}`);
      return null;
    }

    let hours = parseInt(timeParts[0]);
    const minutes = parseInt(timeParts[1]);
    const ampm = timeParts[2] ? timeParts[2].toLowerCase() : null;

    if (ampm === "pm" && hours < 12) hours += 12;
    if (ampm === "am" && hours === 12) hours = 0;

    if (
      isNaN(year) ||
      isNaN(month) ||
      isNaN(day) ||
      isNaN(hours) ||
      isNaN(minutes) ||
      year < 2000 ||
      year > 2100 ||
      month < 1 ||
      month > 12 ||
      day < 1 ||
      day > 31 ||
      hours < 0 ||
      hours > 23 ||
      minutes < 0 ||
      minutes > 59
    ) {
      console.warn(`Invalid date/time components: ${dateStr} ${timeStr}`);
      return null;
    }

    const date = new Date(year, month - 1, day, hours, minutes);
    if (isNaN(date.getTime())) return null;
    return date;
  } catch (e) {
    console.warn(`Error combining date/time: ${dateStr} ${timeStr}`, e);
    return null;
  }
}

// ---------- Process file ----------
function processFile() {
  console.log("processFile called");
  const fileInput = document.getElementById("fileInput");
  if (!fileInput) {
    showToast("File input not found.");
    return;
  }

  const file = fileInput.files[0];
  if (!file) {
    showToast("Please select a file.");
    return;
  }

  const reader = new FileReader();
  const fileName = file.name.toLowerCase();

  if (fileName.endsWith(".csv")) {
    reader.onload = (e) => {
      try {
        const text = e.target.result;
        const lines = text.split(/\r?\n/);
        if (lines.length < 2) {
          showToast("This file appears to be empty.");
          return;
        }

        // Find header row
        let headerRowIndex = -1;
        for (let i = 0; i < Math.min(15, lines.length); i++) {
          const l = lines[i].toLowerCase();
          if (
            l.includes("id") &&
            (l.includes("name") || l.includes("last name")) &&
            l.includes("department")
          ) {
            headerRowIndex = i;
            break;
          }
        }

        if (headerRowIndex === -1) {
          showToast("Could not find data header in the CSV.");
          return;
        }

        const headers = parseCSVLine(lines[headerRowIndex]);
        const jsonData = [];

        for (let i = headerRowIndex + 1; i < lines.length; i++) {
          const line = lines[i].trim();
          if (!line) continue;
          const values = parseCSVLine(line);
          const record = {};
          headers.forEach((h, idx) => {
            record[h.trim()] = values[idx] ? values[idx].trim() : "";
          });
          if (
            record["ID"] &&
            (record["Name"] || record["Last Name"] || record["First Name"])
          ) {
            jsonData.push(record);
          }
        }

        console.log("CSV parsed rows:", jsonData.length);
        analyzeData(jsonData);
      } catch (err) {
        console.error("Error processing CSV:", err);
        showToast("Error processing CSV: " + err.message);
      }
    };
    reader.onerror = () => showToast("Error reading file.");
    reader.readAsText(file);
  } else if (fileName.endsWith(".xlsx") || fileName.endsWith(".xls")) {
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: "array", cellDates: false });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];

        window.originalWorksheet = worksheet;

        // Find header row dynamically (scan first 20 rows)
        const range = XLSX.utils.decode_range(worksheet["!ref"]);
        let headerRowIndex = -1;
        let headers = [];

        for (let r = range.s.r; r <= Math.min(range.s.r + 19, range.e.r); r++) {
          const rowHeaders = [];
          for (let c = range.s.c; c <= range.e.c; c++) {
            const addr = XLSX.utils.encode_cell({ r, c });
            const cell = worksheet[addr];
            rowHeaders.push(cell ? String(cell.v).trim() : "");
          }
          const headerStr = rowHeaders.join("|").toLowerCase();
          if (
            headerStr.includes("id") &&
            (headerStr.includes("name") || headerStr.includes("last name")) &&
            headerStr.includes("department") &&
            (headerStr.includes("date") || headerStr.includes("check-in"))
          ) {
            headerRowIndex = r;
            headers = rowHeaders;
            break;
          }
        }

        if (headerRowIndex === -1) {
          showToast(
            "❌ Could not find valid header row. Required columns: ID, Name (or Last Name/First Name), Department, Date, Check-In Time, Card Swiping Type",
          );
          return;
        }

        console.log("Header row index:", headerRowIndex);
        console.log("Headers:", headers);

        // Extract operator
        let operator = "";
        for (let r = 0; r < headerRowIndex; r++) {
          for (let c = range.s.c; c <= range.e.c; c++) {
            const addr = XLSX.utils.encode_cell({ r, c });
            const cell = worksheet[addr];
            if (cell && cell.v) {
              const val = String(cell.v);
              const m = val.match(/Operator:\s*(.+)/i);
              if (m) {
                operator = m[1].trim();
                break;
              }
            }
          }
          if (operator) break;
        }
        window.currentOperator = operator ? "Operator: " + operator : "";

        // Parse JSON from header row onward
        const jsonData = XLSX.utils.sheet_to_json(worksheet, {
          range: headerRowIndex,
          defval: "",
          raw: true,
        });

        console.log("Excel parsed rows:", jsonData.length);
        if (jsonData.length > 0) {
          console.log("Sample row keys:", Object.keys(jsonData[0]));
          console.log("Sample row:", jsonData[0]);
        }

        if (jsonData.length === 0) {
          showToast("No data rows found in the file.");
          return;
        }

        analyzeData(jsonData);
      } catch (err) {
        console.error("Error processing Excel:", err);
        showToast("Error processing Excel: " + err.message);
      }
    };
    reader.onerror = () => showToast("Error reading file.");
    reader.readAsArrayBuffer(file);
  } else {
    showToast("Please select a CSV or Excel file.");
  }
}

function parseCSVLine(line) {
  const result = [];
  let current = "";
  let inQuotes = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (ch === '"') {
      if (inQuotes && line[i + 1] === '"') {
        current += '"';
        i++;
      } else {
        inQuotes = !inQuotes;
      }
    } else if (ch === "," && !inQuotes) {
      result.push(current);
      current = "";
    } else {
      current += ch;
    }
  }
  result.push(current);
  return result;
}

// ---------- Analyze ----------
function analyzeData(data) {
  if (!data || data.length === 0) {
    showToast("The uploaded file contains no data.");
    return;
  }

  console.log("analyzeData called with", data.length, "records");
  console.log("Sample keys:", Object.keys(data[0]));

  const results = [];
  const employeeRecords = {};

  // Reset max breaks for each new analysis
  window.maxBreaksFound = 0;

  data.forEach((record, idx) => {
    // --- ID ---
    const id = String(record["ID"] || "").trim();
    if (!id) return;

    // --- Name (supports "Name" OR "Last Name"/"First Name") ---
    let name = (record["Name"] || "").toString().trim();
    if (!name) {
      const last = (record["Last Name"] || "").toString().trim();
      const first = (record["First Name"] || "").toString().trim();
      name = [last, first].filter(Boolean).join(", ");
    }
    if (!name) return;

    // --- Department ---
    const department = (record["Department"] || "")
      .toString()
      .replace(/^All Departments>/, "")
      .trim();

    // --- Date ---
    const date = record["Date"];
    if (!date) return;

    // --- Time ---
    const time = record["Check-In Time"] || record["Time"];
    if (!time && time !== 0) return;

    // --- Type ---
    let type = record["Card Swiping Type"];
    if (!type || String(type).trim() === "" || String(type).trim() === "-") {
      type = record["Note"];
    }
    if (!type || String(type).trim() === "" || String(type).trim() === "-")
      return;

    // Convert numeric time
    let formattedTime = time;
    if (typeof time === "number") {
      const totalMin = Math.round(time * 24 * 60);
      const h = Math.floor(totalMin / 60) % 24;
      const m = totalMin % 60;
      formattedTime = `${String(h).padStart(2, "0")}:${String(m).padStart(
        2,
        "0",
      )}`;
    }

    // Format date for later use
    let formattedDate = date;
    if (date instanceof Date) {
      const dd = String(date.getDate()).padStart(2, "0");
      const mm = String(date.getMonth() + 1).padStart(2, "0");
      const yyyy = date.getFullYear();
      formattedDate = `${dd}-${mm}-${yyyy}`;
    } else if (typeof date === "number") {
      const excelEpoch = new Date(Date.UTC(1899, 11, 30));
      const d = new Date(excelEpoch.getTime() + date * 86400000);
      const dd = String(d.getUTCDate()).padStart(2, "0");
      const mm = String(d.getUTCMonth() + 1).padStart(2, "0");
      const yyyy = d.getUTCFullYear();
      formattedDate = `${dd}-${mm}-${yyyy}`;
    } else {
      formattedDate = String(date).trim();
    }

    const dateTime = combineDateTime(formattedDate, formattedTime);
    if (!dateTime) {
      console.warn(
        `Skipping record - invalid datetime: "${date}" + "${formattedTime}"`,
      );
      return;
    }

    // Normalize type
    const lt = String(type).toLowerCase().trim();
    const normalizedType =
      lt === "check out" || lt.includes("checkout") || lt === "out"
        ? "Check Out"
        : lt === "check in" || lt.includes("checkin") || lt === "in"
          ? "Check In"
          : lt.includes("break in")
            ? "Break In"
            : lt.includes("break out")
              ? "Break Out"
              : type;

    if (!employeeRecords[id]) {
      employeeRecords[id] = { name, department, records: [] };
    }

    employeeRecords[id].records.push({
      date: formattedDate,
      time: formattedTime,
      type: normalizedType,
      dateTime,
    });
  });

  console.log("Employees grouped:", Object.keys(employeeRecords).length);

  // Sort + dedupe
  Object.values(employeeRecords).forEach((emp) => {
    emp.records.sort((a, b) => a.dateTime - b.dateTime);
    const seen = new Set();
    emp.records = emp.records.filter((r) => {
      const key = `${r.date}|${r.time}|${r.type}`;
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });
  });

  // Process each employee
  Object.entries(employeeRecords).forEach(([id, emp]) => {
    const usedRecords = new Set();
    const processedDates = new Set();

    const isSanteh =
      emp.department && emp.department.toUpperCase().includes("SANTEH");
    const isKflStaff =
      emp.department && emp.department.toUpperCase().includes("KFL-STAFF");

    for (let i = 0; i < emp.records.length; i++) {
      const record = emp.records[i];
      if (usedRecords.has(i) || processedDates.has(record.date)) continue;
      if (record.type !== "Check In") continue;

      const checkInHour = record.dateTime.getHours();
      const isMorningShift = checkInHour >= 6 && checkInHour < 12;
      const isDayShift = checkInHour >= 6 && checkInHour < 17;
      const isNightShift = checkInHour >= 17;

      let bestCheckout = null;
      let bestCheckoutIndex = -1;
      let bestTimeDiff = Infinity;
      let remarks = "";

      let fallbackCheckout = null;
      let fallbackCheckoutIndex = -1;
      let fallbackTimeDiff = Infinity;

      for (let j = i + 1; j < emp.records.length; j++) {
        if (usedRecords.has(j)) continue;

        const po = emp.records[j];
        const timeDiff = (po.dateTime - record.dateTime) / 3600000;

        const minHours = isMorningShift || isDayShift ? 1 : 4;
        const maxHours = 24;

        if (timeDiff >= minHours && timeDiff <= maxHours) {
          if (po.type === "Check Out") {
            if (timeDiff < bestTimeDiff || bestCheckout?.type !== "Check Out") {
              bestCheckout = po;
              bestCheckoutIndex = j;
              bestTimeDiff = timeDiff;
              remarks = "";
            }
          } else if (po.type === "Check In") {
            const coHour = po.dateTime.getHours();
            if ((isDayShift && coHour >= 16) || !bestCheckout) {
              if (!bestCheckout || timeDiff < bestTimeDiff) {
                bestCheckout = po;
                bestCheckoutIndex = j;
                bestTimeDiff = timeDiff;
                remarks = "Check-in treated as checkout";
              }
            }
          }
        }

        if (timeDiff >= 1 && timeDiff <= 24 && timeDiff < fallbackTimeDiff) {
          fallbackCheckout = po;
          fallbackCheckoutIndex = j;
          fallbackTimeDiff = timeDiff;
        }
      }

      if (!bestCheckout && fallbackCheckout) {
        bestCheckout = fallbackCheckout;
        bestCheckoutIndex = fallbackCheckoutIndex;
        remarks = `Using ${fallbackCheckout.type} as checkout`;
      }

      if (bestCheckout) {
        const duration = (bestCheckout.dateTime - record.dateTime) / 3600000;

        usedRecords.add(i);
        usedRecords.add(bestCheckoutIndex);

        const breakTimes = {};

        if (isSanteh || isKflStaff) {
          const breakRecords = [];
          for (let k = i + 1; k < bestCheckoutIndex; k++) {
            const br = emp.records[k];
            if (br.type === "Break In" || br.type === "Break Out") {
              breakRecords.push(br);
            }
          }
          breakRecords.sort((a, b) => a.dateTime - b.dateTime);

          if (isSanteh) {
            let pairIdx = 1;
            for (let b = 0; b < breakRecords.length - 1; b++) {
              if (
                breakRecords[b].type === "Break In" &&
                breakRecords[b + 1].type === "Break Out"
              ) {
                breakTimes[`BreakIn${pairIdx}`] = breakRecords[b].time;
                breakTimes[`BreakOut${pairIdx}`] = breakRecords[b + 1].time;
                pairIdx++;
                b++;
              }
            }
            window.maxBreaksFound = Math.max(
              window.maxBreaksFound,
              pairIdx - 1,
            );
          } else {
            // KFL-STAFF: only noon break
            if (breakRecords.length >= 2) {
              const first = breakRecords[0];
              const fh = first.dateTime.getHours();
              if (fh >= 11 && fh <= 14) {
                for (let b = 1; b < breakRecords.length; b++) {
                  const second = breakRecords[b];
                  const sh = second.dateTime.getHours();
                  if (sh >= 11 && sh <= 14) {
                    if (first.dateTime < second.dateTime) {
                      breakTimes["BreakIn1"] = first.time;
                      breakTimes["BreakOut1"] = second.time;
                    } else {
                      breakTimes["BreakIn1"] = second.time;
                      breakTimes["BreakOut1"] = first.time;
                    }
                    break;
                  }
                }
              }
            }
            if (!breakTimes["BreakIn1"]) {
              for (let b = 0; b < breakRecords.length - 1; b++) {
                if (
                  breakRecords[b].type === "Break In" &&
                  breakRecords[b + 1].type === "Break Out"
                ) {
                  breakTimes["BreakIn1"] = breakRecords[b].time;
                  breakTimes["BreakOut1"] = breakRecords[b + 1].time;
                  break;
                }
              }
            }
          }
        }

        let shiftType = "";
        let finalRemarks = remarks;
        if (isNightShift) {
          shiftType = "Night Shift";
          finalRemarks =
            finalRemarks ||
            (duration < 4
              ? "Night shift undertime"
              : duration > 12
                ? "Extended night shift"
                : "Night shift");
        } else if (isMorningShift) {
          shiftType = "Morning Shift";
          finalRemarks =
            finalRemarks ||
            (duration < 2
              ? "Short morning shift"
              : duration > 10
                ? "Extended morning shift"
                : "Morning shift");
        } else {
          shiftType = "Day Shift";
          finalRemarks =
            finalRemarks ||
            (duration < 4
              ? "Undertime - early checkout"
              : duration > 12
                ? "Extended day shift"
                : duration > 9
                  ? "Day shift with overtime"
                  : "Day shift");
        }

        const resultRecord = {
          Employee: emp.name,
          Department: emp.department,
          Status: shiftType,
          Duration: `${Math.floor(duration)}h ${Math.round(
            (duration % 1) * 60,
          )}m`,
          Date: record.date,
          CheckIn: record.time,
          CheckOut: bestCheckout.time,
          Remarks: finalRemarks,
        };

        if (isSanteh) {
          for (let i = 1; i <= (window.maxBreaksFound || 2); i++) {
            resultRecord[`BreakIn${i}`] = breakTimes[`BreakIn${i}`] || "-";
            resultRecord[`BreakOut${i}`] = breakTimes[`BreakOut${i}`] || "-";
          }
        } else if (isKflStaff) {
          resultRecord["BreakIn1"] = breakTimes["BreakIn1"] || "-";
          resultRecord["BreakOut1"] = breakTimes["BreakOut1"] || "-";
        }

        results.push(resultRecord);
        processedDates.add(record.date);
      } else {
        usedRecords.add(i);
        results.push({
          Employee: emp.name,
          Department: emp.department,
          Status: "Missing Check Out",
          Duration: "-",
          Date: record.date,
          CheckIn: record.time,
          CheckOut: "-",
          Remarks: "Missing checkout - no additional records",
        });
        processedDates.add(record.date);
      }
    }

    // ✅ FIXED: After processing all Check Ins, find ORPHANED Check Outs
    // (Check Outs that were never matched to a Check In)
    for (let i = 0; i < emp.records.length; i++) {
      if (usedRecords.has(i)) continue;
      const record = emp.records[i];
      if (record.type !== "Check Out") continue;
      if (processedDates.has(record.date)) continue;

      // Found an orphaned Check Out → mark as Missing Check In
      usedRecords.add(i);
      results.push({
        Employee: emp.name,
        Department: emp.department,
        Status: "Missing Check In",
        Duration: "-",
        Date: record.date,
        CheckIn: "-",
        CheckOut: record.time,
        Remarks: "Missing check-in - no preceding Check In record",
      });
      processedDates.add(record.date);
    }

    // ✅ FIXED (optional): Also flag orphaned Break In/Out records
    for (let i = 0; i < emp.records.length; i++) {
      if (usedRecords.has(i)) continue;
      const record = emp.records[i];
      if (record.type === "Break In" || record.type === "Break Out") {
        if (processedDates.has(record.date)) continue;
        usedRecords.add(i);
        results.push({
          Employee: emp.name,
          Department: emp.department,
          Status: `Orphaned ${record.type}`,
          Duration: "-",
          Date: record.date,
          CheckIn: "-",
          CheckOut: "-",
          Remarks: `${record.type} record without a matching shift`,
        });
        processedDates.add(record.date);
      }
    }
  });

  console.log("Total results:", results.length);
  displayResults(results);
}

// ============================================================
// DISPLAY RESULTS — Enhanced UX Preview
// ============================================================
function displayResults(results) {
  const output = document.getElementById("output");
  if (!output) return;
  output.innerHTML = "";

  if (!results || results.length === 0) {
    output.innerHTML = `
      <div class="empty-state">
        <i class="fas fa-inbox"></i>
        <h3>No records found</h3>
        <p>The uploaded file didn't contain any valid attendance records.</p>
      </div>`;
    return;
  }

  // Sort by employee, then date
  results.sort((a, b) => {
    const nc = (a.Employee || "").localeCompare(b.Employee || "");
    if (nc !== 0) return nc;
    return parseDate(a.Date) - parseDate(b.Date);
  });

  const hasSanteh = results.some(
    (r) => r.Department && r.Department.toUpperCase().includes("SANTEH"),
  );
  const hasKflStaff = results.some(
    (r) => r.Department && r.Department.toUpperCase().includes("KFL-STAFF"),
  );
  const hasBreak = hasSanteh || hasKflStaff;

  // Build headers
  let headers;
  if (hasBreak) {
    headers = [
      { key: "Employee", label: "Employee" },
      { key: "Department", label: "Department" },
      { key: "Status", label: "Status" },
      { key: "Duration", label: "Hours" },
      { key: "Date", label: "Date" },
      { key: "CheckIn", label: "Check In" },
    ];
    if (hasSanteh) {
      const mb = window.maxBreaksFound || 2;
      for (let i = 1; i <= mb; i++) {
        const ord =
          i === 1 ? "1st" : i === 2 ? "2nd" : i === 3 ? "3rd" : `${i}th`;
        headers.push({ key: `BreakIn${i}`, label: `${ord} Break In` });
        headers.push({ key: `BreakOut${i}`, label: `${ord} Break Out` });
      }
    } else {
      headers.push({ key: "BreakIn1", label: "Noon Break In" });
      headers.push({ key: "BreakOut1", label: "Noon Break Out" });
    }
    headers.push({ key: "CheckOut", label: "Check Out" });
    headers.push({ key: "Remarks", label: "Remarks" });
  } else {
    headers = [
      { key: "Employee", label: "Employee" },
      { key: "Department", label: "Department" },
      { key: "Status", label: "Status" },
      { key: "Duration", label: "Hours" },
      { key: "Date", label: "Date" },
      { key: "CheckIn", label: "Check In" },
      { key: "CheckOut", label: "Check Out" },
      { key: "Remarks", label: "Remarks" },
    ];
  }

  // ---------- Build summary cards ----------
  const uniqueEmployees = [...new Set(results.map((r) => r.Employee))];
  const statusCounts = {
    "Day Shift": 0,
    "Night Shift": 0,
    "Morning Shift": 0,
    "Missing Check Out": 0,
  };
  results.forEach((r) => {
    if (statusCounts[r.Status] !== undefined) statusCounts[r.Status]++;
  });

  const summaryHTML = `
    <div class="summary-cards">
      <div class="summary-card">
        <div class="summary-icon"><i class="fas fa-users"></i></div>
        <div class="summary-info">
          <div class="summary-value">${uniqueEmployees.length}</div>
          <div class="summary-label">Employees</div>
        </div>
      </div>
      <div class="summary-card">
        <div class="summary-icon"><i class="fas fa-list"></i></div>
        <div class="summary-info">
          <div class="summary-value">${results.length}</div>
          <div class="summary-label">Records</div>
        </div>
      </div>
      <div class="summary-card status-day">
        <div class="summary-icon"><i class="fas fa-sun"></i></div>
        <div class="summary-info">
          <div class="summary-value">${statusCounts["Day Shift"]}</div>
          <div class="summary-label">Day Shift</div>
        </div>
      </div>
      <div class="summary-card status-night">
        <div class="summary-icon"><i class="fas fa-moon"></i></div>
        <div class="summary-info">
          <div class="summary-value">${statusCounts["Night Shift"]}</div>
          <div class="summary-label">Night Shift</div>
        </div>
      </div>
      <div class="summary-card status-missing">
        <div class="summary-icon"><i class="fas fa-exclamation-triangle"></i></div>
        <div class="summary-info">
          <div class="summary-value">${statusCounts["Missing Check Out"]}</div>
          <div class="summary-label">Missing CO</div>
        </div>
      </div>
    </div>`;

  // ---------- Build toolbar ----------
  const employeeOptions = uniqueEmployees
    .map((e) => `<option value="${e}">${e}</option>`)
    .join("");

  const deptOptions = [...new Set(results.map((r) => r.Department))]
    .filter(Boolean)
    .map((d) => `<option value="${d}">${d}</option>`)
    .join("");

  const toolbarHTML = `
    <div class="preview-toolbar">
      <div class="toolbar-row">
        <div class="search-wrap">
          <i class="fas fa-search"></i>
          <input type="text" id="previewSearch" placeholder="Search employees, dates, remarks…" />
        </div>
        <div class="filter-group">
          <select id="filterEmployee">
            <option value="">All Employees</option>
            ${employeeOptions}
          </select>
          <select id="filterDepartment">
            <option value="">All Departments</option>
            ${deptOptions}
          </select>
          <select id="filterStatus">
            <option value="">All Statuses</option>
            <option value="Day Shift">Day Shift</option>
            <option value="Night Shift">Night Shift</option>
            <option value="Morning Shift">Morning Shift</option>
            <option value="Missing Check Out">Missing Check Out</option>
          </select>
          <button type="button" id="clearFiltersBtn" class="clear-filters-btn" title="Clear filters">
            <i class="fas fa-times"></i>
          </button>
        </div>
      </div>
      <div class="toolbar-info">
        <span id="resultCount">Showing ${results.length} of ${results.length} records</span>
        <span class="toolbar-hint"><i class="fas fa-info-circle"></i> Click a column header to sort</span>
      </div>
    </div>`;

  // ---------- Build table ----------
  const tableContainer = document.createElement("div");
  tableContainer.className = "table-container";

  const table = document.createElement("table");
  table.className = "preview-table";

  const headerRow = document.createElement("tr");
  headers.forEach((h, idx) => {
    const th = document.createElement("th");
    th.textContent = h.label;
    th.dataset.key = h.key;
    th.dataset.index = idx;
    th.classList.add("sortable");
    th.innerHTML = `${h.label} <i class="fas fa-sort sort-icon"></i>`;
    headerRow.appendChild(th);
  });
  table.appendChild(headerRow);

  // Rows
  let lastEmployee = null;
  let groupIndex = 0;
  results.forEach((result, rIdx) => {
    const row = document.createElement("tr");
    row.dataset.employee = result.Employee || "";
    row.dataset.department = result.Department || "";
    row.dataset.status = result.Status || "";
    row.dataset.search = headers
      .map((h) => String(result[h.key] || ""))
      .join(" ")
      .toLowerCase();

    // Alternating group shading
    if (result.Employee !== lastEmployee) {
      groupIndex++;
      lastEmployee = result.Employee;
    }
    row.classList.add(groupIndex % 2 === 0 ? "group-even" : "group-odd");

    headers.forEach((h) => {
      const td = document.createElement("td");
      const val = result[h.key];
      const displayVal =
        val === undefined || val === null || val === "" ? "-" : val;

      // Special cell renderers
      if (h.key === "Employee") {
        td.className = "cell-employee";
        td.textContent = displayVal;
      } else if (h.key === "Department") {
        td.className = "cell-department";
        td.textContent = displayVal;
        td.title = displayVal;
      } else if (h.key === "Status") {
        td.className = "cell-status";
        const cls = statusToClass(displayVal);
        td.innerHTML = `<span class="status-badge ${cls}"><i class="fas ${statusIcon(
          displayVal,
        )}"></i> ${displayVal}</span>`;
      } else if (h.key === "Duration") {
        td.className = "cell-duration";
        td.textContent = displayVal;
      } else if (h.key === "Date") {
        td.className = "cell-date";
        td.textContent = displayVal;
      } else if (h.key === "CheckOut" && displayVal === "-") {
        td.className = "cell-missing";
        td.innerHTML = `<span class="missing-badge"><i class="fas fa-exclamation-circle"></i> Missing</span>`;
      } else if (h.key === "Remarks") {
        td.className = "cell-remarks";
        td.textContent = displayVal;
        td.title = displayVal;
      } else {
        td.textContent = displayVal;
      }

      row.appendChild(td);
    });

    table.appendChild(row);
  });

  tableContainer.appendChild(table);

  // ---------- Assemble ----------
  output.innerHTML = summaryHTML + toolbarHTML;
  output.appendChild(tableContainer);

  // ---------- Wire up interactions ----------
  wirePreviewInteractions(output, table, results.length);
}

// ---------- Helpers ----------
function statusToClass(status) {
  if (status === "Day Shift") return "badge-day";
  if (status === "Night Shift") return "badge-night";
  if (status === "Morning Shift") return "badge-morning";
  if (status === "Missing Check Out") return "badge-missing";
  return "badge-default";
}

function statusIcon(status) {
  if (status === "Day Shift") return "fa-sun";
  if (status === "Night Shift") return "fa-moon";
  if (status === "Morning Shift") return "fa-cloud-sun";
  if (status === "Missing Check Out") return "fa-exclamation-triangle";
  return "fa-circle";
}

function wirePreviewInteractions(output, table, totalCount) {
  const search = output.querySelector("#previewSearch");
  const filterEmp = output.querySelector("#filterEmployee");
  const filterDept = output.querySelector("#filterDepartment");
  const filterStatus = output.querySelector("#filterStatus");
  const clearBtn = output.querySelector("#clearFiltersBtn");
  const resultCount = output.querySelector("#resultCount");

  function applyFilters() {
    const q = (search.value || "").toLowerCase().trim();
    const emp = filterEmp.value;
    const dept = filterDept.value;
    const status = filterStatus.value;

    let visible = 0;
    const rows = table.querySelectorAll("tbody tr, tr:not(:first-child)");
    rows.forEach((row) => {
      const matchesSearch = !q || row.dataset.search.includes(q);
      const matchesEmp = !emp || row.dataset.employee === emp;
      const matchesDept = !dept || row.dataset.department === dept;
      const matchesStatus = !status || row.dataset.status === status;
      const show = matchesSearch && matchesEmp && matchesDept && matchesStatus;
      row.style.display = show ? "" : "none";
      if (show) visible++;
    });

    resultCount.textContent = `Showing ${visible} of ${totalCount} records`;
  }

  search.addEventListener("input", applyFilters);
  filterEmp.addEventListener("change", applyFilters);
  filterDept.addEventListener("change", applyFilters);
  filterStatus.addEventListener("change", applyFilters);

  clearBtn.addEventListener("click", () => {
    search.value = "";
    filterEmp.value = "";
    filterDept.value = "";
    filterStatus.value = "";
    applyFilters();
  });

  // ---------- Sorting ----------
  const headerRow = table.querySelector("tr");
  let sortState = { key: null, dir: 1 };

  headerRow.querySelectorAll("th.sortable").forEach((th) => {
    th.addEventListener("click", () => {
      const key = th.dataset.key;
      const idx = parseInt(th.dataset.index);

      if (sortState.key === key) {
        sortState.dir *= -1;
      } else {
        sortState.key = key;
        sortState.dir = 1;
      }

      // Update sort icons
      headerRow.querySelectorAll("th .sort-icon").forEach((icon) => {
        icon.className = "fas fa-sort sort-icon";
      });
      const icon = th.querySelector(".sort-icon");
      icon.className =
        sortState.dir === 1
          ? "fas fa-sort-up sort-icon"
          : "fas fa-sort-down sort-icon";

      // Sort rows
      const rows = Array.from(table.querySelectorAll("tr:not(:first-child)"));
      rows.sort((a, b) => {
        const av = a.querySelectorAll("td")[idx]?.textContent.trim() || "";
        const bv = b.querySelectorAll("td")[idx]?.textContent.trim() || "";
        // Date compare
        if (key === "Date") {
          return (parseDate(av) - parseDate(bv)) * sortState.dir;
        }
        // Numeric compare for durations
        if (key === "Duration") {
          const parse = (s) => {
            const m = String(s).match(/(\d+)h\s*(\d+)?m?/);
            if (!m) return 0;
            return parseInt(m[1]) * 60 + (parseInt(m[2]) || 0);
          };
          return (parse(av) - parse(bv)) * sortState.dir;
        }
        return av.localeCompare(bv) * sortState.dir;
      });

      rows.forEach((r) => table.appendChild(r));
    });
  });
}

// ---------- Toasts ----------
function showToast(message) {
  const toast = document.createElement("div");
  toast.innerHTML = message;
  Object.assign(toast.style, {
    position: "fixed",
    top: "20px",
    left: "50%",
    transform: "translateX(-50%) translateY(-100px)",
    background: "linear-gradient(45deg, #4facfe, #00f2fe)",
    color: "white",
    padding: "15px 25px",
    borderRadius: "25px",
    fontSize: "16px",
    fontWeight: "600",
    boxShadow: "0 8px 32px rgba(79, 172, 254, 0.3)",
    zIndex: "10000",
    transition: "all 0.4s cubic-bezier(0.68, -0.55, 0.265, 1.55)",
    opacity: "0",
  });
  document.body.appendChild(toast);
  setTimeout(() => {
    toast.style.transform = "translateX(-50%) translateY(0)";
    toast.style.opacity = "1";
  }, 10);
  setTimeout(() => {
    toast.style.transform = "translateX(-50%) translateY(-100px)";
    toast.style.opacity = "0";
    setTimeout(() => toast.remove(), 400);
  }, 3000);
}

function showToastMessage() {
  const msg = `
    <strong>Important Notice:</strong><br><br>
    Please double-check the file provided by the system.<br><br>
    <strong>Update 09-14-2026:</strong><br>
    - Fixed missing checkout for long shifts.<br>
    - Added support for <em>ID / Last Name / First Name / Department</em> format.<br>
    - Auto-calculate on file upload.<br>
    <strong>* Always double-check the data.</strong><br><br>
    For inquiries, contact IT Personnel.<br><br>
    <strong>Thank you!</strong>
  `;

  const backdrop = document.createElement("div");
  Object.assign(backdrop.style, {
    position: "fixed",
    top: "0",
    left: "0",
    width: "100%",
    height: "100%",
    backgroundColor: "rgba(0,0,0,0.5)",
    zIndex: "9998",
    opacity: "0",
    transition: "opacity 0.3s ease",
  });

  const toast = document.createElement("div");
  toast.innerHTML = msg;
  Object.assign(toast.style, {
    position: "fixed",
    top: "50%",
    left: "50%",
    transform: "translate(-50%, -50%) scale(0.7)",
    padding: "25px 35px",
    background: "linear-gradient(135deg, #f44336, #e53935)",
    color: "#fff",
    borderRadius: "15px",
    fontSize: "16px",
    textAlign: "center",
    zIndex: "9999",
    lineHeight: "1.6",
    maxWidth: "90vw",
    maxHeight: "80vh",
    overflowY: "auto",
    boxShadow: "0 20px 60px rgba(244,67,54,0.3)",
    transition: "all 0.4s cubic-bezier(0.68, -0.55, 0.265, 1.55)",
    opacity: "0",
  });

  const closeBtn = document.createElement("button");
  closeBtn.innerHTML = "×";
  Object.assign(closeBtn.style, {
    position: "absolute",
    top: "10px",
    right: "15px",
    background: "none",
    border: "none",
    color: "white",
    fontSize: "24px",
    cursor: "pointer",
  });
  closeBtn.onclick = closeModal;
  toast.appendChild(closeBtn);

  document.body.appendChild(backdrop);
  document.body.appendChild(toast);

  setTimeout(() => {
    backdrop.style.opacity = "1";
    toast.style.transform = "translate(-50%, -50%) scale(1)";
    toast.style.opacity = "1";
  }, 10);

  function closeModal() {
    toast.style.transform = "translate(-50%, -50%) scale(0.7)";
    toast.style.opacity = "0";
    backdrop.style.opacity = "0";
    setTimeout(() => {
      toast.remove();
      backdrop.remove();
    }, 400);
  }

  setTimeout(closeModal, 8000);
  backdrop.onclick = closeModal;
}

window.addEventListener("DOMContentLoaded", () => {
  // Only show notice once per session
  if (!sessionStorage.getItem("kfl_notice_shown")) {
    showToastMessage();
    sessionStorage.setItem("kfl_notice_shown", "1");
  }
});

// ============================================================
// EXPORT EXCEL — Enhanced with Summary sheet
// ============================================================
async function exportToExcel() {
  const output = document.getElementById("output");
  const table = output.querySelector("table");

  if (!table) {
    showToast("No data to export.");
    return;
  }

  // ---------- Extract data from HTML table ----------
  const results = [];
  const headers = Array.from(table.querySelectorAll("th")).map((th) =>
    th.textContent.replace(/[⇅]/g, "").trim(),
  );
  const rows = table.querySelectorAll("tr");
  const hasBreakColumns = headers.some((h) => h.includes("Break In"));

  const dates = [];
  rows.forEach((row, ri) => {
    if (ri === 0) return;
    const dateCell = row.cells[4];
    if (dateCell) dates.push(dateCell.textContent);
  });

  const sortedDates = dates.sort((a, b) => parseDate(a) - parseDate(b));
  const minDate = sortedDates[0] || "";
  const maxDate = sortedDates[sortedDates.length - 1] || "";

  const now = new Date();
  const currentDate = `${String(now.getDate()).padStart(2, "0")}-${String(
    now.getMonth() + 1,
  ).padStart(2, "0")}-${now.getFullYear()}`;
  const currentTime = `${String(now.getHours()).padStart(2, "0")}:${String(
    now.getMinutes(),
  ).padStart(2, "0")}`;
  const operator = window.currentOperator || "";

  rows.forEach((row, ri) => {
    if (ri === 0) return;
    const rowData = {};
    const cells = row.querySelectorAll("td");
    cells.forEach((cell, ci) => {
      if (headers[ci] !== "Remarks") {
        let v = cell.textContent;
        if (headers[ci] === "Department")
          v = v.replace(/^All Departments>/, "");
        if (headers[ci] === "Noon Break In") rowData["BreakIn1"] = v;
        else if (headers[ci] === "Noon Break Out") rowData["BreakOut1"] = v;
        else rowData[headers[ci]] = v;
      }
    });
    results.push(rowData);
  });

  // ---------- Create workbook ----------
  const wb = new ExcelJS.Workbook();
  wb.creator = "KFL Manpower Agency";
  wb.created = new Date();

  const maxBreaks = window.maxBreaksFound || 2;
  const lastColIndex = hasBreakColumns ? 6 + maxBreaks * 2 + 1 : 7;

  // ============================================================
  // Helper: duration string "13h 21m" → decimal hours
  // ============================================================
  const durationToHours = (str) => {
    if (!str || str === "-") return 0;
    const m = String(str).match(/(\d+)h\s*(\d+)?m?/);
    if (!m) return 0;
    return parseInt(m[1]) + (parseInt(m[2]) || 0) / 60;
  };

  // ============================================================
  // SHEET 0: SUMMARY (NEW)
  // ============================================================
  const summarySheet = wb.addWorksheet("Summary", {
    properties: { tabColor: { argb: "FF4472C4" } },
  });

  summarySheet.mergeCells(1, 1, 3, 6);
  const st = summarySheet.getCell("A1");
  st.value = "KFL MANPOWER AGENCY SERVER 3";
  st.font = { bold: true, size: 18, color: { argb: "FF1F4E78" } };
  st.alignment = { horizontal: "center", vertical: "middle" };
  st.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFD9E1F2" },
  };
  for (let r = 1; r <= 3; r++) {
    for (let c = 1; c <= 6; c++) {
      summarySheet.getCell(r, c).border = {
        top: { style: "double" },
        bottom: { style: "double" },
        left: { style: "double" },
        right: { style: "double" },
      };
    }
  }

  summarySheet.mergeCells(4, 1, 4, 6);
  const sTitle = summarySheet.getCell("A4");
  sTitle.value = "ATTENDANCE SUMMARY REPORT";
  sTitle.font = { bold: true, size: 14, color: { argb: "FF1F4E78" } };
  sTitle.alignment = { horizontal: "center", vertical: "middle" };
  sTitle.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFD9E1F2" },
  };

  summarySheet.getCell("A6").value = operator || "Operator: —";
  summarySheet.getCell("A6").font = { bold: true, size: 11 };
  summarySheet.getCell("D6").value = `Export: ${currentDate} ${currentTime}`;
  summarySheet.getCell("D6").font = {
    italic: true,
    color: { argb: "FF666666" },
  };
  summarySheet.getCell("A7").value = `Period: ${minDate} - ${maxDate}`;
  summarySheet.getCell("A7").font = {
    italic: true,
    color: { argb: "FF666666" },
  };

  // Compute stats
  const totalRecords = results.length;
  const uniqueEmployees = [...new Set(results.map((r) => r.Employee))];
  const statusCounts = {
    "Day Shift": 0,
    "Night Shift": 0,
    "Morning Shift": 0,
    "Missing Check Out": 0,
    "Missing Check In": 0,
  };
  let totalHours = 0;
  results.forEach((r) => {
    if (statusCounts[r.Status] !== undefined) statusCounts[r.Status]++;
    totalHours += durationToHours(r.Duration);
  });

  let sRow = 9;
  summarySheet.getCell(sRow, 1).value = "OVERVIEW";
  summarySheet.mergeCells(sRow, 1, sRow, 6);
  summarySheet.getCell(sRow, 1).font = {
    bold: true,
    size: 12,
    color: { argb: "FFFFFFFF" },
  };
  summarySheet.getCell(sRow, 1).fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FF305496" },
  };
  summarySheet.getCell(sRow, 1).alignment = {
    horizontal: "center",
    vertical: "middle",
  };
  sRow++;

  const overviewItems = [
    ["Total Employees", uniqueEmployees.length],
    ["Total Records", totalRecords],
    ["Total Hours Rendered", Math.round(totalHours * 100) / 100],
    [
      "Average Hours/Record",
      totalRecords ? Math.round((totalHours / totalRecords) * 100) / 100 : 0,
    ],
    ["Date Range", `${minDate} to ${maxDate}`],
  ];

  overviewItems.forEach(([label, value]) => {
    const lc = summarySheet.getCell(sRow, 1);
    lc.value = label;
    lc.font = { bold: true };
    lc.border = {
      top: { style: "thin" },
      bottom: { style: "thin" },
      left: { style: "thin" },
      right: { style: "thin" },
    };
    summarySheet.mergeCells(sRow, 1, sRow, 3);

    const vc = summarySheet.getCell(sRow, 4);
    vc.value = value;
    vc.alignment = { horizontal: "center" };
    vc.border = {
      top: { style: "thin" },
      bottom: { style: "thin" },
      left: { style: "thin" },
      right: { style: "thin" },
    };
    summarySheet.mergeCells(sRow, 4, sRow, 6);
    sRow++;
  });

  sRow += 2;
  summarySheet.getCell(sRow, 1).value = "SHIFT BREAKDOWN";
  summarySheet.mergeCells(sRow, 1, sRow, 6);
  summarySheet.getCell(sRow, 1).font = {
    bold: true,
    size: 12,
    color: { argb: "FFFFFFFF" },
  };
  summarySheet.getCell(sRow, 1).fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FF305496" },
  };
  summarySheet.getCell(sRow, 1).alignment = {
    horizontal: "center",
    vertical: "middle",
  };
  sRow++;

  // Shift breakdown header
  ["Status", "Count", "Percentage", "", "", ""].forEach((h, i) => {
    const c = summarySheet.getCell(sRow, i + 1);
    c.value = h;
    c.font = { bold: true, size: 10 };
    c.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFD9E1F2" },
    };
    c.alignment = { horizontal: "center" };
    c.border = {
      top: { style: "thin" },
      bottom: { style: "thin" },
      left: { style: "thin" },
      right: { style: "thin" },
    };
  });
  summarySheet.mergeCells(sRow, 3, sRow, 6);
  sRow++;

  const statusColors = {
    "Day Shift": "FF375623",
    "Night Shift": "FF1F4E78",
    "Morning Shift": "FFBF8F00",
    "Missing Check Out": "FFC00000",
    "Missing Check In": "FFE67E22",
  };

  Object.entries(statusCounts).forEach(([label, count]) => {
    const pct = totalRecords ? (count / totalRecords) * 100 : 0;

    const lc = summarySheet.getCell(sRow, 1);
    lc.value = label;
    lc.font = {
      bold: true,
      color: { argb: statusColors[label] || "FF000000" },
    };
    lc.border = {
      top: { style: "thin" },
      bottom: { style: "thin" },
      left: { style: "thin" },
      right: { style: "thin" },
    };

    const cc = summarySheet.getCell(sRow, 2);
    cc.value = count;
    cc.alignment = { horizontal: "center" };
    cc.border = {
      top: { style: "thin" },
      bottom: { style: "thin" },
      left: { style: "thin" },
      right: { style: "thin" },
    };

    const pc = summarySheet.getCell(sRow, 3);
    pc.value = `${pct.toFixed(1)}%`;
    pc.alignment = { horizontal: "center" };
    pc.border = {
      top: { style: "thin" },
      bottom: { style: "thin" },
      left: { style: "thin" },
      right: { style: "thin" },
    };
    summarySheet.mergeCells(sRow, 3, sRow, 6);

    sRow++;
  });

  summarySheet.columns = [
    { width: 30 },
    { width: 15 },
    { width: 15 },
    { width: 15 },
    { width: 15 },
    { width: 15 },
  ];

  // ============================================================
  // SHEET 1: EMPLOYEE SUMMARY (NEW)
  // ============================================================
  const empSheet = wb.addWorksheet("Employee Summary", {
    properties: { tabColor: { argb: "FF00C851" } },
    views: [{ state: "frozen", ySplit: 4 }],
  });

  empSheet.mergeCells(1, 1, 2, 9);
  const es = empSheet.getCell("A1");
  es.value = "EMPLOYEE SUMMARY";
  es.font = { bold: true, size: 16, color: { argb: "FF1F4E78" } };
  es.alignment = { horizontal: "center", vertical: "middle" };
  es.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFD9E1F2" },
  };

  empSheet.getCell("A3").value = operator;
  empSheet.getCell("A3").font = { bold: true, size: 10 };
  empSheet.getCell("F3").value = `Period: ${minDate} - ${maxDate}`;
  empSheet.getCell("F3").font = { italic: true, size: 10 };
  empSheet.getCell("H3").value = `Export: ${currentDate} ${currentTime}`;
  empSheet.getCell("H3").font = { italic: true, size: 10 };

  const empHeaders = [
    "Employee",
    "Department",
    "Total Days",
    "Day Shift",
    "Night Shift",
    "Morning Shift",
    "Missing CO",
    "Missing CI",
    "Total Hours",
  ];
  empHeaders.forEach((h, i) => {
    const c = empSheet.getCell(5, i + 1);
    c.value = h;
    c.font = { bold: true, color: { argb: "FFFFFFFF" }, size: 11 };
    c.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FF305496" },
    };
    c.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
    c.border = {
      top: { style: "thin" },
      bottom: { style: "thin" },
      left: { style: "thin" },
      right: { style: "thin" },
    };
  });

  // Group by employee
  const byEmp = {};
  results.forEach((r) => {
    if (!byEmp[r.Employee]) {
      byEmp[r.Employee] = {
        department: r.Department,
        days: 0,
        dayShift: 0,
        nightShift: 0,
        morningShift: 0,
        missingCO: 0,
        missingCI: 0,
        hours: 0,
      };
    }
    const e = byEmp[r.Employee];
    e.days++;
    e.hours += durationToHours(r.Duration);
    if (r.Status === "Day Shift") e.dayShift++;
    else if (r.Status === "Night Shift") e.nightShift++;
    else if (r.Status === "Morning Shift") e.morningShift++;
    else if (r.Status === "Missing Check Out") e.missingCO++;
    else if (r.Status === "Missing Check In") e.missingCI++;
  });

  let er = 6;
  Object.entries(byEmp)
    .sort((a, b) => a[0].localeCompare(b[0]))
    .forEach(([emp, e]) => {
      const vals = [
        emp,
        e.department,
        e.days,
        e.dayShift,
        e.nightShift,
        e.morningShift,
        e.missingCO,
        e.missingCI,
        Math.round(e.hours * 100) / 100,
      ];
      vals.forEach((v, i) => {
        const c = empSheet.getCell(er, i + 1);
        c.value = v;
        c.border = {
          top: { style: "thin" },
          bottom: { style: "thin" },
          left: { style: "thin" },
          right: { style: "thin" },
        };
        c.alignment = {
          horizontal: i <= 1 ? "left" : "center",
          vertical: "middle",
        };
        if (i === 0) c.font = { bold: true };
      });
      er++;
    });

  // Totals row
  const totalsRow = er;
  empSheet.getCell(totalsRow, 1).value = "TOTAL";
  empSheet.getCell(totalsRow, 1).font = {
    bold: true,
    color: { argb: "FFFFFFFF" },
  };
  empSheet.getCell(totalsRow, 1).fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FF4472C4" },
  };
  empSheet.mergeCells(totalsRow, 1, totalsRow, 2);
  const colTotals = [0, 0, 0, 0, 0, 0, 0]; // days, day, night, morning, mCO, mCI, hours
  Object.values(byEmp).forEach((e) => {
    colTotals[0] += e.days;
    colTotals[1] += e.dayShift;
    colTotals[2] += e.nightShift;
    colTotals[3] += e.morningShift;
    colTotals[4] += e.missingCO;
    colTotals[5] += e.missingCI;
    colTotals[6] += e.hours;
  });
  colTotals.forEach((v, i) => {
    const c = empSheet.getCell(totalsRow, i + 3);
    c.value = i === 6 ? Math.round(v * 100) / 100 : v;
    c.font = { bold: true, color: { argb: "FFFFFFFF" } };
    c.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FF4472C4" },
    };
    c.alignment = { horizontal: "center" };
    c.border = {
      top: { style: "thin" },
      bottom: { style: "thin" },
      left: { style: "thin" },
      right: { style: "thin" },
    };
  });

  empSheet.columns = [
    { width: 28 },
    { width: 35 },
    { width: 12 },
    { width: 12 },
    { width: 13 },
    { width: 15 },
    { width: 12 },
    { width: 12 },
    { width: 14 },
  ];

  empSheet.autoFilter = {
    from: { row: 5, column: 1 },
    to: { row: 5, column: 9 },
  };

  // ============================================================
  // SHEET 2: Attendance Pivot (existing, enhanced)
  // ============================================================
  const pivotSheet = wb.addWorksheet("Attendance Pivot", {
    views: [{ state: "frozen", ySplit: 6 }],
    properties: { tabColor: { argb: "FF4472C4" } },
  });

  pivotSheet.mergeCells(1, 1, 3, 7);
  const cc = pivotSheet.getCell("A1");
  cc.value = "KFL MANPOWER AGENCY SERVER 3";
  cc.font = { bold: true, size: 16, color: { argb: "FF1F4E78" } };
  cc.alignment = { horizontal: "center", vertical: "middle" };
  cc.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFD9E1F2" },
  };
  for (let r = 1; r <= 3; r++) {
    for (let c = 1; c <= 7; c++) {
      pivotSheet.getCell(r, c).border = {
        top: { style: "double" },
        bottom: { style: "double" },
        left: { style: "double" },
        right: { style: "double" },
      };
    }
  }

  pivotSheet.getCell("A4").value = operator;
  pivotSheet.getCell("A4").font = { bold: true, size: 11 };
  pivotSheet.getCell("A5").value = `Export Time: ${currentDate} ${currentTime}`;
  pivotSheet.getCell("A6").value = `Time Period: ${minDate} - ${maxDate}`;

  const employeeSummary = {};
  results.forEach((r) => {
    if (!employeeSummary[r.Employee]) {
      employeeSummary[r.Employee] = {
        records: [],
        isSanteh: r.Department && r.Department.toUpperCase().includes("SANTEH"),
        isKflStaff:
          r.Department && r.Department.toUpperCase().includes("KFL-STAFF"),
      };
    }
    const rd = { Date: r.Date };
    rd.CheckIn = r.CheckIn || r["Check In"] || "-";
    rd.CheckOut = r.CheckOut || r["Check Out"] || "-";
    if (employeeSummary[r.Employee].isSanteh) {
      for (let i = 1; i <= maxBreaks; i++) {
        const ord =
          i === 1 ? "1st" : i === 2 ? "2nd" : i === 3 ? "3rd" : `${i}th`;
        rd[`BreakIn${i}`] = r[`BreakIn${i}`] || r[`${ord} Break In`] || "-";
        rd[`BreakOut${i}`] = r[`BreakOut${i}`] || r[`${ord} Break Out`] || "-";
      }
    } else if (employeeSummary[r.Employee].isKflStaff) {
      rd.BreakIn1 = r.BreakIn1 || r["Noon Break In"] || "-";
      rd.BreakOut1 = r.BreakOut1 || r["Noon Break Out"] || "-";
    }
    employeeSummary[r.Employee].records.push(rd);
  });

  let pr = 8;
  Object.entries(employeeSummary).forEach(([emp, ed]) => {
    const sorted = ed.records.sort(
      (a, b) => parseDate(a.Date) - parseDate(b.Date),
    );

    pivotSheet.mergeCells(pr, 1, pr, 7);
    const ec = pivotSheet.getCell(pr, 1);
    ec.value = emp;
    ec.font = { bold: true, size: 12, color: { argb: "FFFFFFFF" } };
    ec.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FF4472C4" },
    };
    pr++;

    let hc = [];
    if (ed.isSanteh) {
      hc = ["Date", "Check In"];
      for (let i = 1; i <= maxBreaks; i++) {
        const o = i === 1 ? "st" : i === 2 ? "nd" : i === 3 ? "rd" : "th";
        hc.push(`${i}${o} Break In`, `${i}${o} Break Out`);
      }
      hc.push("Check Out");
    } else if (ed.isKflStaff) {
      hc = ["Date", "Check In", "Noon Break In", "Noon Break Out", "Check Out"];
    } else {
      hc = ["Date", "CheckIn", "CheckOut"];
    }

    hc.forEach((h, i) => {
      const c = pivotSheet.getCell(pr, i + 1);
      c.value = h;
      c.font = { bold: true, size: 10 };
      c.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFD9E1F2" },
      };
      c.alignment = { horizontal: "center", vertical: "middle" };
      c.border = {
        top: { style: "thin" },
        bottom: { style: "thin" },
        left: { style: "thin" },
        right: { style: "thin" },
      };
    });
    pr++;

    sorted.forEach((rec) => {
      let dr = [];
      if (ed.isSanteh) {
        dr = [rec.Date, rec.CheckIn];
        for (let i = 1; i <= maxBreaks; i++) {
          dr.push(rec[`BreakIn${i}`] || "-", rec[`BreakOut${i}`] || "-");
        }
        dr.push(rec.CheckOut);
      } else if (ed.isKflStaff) {
        dr = [
          rec.Date,
          rec.CheckIn,
          rec.BreakIn1 || "-",
          rec.BreakOut1 || "-",
          rec.CheckOut,
        ];
      } else {
        dr = [rec.Date, rec.CheckIn, rec.CheckOut];
      }
      dr.forEach((v, i) => {
        const c = pivotSheet.getCell(pr, i + 1);
        c.value = v;
        c.alignment = { horizontal: "center", vertical: "middle" };
        c.border = {
          top: { style: "thin" },
          bottom: { style: "thin" },
          left: { style: "thin" },
          right: { style: "thin" },
        };
        // Highlight missing
        if (v === "-") {
          c.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FFFFE5E5" },
          };
          c.font = { color: { argb: "FFC00000" }, bold: true };
        }
      });
      pr++;
    });
    pr++;
  });

  pivotSheet.columns = Array(7).fill({ width: 15 });

  // ============================================================
  // SHEET 3: Attendance Data (existing, enhanced)
  // ============================================================
  const ds = wb.addWorksheet("Attendance Data", {
    views: [{ state: "frozen", ySplit: 8 }],
    properties: { tabColor: { argb: "FFFFBB33" } },
  });
  ds.mergeCells(1, 1, 3, lastColIndex + 1);
  const dt = ds.getCell("A1");
  dt.value = "KFL MANPOWER AGENCY SERVER 3";
  dt.font = { bold: true, size: 16, color: { argb: "FF1F4E78" } };
  dt.alignment = { horizontal: "center", vertical: "middle" };
  dt.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFD9E1F2" },
  };

  ds.getCell("A4").value = operator;
  ds.getCell("A5").value = `Export Time: ${currentDate} ${currentTime}`;
  ds.getCell("A6").value = `Time Period: ${minDate} - ${maxDate}`;

  headers.forEach((h, i) => {
    const c = ds.getCell(8, i + 1);
    c.value = h;
    c.font = { bold: true, color: { argb: "FFFFFFFF" }, size: 11 };
    c.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FF305496" },
    };
    c.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
    c.border = {
      top: { style: "thin" },
      bottom: { style: "thin" },
      left: { style: "thin" },
      right: { style: "thin" },
    };
  });

  let dr2 = 9;
  results.forEach((r) => {
    let vals = [];
    if (hasBreakColumns) {
      const keys = [
        "Employee",
        "Department",
        "Status",
        "Duration",
        "Date",
        "CheckIn",
      ];
      if (headers.some((h) => h.includes("1st Break In"))) {
        for (let i = 1; i <= maxBreaks; i++)
          keys.push(`BreakIn${i}`, `BreakOut${i}`);
      } else {
        keys.push("BreakIn1", "BreakOut1");
      }
      keys.push("CheckOut");
      vals = keys.map((k) => r[k] || "-");
    } else {
      vals = [
        "Employee",
        "Department",
        "Status",
        "Duration",
        "Date",
        "CheckIn",
        "CheckOut",
      ].map((k) => r[k] || "-");
    }
    vals.push(r.Remarks || "Normal");

    vals.forEach((v, i) => {
      const c = ds.getCell(dr2, i + 1);
      c.value = v;
      c.alignment = { horizontal: "left", vertical: "middle" };
      c.border = {
        top: { style: "thin" },
        bottom: { style: "thin" },
        left: { style: "thin" },
        right: { style: "thin" },
      };
      if (v === "Missing Check Out") {
        c.font = { color: { argb: "FFC00000" }, bold: true };
        c.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FFFFE5E5" },
        };
      } else if (v === "Missing Check In") {
        c.font = { color: { argb: "FFE67E22" }, bold: true };
        c.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FFFFF3E0" },
        };
      } else if (v === "Day Shift") {
        c.font = { color: { argb: "FF375623" }, bold: true };
      } else if (v === "Night Shift") {
        c.font = { color: { argb: "FF1F4E78" }, bold: true };
      }
    });
    dr2++;
  });

  ds.columns = Array(headers.length)
    .fill(null)
    .map((_, i) => ({ width: i === 1 ? 40 : 15 }));

  ds.autoFilter = {
    from: { row: 8, column: 1 },
    to: { row: 8, column: headers.length },
  };

  // ============================================================
  // SHEET 4: Original Data
  // ============================================================
  if (window.originalWorksheet) {
    const os = wb.addWorksheet("Original Data", {
      properties: { tabColor: { argb: "FF999999" } },
    });
    const od = XLSX.utils.sheet_to_json(window.originalWorksheet, {
      header: 1,
    });
    od.forEach((row, ri) => {
      (row || []).forEach((v, ci) => {
        const c = os.getCell(ri + 1, ci + 1);
        c.value = v;
        if (ri === 7) {
          c.font = { bold: true, color: { argb: "FFFFFFFF" } };
          c.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FF305496" },
          };
        }
      });
    });
    os.columns = [
      { width: 12 },
      { width: 25 },
      { width: 30 },
      { width: 12 },
      { width: 12 },
      { width: 20 },
    ];
  }

  // ---------- Save ----------
  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  saveAs(
    blob,
    `attendance_report_${currentDate.replace(/-/g, "")}_${currentTime.replace(
      /:/g,
      "",
    )}.xlsx`,
  );
  showToast("📊 Excel exported with Summary + Employee Summary sheets!");
}

// ============================================================
// EXPORT PDF — Enhanced with logo, stats, summary, footer
// ============================================================
async function exportToPDF() {
  const output = document.getElementById("output");
  const table = output.querySelector("table");

  if (!table) {
    showToast("No data to export.");
    return;
  }

  try {
    // ---------- Extract data from HTML table ----------
    const headers = Array.from(table.querySelectorAll("th")).map((th) =>
      th.textContent.replace(/[⇅]/g, "").trim(),
    );
    const rows = Array.from(table.querySelectorAll("tr")).slice(1);

    const results = rows.map((row) => {
      const cells = Array.from(row.querySelectorAll("td")).map((td) =>
        td.textContent.trim(),
      );
      const obj = {};
      headers.forEach((h, i) => {
        if (h === "Noon Break In") obj["BreakIn1"] = cells[i];
        else if (h === "Noon Break Out") obj["BreakOut1"] = cells[i];
        else obj[h] = cells[i];
      });
      return obj;
    });

    if (results.length === 0) {
      showToast("No data to export.");
      return;
    }

    // ---------- Metadata ----------
    const dates = results
      .map((r) => r.Date)
      .filter(Boolean)
      .sort((a, b) => parseDate(a) - parseDate(b));
    const minDate = dates[0] || "";
    const maxDate = dates[dates.length - 1] || "";

    const now = new Date();
    const currentDate = `${String(now.getDate()).padStart(2, "0")}-${String(
      now.getMonth() + 1,
    ).padStart(2, "0")}-${now.getFullYear()}`;
    const currentTime = `${String(now.getHours()).padStart(2, "0")}:${String(
      now.getMinutes(),
    ).padStart(2, "0")}`;
    const operator = window.currentOperator || "";

    const maxBreaks = window.maxBreaksFound || 2;

    // ---------- Create PDF ----------
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({
      orientation: "landscape",
      unit: "pt",
      format: "a4",
    });

    const pageWidth = doc.internal.pageSize.getWidth();
    const pageHeight = doc.internal.pageSize.getHeight();
    const margin = 28;

    // ---------- Colors ----------
    const COLORS = {
      titleText: [31, 78, 120],
      titleBg: [217, 225, 242],
      headerBg: [48, 84, 150],
      employeeBar: [68, 114, 196],
      subHeaderBg: [217, 225, 242],
      gray: [100, 100, 100],
      black: [0, 0, 0],
      white: [255, 255, 255],
      dayGreen: [55, 86, 35],
      nightBlue: [31, 78, 120],
      morningAmber: [191, 143, 0],
      missingRed: [192, 0, 0],
      missingInOrange: [230, 126, 34],
    };

    // ---------- Load logo (async, with fallback) ----------
    let logoData = null;
    try {
      logoData = await loadImageAsDataURL("logo1.png");
    } catch (e) {
      console.warn("Logo not available:", e);
    }

    // ---------- Helper: draw header banner ----------
    function drawBanner(subtitle) {
      // Title bar
      doc.setFillColor(...COLORS.titleBg);
      doc.rect(0, 0, pageWidth, 48, "F");

      // Logo (left side)
      if (logoData) {
        try {
          doc.addImage(logoData, "PNG", margin, 6, 36, 36);
        } catch (e) {}
      }

      doc.setTextColor(...COLORS.titleText);
      doc.setFont("helvetica", "bold");
      doc.setFontSize(15);
      doc.text("KFL MANPOWER AGENCY SERVER 3", pageWidth / 2, 20, {
        align: "center",
      });

      doc.setFontSize(10);
      doc.setFont("helvetica", "normal");
      doc.text(subtitle, pageWidth / 2, 38, { align: "center" });

      // Metadata row
      let metaY = 62;
      doc.setTextColor(...COLORS.black);
      doc.setFontSize(9);

      if (operator) {
        doc.setFont("helvetica", "bold");
        doc.text(operator, margin, metaY);
        metaY += 12;
      }

      doc.setFont("helvetica", "normal");
      doc.setTextColor(...COLORS.gray);
      doc.text(`Export Time: ${currentDate} ${currentTime}`, margin, metaY);
      if (minDate && maxDate) {
        doc.text(`Time Period: ${minDate} - ${maxDate}`, margin + 240, metaY);
      }
      doc.text(`Total: ${results.length} records`, pageWidth - margin, metaY, {
        align: "right",
      });

      return metaY + 16;
    }

    // ============================================================
    // PAGE 1: SUMMARY STATS
    // ============================================================
    let yPos = drawBanner("Attendance Summary Report");

    // Compute stats
    const uniqueEmployees = [...new Set(results.map((r) => r.Employee))];
    const statusCounts = {
      "Day Shift": 0,
      "Night Shift": 0,
      "Morning Shift": 0,
      "Missing Check Out": 0,
      "Missing Check In": 0,
    };
    const durationToHours = (str) => {
      if (!str || str === "-") return 0;
      const m = String(str).match(/(\d+)h\s*(\d+)?m?/);
      if (!m) return 0;
      return parseInt(m[1]) + (parseInt(m[2]) || 0) / 60;
    };
    let totalHours = 0;
    results.forEach((r) => {
      if (statusCounts[r.Status] !== undefined) statusCounts[r.Status]++;
      totalHours += durationToHours(r.Duration);
    });

    // ---- Overview card ----
    doc.setFillColor(...COLORS.headerBg);
    doc.rect(margin, yPos, pageWidth - margin * 2, 18, "F");
    doc.setTextColor(...COLORS.white);
    doc.setFont("helvetica", "bold");
    doc.setFontSize(11);
    doc.text("OVERVIEW", margin + 8, yPos + 12);
    yPos += 22;

    const overviewRows = [
      ["Total Employees", String(uniqueEmployees.length)],
      ["Total Records", String(results.length)],
      ["Total Hours Rendered", `${Math.round(totalHours * 100) / 100} hours`],
      [
        "Average per Record",
        `${
          results.length
            ? Math.round((totalHours / results.length) * 100) / 100
            : 0
        } hours`,
      ],
      ["Date Range", `${minDate}  →  ${maxDate}`],
    ];

    doc.setFontSize(10);
    overviewRows.forEach(([label, value], i) => {
      const bg = i % 2 === 0 ? [245, 248, 252] : [255, 255, 255];
      doc.setFillColor(...bg);
      doc.rect(margin, yPos, pageWidth - margin * 2, 16, "F");
      doc.setDrawColor(220, 220, 220);
      doc.setLineWidth(0.3);
      doc.rect(margin, yPos, pageWidth - margin * 2, 16);
      doc.setTextColor(...COLORS.black);
      doc.setFont("helvetica", "bold");
      doc.text(label, margin + 8, yPos + 11);
      doc.setFont("helvetica", "normal");
      doc.text(value, pageWidth - margin - 8, yPos + 11, { align: "right" });
      yPos += 16;
    });

    yPos += 12;

    // ---- Shift breakdown ----
    doc.setFillColor(...COLORS.headerBg);
    doc.rect(margin, yPos, pageWidth - margin * 2, 18, "F");
    doc.setTextColor(...COLORS.white);
    doc.setFont("helvetica", "bold");
    doc.setFontSize(11);
    doc.text("SHIFT BREAKDOWN", margin + 8, yPos + 12);
    yPos += 22;

    const statusColors = {
      "Day Shift": COLORS.dayGreen,
      "Night Shift": COLORS.nightBlue,
      "Morning Shift": COLORS.morningAmber,
      "Missing Check Out": COLORS.missingRed,
      "Missing Check In": COLORS.missingInOrange,
    };

    doc.setFontSize(10);
    Object.entries(statusCounts).forEach(([label, count], i) => {
      const bg = i % 2 === 0 ? [245, 248, 252] : [255, 255, 255];
      doc.setFillColor(...bg);
      doc.rect(margin, yPos, pageWidth - margin * 2, 16, "F");
      doc.setDrawColor(220, 220, 220);
      doc.rect(margin, yPos, pageWidth - margin * 2, 16);

      doc.setTextColor(...(statusColors[label] || COLORS.black));
      doc.setFont("helvetica", "bold");
      doc.text(label, margin + 8, yPos + 11);

      doc.setTextColor(...COLORS.black);
      doc.setFont("helvetica", "normal");
      const pct = results.length
        ? ((count / results.length) * 100).toFixed(1)
        : "0.0";
      doc.text(`${count}  (${pct}%)`, pageWidth - margin - 8, yPos + 11, {
        align: "right",
      });
      yPos += 16;
    });

    // ============================================================
    // PAGE 2+: ATTENDANCE PIVOT
    // ============================================================
    doc.addPage("a4", "landscape");
    yPos = drawBanner("Attendance Pivot");

    // Group results per employee
    const employeeSummary = {};
    results.forEach((r) => {
      if (!employeeSummary[r.Employee]) {
        employeeSummary[r.Employee] = {
          records: [],
          department: r.Department,
          isSanteh:
            r.Department && r.Department.toUpperCase().includes("SANTEH"),
          isKflStaff:
            r.Department && r.Department.toUpperCase().includes("KFL-STAFF"),
        };
      }
      const record = {
        Date: r.Date,
        CheckIn: r.CheckIn || r["Check In"] || "-",
        CheckOut: r.CheckOut || r["Check Out"] || "-",
        BreakIn1: r.BreakIn1 || r["Noon Break In"] || "-",
        BreakOut1: r.BreakOut1 || r["Noon Break Out"] || "-",
      };
      for (let i = 1; i <= maxBreaks; i++) {
        record[`BreakIn${i}`] = r[`BreakIn${i}`] || "-";
        record[`BreakOut${i}`] = r[`BreakOut${i}`] || "-";
      }
      employeeSummary[r.Employee].records.push(record);
    });

    const employeeNames = Object.keys(employeeSummary).sort((a, b) =>
      a.localeCompare(b),
    );

    employeeNames.forEach((employee, idx) => {
      const empData = employeeSummary[employee];
      const sortedRecords = empData.records.sort(
        (a, b) => parseDate(a.Date) - parseDate(b.Date),
      );

      let subHeaders;
      if (empData.isSanteh) {
        subHeaders = ["Date", "Check In"];
        for (let i = 1; i <= maxBreaks; i++) {
          const ord = i === 1 ? "st" : i === 2 ? "nd" : i === 3 ? "rd" : "th";
          subHeaders.push(`${i}${ord} Break In`, `${i}${ord} Break Out`);
        }
        subHeaders.push("Check Out");
      } else if (empData.isKflStaff) {
        subHeaders = [
          "Date",
          "Check In",
          "Noon Break In",
          "Noon Break Out",
          "Check Out",
        ];
      } else {
        subHeaders = ["Date", "CheckIn", "CheckOut"];
      }

      const bodyRows = sortedRecords.map((rec) => {
        if (empData.isSanteh) {
          const row = [rec.Date, rec.CheckIn];
          for (let i = 1; i <= maxBreaks; i++) {
            row.push(rec[`BreakIn${i}`] || "-", rec[`BreakOut${i}`] || "-");
          }
          row.push(rec.CheckOut);
          return row;
        } else if (empData.isKflStaff) {
          return [
            rec.Date,
            rec.CheckIn,
            rec.BreakIn1 || "-",
            rec.BreakOut1 || "-",
            rec.CheckOut,
          ];
        } else {
          return [rec.Date, rec.CheckIn, rec.CheckOut];
        }
      });

      // Estimate block height: employee bar + head + rows
      const estimatedHeight = 18 + 22 + bodyRows.length * 16 + 14;

      // Page-break if the block won't fit
      if (yPos + estimatedHeight > pageHeight - 40) {
        doc.addPage("a4", "landscape");
        yPos = drawBanner("Attendance Pivot (continued)");
      }

      // Employee name bar
      doc.setFillColor(...COLORS.employeeBar);
      doc.rect(margin, yPos, pageWidth - margin * 2, 18, "F");
      doc.setTextColor(...COLORS.white);
      doc.setFont("helvetica", "bold");
      doc.setFontSize(10);
      doc.text(employee, margin + 6, yPos + 12);

      // Right-side meta on employee bar
      doc.setFontSize(8.5);
      doc.setFont("helvetica", "normal");
      const totalHrs = sortedRecords
        .map((r) => 0) // Duration isn't in the pivot record, skip
        .reduce((a, b) => a + b, 0);
      doc.text(
        `${sortedRecords.length} day(s)`,
        pageWidth - margin - 6,
        yPos + 12,
        { align: "right" },
      );

      yPos += 20;

      doc.autoTable({
        head: [subHeaders],
        body: bodyRows,
        startY: yPos,
        margin: { left: margin, right: margin },
        theme: "grid",
        styles: {
          font: "helvetica",
          fontSize: 8.5,
          cellPadding: 3.5,
          textColor: COLORS.black,
          lineColor: [180, 180, 180],
          lineWidth: 0.4,
          valign: "middle",
          halign: "center",
        },
        headStyles: {
          fillColor: COLORS.subHeaderBg,
          textColor: COLORS.black,
          fontStyle: "bold",
          halign: "center",
          fontSize: 8.5,
        },
        alternateRowStyles: {
          fillColor: [248, 250, 253],
        },
        didParseCell: function (data) {
          // Highlight missing cells
          if (data.section === "body") {
            const v = String(data.cell.raw || "").trim();
            if (v === "-") {
              data.cell.styles.textColor = COLORS.missingRed;
              data.cell.styles.fontStyle = "bold";
            }
          }
        },
        didDrawPage: function (data) {
          if (data.pageNumber > 1) {
            // Continuation header
            doc.setFillColor(...COLORS.titleBg);
            doc.rect(0, 0, pageWidth, 20, "F");
            doc.setTextColor(...COLORS.titleText);
            doc.setFont("helvetica", "bold");
            doc.setFontSize(10);
            doc.text(
              "KFL MANPOWER AGENCY SERVER 3 — Attendance Pivot (cont.)",
              pageWidth / 2,
              13,
              { align: "center" },
            );
          }
        },
      });

      yPos = doc.lastAutoTable.finalY + 14;
    });

    // ============================================================
    // LAST PAGE: SIGNATURE BLOCK + FOOTER
    // ============================================================
    if (yPos + 80 > pageHeight - 30) {
      doc.addPage("a4", "landscape");
      yPos = 60;
    }

    doc.setDrawColor(...COLORS.titleText);
    doc.setLineWidth(0.8);
    doc.line(margin, yPos + 30, margin + 200, yPos + 30);
    doc.setFontSize(9);
    doc.setTextColor(...COLORS.gray);
    doc.setFont("helvetica", "normal");
    doc.text("Prepared By (Operator)", margin, yPos + 44);

    doc.line(pageWidth / 2 - 100, yPos + 30, pageWidth / 2 + 100, yPos + 30);
    doc.text("Verified By", pageWidth / 2, yPos + 44, { align: "center" });

    doc.line(
      pageWidth - margin - 200,
      yPos + 30,
      pageWidth - margin,
      yPos + 30,
    );
    doc.text("Approved By", pageWidth - margin, yPos + 44, {
      align: "right",
    });

    // ============================================================
    // FOOTER on every page
    // ============================================================
    const pageCount = doc.internal.getNumberOfPages();
    for (let i = 1; i <= pageCount; i++) {
      doc.setPage(i);
      doc.setDrawColor(...COLORS.titleText);
      doc.setLineWidth(0.8);
      doc.line(margin, pageHeight - 22, pageWidth - margin, pageHeight - 22);

      doc.setFontSize(8);
      doc.setTextColor(...COLORS.gray);
      doc.setFont("helvetica", "normal");
      doc.text(
        `KFL Manpower Agency  |  Generated ${currentDate} ${currentTime}  |  ${
          operator || "—"
        }`,
        margin,
        pageHeight - 10,
      );

      doc.text(
        `Page ${i} of ${pageCount}`,
        pageWidth - margin,
        pageHeight - 10,
        { align: "right" },
      );
    }

    // ---------- Save ----------
    const fileName = `attendance_report_${currentDate.replace(
      /-/g,
      "",
    )}_${currentTime.replace(/:/g, "")}.pdf`;
    doc.save(fileName);
    showToast("📄 PDF exported with summary + signature block!");
  } catch (err) {
    console.error("Error exporting PDF:", err);
    showToast("Error exporting PDF: " + err.message);
  }
}

// ---------- Helper: load image as data URL ----------
function loadImageAsDataURL(url) {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.crossOrigin = "anonymous";
    img.onload = () => {
      try {
        const canvas = document.createElement("canvas");
        canvas.width = img.width;
        canvas.height = img.height;
        canvas.getContext("2d").drawImage(img, 0, 0);
        resolve(canvas.toDataURL("image/png"));
      } catch (e) {
        reject(e);
      }
    };
    img.onerror = () => reject(new Error("Image load failed"));
    img.src = url;
  });
}
