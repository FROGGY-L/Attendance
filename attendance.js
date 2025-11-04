document.getElementById('fileInput').addEventListener('change', function(event) {
    const file = event.target.files[0];
    const fileNameSpan = document.getElementById('fileName');
    
    if (file) {
        fileNameSpan.textContent = `Selected File: ${file.name}`;
    } else {
        showToast('No File Chosen');
        fileNameSpan.textContent = '';
    }
});
function clearFile() {
    const fileInput = document.getElementById('fileInput');
    const fileNameSpan = document.getElementById('fileName');
    const output = document.getElementById('output');

    fileInput.value = '';
    fileNameSpan.textContent = '';
    output.innerHTML = '';
    showToast('Cleared');
}

function processFile() {
    const fileInput = document.getElementById('fileInput');
    const file = fileInput.files[0];

    if (!file) {
        showToast('Please select an Excel file.');
        return;
    }

    const reader = new FileReader();
    reader.onload = (e) => {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });

        // Assuming the first sheet contains the data
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];

        // Read the first row of the sheet to check the content
        const firstRow = XLSX.utils.sheet_to_json(worksheet, { range: 0, header: 1 })[0];

        // Check if the first row contains "KFL Manpower Agency"
        function handleFileUpload(file) {
            const reader = new FileReader();

            reader.onload = function (event) {
                const text = event.target.result;
                const rows = text.split("\n").map(row => row.split(",")); // Splitting into rows and columns

                if (rows.length < 3) {
                    showToast("Not enough rows in the file.");
                    return;
                }

                // Checking first three rows for "KFL Manpower Agency" in the fourth column (index 3)
                const isValid = rows.slice(0, 3).every(row => row[3] === "KFL Manpower Agency");

                if (!isValid) {
                    showToast("This is not a valid KFL Attendance file.");
                    return;
                }

                console.log("Valid KFL Attendance file");
            };

            reader.readAsText(file);
        }


        // Convert the worksheet to JSON, starting from row 8
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { range: 7 }); // 0-based index, row 8 is index 7

        analyzeData(jsonData);
    };
    // toast if file is not valid
    reader.readAsArrayBuffer(file);
                }
                function showToast(message) {
        const toast = document.createElement('div');
        toast.className = 'toastfile';
        toast.textContent = message;
        
        document.body.appendChild(toast);
        
        // Show the toast and hide it after 3 seconds
        toast.style.display = 'block';
        setTimeout(() => {
            toast.style.display = 'none';
            document.body.removeChild(toast);
        }, 3000);
    }

    // Important notice
    function showToastMessage() {
        const toastMessage = `
       <strong>Important Notice:</strong><br><br>
        Please ensure you double-check the file provided by the system.<br><br>
        <strong>Update 04-08-2025:</strong><br>
        - Fix issues on night shift.<br>
        <strong>Update 04-16-2025:</strong><br>
        - Ensure check-out from next day is included even if reused for night shift.<br>
        <strong>Update 04-24-2025:</strong><br>
        - Fixed night shift and morning shift also added mid shift for calculating. Also reduced time on Important Notice.<br>
        <strong>Update 04-30-2025:</strong><br>
        - Added missing checkin. <br>
        <strong>* Always double-check the data.</strong><br><br>
        For any inquiries, feel free to contact IT Personnel.<br><br>
        <strong>Thank you!</strong>
        `;
    
        const toast = document.createElement('div');
        toast.classList.add('toast-message');
        toast.innerHTML = toastMessage;
    
        // Style the toast message (you can modify this style as per your preference)
        toast.style.position = 'fixed';
        toast.style.top = '50%'; // Center vertically
        toast.style.left = '50%'; // Center horizontally
        toast.style.transform = 'translate(-50%, -50%)'; // Ensure exact center
        toast.style.padding = '20px 30px';
        toast.style.backgroundColor = '#f44336'; // Red background for warning
        toast.style.color = '#fff';
        toast.style.borderRadius = '10px'; // More rounded corners
        toast.style.fontSize = '16px';
        toast.style.textAlign = 'center'; // Center align the text
        toast.style.zIndex = '9999';
        toast.style.lineHeight = '1.5'; // Add space between lines for better readability
    
        // Append the toast to the body
        document.body.appendChild(toast);
    
        // Hide the toast after 5 seconds
        setTimeout(() => {
            toast.style.display = 'none';
        }, 3000);
    }
    
    // Ensure the toast shows when the page is fully loaded
    window.addEventListener('DOMContentLoaded', () => {
        showToastMessage();
    });
    

    function combineDateTime(dateStr, timeStr) {
        if (!dateStr || !timeStr) return null;
        
        try {
            // Parse date in DD-MM-YYYY format
            const [day, month, year] = dateStr.split('-').map(Number);
            
            // Parse time in HH:mm format
            const [hours, minutes] = timeStr.split(':').map(Number);
            
            // Validate all components
            if (isNaN(year) || isNaN(month) || isNaN(day) || 
                isNaN(hours) || isNaN(minutes)) {
                console.warn(`Invalid date/time components: ${dateStr} ${timeStr}`);
                return null;
            }
    
            // Create Date object (months are 0-based in JavaScript)
            const date = new Date(year, month - 1, day, hours, minutes);
            
            if (isNaN(date.getTime())) {
                console.warn(`Invalid date/time: ${dateStr} ${timeStr}`);
                return null;
            }
            
            return date;
        } catch (e) {
            console.warn(`Error combining date/time: ${dateStr} ${timeStr}`, e);
            return null;
        }
    }
    
    // time in and out analyze
    function analyzeData(data) {
        if (data.length === 0) {
            showToast('The uploaded file contains no data.');
            return;
        }
    
        const results = [];
        const groupedData = {};
    
        // Group by ID + Date
        data.forEach(record => {
            const id = record['ID'];
            const name = record['Name'];
            const department = record['Department'];
            const date = record['Date'];
            const time = record['Time'] || record['Check-In Time'];
            const dateTime = combineDateTime(date, time);
            const type = record['Card Swiping Type'];
    
            if (!dateTime) {
                console.warn(`⛔ Skipping invalid datetime: ${date} ${time}`);
                return;
            }            
    
            const key = `${id}_${date}`;
            if (!groupedData[key]) {
                groupedData[key] = {
                    records: [],
                    name,
                    department,
                    date,
                    id,
                };
            }
    
            groupedData[key].records.push({ type, dateTime, time });
        });
    
        function parseDMY(dateStr) {
            const [day, month, year] = dateStr.split('-');
            return new Date(`${year}-${month.padStart(2, '0')}-${day.padStart(2, '0')}`);
        }        
    
        const sortedKeys = Object.keys(groupedData).sort((a, b) => {
            const [idA, dateA] = a.split('_');
            const [idB, dateB] = b.split('_');
            return parseDMY(dateA) - parseDMY(dateB) || idA.localeCompare(idB);
        });
    
        let reusedCheckOuts = new Set();
    
        for (let i = 0; i < sortedKeys.length; i++) {
            const key = sortedKeys[i];
            const { records, name, department, date, id } = groupedData[key];
    
            const checkIns = [], checkOuts = [], breakIns = [], breakOuts = [];
    
            records.forEach(r => {
                if (r.type === 'Check In') checkIns.push(r);
                else if (r.type === 'Check Out') checkOuts.push(r);
                else if (r.type === 'Break In' && department === 'All Departments>SANTEH') breakIns.push(r.time);
                else if (r.type === 'Break Out' && department === 'All Departments>SANTEH') breakOuts.push(r.time);
            });
    
            const firstCheckIn = checkIns.length > 0 ? checkIns.reduce((a, b) => a.dateTime < b.dateTime ? a : b) : null;
            let lastCheckOut = checkOuts.length > 0 ? checkOuts.reduce((a, b) => (a.dateTime > b.dateTime ? a : b)) : null;
            let remarks = 'Normal';
    
            // Check for missing check-in
            if (!firstCheckIn && checkOuts.length > 0) {
                // This is a case where there are check-outs but no check-ins
                const earliestCheckOut = checkOuts.reduce((a, b) => a.dateTime < b.dateTime ? a : b);
                
                results.push({
                    Employee: name,
                    Department: department,
                    Status: 'Missing Check In',
                    Duration: '-',
                    Date: date,
                    CheckIn: '-',
                    CheckOut: earliestCheckOut.time,
                    BreakIn1: '-',
                    BreakOut1: '-',
                    BreakIn2: '-',
                    BreakOut2: '-',
                    BreakIn3: '-',
                    BreakOut3: '-',
                    Remarks: 'Check In Missing'
                });
                continue; // Skip to next record
            }
    
            if (!firstCheckIn) continue; // Skip if no check-in and no check-out
    
            const checkInDateTime = new Date(firstCheckIn.dateTime);
            const checkInHour = checkInDateTime.getHours();
            const shiftType = checkInHour >= 17 || checkInHour < 5 ? 'Night Shift' :
                              checkInHour >= 5 && checkInHour < 13 ? 'Morning Shift' : 'Mid Shift';
    
            // Validate same-day check-out
            if (lastCheckOut) {
                const outTime = lastCheckOut.dateTime;
                console.log("🔍 Raw outTime:", outTime);
    
                // Attempt to parse the outTime
                let parsedOutTime = new Date(outTime);
    
                // Check if parsing failed
                if (isNaN(parsedOutTime.getTime())) {
                    console.warn(`❌ Invalid outTime for ${name} on ${date}:`, outTime);
                } else {
                    const diff = (parsedOutTime - checkInDateTime) / (1000 * 60 * 60);
                    console.log(`   ⏱️ Trying out at ${outTime} → diff = ${diff.toFixed(2)} hrs`);
    
                    // Logic for checking time difference
                    if (diff <= 0 || diff > 16) {
                        lastCheckOut = null;
                    } else {
                        reusedCheckOuts.add(outTime);
                    }
                }
            }
    
            // 🔁 Try next-day checkout if night shift and no same-day checkout found
            if (!lastCheckOut && shiftType === 'Night Shift') {
                console.log(`🕒 Looking for next-day checkout for ${name} on ${date}...`);
    
                // Helper: Get next date in DD-MM-YYYY format
                function getNextDateStr(d) {
                    const [day, month, year] = d.split('-').map(Number);
                    const current = new Date(year, month - 1, day);
                    current.setDate(current.getDate() + 1);
                    const d2 = current.getDate().toString().padStart(2, '0');
                    const m2 = (current.getMonth() + 1).toString().padStart(2, '0');
                    const y2 = current.getFullYear();
                    return `${d2}-${m2}-${y2}`;
                }
    
                const nextDateStr = getNextDateStr(date);
                const nextKey = `${id}_${nextDateStr}`;
                const nextGroup = groupedData[nextKey];
    
                if (nextGroup) {
                    const possibleOuts = nextGroup.records.filter(r =>
                        r.type === 'Check Out' && !reusedCheckOuts.has(r.dateTime)
                    );
    
                    for (const r of possibleOuts) {
                        // r.dateTime is already a Date object, no need to parse
                        const outTime = r.dateTime;
                        
                        if (!(outTime instanceof Date) || isNaN(outTime.getTime())) {
                            console.warn(`❌ Invalid outTime:`, outTime);
                            continue;
                        }
    
                        const diffHours = (outTime - checkInDateTime) / (1000 * 60 * 60);
                        console.log(`   ⏱️ Trying out at ${outTime.toString()} → diff = ${diffHours.toFixed(2)} hrs`);
    
                        if (diffHours > 0 && diffHours <= 14) {
                            lastCheckOut = r;
                            reusedCheckOuts.add(r.dateTime);
                            remarks = `Reused check-out from ${nextGroup.date}`;
                            console.log(`✅ Found valid next-day check-out: ${outTime.toString()}`);
                            break;
                        }
                    }
                } else {
                    console.log(`📭 No next-day record group found for ${nextKey}`);
                }
    
                if (!lastCheckOut) {
                    console.log(`❌ No valid next-day check-out found for ${name} on ${date}`);
                }
            }
             
    
            const checkOutDateTime = lastCheckOut ? new Date(lastCheckOut.dateTime) : null;
            let duration = '-';
            let status = shiftType;
            let checkOutTime = '-';
    
            if (checkOutDateTime) {
                const hours = (checkOutDateTime - checkInDateTime) / (1000 * 60 * 60);
                if (hours > 0 && hours <= 24) {
                    duration = `${Math.floor(hours)} hours ${Math.round((hours % 1) * 60)} minutes`;
                } else {
                    remarks = 'Not Normal';
                }
                checkOutTime = lastCheckOut.time;
            } else {
                status = 'Missing Check Out';
                remarks = 'Check Out Missing';
            }
    
            const formattedBreakIns = [...breakIns, '-', '-', '-'].slice(0, 3);
            const formattedBreakOuts = [...breakOuts, '-', '-', '-'].slice(0, 3);
    
            results.push({
                Employee: name,
                Department: department,
                Status: status,
                Duration: duration,
                Date: date,
                CheckIn: firstCheckIn.time,
                CheckOut: checkOutTime,
                BreakIn1: formattedBreakIns[0],
                BreakOut1: formattedBreakOuts[0],
                BreakIn2: formattedBreakIns[1],
                BreakOut2: formattedBreakOuts[1],
                BreakIn3: formattedBreakIns[2],
                BreakOut3: formattedBreakOuts[2],
                Remarks: remarks
            });
        }
    
        displayResults(results);
    }
    
   function displayResults(results) {
        const output = document.getElementById('output');
        output.innerHTML = '';

        // Create search bar
        const searchContainer = document.createElement('div');
        searchContainer.style.marginBottom = '10px';
        const searchBar = document.createElement('input');
        searchBar.type = 'text';
        searchBar.placeholder = 'Search...';
        searchBar.style.width = '20%';
        searchBar.style.padding = '8px';
        searchBar.style.marginBottom = '10px';
        searchBar.style.border = '1px solid #ccc';
        searchBar.style.borderRadius = '4px';
        searchBar.style.marginLeft = '0';
        searchBar.style.textAlign = 'left';
        searchBar.style.display = 'block';

        // Add search functionality
        searchBar.addEventListener('input', () => {
            const filter = searchBar.value.toLowerCase();
            const rows = table.querySelectorAll('tr:not(:first-child)'); // Exclude header row

            rows.forEach(row => {
                const cells = row.querySelectorAll('td');
                const match = Array.from(cells).some(cell => cell.textContent.toLowerCase().includes(filter));
                row.style.display = match ? '' : 'none';
            });
        });

        searchContainer.appendChild(searchBar);
        output.appendChild(searchContainer);

        if (!results || results.length === 0) {
            const errorMessage = document.createElement('p');
            errorMessage.textContent = showToast('Invalid file: No data to display.');
            errorMessage.style.color = 'red';
            errorMessage.style.fontWeight = 'bold';
            output.appendChild(errorMessage);
            return;
        }
        results.sort((a, b) => {
            // First compare by Employee name
            const nameCompare = a.Employee.localeCompare(b.Employee);
            if (nameCompare !== 0) return nameCompare;
            
            // If same employee, compare by Date
            const dateA = new Date(a.Date);
            const dateB = new Date(b.Date);
            return dateA - dateB;
        });

        const table = document.createElement('table');
        const headerRow = document.createElement('tr');

        const hasSANTEH = results.some(result => result.Department === 'All Departments>SANTEH');

        // Define headers based on department
        const headers = hasSANTEH
        ? ['Employee', 'Department', 'Status', 'Hours Rendered', 'Date', 'CheckIn', 'BreakIn1', 'BreakOut1','BreakIn2', 'BreakOut2','BreakIn3', 'BreakOut3', 'CheckOut', 'Reg Late', 'Break Late', 'Total Late', 'Reg OT', 'NDOT', 'ND', 'Remarks']
        : ['Employee', 'Department', 'Status', 'Hours Rendered', 'Date', 'CheckIn', 'CheckOut', 'Remarks'];

        // Add headers
        headers.forEach(header => {
            const th = document.createElement('th');
            th.textContent = header;
            headerRow.appendChild(th);
        });

        table.appendChild(headerRow);

        results.forEach(result => {
            const row = document.createElement('tr');

            // Parse and convert Duration to decimal format
            function parseDurationToDecimal(durationStr) {
                const hoursMatch = durationStr.match(/(\d+)\s*hours?/);
                const minutesMatch = durationStr.match(/(\d+)\s*minutes?/);

                const hours = hoursMatch ? parseInt(hoursMatch[1], 10) : 0;
                const minutes = minutesMatch ? parseInt(minutesMatch[1], 10) : 0;

                return hours + minutes / 60; // Convert minutes to fractional hours
            }

            if (hasSANTEH) {
                // Define all columns dynamically, ensuring unique keys for breaks
                const columns = ['Employee', 'Department', 'Status', 'Duration', 'Date', 'CheckIn',
                                'BreakIn1', 'BreakOut1', 'BreakIn2', 'BreakOut2', 'BreakIn3', 'BreakOut3', 'CheckOut'];

                columns.forEach(key => {
                    const td = document.createElement('td');
                    td.textContent = result[key] || '-'; // Ensure missing values are displayed as '-'
                    row.appendChild(td);
                });

                // Calculate Reg Late
                let regLate = 0;

                if (result.CheckIn) {
                    const checkInTime = new Date(`1970-01-01T${result.CheckIn}`);

                    // Determine shift start time
                    let shiftStartTime = checkInTime.getHours() >= 19 || (checkInTime.getHours() === 18 && checkInTime.getMinutes() >= 30)
                    ? new Date("1970-01-01T19:00:00")
                    : new Date("1970-01-01T07:00:00");
                    if (checkInTime.getHours() >= 19 || (checkInTime.getHours() === 18 && checkInTime.getMinutes() >= 30)) {
                        shiftStartTime = new Date(`1970-01-01T19:00:00`);
                        if (checkInTime.getHours() === 18 && checkInTime.getMinutes() >= 30) {
                            shiftStartTime = new Date(`1970-01-01T19:00:00`);
                        }

                        // 15-minute grace period
                        const gracePeriodEnd = new Date(shiftStartTime);
                        gracePeriodEnd.setMinutes(gracePeriodEnd.getMinutes() + 15);
                        if (checkInTime >= shiftStartTime && checkInTime <= gracePeriodEnd) {
                            shiftStartTime = new Date(`1970-01-01T19:00:00`);
                        }
                    } else {
                        shiftStartTime = new Date(`1970-01-01T07:00:00`);
                        if (checkInTime.getHours() === 6 && checkInTime.getMinutes() === 30) {
                            shiftStartTime = checkInTime;
                        }
                    }

                    // Calculate Regular Late
                    if (!isNaN(checkInTime) && checkInTime > shiftStartTime) {
                        const minutesLate = (checkInTime - shiftStartTime) / (1000 * 60);
                        if (checkInTime.getHours() >= 19 && minutesLate > 0) {
                            regLate = 0;
                        } else if (minutesLate > 5 && minutesLate <= 15) {
                            regLate = 0.5;
                        } else if (minutesLate > 15) {
                            regLate = Math.ceil(minutesLate / 60);
                        }
                    }
                }

                const regLateCell = document.createElement('td');
                regLateCell.textContent = regLate.toFixed(2);
                row.appendChild(regLateCell);

                // Function to validate breaks
                function validateBreakTime(breakIn, breakOut, maxMinutes, isFirstBreak = false, isThirdBreak = false) {
                    if (breakIn && breakOut) {
                        const breakInTime = new Date(`1970-01-01T${breakIn}`);
                        const breakOutTime = new Date(`1970-01-01T${breakOut}`);
                        const breakDuration = (breakOutTime - breakInTime) / (1000 * 60);
                
                        // Warn if first break is not in the AM
                        if (isFirstBreak && breakInTime.getHours() >= 12) {
                            console.warn(`First Break In (${breakIn}) should be in the AM!`);
                        }
                
                        // Warn if third break is not in the PM
                        if (isThirdBreak && breakInTime.getHours() < 12) {
                            console.warn(`Third Break In (${breakIn}) should be in the PM!`);
                        }
                
                        if (breakDuration > maxMinutes) {
                            return (breakDuration - maxMinutes) / 60; // Convert excess minutes to decimal hours
                        }
                    }
                    return 0;
                }

                function assignBreaks(record) {
                    const checkInTime = new Date(`1970-01-01T${record.CheckIn}`);
                    const checkInHour = checkInTime.getHours();
                    const isMorningShift = checkInHour < 12; // True if morning, False if night
                
                    if (isMorningShift) {
                        // Day Shift (AM Check-In)
                        // No auto-fill, just validate if data exists
                        validateBreakTime(record.BreakIn1, record.BreakOut1, 60, true, false);  // 1st Break
                        validateBreakTime(record.AMBreakIn, record.AMBreakOut, 60);             // 2nd Break
                        validateBreakTime(record.BreakIn2, record.BreakOut2, 60, false, true);  // 3rd Break
                    } else {
                        // Night Shift (PM Check-In)
                        validateBreakTime(record.BreakIn1, record.BreakOut1, 60, true, false);  // 1st Break
                        validateBreakTime(record.PMBreakIn, record.PMBreakOut, 60);             // 2nd Break
                        validateBreakTime(record.BreakIn2, record.BreakOut2, 60, false, true);  // 3rd Break
                    }
                
                    return record;
                }

                // Apply function to all records
                results = results.map(assignBreaks);


                // Apply to all records
                results = results.map(assignBreaks);

                // Calculate Break Late
                let breakLate = 0;
                breakLate += validateBreakTime(result.BreakIn1, result.BreakOut1, 15);
                breakLate += validateBreakTime(result.BreakIn2, result.BreakOut2, 60);
                breakLate += validateBreakTime(result.BreakIn3, result.BreakOut3, 15);

                const breakLateCell = document.createElement('td');
                breakLateCell.textContent = breakLate > 0 ? breakLate.toFixed(2) : '-';
                row.appendChild(breakLateCell);

                // Calculate Total Late
                let totalLate = regLate + breakLate;
                const totalLateCell = document.createElement('td');
                totalLateCell.textContent = totalLate > 0 ? totalLate.toFixed(2) : '-';
                row.appendChild(totalLateCell);

                // Calculate Reg OT and NDOT
                let regOT = 0;
                let ndot = 0;

                if (result.CheckIn && result.CheckOut) {
                    const checkIn = new Date(`1970-01-01T${result.CheckIn}`);
                    const checkOut = new Date(`1970-01-01T${result.CheckOut}`);

                    if (!isNaN(checkIn) && !isNaN(checkOut)) {
                        const nightStart = new Date('1970-01-01T22:00');
                        const nightEnd = new Date('1970-01-02T06:00');
                        const dayStart = new Date('1970-01-01T06:00');
                        const dayEnd = new Date('1970-01-01T22:00');

                        if (checkOut < checkIn) {
                            checkOut.setDate(checkOut.getDate() + 1);
                        }

                        const totalDuration = (checkOut - checkIn) / (1000 * 60 * 60);

                        if (totalDuration > 8) {
                            if (checkOut > dayStart && checkIn < dayEnd) {
                                const regOTStart = checkIn < dayStart ? dayStart : checkIn;
                                const regOTEnd = checkOut > dayEnd ? dayEnd : checkOut;
                                if (regOTStart < regOTEnd) {
                                    regOT = parseInt((regOTEnd - regOTStart) / (1000 * 60 * 60));
                                }
                            }

                            if (checkOut > nightStart || checkIn < nightEnd) {
                                const ndotStart = checkIn < nightStart ? nightStart : checkIn;
                                const ndotEnd = checkOut > nightEnd ? nightEnd : checkOut;
                                if (ndotStart < ndotEnd) {
                                    ndot = (ndotEnd - ndotStart) / (1000 * 60 * 60);
                                }
                            }

                            const otStart = new Date(checkIn);
                            otStart.setHours(otStart.getHours() + 8);
                            if (checkOut > otStart) {
                                const totalOT = (checkOut - otStart) / (1000 * 60 * 60);
                                const overlap = Math.max(0, regOT + ndot - totalOT);
                                if (overlap > 0) {
                                    const adjustFactor = overlap / (regOT + ndot);
                                    regOT *= 1 - adjustFactor;
                                    ndot *= 1 - adjustFactor;
                                }
                            } else {
                                regOT = 0;
                                ndot = 0;
                            }
                        }
                    }
                }

                const regOTCell = document.createElement('td');
                regOTCell.textContent = regOT > 0 ? regOT.toFixed(2) : '-';
                row.appendChild(regOTCell);

                const ndotCell = document.createElement('td');
                ndotCell.textContent = ndot > 0 ? ndot.toFixed(2) : '-';
                row.appendChild(ndotCell);

                // Calculate ND (10:00 PM to 6:00 AM excluding NDOT)
                let nd = 0;

                if (result.CheckIn && result.CheckOut) {
                    const checkIn = new Date(`1970-01-01T${result.CheckIn}`);
                    const checkOut = new Date(`1970-01-01T${result.CheckOut}`);

                    if (!isNaN(checkIn) && !isNaN(checkOut)) {
                            const nightStart = new Date('1970-01-01T22:00');
                            const nightEnd = new Date('1970-01-02T06:00'); // Adjust to next day

                            // Adjust for overnight shifts
                            if (checkOut < checkIn) {
                                checkOut.setDate(checkOut.getDate() + 1);
                            }

                            // Calculate ND (total hours between 10 PM and 6 AM)
                            if (checkOut > nightStart || checkIn < nightEnd) {
                                const ndStart = checkIn < nightStart ? nightStart : checkIn;
                                const ndEnd = checkOut > nightEnd ? nightEnd : checkOut;

                                if (ndStart < ndEnd) {
                                    nd = (ndEnd - ndStart) / (1000 * 60 * 60); // Convert ms to hours
                                }
                            }
                        }
                }

                // Subtract NDOT from ND
                nd = Math.max(0, nd - ndot);

                const ndCell = document.createElement('td');
                ndCell.textContent = nd > 0 ? nd.toFixed(2) : '-';
                row.appendChild(ndCell);

                // Add Remarks column
                const RemarksCell = document.createElement('td');
                const durationDecimal = parseDurationToDecimal(result.Duration || '');
                RemarksCell.textContent = durationDecimal > 13 ? 'Not Normal' : 'Normal';
                row.appendChild(RemarksCell);
            } else {
                // Add limited columns if department is not All Departments>SANTEH
                ['Employee', 'Department', 'Status', 'Duration', 'Date', 'CheckIn', 'CheckOut'].forEach(key => {
                    const td = document.createElement('td');
                    td.textContent = result[key] || '-';
                    row.appendChild(td);
                });

                // Add Remarks column
                const RemarksCell = document.createElement('td');
                const durationDecimal = parseDurationToDecimal(result.Duration || '');
                RemarksCell.textContent = durationDecimal > 13 ? 'Not Normal' : 'Normal';
                row.appendChild(RemarksCell);
                }

            table.appendChild(row);
        });

        output.appendChild(table);
    }

function adjustTableColumnWidths(table) {
    const rows = table.rows;
    const columnCount = rows[0].cells.length;

    // Initialize column widths based on the longest content in each column
    const columnWidths = new Array(columnCount).fill(0);

    // Iterate over the rows and adjust column widths
    for (let i = 0; i < columnCount; i++) {
        for (let row of rows) {
            const cell = row.cells[i];
            const cellWidth = cell.textContent.length;
            columnWidths[i] = Math.max(columnWidths[i], cellWidth);
        }
    }

    // Apply the computed widths to the table columns
    for (let i = 0; i < columnCount; i++) {
        table.columns[i] = columnWidths[i];
        const col = table.getElementsByTagName('col')[i];
        if (col) col.style.width = `${columnWidths[i] * 8}px`; // Adjust width
    }
}

function exportToExcel() {
    const output = document.getElementById('output');
    const table = output.querySelector('table');

    if (!table) {
        showToast('No data to export.');
        return;
    }

    // Helper function to parse dates correctly (DD-MM-YYYY to YYYY-MM-DD)
    function parseDate(dateStr) {
        const parts = dateStr.split("-");
        if (parts.length === 3) {
            return new Date(`${parts[2]}-${parts[1]}-${parts[0]}`); // Convert to YYYY-MM-DD
        }
        return new Date(dateStr); // Fallback in case it's already valid
    }

    // Parse the data into a structured format
    const results = [];
    const headers = Array.from(table.querySelectorAll('th')).map(th => th.textContent);
    const rows = table.querySelectorAll('tr');

    rows.forEach((row, rowIndex) => {
        if (rowIndex === 0) return; // Skip header row

        const rowData = {};
        const cells = row.querySelectorAll('td');
        cells.forEach((cell, colIndex) => {
            rowData[headers[colIndex]] = cell.textContent;
        });
        results.push(rowData);
    });

    // Fix: Properly parse and sort dates
    results.sort((a, b) => parseDate(a.Date) - parseDate(b.Date));

    // Group data by employee
    const groupedData = results.reduce((acc, row) => {
        const employee = row.Employee;
        if (!acc[employee]) acc[employee] = [];
        acc[employee].push({
            Date: row.Date,
            'Check In': row.CheckIn,
            'Break In 1': row['BreakIn1'] || '',
            'Break Out 1': row['BreakOut1'] || '',
            'Break In 2': row['BreakIn2'] || '',
            'Break Out 2': row['BreakOut2'] || '',
            'Break In 3': row['BreakIn3'] || '',
            'Break Out 3': row['BreakOut3'] || '',
            'Check Out': row.CheckOut,
            'Reg Late': row['Reg Late'] || '',
            'Break Late': row['Break Late'] || '',
            'Total Late': row['Total Late'] || '',
            'Reg OT': row['Reg OT'] || '',
            'NDOT': row['NDOT'] || '',
            'ND': row['ND'] || '',
            'Hours Rendered': row['Hours Rendered'] || '',
            'Remarks': row['Remarks'] || ''
        });
        return acc;
    }, {});

    // Create a new workbook
    const wb = XLSX.utils.book_new();

    // Add the main "Attendance Data" sheet
    const ws = XLSX.utils.json_to_sheet([]);
    const department = results[0]?.Department || 'N/A';
    const startDate = results.length > 0 ? results[0].Date : 'N/A';
    const endDate = results.length > 0 ? results[results.length - 1].Date : 'N/A';
    const mainHeader = [["KFL MANPOWER AGENCY ATTENDANCE DATA PIVOT"], 
                        [`Date Range: ${startDate} - ${endDate}`], 
                        [`Department: ${department}`]];

    XLSX.utils.sheet_add_aoa(ws, mainHeader, { origin: 'A1' });

    let rowOffset = 4;
    for (const employee in groupedData) {
        const employeeData = groupedData[employee];

        const headers = department === 'All Departments>SANTEH'
        ? ['Date', 'Check In', 'Break In 1', 'Break Out 1', 'Break In 2', 'Break Out 2', 'Break In 3', 'Break Out 3', 'Check Out', 'Reg Late', 'Break Late', 'Total Late', 'Reg OT', 'NDOT', 'ND', 'Hours Rendered', 'Remarks']
        : ['Date', 'Check In', 'Check Out'];

        const headerRow = [[`${employee}`], headers];
        XLSX.utils.sheet_add_aoa(ws, headerRow, { origin: `A${rowOffset + 1}` });

        const formattedData = employeeData.map(record => {
            const row = {};
            headers.forEach(header => {
                row[header] = record[header] || ''; // ✅ Ensures all break columns are included even if empty
            });
            return row;
        });
        XLSX.utils.sheet_add_json(ws, formattedData, { origin: `A${rowOffset + 3}`, skipHeader: true });
        rowOffset += formattedData.length + 3;
    }

    wb.SheetNames.push('Attendance Data');
    wb.Sheets['Attendance Data'] = ws;

    // Extract unique employee names for the Daily Time Record
    const uniqueEmployees = [...new Set(results.map(row => row.Employee))];

    // Initialize worksheet data storage for Daily Time Record
    let wsData = [];

    // Loop through each employee and format their Daily Time Record
    uniqueEmployees.forEach((employee, index) => {
        const employeeRecords = results.filter(row => row.Employee === employee);

        if (index > 0) {
            wsData.push([]); // Blank row for separation
        }

        wsData.push(["DAILY TIME RECORD"]);
        wsData.push([employee]);
        wsData.push(["DEPARTMENT:", employeeRecords[0]?.Department || "N/A"]);
        wsData.push(["COMPANY:", employeeRecords[0]?.Company || "N/A"]);
        wsData.push([]);

        wsData.push([
            "DATE", "", "AM", "", "PM", "", "DEDUCTION", "HOURS CREDIT", "", "TOTAL OF HOURS"
        ]);
        wsData.push([
            "", "", "TIME IN", "TIME OUT", "TIME IN", "TIME OUT", "UNDERTIME", "REG HOURS", "REG OT", ""
        ]);

        // Get the earliest attendance date dynamically
        const minDate = results.length > 0 ? parseDate(results[0].Date) : new Date();
        const startDate = new Date(minDate);

        // Format month and year dynamically
        const monthYear = startDate.toLocaleDateString("en-US", { month: "long", year: "numeric" });
        wsData.push([monthYear]);

        for (let i = 0; i < 15; i++) {
            let currentDate = new Date(startDate);
            currentDate.setDate(startDate.getDate() + i);

            let dayName = currentDate.toLocaleDateString("en-US", { weekday: "long" });
            let formattedDate = currentDate.toLocaleDateString("en-US", { day: "2-digit" });

            let record = employeeRecords.find(r => parseDate(r.Date).getTime() === currentDate.getTime());

            wsData.push([
                formattedDate, dayName,
                record?.CheckIn || "",record?.BreakIn2 || "" ,
                record?.BreakOut2 || "", record?.CheckOut || "",
                record?.["Reg Late"] || "", record?.["Hours Rendered"] || "",
                record?.["Reg OT"] || "", record?.["Total Hours"] || ""
            ]);
        }

        wsData.push(["TOTAL", "", "", "", "", "", "", "", "", ""]);
    });

    const dailyTimeRecordWs = XLSX.utils.aoa_to_sheet(wsData);
    wb.SheetNames.push("Daily Time Record");
    wb.Sheets["Daily Time Record"] = dailyTimeRecordWs;
    

    // Export to Excel
    XLSX.writeFile(wb, 'attendance_data.xlsx');
}

function openModal() {
    document.getElementById('myModal').style.display = 'block';
    document.getElementById('popupContent').src = 'froggy.php'; // Load PHP page
}

function closeModal() {
    document.getElementById('myModal').style.display = 'none';
}

// Close the modal if user clicks outside the content
window.onclick = function(event) {
    if (event.target === document.getElementById('myModal')) {
        closeModal();
    }
} 