document.addEventListener('DOMContentLoaded', function() {
    const fileInput = document.getElementById('fileInput');
    const fileNameSpan = document.getElementById('fileName');
    
    if (fileInput && fileNameSpan) {
        fileInput.addEventListener('change', function(event) {
            const file = event.target.files[0];
            
            if (file) {
                const fileName = file.name.toLowerCase();
                if (fileName.endsWith('.xlsx') || fileName.endsWith('.xls') || fileName.endsWith('.csv')) {
                    fileNameSpan.textContent = `Selected File: ${file.name}`;
                    console.log('File selected:', file.name);
                    showToast('📁 File uploaded successfully!');
                } else {
                    fileNameSpan.textContent = 'Invalid file type';
                    fileInput.value = '';
                    showToast('❌ Please select a valid Excel (.xlsx, .xls) or CSV file!');
                }
            } else {
                fileNameSpan.textContent = 'No file chosen';
                console.log('No file selected');
            }
        });
    } else {
        console.error('File input or fileName span not found');
    }
});

function clearFile() {
    const fileInput = document.getElementById('fileInput');
    const fileNameSpan = document.getElementById('fileName');
    const output = document.getElementById('output');

    fileInput.value = '';
    fileNameSpan.textContent = 'No file selected';
    output.innerHTML = '';
    showToast('✅ Data cleared successfully!');
}

function processFile() {
    console.log('processFile called');
    const fileInput = document.getElementById('fileInput');
    
    if (!fileInput) {
        console.error('File input element not found');
        showToast('File input not found.');
        return;
    }
    
    const file = fileInput.files[0];
    console.log('File object:', file);

    if (!file) {
        console.log('No file selected');
        showToast('Please select a file.');
        return;
    }
    
    console.log('Processing file:', file.name, 'Size:', file.size, 'Type:', file.type);

    const reader = new FileReader();
    const fileName = file.name.toLowerCase();
    
    if (fileName.endsWith('.csv')) {
        reader.onload = (e) => {
            try {
                console.log('CSV file loaded successfully');
                const text = e.target.result;
                const lines = text.split('\n');
                console.log('Number of lines:', lines.length);
            
            if (lines.length < 6 || !lines[0].includes('KFL MANPOWER AGENCY')) {
                showToast('This is not a valid KFL Attendance file.');
                return;
            }
            
            let headerRowIndex = -1;
            for (let i = 0; i < Math.min(10, lines.length); i++) {
                if (lines[i].includes('ID,Name,Department')) {
                    headerRowIndex = i;
                    break;
                }
            }
            
            if (headerRowIndex === -1) {
                showToast('Could not find data header in the file.');
                return;
            }
            
            const headers = lines[headerRowIndex].split(',');
            const jsonData = [];
            
            for (let i = headerRowIndex + 1; i < lines.length; i++) {
                const line = lines[i].trim();
                if (line) {
                    const values = line.split(',');
                    const record = {};
                    headers.forEach((header, index) => {
                        record[header.trim()] = values[index] ? values[index].trim() : '';
                    });
                    if (record.ID && record.Name) {
                        jsonData.push(record);
                    }
                }
            }
            
                analyzeData(jsonData);
                showToast('✅ Attendance data calculated successfully!');
            } catch (error) {
                console.error('Error processing CSV file:', error);
                showToast('Error processing CSV file: ' + error.message);
            }
        };
        
        reader.onerror = (error) => {
            console.error('Error reading CSV file:', error);
            showToast('Error reading file.');
        };
        
        reader.readAsText(file);
    } else if (fileName.endsWith('.xlsx') || fileName.endsWith('.xls')) {
        reader.onload = (e) => {
            try {
                console.log('Excel file loaded successfully');
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const sheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[sheetName];
                
                // Check for KFL MANPOWER AGENCY SERVER 3 in cells A1-A3
                const a1 = worksheet['A1'] ? worksheet['A1'].v : '';
                const a2 = worksheet['A2'] ? worksheet['A2'].v : '';
                const a3 = worksheet['A3'] ? worksheet['A3'].v : '';
                
                const headerText = `${a1}${a2}${a3}`.toUpperCase();
                if (!headerText.includes('KFL MANPOWER AGENCY SERVER 3')) {
                    showToast('❌ Invalid file format. Please upload a valid KFL attendance file.');
                    return;
                }
                
                const jsonData = XLSX.utils.sheet_to_json(worksheet, { range: 7 });
                console.log('Excel data parsed, rows:', jsonData.length);
                analyzeData(jsonData);
                showToast('✅ Attendance data calculated successfully!');
            } catch (error) {
                console.error('Error processing Excel file:', error);
                showToast('Error processing Excel file: ' + error.message);
            }
        };
        
        reader.onerror = (error) => {
            console.error('Error reading Excel file:', error);
            showToast('Error reading file.');
        };
        
        reader.readAsArrayBuffer(file);
    } else {
        showToast('Please select a CSV or Excel file.');
        return;
    }
}

function showToast(message) {
    const toast = document.createElement('div');
    toast.className = 'modern-toast';
    toast.innerHTML = message;
    
    // Modern toast styling
    toast.style.position = 'fixed';
    toast.style.top = '20px';
    toast.style.left = '50%';
    toast.style.transform = 'translateX(-50%) translateY(-100px)';
    toast.style.background = 'linear-gradient(45deg, #4facfe, #00f2fe)';
    toast.style.color = 'white';
    toast.style.padding = '15px 25px';
    toast.style.borderRadius = '25px';
    toast.style.fontSize = '16px';
    toast.style.fontWeight = '600';
    toast.style.boxShadow = '0 8px 32px rgba(79, 172, 254, 0.3)';
    toast.style.zIndex = '10000';
    toast.style.backdropFilter = 'blur(10px)';
    toast.style.border = '1px solid rgba(255, 255, 255, 0.2)';
    toast.style.transition = 'all 0.4s cubic-bezier(0.68, -0.55, 0.265, 1.55)';
    toast.style.opacity = '0';
    
    document.body.appendChild(toast);
    
    // Animate in
    setTimeout(() => {
        toast.style.transform = 'translateX(-50%) translateY(0)';
        toast.style.opacity = '1';
    }, 10);
    
    // Animate out and remove
    setTimeout(() => {
        toast.style.transform = 'translateX(-50%) translateY(-100px)';
        toast.style.opacity = '0';
        setTimeout(() => {
            if (document.body.contains(toast)) {
                document.body.removeChild(toast);
            }
        }, 400);
    }, 3000);
}

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
    <strong>Update 01-20-2025:</strong><br>
    - Enhanced UI with modern green-blue gradient design and improved mobile responsiveness.<br>
    - Redesigned layout with compact mode for better data visibility.<br>
    - Added proper modal dialogs for footer links and improved scrolling behavior.<br>
    - Optimized button alignment and file input positioning for better user experience.<br>
    <strong>* Always double-check the data.</strong><br><br>
    For any inquiries, feel free to contact IT Personnel.<br><br>
    <strong>Thank you!</strong>
    `;

    // Create backdrop
    const backdrop = document.createElement('div');
    backdrop.style.position = 'fixed';
    backdrop.style.top = '0';
    backdrop.style.left = '0';
    backdrop.style.width = '100%';
    backdrop.style.height = '100%';
    backdrop.style.backgroundColor = 'rgba(0, 0, 0, 0.5)';
    backdrop.style.zIndex = '9998';
    backdrop.style.opacity = '0';
    backdrop.style.transition = 'opacity 0.3s ease';

    // Create modal
    const toast = document.createElement('div');
    toast.classList.add('toast-message');
    toast.innerHTML = toastMessage;

    toast.style.position = 'fixed';
    toast.style.top = '50%';
    toast.style.left = '50%';
    toast.style.transform = 'translate(-50%, -50%) scale(0.7)';
    toast.style.padding = '25px 35px';
    toast.style.background = 'linear-gradient(135deg, #f44336, #e53935)';
    toast.style.color = '#fff';
    toast.style.borderRadius = '15px';
    toast.style.fontSize = '16px';
    toast.style.textAlign = 'center';
    toast.style.zIndex = '9999';
    toast.style.lineHeight = '1.6';
    toast.style.maxWidth = '90vw';
    toast.style.maxHeight = '80vh';
    toast.style.overflowY = 'auto';
    toast.style.boxShadow = '0 20px 60px rgba(244, 67, 54, 0.3)';
    toast.style.transition = 'all 0.4s cubic-bezier(0.68, -0.55, 0.265, 1.55)';
    toast.style.opacity = '0';

    // Add close button
    const closeBtn = document.createElement('button');
    closeBtn.innerHTML = '×';
    closeBtn.style.position = 'absolute';
    closeBtn.style.top = '10px';
    closeBtn.style.right = '15px';
    closeBtn.style.background = 'none';
    closeBtn.style.border = 'none';
    closeBtn.style.color = 'white';
    closeBtn.style.fontSize = '24px';
    closeBtn.style.cursor = 'pointer';
    closeBtn.style.transition = 'transform 0.2s ease';
    
    closeBtn.onmouseover = () => closeBtn.style.transform = 'scale(1.2)';
    closeBtn.onmouseout = () => closeBtn.style.transform = 'scale(1)';
    closeBtn.onclick = closeModal;

    toast.appendChild(closeBtn);
    document.body.appendChild(backdrop);
    document.body.appendChild(toast);

    // Animate in
    setTimeout(() => {
        backdrop.style.opacity = '1';
        toast.style.transform = 'translate(-50%, -50%) scale(1)';
        toast.style.opacity = '1';
    }, 10);

    function closeModal() {
        toast.style.transform = 'translate(-50%, -50%) scale(0.7)';
        toast.style.opacity = '0';
        backdrop.style.opacity = '0';
        
        setTimeout(() => {
            if (document.body.contains(toast)) {
                document.body.removeChild(toast);
            }
            if (document.body.contains(backdrop)) {
                document.body.removeChild(backdrop);
            }
        }, 400);
    }

    // Auto close after 8 seconds
    setTimeout(closeModal, 8000);

    // Close on backdrop click
    backdrop.onclick = closeModal;
}

window.addEventListener('DOMContentLoaded', () => {
    showToastMessage();
});

function combineDateTime(dateStr, timeStr) {
    if (!dateStr || !timeStr) return null;
    
    try {
        let day, month, year;
        
        if (dateStr.includes('-')) {
            const parts = dateStr.split('-');
            if (parts.length === 3) {
                day = parseInt(parts[0]);
                month = parseInt(parts[1]);
                year = parseInt(parts[2]);
            }
        } else if (dateStr.includes('/')) {
            const parts = dateStr.split('/');
            if (parts.length === 3) {
                day = parseInt(parts[0]);
                month = parseInt(parts[1]);
                year = parseInt(parts[2]);
            }
        }
        
        // Handle numeric time values (Excel decimal format)
        if (typeof timeStr === 'number') {
            const totalMinutes = Math.round(timeStr * 24 * 60);
            const hours = Math.floor(totalMinutes / 60);
            const minutes = totalMinutes % 60;
            timeStr = `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}`;
        }
        
        const timeParts = timeStr.split(':');
        if (timeParts.length < 2) {
            console.warn(`Invalid time format: ${timeStr}`);
            return null;
        }
        
        const hours = parseInt(timeParts[0]);
        const minutes = parseInt(timeParts[1]);
        
        if (isNaN(year) || isNaN(month) || isNaN(day) || 
            isNaN(hours) || isNaN(minutes) ||
            year < 2000 || year > 2100 ||
            month < 1 || month > 12 ||
            day < 1 || day > 31 ||
            hours < 0 || hours > 23 ||
            minutes < 0 || minutes > 59) {
            console.warn(`Invalid date/time components: ${dateStr} ${timeStr}`);
            return null;
        }

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

function analyzeData(data) {
    if (data.length === 0) {
        showToast('The uploaded file contains no data.');
        return;
    }

    const results = [];
    const employeeRecords = {};

    // Group all records by employee ID
    data.forEach(record => {
        const id = record['ID'];
        const name = record['Name'];
        const department = record['Department'];
        const date = record['Date'];
        const time = record['Check-In Time'] || record['Time'];
        const type = record['Card Swiping Type'];
        
        if (!id || !name || !date || !time || !type) return;
        
        // Convert numeric time to string format
        let formattedTime = time;
        if (typeof time === 'number') {
            const totalMinutes = Math.round(time * 24 * 60);
            const hours = Math.floor(totalMinutes / 60);
            const minutes = totalMinutes % 60;
            formattedTime = `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}`;
        }
        
        const dateTime = combineDateTime(date, formattedTime);
        if (!dateTime) return;

        // Normalize type to handle case variations
        const lowerType = type.toLowerCase().trim();
        const normalizedType = lowerType.includes('check out') || lowerType.includes('checkout') ? 'Check Out' : 
                              lowerType.includes('check in') || lowerType.includes('checkin') ? 'Check In' : type;

        if (!employeeRecords[id]) {
            employeeRecords[id] = {
                name,
                department,
                records: []
            };
        }

        employeeRecords[id].records.push({
            date,
            time: formattedTime,
            type: normalizedType,
            dateTime
        });
    });

    // Sort records by datetime for each employee
    Object.values(employeeRecords).forEach(emp => {
        emp.records.sort((a, b) => a.dateTime - b.dateTime);
    });

    // Get all unique dates
    const allDates = new Set();
    Object.values(employeeRecords).forEach(emp => {
        emp.records.forEach(record => {
            allDates.add(record.date);
        });
    });

    const sortedDates = Array.from(allDates).sort((a, b) => {
        const [dayA, monthA, yearA] = a.split('-').map(Number);
        const [dayB, monthB, yearB] = b.split('-').map(Number);
        return new Date(yearA, monthA - 1, dayA) - new Date(yearB, monthB - 1, dayB);
    });

    // Process each employee
    Object.entries(employeeRecords).forEach(([id, emp]) => {
        const usedRecords = new Set();
        
        // Debug: Log all records for this employee
        console.log(`Processing employee: ${emp.name}`);
        emp.records.forEach((r, idx) => {
            console.log(`  Record ${idx}: ${r.date} ${r.time} ${r.type}`);
        });
        
        // Process all records chronologically to find valid shifts
        for (let i = 0; i < emp.records.length; i++) {
            const record = emp.records[i];
            
            // Skip if already used
            if (usedRecords.has(i)) {
                console.log(`  Skipping record ${i} (already used): ${record.date} ${record.time} ${record.type}`);
                continue;
            }
            
            // Only process Check In records as potential shift starts
            if (record.type !== 'Check In') {
                console.log(`  Skipping record ${i} (not Check In): ${record.date} ${record.time} ${record.type}`);
                continue;
            }
            
            console.log(`  Processing Check In: ${record.date} ${record.time}`);
            
            const checkInHour = record.dateTime.getHours();
            const isNightShift = checkInHour >= 17; // 5 PM or later is night shift
            let checkout = null;
            let checkoutIndex = -1;
            let remarks = '';
            
            // Look for matching checkout
            for (let j = i + 1; j < emp.records.length; j++) {
                if (usedRecords.has(j)) continue;
                
                const potentialCheckout = emp.records[j];
                const timeDiff = (potentialCheckout.dateTime - record.dateTime) / (1000 * 60 * 60);
                
                console.log(`    Checking potential checkout: ${potentialCheckout.date} ${potentialCheckout.time} ${potentialCheckout.type}, timeDiff: ${timeDiff.toFixed(2)}h`);
                
                // Valid checkout conditions
                if (timeDiff >= 1 && timeDiff <= 16) {
                    if (potentialCheckout.type === 'Check Out') {
                        console.log(`    Found Check Out match!`);
                        checkout = potentialCheckout;
                        checkoutIndex = j;
                        break;
                    } else if (potentialCheckout.type === 'Check In') {
                        // For night shift, allow shorter duration (4+ hours)
                        // For day shift, require minimum 4 hours to capture valid shifts
                        const minHours = 4;
                        if (timeDiff >= minHours) {
                            console.log(`    Found Check In match (treated as checkout)!`);
                            checkout = potentialCheckout;
                            checkoutIndex = j;
                            remarks = 'Check-in treated as checkout';
                            break;
                        }
                    }
                }
            }
            
            if (checkout) {
                const duration = (checkout.dateTime - record.dateTime) / (1000 * 60 * 60);
                console.log(`    Creating shift: ${record.date} ${record.time} to ${checkout.time}, duration: ${duration.toFixed(2)}h`);
                
                // Mark records as used
                usedRecords.add(i);
                usedRecords.add(checkoutIndex);
                
                // Determine shift type and remarks
                if (isNightShift) {
                    if (duration < 4) {
                        remarks = remarks || 'Night shift undertime';
                    } else if (duration > 12) {
                        remarks = remarks || 'Extended night shift';
                    } else {
                        remarks = remarks || 'Night shift';
                    }
                    
                    results.push({
                        Employee: emp.name,
                        Department: emp.department,
                        Status: 'Night Shift',
                        Duration: `${Math.floor(duration)}h ${Math.round((duration % 1) * 60)}m`,
                        Date: record.date,
                        CheckIn: record.time,
                        CheckOut: checkout.time,
                        Remarks: remarks
                    });
                } else {
                    if (duration < 4) {
                        remarks = remarks || 'Undertime - early checkout';
                    } else if (duration > 12) {
                        remarks = remarks || 'Extended shift - long hours';
                    } else if (duration > 9) {
                        remarks = remarks || 'Overtime - late checkout';
                    } else {
                        remarks = remarks || 'Day shift';
                    }
                    
                    results.push({
                        Employee: emp.name,
                        Department: emp.department,
                        Status: checkInHour >= 12 ? 'Day Shift' : 'Morning Shift',
                        Duration: `${Math.floor(duration)}h ${Math.round((duration % 1) * 60)}m`,
                        Date: record.date,
                        CheckIn: record.time,
                        CheckOut: checkout.time,
                        Remarks: remarks
                    });
                }
            } else {
                // No checkout found - mark as missing
                console.log(`    No checkout found for: ${record.date} ${record.time}`);
                usedRecords.add(i);
                results.push({
                    Employee: emp.name,
                    Department: emp.department,
                    Status: 'Missing Check Out',
                    Duration: '-',
                    Date: record.date,
                    CheckIn: record.time,
                    CheckOut: 'Missing',
                    Remarks: 'Missing checkout - no additional records'
                });
            }
        }
    });

    displayResults(results);
}

function getNextDay(dateStr) {
    const [day, month, year] = dateStr.split('-').map(Number);
    const date = new Date(year, month - 1, day);
    date.setDate(date.getDate() + 1);
    
    const nextDay = String(date.getDate()).padStart(2, '0');
    const nextMonth = String(date.getMonth() + 1).padStart(2, '0');
    const nextYear = date.getFullYear();
    
    return `${nextDay}-${nextMonth}-${nextYear}`;
}



function displayResults(results) {
    const output = document.getElementById('output');
    output.innerHTML = '';

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

    searchBar.addEventListener('input', () => {
        const filter = searchBar.value.toLowerCase();
        const rows = table.querySelectorAll('tr:not(:first-child)');

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
        errorMessage.textContent = 'No data to display.';
        errorMessage.style.color = 'red';
        errorMessage.style.fontWeight = 'bold';
        output.appendChild(errorMessage);
        return;
    }

    results.sort((a, b) => {
        const nameCompare = a.Employee.localeCompare(b.Employee);
        if (nameCompare !== 0) return nameCompare;
        
        const [dayA, monthA, yearA] = a.Date.split('-').map(Number);
        const [dayB, monthB, yearB] = b.Date.split('-').map(Number);
        return new Date(yearA, monthA - 1, dayA) - new Date(yearB, monthB - 1, dayB);
    });

    const tableContainer = document.createElement('div');
    tableContainer.className = 'table-container';
    
    const table = document.createElement('table');
    const headerRow = document.createElement('tr');

    const headers = ['Employee', 'Department', 'Status', 'Hours Rendered', 'Date', 'CheckIn', 'CheckOut', 'Remarks'];

    headers.forEach(header => {
        const th = document.createElement('th');
        th.textContent = header;
        headerRow.appendChild(th);
    });

    table.appendChild(headerRow);

    results.forEach(result => {
        const row = document.createElement('tr');

        ['Employee', 'Department', 'Status', 'Duration', 'Date', 'CheckIn', 'CheckOut'].forEach(key => {
            const td = document.createElement('td');
            td.textContent = result[key] || '-';
            row.appendChild(td);
        });

        const remarksCell = document.createElement('td');
        remarksCell.textContent = result.Remarks || 'Normal';
        row.appendChild(remarksCell);

        table.appendChild(row);
    });

    tableContainer.appendChild(table);
    output.appendChild(tableContainer);
}

function exportToExcel() {
    const output = document.getElementById('output');
    const table = output.querySelector('table');

    if (!table) {
        showToast('No data to export.');
        return;
    }

    const results = [];
    const headers = Array.from(table.querySelectorAll('th')).map(th => th.textContent);
    const rows = table.querySelectorAll('tr');

    // Get date range from results
    const dates = [];
    rows.forEach((row, rowIndex) => {
        if (rowIndex === 0) return;
        const dateCell = row.cells[4]; // Date column
        if (dateCell) dates.push(dateCell.textContent);
    });
    
    const sortedDates = dates.sort((a, b) => {
        const [dayA, monthA, yearA] = a.split('-').map(Number);
        const [dayB, monthB, yearB] = b.split('-').map(Number);
        return new Date(yearA, monthA - 1, dayA) - new Date(yearB, monthB - 1, dayB);
    });
    
    const minDate = sortedDates[0] || '';
    const maxDate = sortedDates[sortedDates.length - 1] || '';
    
    // Current date and time
    const now = new Date();
    const currentDate = `${String(now.getDate()).padStart(2, '0')}-${String(now.getMonth() + 1).padStart(2, '0')}-${now.getFullYear()}`;
    const currentTime = `${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}`;

    // Create header data with centered company name using spaces
    const headerData = [
        ['                                                                                                                              KFL MANPOWER AGENCY SERVER 3                    '],
        [''],
        [''],
        ['Transaction'],
        [`Export Time: ${currentDate} ${currentTime}`],
        [`Time Period: ${minDate} - ${maxDate}`],
        ['']
    ];

    rows.forEach((row, rowIndex) => {
        if (rowIndex === 0) return;

        const rowData = {};
        const cells = row.querySelectorAll('td');
        cells.forEach((cell, colIndex) => {
            if (headers[colIndex] !== 'Remarks') {
                let cellValue = cell.textContent;
                // Remove "All Departments>" from department
                if (headers[colIndex] === 'Department') {
                    cellValue = cellValue.replace(/^All Departments&gt;/, '');
                }
                rowData[headers[colIndex]] = cellValue;
            }
        });
        results.push(rowData);
    });

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(headerData);
    XLSX.utils.sheet_add_json(ws, results, { origin: 'A8' });
    
    // Set column widths with text wrapping
    ws['!cols'] = [
        { wch: 25, wrapText: true }, // Employee
        { wch: 50, wrapText: true }, // Department
        { wch: 15 }, // Status
        { wch: 20 }, // Hours Rendered
        { wch: 12 }, // Date
        { wch: 10 }, // CheckIn
        { wch: 10 }  // CheckOut
    ];
    
    // Merge A1:G3 and add border
    ws['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 2, c: 6 } }];
    
    // Apply styles manually to specific cells
    if (!ws['A1']) ws['A1'] = { v: 'KFL MANPOWER AGENCY SERVER 3', t: 's' };
    ws['A1'].s = {
        alignment: { horizontal: 'center', vertical: 'center' },
        font: { bold: true, sz: 14 },
        border: {
            top: { style: 'double' },
            bottom: { style: 'double' },
            left: { style: 'double' },
            right: { style: 'double' }
        }
    };
    
    // Apply border to merged cells A1:G3
    ['B1', 'C1', 'D1', 'E1', 'F1', 'G1', 'A2', 'B2', 'C2', 'D2', 'E2', 'F2', 'G2', 'A3', 'B3', 'C3', 'D3', 'E3', 'F3', 'G3'].forEach(cell => {
        if (!ws[cell]) ws[cell] = { v: '', t: 's' };
        ws[cell].s = {
            border: {
                top: { style: 'double' },
                bottom: { style: 'double' },
                left: { style: 'double' },
                right: { style: 'double' }
            }
        };
    });
    
    // Apply sky blue background to header row A8:G8
    ['A8', 'B8', 'C8', 'D8', 'E8', 'F8', 'G8'].forEach(cell => {
        if (!ws[cell]) ws[cell] = { v: '', t: 's' };
        ws[cell].s = {
            fill: { fgColor: { rgb: '87CEEB' } },
            font: { bold: true }
        };
    });
    
    wb.SheetNames.push('Attendance Data');
    wb.Sheets['Attendance Data'] = ws;

    XLSX.writeFile(wb, 'attendance_data.xlsx');
    showToast('📊 Data exported to Excel successfully!');
}

