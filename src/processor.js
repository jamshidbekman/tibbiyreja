const ExcelJS = require('exceljs');
const xlsx = require('xlsx'); // Still used for reading input
const dayjs = require('dayjs');
const archiver = require('archiver');
const fs = require('fs');
const path = require('path');

// Uzbek Month Names
const MONTH_NAMES = [
    "Yanvar", "Fevral", "Mart", "Aprel", "May", "Iyun",
    "Iyul", "Avgust", "Sentabr", "Oktabr", "Noyabr", "Dekabr"
];

const CUSTOM_DATE_FORMAT = 'D.MM.YYYY';

/**
 * Detect birth date column
 */
function findBirthDateColumn(headerRow) {
    // headerRow is Array of strings or objects {header:..., key:...}
    // We assume row values are strings
    const keywords = ['tug\'ilgan', 'birth', 'd.o.b', 'data rojdeniya', 'sana'];

    // 1. Try key name match
    // headerRow is just values: ["Id", "Name", "Tug'ilgan sanasi"]
    for (let i = 0; i < headerRow.length; i++) {
        const val = String(headerRow[i]).toLowerCase();
        if (keywords.some(k => val.includes(k))) {
            return i + 1; // Return 1-based index for ExcelJS
        }
    }
    return 1; // Fallback to col 1
}

function isValidDate(val) {
    if (!val) return false;
    return dayjs(val).isValid() || (typeof val === 'object' && val instanceof Date);
}

/**
 * Parse Excel Date
 */
function parseDate(val) {
    if (val === null || val === undefined || val === '') return dayjs(null);

    // ExcelJS usually returns Date objects for date cells
    if (val instanceof Date) return dayjs(val);

    if (typeof val === 'number') {
        // If it looks like a Year (e.g. 1990, 2005), treat as Year
        if (val >= 1900 && val <= 2100) {
            return dayjs(`${val}-01-01`);
        }
        // If it looks like a small number (e.g. count 1025), but not a year, 
        // it might be interpreted as serial (1025 -> 1902). 
        // We defer filtering to the validator, but usually < 10000 is suspicious for DOBS (10000 = 1927).
        // Let's assume valid serials are > 4000 (1910).
        if (val < 4000) return dayjs(null); // Treat small numbers as invalid to avoid "Count" rows

        // Excel date
        return dayjs(new Date((val - (25567 + 2)) * 86400 * 1000));
    }

    if (typeof val === 'string') {
        const trimmed = val.trim();
        if (!trimmed) return dayjs(null);

        // Handle DD.MM.YYYY
        if (trimmed.includes('.')) {
            const parts = trimmed.split('.');
            if (parts.length === 3) {
                return dayjs(`${parts[2]}-${parts[1]}-${parts[0]}`);
            }
        }
        return dayjs(trimmed);
    }

    return dayjs(val);
}

/**
 * Get Working Days for a Month
 */
function getWorkingDays(year, monthIndex, holidays, saturdayWorking) {
    const days = [];
    // monthIndex is 0-11
    const start = dayjs().year(year).month(monthIndex).date(1);
    const end = start.endOf('month');

    let curr = start;
    while (curr.isBefore(end) || curr.isSame(end, 'day')) {
        const dWeek = curr.day(); // 0 is Sunday
        const isSunday = dWeek === 0;
        const isSaturday = dWeek === 6;

        // Formatted date string to check holidays
        const dateStr = curr.format('YYYY-MM-DD');
        const isHoliday = holidays.includes(dateStr);

        let isWorking = !isSunday && !isHoliday;
        if (isSaturday && !saturdayWorking) {
            isWorking = false;
        }

        if (isWorking) {
            days.push(curr);
        }
        curr = curr.add(1, 'day');
    }
    return days;
}

// Helper to check working day validity for SINGLE date
function isWorkingDay(date, holidays, saturdayWorking) {
    const dWeek = date.day();
    const dateStr = date.format('YYYY-MM-DD');
    if (dWeek === 0) return false;
    if (dWeek === 6 && !saturdayWorking) return false;
    if (holidays.includes(dateStr)) return false;
    return true;
}

function findNextWorkingDay(date, holidays, saturdayWorking) {
    let curr = date;
    // Search limit 60 days
    for (let i = 0; i < 60; i++) {
        if (isWorkingDay(curr, holidays, saturdayWorking)) return curr;
        curr = curr.add(1, 'day');
    }
    return date;
}


/**
 * Apply Styles to Worksheet from Template
 */
function applyTemplateStyles(targetSheet, headers, columnWidths, headerStyle) {
    // Set Columns with Keys and Widths
    targetSheet.columns = headers.map((h, i) => ({
        header: h,
        key: h,
        width: columnWidths[i] || 20
    }));

    // Apply Header Style
    const headerRow = targetSheet.getRow(1);

    // Explicitly loop through cells to apply style ONLY to data cells
    // Avoids infinite row styling
    for (let i = 1; i <= headers.length; i++) {
        const cell = headerRow.getCell(i);
        if (headerStyle) {
            cell.font = headerStyle.font;
            cell.fill = headerStyle.fill;
            cell.border = headerStyle.border;
            cell.alignment = headerStyle.alignment;
        } else {
            // Default Style
            cell.font = { bold: true };
            cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
            cell.alignment = { vertical: 'middle', horizontal: 'center' };
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE0E0E0' } };
        }
    }
    headerRow.commit();
}

/**
 * Find Header Row Index
 */
function findHeaderRow(sheet) {
    const keywords = ['tug\'ilgan', 'birth', 'd.o.b', 'sana', 'yil', 'date'];
    // Scan first 15 rows
    const limit = Math.min(15, sheet.rowCount);
    for (let r = 1; r <= limit; r++) {
        const row = sheet.getRow(r);
        let matchCount = 0;
        row.eachCell((cell) => {
            let val = cell.value;
            // Handle Rich Text
            if (typeof val === 'object' && val !== null) {
                if (val.richText) val = val.richText.map(t => t.text).join('');
                else if (val.text) val = val.text;
            }
            const str = String(val || '').toLowerCase();
            if (keywords.some(k => str.includes(k))) matchCount++;
        });
        if (matchCount > 0) return r;
    }
    return 1; // Default
}

/**
 * Main Process Function
 */
async function processExcel(buffer, config) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);

    const inputSheet = workbook.worksheets[0];
    if (inputSheet.rowCount < 1) throw new Error("Excel file is empty");

    // Detect Header Row
    const headerRowIdx = findHeaderRow(inputSheet);
    console.log(`Detected Header Row at index: ${headerRowIdx}`);

    const headerRow = inputSheet.getRow(headerRowIdx);
    const headers = [];
    headerRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        let val = cell.value;
        if (typeof val === 'object' && val !== null) {
            if (val.richText) val = val.richText.map(t => t.text).join('');
            else if (val.text) val = val.text;
        }
        headers[colNumber - 1] = val ? String(val) : `Col${colNumber}`;
    });

    console.log('Headers:', headers);

    // Extract Styles (from Column 1 and Header Row)
    const columnWidths = [];
    for (let i = 1; i <= inputSheet.columnCount; i++) {
        const col = inputSheet.getColumn(i);
        columnWidths[i - 1] = col.width || 20;
    }

    const firstHeaderCell = headerRow.getCell(1);
    const templateHeaderStyle = {
        font: firstHeaderCell.font,
        fill: firstHeaderCell.fill,
        border: firstHeaderCell.border,
        alignment: firstHeaderCell.alignment
    };

    const birthColIdx = findBirthDateColumn(headers);

    // Find ID Column (No, №, T/r)
    let idColKey = null;
    const idKeywords = ['№', 't/r', 'no', 'tartib'];
    for (const h of headers) {
        if (idKeywords.some(k => h.toLowerCase().includes(k))) {
            idColKey = h;
            break;
        }
    }
    // If not found, maybe first column? Usually it is.
    if (!idColKey && headers.length > 0) {
        idColKey = headers[0];
    }

    console.log('Detected Birth Column Index:', birthColIdx);
    console.log('Detected ID Column Key:', idColKey);


    // Detect Name Column for Gender Filter
    const keywordsName = ['f.i.sh', 'fish', 'ism', 'name', 'familiya'];
    let nameColKey = null;
    for (const h of headers) {
        if (keywordsName.some(k => h.toLowerCase().includes(k))) {
            nameColKey = h;
            break;
        }
    }
    // Fallback to col 2 if not found (Col 1 is usually ID)
    if (!nameColKey && headers.length > 1) {
        nameColKey = headers[1];
    }
    console.log('Detected Name Column Key:', nameColKey);


    // Trim Headers: Remove trailing empty columns
    let lastNonEmptyIdx = headers.length;
    while (lastNonEmptyIdx > 0 && (!headers[lastNonEmptyIdx - 1] || headers[lastNonEmptyIdx - 1].startsWith('Col'))) {
        lastNonEmptyIdx--;
    }
    // If we cut off too much (unlikely if header detection is good), ensure we keep up to birth col
    if (lastNonEmptyIdx < birthColIdx) lastNonEmptyIdx = headers.length;

    const trimmedHeaders = headers.slice(0, lastNonEmptyIdx);
    const trimmedWidths = columnWidths.slice(0, lastNonEmptyIdx);

    const rowsData = [];
    inputSheet.eachRow((row, rowNumber) => {
        if (rowNumber <= headerRowIdx) return; // Skip Header and Title rows

        const rowObj = {};
        row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
            if (colNumber > lastNonEmptyIdx) return;

            const key = headers[colNumber - 1];
            if (key) {
                let val = cell.value;
                if (typeof val === 'object' && val !== null) {
                    if (val.richText) val = val.richText.map(t => t.text).join('');
                    else if (val.text) val = val.text;
                }

                // FORMATTING FIXES
                // 1. JSHSHIR / Passport ID -> Force String
                const lowerKey = key.toLowerCase();
                if (lowerKey.includes('jshshir') || lowerKey.includes('shaxsiy') || lowerKey.includes('hujjat')) {
                    if (val) val = String(val);
                }

                rowObj[key] = val;
            }
        });

        const birthCell = row.getCell(birthColIdx);
        rowObj._birthDate = parseDate(birthCell.value);
        rowsData.push(rowObj);
    });

    const zipName = `Schedules_${Date.now()}.zip`;
    const zipPath = path.join(__dirname, '../dist', zipName);
    const outputStream = fs.createWriteStream(zipPath);
    const archive = archiver('zip', { zlib: { level: 9 } });

    archive.pipe(outputStream);

    const allRangesRows = [];
    const warnings = [];

    // Determine Global Max Visits for Consolidated Sheet
    let maxVisitsGlobal = 1;
    config.ranges.forEach(r => {
        if ((r.visitCount || 1) > maxVisitsGlobal) maxVisitsGlobal = r.visitCount || 1;
    });

    // Helper: Determine Gender
    // Returns 'female' if first word ends in 'a' (Case Insensitive), else 'male'
    function getGender(row) {
        if (!nameColKey) return 'male'; // Default
        const storedVal = row[nameColKey];
        if (!storedVal) return 'male';

        const valStr = String(storedVal).trim();
        // Take First Word (Surname usually)
        const firstWord = valStr.split(' ')[0];
        if (!firstWord) return 'male';

        if (firstWord.toLowerCase().endsWith('a')) return 'female';
        return 'male';
    }

    // Process each range
    for (const range of config.ranges) {
        const visitCount = range.visitCount || 1;

        // Dynamic Range Headers - REMOVED 'Tashrif oyi'
        let rangeHeaders = [...trimmedHeaders];
        let visitKeys = [];

        if (visitCount === 1) {
            rangeHeaders.push('Tashrif sanasi');
            visitKeys.push('Tashrif sanasi');
        } else {
            for (let v = 1; v <= visitCount; v++) {
                rangeHeaders.push(`${v}-tashrif sanasi`);
                visitKeys.push(`${v}-tashrif sanasi`);
            }
        }

        // Extend widths - REMOVED 15 for 'Tashrif oyi'
        const rangeColWidths = [...trimmedWidths];
        visitKeys.forEach(() => rangeColWidths.push(15));

        const rangePatients = rowsData.filter(r => {
            if (!r._birthDate.isValid()) return false;
            const y = r._birthDate.year();
            if (y < 1900) return false;

            // Year Filter
            if (y < range.startYear || y > range.endYear) return false;

            // Gender Filter
            if (range.gender && range.gender !== 'all') {
                const detectedGender = getGender(r);
                if (detectedGender !== range.gender) return false;
            }

            return true;
        });

        const totalPatients = rangePatients.length;

        // AUTO PLAN Logic for > 1 visit (or manually provided counts)
        // If visitCount > 1, we ignore 'counts' validation and process ALL patients
        let processAll = false;

        // If Birthday Mode is ON, we definitely process all
        if (range.useBirthday) processAll = true;
        else if (visitCount > 1) processAll = true;


        const totalPlanned = range.counts.reduce((a, b) => a + b, 0);

        if (totalPatients === 0) {
            warnings.push(`${range.startYear}-${range.endYear} (Jins: ${range.gender || 'all'}): Aholi topilmadi.`);
            continue;
        }

        if (!processAll) {
            if (totalPlanned > totalPatients) {
                throw new Error(`Xato: ${range.startYear}-${range.endYear} oralig'ida reja (${totalPlanned}) aholi sonidan (${totalPatients}) ko'p!`);
            }
            if (totalPlanned < totalPatients) {
                warnings.push(`${range.startYear}-${range.endYear}: Reja (${totalPlanned}) aholi sonidan (${totalPatients}) kam.`);
            }
        }

        // Create Output Workbook
        const outWb = new ExcelJS.Workbook();
        const rangeName = `${range.endYear}-${range.startYear}_${range.gender || 'all'}.xlsx`;
        const rangeAllRows = [];

        let patientPool = [...rangePatients];
        let pIndex = 0;

        // ---------------------------------------------------------
        // LOGIC A: Birthday Mode (Specific Date Scheduling)
        // ---------------------------------------------------------
        if (range.useBirthday) {
            // We don't loop by month counts. We simply iterate ALL patients and generate their timeline.
            // We can dump them all into one sheet or split by birth month?
            // Requirement says "months" output usually implies monthly sheets.
            // But if we simply output them by their expected visit month...

            // Let's stick to the structure: Create Month Sheets.
            // Assign patients to their "Birth Month" sheet (for first visit? or spread?)
            // Actually, if it's 12 visits, they appear in EVERY month. 
            // If we generate sheets Jan..Dec, this patient should appear in Jan sheet (1st visit), Feb sheet (2nd visit)?
            // Or just one big list?

            // Current structure: Generates sheets "Yanvar", "Fevral"... containing visits for that month.
            // If a patient has 12 visits, do we list them in 12 sheets? 
            // The old logic listed a patient once in the month they were "scheduled" (for 1 visit).
            // For multi-visit (e.g. 4), they were listed in the month of their FIRST visit, and other columns showed future dates.

            // So for Birthday Mode (Cycle), we should group them by their FIRST visit month.
            // Since it's 12 visits starting Jan (or relative to birthday?), usually "Annual Plan" starts Jan.
            // If child born Feb, does visit start Feb?
            // "1 yil davomida 12 marta".
            // Let's assume standard: We list them in the month of their FIRST visit.
            // If "12 visits" (monthly), everyone starts in Jan (or first relevant month).
            // Actually, simplest is to distribute them evenly or just put everyone in "Yanvar" if they start then?

            // Wait, if I have 100 kids. 12 visits/year.
            // I want to see them in Jan sheet?
            // Let's stick to: Group by Birth Month? No, that's irrelevant for scheduling usually.

            // PROPOSAL: Since they visit EVERY month (12 visits), they basically belong to ALL months.
            // But the Excel structure has separate sheets for months.
            // Usually this implies "Patients scheduled for PRIMARY visit in Jan".
            // For 12 visits, it's a cycle.
            // Let's put them in the month corresponding to their Birth Month? 
            // Example: Born in March. First visit March? (Then until next Feb?)
            // Or Born March => Visits Jan 7 (if already born), Feb 7...
            // Request: "tashrif sanasini tug'ilgan kuniga belgilash".
            // If born 7-March.
            // Visits: Jan 7, Feb 7, March 7...

            // Implementation: We iterate ALL patients. 
            // For each patient, we generate the date list.
            // Date List: for m=0..11, Date = dayjs(targetYear, m, birthDay).
            // We add this patient to the result list.
            // To fit into "Monthly Sheets" paradigm:
            // Maybe we can just put everyone in "Umumiy" and empty monthly sheets? 
            // OR, better: We split them based on their Birth Month? 
            // No, let's split them by their FIRST visit date's month? (Jan).
            // If we have 1000 kids, putting all in Jan sheet is big.
            // BUT, if they are visited every month, they are technically "Active" every month.

            // Let's do this: Iterate all patients. Calculate their dates.
            // Add row to `rangeAllRows`.
            // Also add valid rows to monthly sheets based on their "First Visit Month" (usually Jan).
            // This keeps it consistent.

            processAll = true; // Ensure we don't slice pool

            const patientsByMonth = Array.from({ length: 12 }, () => []);

            patientPool.forEach(p => {
                const birthDay = p._birthDate.date(); // 1-31
                const birthMonth = p._birthDate.month(); // 0-11

                // Generate Dates
                let generatedDates = [];

                // We need 12 visits (or visitCount)
                // Interval logic is standard 12/Count.
                // Base date: Should be related to target year.
                // If born 7th. Visit 7th of Jan, 7th of Feb...

                const interval = 12 / visitCount;

                for (let v = 0; v < visitCount; v++) {
                    // Month index: (0 + v*interval) ? 
                    // Or should we start from Birth Month?
                    // Usually "Annual Plan" = Jan-Dec.
                    // So we start Jan.

                    let targetMonth = Math.floor(0 + v * interval); // 0, 1, 2... for 12 visits
                    // If visit count 4 (interval 3): 0, 3, 6, 9 (Jan, Apr, Jul, Oct)

                    // Construct Date
                    // Watch out for Feb 30 etc.
                    let d = dayjs().year(config.targetYear).month(targetMonth).date(birthDay);

                    // If invalid (e.g. Feb 30 -> March 2 automatically by dayjs? No, dayjs handles overflow typically)
                    // dayjs('2026-02-31') -> ?
                    // Better check end of month
                    const daysInMonth = dayjs().year(config.targetYear).month(targetMonth).daysInMonth();
                    if (birthDay > daysInMonth) {
                        d = dayjs().year(config.targetYear).month(targetMonth).date(daysInMonth);
                    }

                    // Adjust Working Day
                    d = findNextWorkingDay(d, config.holidays || [], config.saturdayWorking);

                    generatedDates.push(d);
                }

                // Sort
                generatedDates.sort((a, b) => a.valueOf() - b.valueOf());

                const pRow = { ...p };
                if (trimmedHeaders[birthColIdx - 1]) {
                    pRow[trimmedHeaders[birthColIdx - 1]] = p._birthDate.toDate();
                }
                delete pRow._birthDate;

                // Add dates to row
                generatedDates.forEach((date, idx) => {
                    if (idx < visitCount) {
                        pRow[visitKeys[idx]] = date.toDate();
                    }
                });

                rangeAllRows.push(pRow);

                // Determine which "Sheet" to put them in.
                // Let's put them in the sheet of the FIRST visit.
                const firstMonth = generatedDates[0].month(); // 0-11
                patientsByMonth[firstMonth].push(pRow);
            });

            // Mark all encountered as processed
            pIndex = patientPool.length;

            // Render Monthly Sheets
            for (let m = 0; m < 12; m++) {
                const pList = patientsByMonth[m];
                if (pList.length === 0) continue;

                const sheet = outWb.addWorksheet(MONTH_NAMES[m]);
                templateHeaderStyle.font = { name: 'Times New Roman', size: 11, bold: true };
                applyTemplateStyles(sheet, rangeHeaders, rangeColWidths, templateHeaderStyle);

                let mId = 1;
                const borderStyle = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };

                pList.forEach(r => {
                    const rowData = { ...r };
                    if (idColKey) rowData[idColKey] = mId++;

                    const values = rangeHeaders.map(h => rowData[h]);
                    const newRow = sheet.addRow(values);
                    for (let c = 1; c <= rangeHeaders.length; c++) {
                        const cell = newRow.getCell(c);
                        cell.border = borderStyle;
                        cell.font = { name: 'Times New Roman', size: 11 };
                        cell.alignment = { wrapText: true, vertical: 'middle', horizontal: 'left' };
                        if (cell.value instanceof Date) cell.numFmt = 'dd.mm.yyyy';
                    }
                });
            }

        }
        // ---------------------------------------------------------
        // LOGIC B: Standard Logic (Distribution)
        // ---------------------------------------------------------
        else {

            // If Auto Plan (Process All), we distribute them evenly across 12 months?
            // Or if logic calls for "Cycle", we might want to spread them?
            // If counts are provided (manual), we use them.
            // If counts are NOT provided (Auto Plan > 1 visit), we assume Uniform Distribution
            // OR we just dump everyone in Jan?
            // Better: Uniform Distribution across 12 months for the "First Visit".

            let countsToUse = range.counts;
            if (processAll) {
                // Distribute totalPatients across 12 months
                const base = Math.floor(totalPatients / 12);
                const rem = totalPatients % 12;
                countsToUse = [];
                for (let i = 0; i < 12; i++) {
                    countsToUse.push(i < rem ? base + 1 : base);
                }
            }

            // Proceed with Standard Month Loop
            for (let m = 0; m < 12; m++) {
                const targetCount = countsToUse[m] || 0;
                if (targetCount === 0) continue;

                const workingDays = getWorkingDays(config.targetYear, m, config.holidays || [], config.saturdayWorking);
                if (workingDays.length === 0) continue;

                const patientsForMonth = [];
                for (let i = 0; i < targetCount; i++) {
                    if (pIndex < patientPool.length) {
                        patientsForMonth.push(patientPool[pIndex]);
                        pIndex++;
                    } else break;
                }
                if (patientsForMonth.length === 0) continue;

                // Format Birth Date logic
                const birthKey = trimmedHeaders[birthColIdx - 1];

                // Distribute
                const N = patientsForMonth.length;
                const D = workingDays.length;
                const base = Math.floor(N / D);
                const remainder = N % D;

                let currentPatientIdx = 0;
                const monthlyRows = [];

                let monthlyIdCounter = 1;

                for (let d = 0; d < D; d++) {
                    const dayDate = workingDays[d];
                    const countForDay = d < remainder ? base + 1 : base;

                    for (let k = 0; k < countForDay; k++) {
                        if (currentPatientIdx >= patientsForMonth.length) break;

                        const p = patientsForMonth[currentPatientIdx];
                        const outRow = { ...p };

                        // Normalize Birth Date
                        if (birthKey && p._birthDate && p._birthDate.isValid()) {
                            outRow[birthKey] = p._birthDate.toDate();
                        }
                        delete outRow._birthDate;

                        // Renumber ID
                        if (idColKey) {
                            outRow[idColKey] = monthlyIdCounter++;
                        }

                        // --- MULTI-VISIT CIRCULAR LOGIC START ---
                        let generatedDates = [dayDate];

                        if (visitCount > 1) {
                            const interval = 12 / visitCount;
                            for (let v = 1; v < visitCount; v++) {
                                // Add interval
                                let nextDate = dayDate.add(interval * v, 'month');
                                // Ensure within target year (Circular)
                                if (nextDate.year() > config.targetYear) {
                                    nextDate = nextDate.subtract(1, 'year');
                                }
                                generatedDates.push(nextDate);
                            }
                        }

                        // Sort Dates (Jan -> Dec) because we wrapped around
                        generatedDates.sort((a, b) => a.valueOf() - b.valueOf());

                        // Validate Working Days
                        generatedDates = generatedDates.map(d => findNextWorkingDay(d, config.holidays || [], config.saturdayWorking));

                        // Assign to Columns
                        generatedDates.forEach((date, idx) => {
                            if (idx < visitCount) {
                                outRow[visitKeys[idx]] = date.toDate();
                            }
                        });
                        // --- MULTI-VISIT CIRCULAR LOGIC END ---

                        monthlyRows.push(outRow);
                        rangeAllRows.push(outRow);

                        currentPatientIdx++;
                    }
                }

                const sheet = outWb.addWorksheet(MONTH_NAMES[m]);
                templateHeaderStyle.font = { name: 'Times New Roman', size: 11, bold: true };
                applyTemplateStyles(sheet, rangeHeaders, rangeColWidths, templateHeaderStyle);

                const borderStyle = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };

                monthlyRows.forEach(row => {
                    const values = rangeHeaders.map(h => row[h]);
                    const newRow = sheet.addRow(values);
                    for (let c = 1; c <= rangeHeaders.length; c++) {
                        const cell = newRow.getCell(c);
                        cell.border = borderStyle;
                        cell.font = { name: 'Times New Roman', size: 11 };
                        cell.alignment = { wrapText: true, vertical: 'middle', horizontal: 'left' };
                        if (cell.value instanceof Date) cell.numFmt = 'dd.mm.yyyy';
                    }
                });
            }
        }

        // Umumiy Sheet (Range specific)
        if (rangeAllRows.length > 0) {
            const uSheet = outWb.addWorksheet("Umumiy");
            applyTemplateStyles(uSheet, rangeHeaders, rangeColWidths, templateHeaderStyle);

            let uIdCounter = 1;
            const uRows = rangeAllRows.map(r => {
                const newR = { ...r };
                if (idColKey) newR[idColKey] = uIdCounter++;
                return newR;
            });

            uRows.forEach(row => {
                const values = rangeHeaders.map(h => row[h]);
                const newRow = uSheet.addRow(values);
                const borderStyle = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
                for (let c = 1; c <= rangeHeaders.length; c++) {
                    const cell = newRow.getCell(c);
                    cell.border = borderStyle;
                    cell.font = { name: 'Times New Roman', size: 11 };
                    cell.alignment = { wrapText: true, vertical: 'middle', horizontal: 'left' };
                    if (cell.value instanceof Date) {
                        cell.numFmt = 'dd.mm.yyyy';
                    }

                    // --- HIGHLIGHT YELLOW FOR UNPLANNED ---
                    if (row._isUnplanned) {
                        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } }; // Yellow
                    }
                }
            });

            // Handle Unplanned (Remaining in Pool)
            // pIndex points to the next available patient. 
            // So everyone from pIndex to end was NOT processed.
            if (pIndex < patientPool.length) {
                const remainingPatients = patientPool.slice(pIndex);
                let uIdCounterUnplanned = uIdCounter; // Continue counter? Yes.

                remainingPatients.forEach(p => {
                    const outRow = { ...p };
                    // Format Birth Date
                    const birthKey = trimmedHeaders[birthColIdx - 1];
                    if (birthKey && p._birthDate && p._birthDate.isValid()) {
                        outRow[birthKey] = p._birthDate.toDate();
                    }
                    delete outRow._birthDate;

                    if (idColKey) {
                        outRow[idColKey] = uIdCounterUnplanned++;
                    }

                    // Mark as unplanned
                    outRow._isUnplanned = true;

                    // Add to array for consolidated
                    rangeAllRows.push(outRow);

                    // Add to Sheet
                    const values = rangeHeaders.map(h => outRow[h]);
                    const newRow = uSheet.addRow(values);
                    const borderStyle = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };

                    for (let c = 1; c <= rangeHeaders.length; c++) {
                        const cell = newRow.getCell(c);
                        cell.border = borderStyle;
                        cell.font = { name: 'Times New Roman', size: 11 };
                        cell.alignment = { wrapText: true, vertical: 'middle', horizontal: 'left' };
                        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } }; // Yellow

                        if (cell.value instanceof Date) {
                            cell.numFmt = 'dd.mm.yyyy';
                        }
                    }
                });
            }
        }

        // Store for global
        allRangesRows.push({ rows: rangeAllRows, visitCount: visitCount, rangeHeaders: rangeHeaders });

        const buffer = await outWb.xlsx.writeBuffer();
        archive.append(buffer, { name: rangeName });
    }

    // Consolidated Global Sheet
    if (allRangesRows.length > 0) {
        const uWb = new ExcelJS.Workbook();
        const uSheet = uWb.addWorksheet("Umumiy Reja");

        // Determine Global Headers based on Max Visits - REMOVED 'Tashrif oyi'
        let globalHeaders = [...trimmedHeaders];
        if (maxVisitsGlobal === 1) {
            globalHeaders.push('Tashrif sanasi');
        } else {
            for (let v = 1; v <= maxVisitsGlobal; v++) {
                globalHeaders.push(`${v}-tashrif sanasi`);
            }
        }

        // REMOVED 15 for 'Tashrif oyi'
        const globalColWidths = [...trimmedWidths];
        for (let i = 0; i < maxVisitsGlobal; i++) globalColWidths.push(15);

        applyTemplateStyles(uSheet, globalHeaders, globalColWidths, templateHeaderStyle);

        // Merge all ranges rows and re-normalize headers
        let gIdCounter = 1;
        const allFlatRows = [];

        allRangesRows.forEach(group => {
            group.rows.forEach(r => {
                // Map range-specific keys (1-tashrif, 2-tashrif) to Global Keys
                // If this range has 1 visit, map 'Tashrif sanasi' to '1-tashrif sanasi' IF global has > 1
                const newR = { ...r };

                if (group.visitCount === 1 && maxVisitsGlobal > 1) {
                    newR['1-tashrif sanasi'] = r['Tashrif sanasi'];
                    delete newR['Tashrif sanasi'];
                }

                if (idColKey) newR[idColKey] = gIdCounter++;
                allFlatRows.push(newR);
            });
        });

        const borderStyle = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };

        allFlatRows.forEach(row => {
            const values = globalHeaders.map(h => row[h]);
            const newRow = uSheet.addRow(values);
            for (let c = 1; c <= globalHeaders.length; c++) {
                const cell = newRow.getCell(c);
                cell.border = borderStyle;
                cell.font = { name: 'Times New Roman', size: 11 };
                cell.alignment = { wrapText: true, vertical: 'middle', horizontal: 'left' };
                if (cell.value instanceof Date) {
                    cell.numFmt = 'dd.mm.yyyy';
                }

                // --- HIGHLIGHT YELLOW FOR UNPLANNED ---
                if (row._isUnplanned) {
                    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } }; // Yellow
                }
            }
        });
        const uBuff = await uWb.xlsx.writeBuffer();
        archive.append(uBuff, { name: "Umumiy_Reja.xlsx" });
    }

    await archive.finalize();

    return new Promise((resolve, reject) => {
        outputStream.on('close', () => resolve({ downloadUrl: `/output/${zipName}`, warnings }));
        outputStream.on('error', reject);
    });
}

/**
 * Analyze Excel
 */
async function analyzeExcel(buffer) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);
    const sheet = workbook.worksheets[0];

    // Detect Header Row
    const headerRowIdx = findHeaderRow(sheet);

    const headers = [];
    sheet.getRow(headerRowIdx).eachCell({ includeEmpty: true }, (cell, col) => {
        let val = cell.value;
        if (typeof val === 'object' && val !== null) {
            if (val.richText) val = val.richText.map(t => t.text).join('');
            else if (val.text) val = val.text;
        }
        headers[col - 1] = val ? String(val) : '';
    });

    const birthColIdx = findBirthDateColumn(headers);
    let totalPatients = 0;

    // Detect Name Column for Gender
    const keywordsName = ['f.i.sh', 'fish', 'ism', 'name', 'familiya'];
    let nameColKey = null;
    let nameColIdx = -1;

    // We need index for cell access
    headers.forEach((h, i) => {
        if (!nameColKey && keywordsName.some(k => h.toLowerCase().includes(k))) {
            nameColKey = h;
            nameColIdx = i + 1;
        }
    });
    if (!nameColKey && headers.length > 1) {
        nameColIdx = 2; // Fallback
    }

    const yearCounts = {}; // { 2000: { total: 0, male: 0, female: 0 } }

    sheet.eachRow((row, rowNum) => {
        if (rowNum <= headerRowIdx) return; // Skip header and above
        const cell = row.getCell(birthColIdx);
        const date = parseDate(cell.value);
        if (date.isValid()) {
            const y = date.year();
            if (y >= 1900 && y < dayjs().year() + 1) {
                if (!yearCounts[y]) yearCounts[y] = { total: 0, male: 0, female: 0 };

                yearCounts[y].total++;

                // Gender Check
                let isFemale = false;
                if (nameColIdx > 0) {
                    const val = row.getCell(nameColIdx).value;
                    const str = val ? String(val).trim() : '';
                    const firstWord = str.split(' ')[0];
                    if (firstWord && firstWord.toLowerCase().endsWith('a')) {
                        isFemale = true;
                    }
                }

                if (isFemale) yearCounts[y].female++;
                else yearCounts[y].male++;

                totalPatients++;
            }
        }
    });

    return { yearCounts, totalPatients };
}

module.exports = { processExcel, analyzeExcel };
