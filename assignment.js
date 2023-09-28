const XLSX = require('xlsx');
const workbook = XLSX.readFile('Assignment_Timecard.xlsx');
const sheet = workbook.Sheets[workbook.SheetNames[0]];

let previousEmployee = null;
let previousShiftEnd = null;
let consecutiveDays = 0;

for (const cell in sheet) {
    if (cell.startsWith('A')) continue; // Skip the header row

    const cellValue = sheet[cell].v || sheet[cell].w || null;
    if (cellValue === null) {
        console.log(`Cell value not found for cell ${cell}`);
        continue;
    }

    const employee = cellValue;
    const timeIn = new Date(sheet['B' + cell].v || sheet['B' + cell].w || null);
    const timeOut = new Date(sheet['C' + cell].v || sheet['C' + cell].w || null);

    if (previousEmployee && previousEmployee === employee) {
        const timeDiff = (timeIn - new Date(previousShiftEnd)) / (60 * 60 * 1000);

        if (timeDiff > 1 && timeDiff < 10) {
            console.log(`Employee: ${employee}, Position: ${sheet['A' + cell].v}, Reason: Less than 10 hours between shifts`);
        }

        const shiftDuration = (timeOut - timeIn) / (60 * 60 * 1000);
        if (shiftDuration > 14) {
            console.log(`Employee: ${employee}, Position: ${sheet['A' + cell].v}, Reason: Shift duration > 14 hours`);
        }

        if (consecutiveDays === 6) {
            console.log(`Employee: ${employee}, Position: ${sheet['A' + cell].v}, Reason: Worked for 7 consecutive days`);
        }

        if (isConsecutiveDay(new Date(previousShiftEnd), timeIn)) {
            consecutiveDays++;
        } else {
            consecutiveDays = 0;
        }
    }

    previousEmployee = employee;
    previousShiftEnd = timeOut;
}

function isConsecutiveDay(date1, date2) {
    return (date2 - date1) / (24 * 60 * 60 * 1000) === 1;
}
