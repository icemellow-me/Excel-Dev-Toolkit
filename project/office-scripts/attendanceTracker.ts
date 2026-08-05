/**
 * Office Script: Attendance Tracker
 * Excel Online (Microsoft 365)
 * ═══════════════════════════════════════════════════════════════
 * Flags students with low attendance and generates a summary.
 */

function main(workbook: ExcelScript.Workbook) {
    const SHEET_NAME = "Students";
    const FIRST_ROW = 2;
    const COL_ID = 1;       // A
    const COL_NAME = 2;     // B
    const COL_ATTENDANCE = 12; // L

    const sheet = workbook.getWorksheet(SHEET_NAME);
    if (!sheet) {
        console.log(`Sheet "${SHEET_NAME}" not found!`);
        return;
    }

    const usedRange = sheet.getUsedRange();
    const lastRow = usedRange.getRowCount() + 1;

    console.log(`Checking attendance for ${lastRow - 1} students...\n`);

    let lowAttendance: string[] = [];
    let goodAttendance: string[] = [];
    let totalAttendance = 0;
    let count = 0;

    for (let row = FIRST_ROW; row <= lastRow; row++) {
        const id = sheet.getCell(row - 1, COL_ID - 1).getValue() as string;
        const name = sheet.getCell(row - 1, COL_NAME - 1).getValue() as string;
        const attendance = sheet.getCell(row - 1, COL_ATTENDANCE - 1).getValue() as number;

        totalAttendance += attendance;
        count++;

        if (attendance < 90) {
            lowAttendance.push(`${id} — ${name} (${attendance}%) ⚠️`);
            // Highlight cell red
            sheet.getCell(row - 1, COL_ATTENDANCE - 1)
                .getFormat().getFill().setColor("FF0000");
            sheet.getCell(row - 1, COL_ATTENDANCE - 1)
                .getFormat().getFont().setColor("FFFFFF");
        } else if (attendance >= 95) {
            goodAttendance.push(`${id} — ${name} (${attendance}%) ✅`);
            // Highlight cell green
            sheet.getCell(row - 1, COL_ATTENDANCE - 1)
                .getFormat().getFill().setColor("00B050");
        }
    }

    const avgAttendance = Math.round(totalAttendance / count * 10) / 10;

    console.log("──────────────────────────────────────────");
    console.log(`📊 ATTENDANCE REPORT`);
    console.log(`   Total Students: ${count}`);
    console.log(`   Average Attendance: ${avgAttendance}%`);
    console.log(`   High Attendance (≥95%): ${goodAttendance.length}`);
    console.log(`   Low Attendance (<90%): ${lowAttendance.length}`);
    console.log("──────────────────────────────────────────\n");

    if (lowAttendance.length > 0) {
        console.log("⚠️  STUDENTS NEEDING ATTENTION:");
        for (const s of lowAttendance) {
            console.log(`   ${s}`);
        }
    }

    if (goodAttendance.length > 0) {
        console.log("\n🌟 EXCELLENT ATTENDANCE:");
        for (const s of goodAttendance) {
            console.log(`   ${s}`);
        }
    }
}
