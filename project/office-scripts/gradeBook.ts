/**
 * Office Script: Grade Book Automation
 * Excel Online (Microsoft 365)
 * ═══════════════════════════════════════════════════════════════
 * 
 * HOW TO USE:
 * 1. Open your workbook in Excel Online (office.com or onedrive.live.com)
 * 2. Go to the "Automate" tab
 * 3. Click "Code Editor" → "New Script"
 * 4. Paste this entire script
 * 5. Click "Run" or assign to a button / Power Automate flow
 */

function main(workbook: ExcelScript.Workbook) {
    // ═══════════════════════════════════════════════════════════
    // CONFIGURATION
    // ═══════════════════════════════════════════════════════════
    const SHEET_NAME = "Grades";
    const FIRST_DATA_ROW = 2;
    const SUBJECT_COLS = 5;  // C through G (English, Math, Science, ICT, Social)
    const COL_TOTAL = 8;     // H
    const COL_AVG = 9;       // I
    const COL_GRADE = 10;    // J
    const COL_REMARKS = 11;   // K

    // ═══════════════════════════════════════════════════════════
    // GET SHEET & RANGE
    // ═══════════════════════════════════════════════════════════
    const sheet = workbook.getWorksheet(SHEET_NAME);
    if (!sheet) {
        console.log(`Sheet "${SHEET_NAME}" not found!`);
        return;
    }

    const usedRange = sheet.getUsedRange();
    const lastRow = usedRange.getRowCount() + 1;

    console.log(`Processing ${lastRow - 1} students...`);

    // ═══════════════════════════════════════════════════════════
    // PROCESS EACH STUDENT
    // ═══════════════════════════════════════════════════════════
    for (let row = FIRST_DATA_ROW; row <= lastRow; row++) {
        // Read subject scores
        const scoreRange = sheet.getRange(`C${row}:G${row}`);
        const scores = scoreRange.getValues()[0] as number[];

        // Calculate Total
        let total = 0;
        for (let i = 0; i < scores.length; i++) {
            total += scores[i] as number;
        }

        // Calculate Average
        const average = Math.round(total / scores.length * 100) / 100;

        // Determine Letter Grade
        let grade: string;
        let remarks: string;
        if (average >= 80) {
            grade = "A";
            remarks = "Excellent — Outstanding performance";
        } else if (average >= 70) {
            grade = "B";
            remarks = "Very Good — Above average performance";
        } else if (average >= 60) {
            grade = "C";
            remarks = "Good — Satisfactory performance";
        } else if (average >= 50) {
            grade = "D";
            remarks = "Pass — Needs improvement";
        } else if (average >= 40) {
            grade = "E";
            remarks = "Weak — Requires significant improvement";
        } else {
            grade = "F";
            remarks = "Fail — Must repeat examination";
        }

        // Write results
        sheet.getCell(row - 1, COL_TOTAL - 1).setValue(total);
        const avgCell = sheet.getCell(row - 1, COL_AVG - 1);
        avgCell.setValue(average);
        avgCell.getNumberFormat()[0][0] = "0.00";

        const gradeCell = sheet.getCell(row - 1, COL_GRADE - 1);
        gradeCell.setValue(grade);
        sheet.getCell(row - 1, COL_REMARKS - 1).setValue(remarks);

        // Color-code the grade cell
        const gradeFillColors: Record<string, string> = {
            "A": "00B050", "B": "92D050", "C": "FFFF00",
            "D": "FFC000", "E": "FF8000", "F": "FF0000"
        };
        const fontColors: Record<string, string> = {
            "A": "FFFFFF", "B": "000000", "C": "000000",
            "D": "000000", "E": "FFFFFF", "F": "FFFFFF"
        };
        gradeCell.getFormat().getFill().setColor(gradeFillColors[grade]);
        gradeCell.getFormat().getFont().setColor(fontColors[grade]);
        gradeCell.getFormat().getFont().setBold(true);
        gradeCell.getFormat().setHorizontalAlignment("center");
    }

    // ═══════════════════════════════════════════════════════════
    // CLASS STATISTICS
    // ═══════════════════════════════════════════════════════════
    const statsRow = lastRow + 3;
    const statsCell = sheet.getCell(statsRow - 1, 0);
    statsCell.setValue("CLASS STATISTICS");
    statsCell.getFormat().getFont().setBold(true);
    statsCell.getFormat().getFont().setSize(13);

    // Subject averages
    const subjects = ["English", "Mathematics", "Science", "ICT", "Social Studies"];
    sheet.getCell(statsRow, 1).setValue("Subject Averages:");
    sheet.getCell(statsRow, 1).getFormat().getFont().setBold(true);

    for (let i = 0; i < subjects.length; i++) {
        const col = String.fromCharCode(67 + i); // C through G
        const range = sheet.getRange(`${col}${FIRST_DATA_ROW}:${col}${lastRow}`);
        const values = range.getValues();
        let sum = 0;
        let count = 0;
        for (let j = 0; j < values.length; j++) {
            const val = values[j][0] as number;
            if (!isNaN(val)) {
                sum += val;
                count++;
            }
        }
        const avg = Math.round(sum / count * 100) / 100;
        sheet.getCell(statsRow + 1, 1 + i).setValue(subjects[i]);
        sheet.getCell(statsRow + 2, 1 + i).setValue(avg);
    }

    // Overall average
    const avgRange = sheet.getRange(`I${FIRST_DATA_ROW}:I${lastRow}`);
    const avgValues = avgRange.getValues();
    let totalAvg = 0;
    const studentCount = avgValues.length;
    for (let i = 0; i < avgValues.length; i++) {
        totalAvg += avgValues[i][0] as number;
    }
    const classAvg = Math.round(totalAvg / studentCount * 100) / 100;

    sheet.getCell(statsRow + 4, 1).setValue("Class Average:");
    sheet.getCell(statsRow + 4, 1).getFormat().getFont().setBold(true);
    sheet.getCell(statsRow + 4, 2).setValue(classAvg);

    // Pass rate
    const gradeRange = sheet.getRange(`J${FIRST_DATA_ROW}:J${lastRow}`);
    const gradeValues = gradeRange.getValues();
    let passCount = 0;
    for (let i = 0; i < gradeValues.length; i++) {
        if (gradeValues[i][0] !== "F") passCount++;
    }
    const passRate = Math.round(passCount / studentCount * 1000) / 10;

    sheet.getCell(statsRow + 5, 1).setValue("Pass Rate:");
    sheet.getCell(statsRow + 5, 1).getFormat().getFont().setBold(true);
    sheet.getCell(statsRow + 5, 2).setValue(passRate + "%");

    console.log(`✅ Grading complete! ${studentCount} students processed.`);
    console.log(`   Class Average: ${classAvg}`);
    console.log(`   Pass Rate: ${passRate}%`);
}
