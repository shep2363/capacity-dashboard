/**
 * CleanupWorkbook — Excel Office Script
 *
 * Run this script in Excel Online (Automate → New Script) before uploading
 * the workbook to the Capacity Dashboard.
 *
 * What it does:
 *  1. Normalises Project (col 0), Resource (col 9), Hours (col 4), and Name (col 1) values.
 *  2. Deletes rows where the resource is blank, bold, struck-through, or in the
 *     blocked list (Inventory, Purchasing, Detailing).
 *  3. Moves the Job # column (col 0) to the end (col 10) then removes the original.
 *  4. Removes the two Baseline columns that appear at index 4 after the shift.
 */

function main(workbook: ExcelScript.Workbook) {
    const sheet = workbook.getActiveWorksheet();

    const used = sheet.getUsedRange();
    if (!used) return;

    const rowCount = used.getRowCount();
    const colCount = used.getColumnCount();

    // Helpers (Office Scripts-safe types; no `any`)
    const asText = (v: string | number | boolean | null): string => {
        if (v === null) return "";
        return v.toString();
    };

    const isBold = (r: number, c: number): boolean =>
        sheet.getCell(r, c).getFormat().getFont().getBold();

    const isStrikethrough = (r: number, c: number): boolean =>
        sheet.getCell(r, c).getFormat().getFont().getStrikethrough();

    // 1) Cleanup/normalize fields (bottom-up)
    for (let i = rowCount - 1; i >= 0; i--) {
        const project = asText(sheet.getCell(i, 0).getValue() as string | number | boolean | null);
        const resource = asText(sheet.getCell(i, 9).getValue() as string | number | boolean | null);
        const hours = asText(sheet.getCell(i, 4).getValue() as string | number | boolean | null);
        const name = asText(sheet.getCell(i, 1).getValue() as string | number | boolean | null);

        sheet.getCell(i, 0).setValue(project.substring(0, 5));
        sheet.getCell(i, 9).setValue(resource.split("[")[0].trim());
        sheet.getCell(i, 4).setValue(hours.split(" ")[0]);
        sheet.getCell(i, 1).setValue(name.trim());
    }

    // 2) Delete unwanted rows (Assembly is kept)
    const blockedResources = new Set<string>(["Inventory", "Purchasing", "Detailing"]);

    for (let i = rowCount - 1; i >= 0; i--) {
        const resourceVal = asText(sheet.getCell(i, 9).getValue() as string | number | boolean | null).trim();

        const shouldDelete =
            resourceVal === "" ||
            isBold(i, 9) ||
            isStrikethrough(i, 9) ||
            blockedResources.has(resourceVal);

        if (shouldDelete) {
            sheet
                .getRangeByIndexes(i, 0, 1, colCount)
                .delete(ExcelScript.DeleteShiftDirection.up);
        }
    }

    // Refresh used range after deletions
    const used2 = sheet.getUsedRange();
    if (!used2) return;

    const rowCount2 = used2.getRowCount();

    // 3) Move job # (col 0) to end (col 10), then delete original first column
    sheet
        .getRangeByIndexes(0, 0, rowCount2, 1)
        .moveTo(sheet.getRangeByIndexes(0, 10, rowCount2, 1));

    sheet
        .getRangeByIndexes(0, 0, rowCount2, 1)
        .delete(ExcelScript.DeleteShiftDirection.left);

    // 4) Remove baseline start and finish columns (twice at index 4)
    sheet
        .getRangeByIndexes(0, 4, rowCount2, 1)
        .delete(ExcelScript.DeleteShiftDirection.left);

    sheet
        .getRangeByIndexes(0, 4, rowCount2, 1)
        .delete(ExcelScript.DeleteShiftDirection.left);
}
