function main(workbook: ExcelScript.Workbook) {
    const ws = workbook.getWorksheet("Summary");
    const range = ws.getRange("A3:D14");

    // Use text so dates are stable and match what you see in the cell
    const rows = range.getTexts(); // 12 x 4: [Date, Brand, Created, Processed]

    const isIsoDate = (s: string) => /^\d{4}-\d{2}-\d{2}$/.test(s.trim());

    // Convert "455", "1,234", "" safely to number or null
    const toNumberOrNull = (s: string): number | null => {
        const t = (s ?? "").trim();
        if (!t) return null;
        const cleaned = t.replace(/,/g, "");
        const n = Number(cleaned);
        return Number.isFinite(n) ? n : null;
    };

    let currentDate: string | null = null;

    const output: { date: string; brand: string; created: number | null; processed: number | null }[] = [];

    for (let i = 0; i < rows.length; i++) {
        const dateCell = (rows[i][0] ?? "").trim();   // column R
        const brandCell = (rows[i][1] ?? "").trim();  // column S
        const createdCell = (rows[i][2] ?? "").trim();// column T
        const processedCell = (rows[i][3] ?? "").trim();// column U

        // If this row has a date, update the current date
        if (dateCell && isIsoDate(dateCell)) {
            currentDate = dateCell;
        }

        // Skip rows until we’ve found the first date block
        if (!currentDate) continue;

        // Skip empty brand rows (if any)
        if (!brandCell) continue;

        output.push({
            date: currentDate,
            brand: brandCell,
            created: toNumberOrNull(createdCell),
            processed: toNumberOrNull(processedCell)
        });
    }

    console.log(output);
    return output;
}
