package nodes;

import axiom.AxiomContext;
import gen.Messages.OfficeFile;
import gen.Messages.TextResult;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.Map;

public class ExtractSpreadsheetText {

    private static final long MAX_TEXT_CHARS = 20_000_000; // 20M chars (~20-40MB)

    /**
     * Extract all cell text from every sheet in a spreadsheet as one plain-
     * text document — cells tab-separated within a row, rows newline-
     * separated, a blank line between sheets. Uses each cell's Excel-
     * formatted display value (dates/numbers rendered as Excel would show
     * them), including formula cells' cached results. Bounded by sheet
     * count, per-sheet row/column extent, and total output size; truncated
     * reports if any bound was hit.
     *
     * @param ax    The AxiomContext: logging, secrets, reflection, mutation.
     * @param input The decoded OfficeFile for this invocation.
     */
    public static TextResult extractSpreadsheetText(AxiomContext ax, OfficeFile input) {
        ax.log().info("extractSpreadsheetText handling", Map.of());
        try {
            byte[] bytes = OfficeUtil.loadBytes(input);
            try (Workbook wb = OfficeUtil.openWorkbook(bytes, input.getPassword())) {
                DataFormatter formatter = new DataFormatter();
                StringBuilder sb = new StringBuilder();
                boolean truncated = false;
                int sheetsProcessed = 0;
                int numSheets = Math.min(wb.getNumberOfSheets(), OfficeUtil.MAX_SHEETS_FOR_TEXT);
                truncated |= wb.getNumberOfSheets() > numSheets;

                outer:
                for (int si = 0; si < numSheets; si++) {
                    Sheet sheet = wb.getSheetAt(si);
                    if (si > 0) sb.append('\n');
                    int firstRow = Math.max(sheet.getFirstRowNum(), 0);
                    int lastRowAvail = sheet.getLastRowNum();
                    int lastRow = Math.min(lastRowAvail, firstRow + OfficeUtil.HARD_MAX_ROWS - 1);
                    if (lastRowAvail > lastRow) truncated = true;
                    for (int r = firstRow; r <= lastRow; r++) {
                        org.apache.poi.ss.usermodel.Row row = sheet.getRow(r);
                        if (row == null) { sb.append('\n'); continue; }
                        int lastCellAvail = row.getLastCellNum() - 1;
                        int lastCol = Math.min(lastCellAvail, OfficeUtil.HARD_MAX_COLS - 1);
                        if (lastCellAvail > lastCol) truncated = true;
                        for (int c = 0; c <= lastCol; c++) {
                            if (c > 0) sb.append('\t');
                            org.apache.poi.ss.usermodel.Cell cell = row.getCell(c);
                            if (cell != null) {
                                try {
                                    sb.append(OfficeUtil.formatCellText(cell, formatter));
                                } catch (Exception ignored) {
                                    // best-effort formatting
                                }
                            }
                            if (sb.length() > MAX_TEXT_CHARS) {
                                truncated = true;
                                break outer;
                            }
                        }
                        sb.append('\n');
                    }
                    sheetsProcessed++;
                }
                return TextResult.newBuilder()
                        .setText(sb.toString())
                        .setUnitCount(sheetsProcessed)
                        .setTruncated(truncated)
                        .build();
            }
        } catch (OfficeUtil.OfficeError e) {
            return TextResult.newBuilder().setError(e.getMessage()).build();
        } catch (Exception e) {
            return TextResult.newBuilder().setError("could not extract text: " + e.getMessage()).build();
        }
    }
}
