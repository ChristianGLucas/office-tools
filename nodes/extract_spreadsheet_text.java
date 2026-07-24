package nodes;

import axiom.AxiomContext;
import gen.Messages.OfficeFile;
import gen.Messages.TextResult;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.Map;

public class ExtractSpreadsheetText {

    /**
     * Extract all cell text from every sheet in a spreadsheet as one plain-
     * text document — cells tab-separated within a row, rows newline-
     * separated, a blank line between sheets. Uses each cell's Excel-
     * formatted display value (dates/numbers rendered as Excel would show
     * them), including formula cells' cached results.
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
                int sheetsProcessed = 0;
                int numSheets = wb.getNumberOfSheets();

                for (int si = 0; si < numSheets; si++) {
                    Sheet sheet = wb.getSheetAt(si);
                    if (si > 0) sb.append('\n');
                    int firstRow = Math.max(sheet.getFirstRowNum(), 0);
                    int lastRow = sheet.getLastRowNum();
                    for (int r = firstRow; r <= lastRow; r++) {
                        org.apache.poi.ss.usermodel.Row row = sheet.getRow(r);
                        if (row == null) { sb.append('\n'); continue; }
                        int lastCol = row.getLastCellNum() - 1;
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
                        }
                        sb.append('\n');
                    }
                    sheetsProcessed++;
                }
                return TextResult.newBuilder()
                        .setText(sb.toString())
                        .setUnitCount(sheetsProcessed)
                        .build();
            }
        } catch (OfficeUtil.OfficeError e) {
            return TextResult.newBuilder().setError(e.getMessage()).build();
        } catch (Exception e) {
            return TextResult.newBuilder().setError("could not extract text: " + e.getMessage()).build();
        }
    }
}
