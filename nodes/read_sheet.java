package nodes;

import axiom.AxiomContext;
import gen.Messages.Cell;
import gen.Messages.GridResult;
import gen.Messages.ReadSheetInput;
import gen.Messages.Row;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.Map;

public class ReadSheet {

    /**
     * Read a spreadsheet sheet's used cell grid — every row from the sheet's
     * first used row through its last, each cell typed (STRING/NUMERIC/
     * BOOLEAN/FORMULA/BLANK/ERROR) with its value, formula text (for FORMULA
     * cells, reporting the cell's last-saved cached result — never
     * live-recomputed), and Excel-formatted display string. Bounded by
     * max_rows/max_cols (0 = default of 10,000 rows / 1,000 columns; pass a
     * larger value for more); truncated reports whether the sheet's real
     * extent exceeded the cap that was applied. Select the sheet by name
     * (checked first) or 0-based index.
     *
     * @param ax    The AxiomContext: logging, secrets, reflection, mutation.
     * @param input The decoded ReadSheetInput for this invocation.
     */
    public static GridResult readSheet(AxiomContext ax, ReadSheetInput input) {
        ax.log().info("readSheet handling", Map.of());
        try {
            byte[] bytes = OfficeUtil.loadBytes(input.getFile());
            try (Workbook wb = OfficeUtil.openWorkbook(bytes, input.getFile().getPassword())) {
                Sheet sheet = OfficeUtil.resolveSheet(wb, input.getSheetName(), input.getSheetIndex());

                int maxRows = withDefault(input.getMaxRows(), OfficeUtil.DEFAULT_MAX_ROWS);
                int maxCols = withDefault(input.getMaxCols(), OfficeUtil.DEFAULT_MAX_COLS);

                DataFormatter formatter = new DataFormatter();
                GridResult.Builder result = GridResult.newBuilder().setSheetName(sheet.getSheetName());

                int firstRow = sheet.getFirstRowNum();
                int lastRowAvailable = sheet.getLastRowNum();
                if (firstRow < 0) firstRow = 0;
                int lastRowToRead = Math.min(lastRowAvailable, firstRow + maxRows - 1);
                boolean truncated = lastRowAvailable > lastRowToRead;

                int widestCol = 0;
                for (int r = firstRow; r <= lastRowToRead; r++) {
                    org.apache.poi.ss.usermodel.Row poiRow = sheet.getRow(r);
                    Row.Builder rowBuilder = Row.newBuilder();
                    int lastCellAvailable = poiRow == null ? -1 : poiRow.getLastCellNum() - 1;
                    int lastColToRead = Math.min(lastCellAvailable, maxCols - 1);
                    if (lastCellAvailable > lastColToRead) truncated = true;
                    for (int c = 0; c <= lastColToRead; c++) {
                        org.apache.poi.ss.usermodel.Cell poiCell = poiRow == null ? null : poiRow.getCell(c);
                        rowBuilder.addCells(OfficeUtil.cellToProto(poiCell, formatter, r, c));
                    }
                    widestCol = Math.max(widestCol, lastColToRead + 1);
                    result.addRows(rowBuilder.build());
                }
                result.setRowCount(lastRowToRead - firstRow + 1 < 0 ? 0 : lastRowToRead - firstRow + 1);
                result.setColCount(Math.max(widestCol, 0));
                result.setTruncated(truncated);
                return result.build();
            }
        } catch (OfficeUtil.OfficeError e) {
            return GridResult.newBuilder().setError(e.getMessage()).build();
        } catch (Exception e) {
            return GridResult.newBuilder().setError("could not read sheet: " + e.getMessage()).build();
        }
    }

    static int withDefault(int requested, int defaultVal) {
        return requested <= 0 ? defaultVal : requested;
    }
}
