package nodes;

import axiom.AxiomContext;
import gen.Messages.GridResult;
import gen.Messages.ReadRangeInput;
import gen.Messages.Row;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.Map;

public class ReadRange {

    /**
     * Read exactly one explicit A1-notation cell range from a sheet — e.g.
     * "B2:D10" or a single cell "B2" — returning the same typed-cell grid
     * shape as ReadSheet, but addressed by range rather than by bounded
     * bulk read. Apache POI's CellRangeAddress owns the A1-notation parsing.
     * A malformed range reference or an out-of-bounds sheet returns a
     * structured error.
     *
     * @param ax    The AxiomContext: logging, secrets, reflection, mutation.
     * @param input The decoded ReadRangeInput for this invocation.
     */
    public static GridResult readRange(AxiomContext ax, ReadRangeInput input) {
        ax.log().info("readRange handling", Map.of());
        try {
            if (input.getRangeRef() == null || input.getRangeRef().isEmpty()) {
                return GridResult.newBuilder().setError("range_ref is required").build();
            }
            CellRangeAddress range;
            try {
                range = CellRangeAddress.valueOf(input.getRangeRef());
            } catch (Exception e) {
                return GridResult.newBuilder().setError("invalid range_ref: " + input.getRangeRef()).build();
            }
            if (range.getFirstRow() < 0 || range.getFirstColumn() < 0
                    || range.getLastRow() < 0 || range.getLastColumn() < 0) {
                return GridResult.newBuilder()
                        .setError("invalid range_ref (must name both a column and a row on each side): "
                                + input.getRangeRef())
                        .build();
            }

            byte[] bytes = OfficeUtil.loadBytes(input.getFile());
            try (Workbook wb = OfficeUtil.openWorkbook(bytes, input.getFile().getPassword())) {
                Sheet sheet = OfficeUtil.resolveSheet(wb, input.getSheetName(), input.getSheetIndex());

                int firstRow = range.getFirstRow();
                int lastRow = range.getLastRow();
                int firstCol = range.getFirstColumn();
                int lastCol = range.getLastColumn();

                DataFormatter formatter = new DataFormatter();
                GridResult.Builder result = GridResult.newBuilder().setSheetName(sheet.getSheetName());
                for (int r = firstRow; r <= lastRow; r++) {
                    org.apache.poi.ss.usermodel.Row poiRow = sheet.getRow(r);
                    Row.Builder rowBuilder = Row.newBuilder();
                    for (int c = firstCol; c <= lastCol; c++) {
                        org.apache.poi.ss.usermodel.Cell poiCell = poiRow == null ? null : poiRow.getCell(c);
                        rowBuilder.addCells(OfficeUtil.cellToProto(poiCell, formatter, r, c));
                    }
                    result.addRows(rowBuilder.build());
                }
                result.setRowCount(lastRow - firstRow + 1);
                result.setColCount(lastCol - firstCol + 1);
                return result.build();
            }
        } catch (OfficeUtil.OfficeError e) {
            return GridResult.newBuilder().setError(e.getMessage()).build();
        } catch (Exception e) {
            return GridResult.newBuilder().setError("could not read range: " + e.getMessage()).build();
        }
    }
}
