package nodes;

import axiom.AxiomContext;
import gen.Messages.CellResult;
import gen.Messages.ReadCellInput;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;

import java.util.Map;

public class ReadCell {

    /**
     * Look up one cell by A1-notation reference (e.g. "C7") on a sheet.
     * Returns the same typed Cell shape as ReadSheet/ReadRange rows. A
     * reference past the sheet's used range is not an error: found is
     * false and cell is returned with type BLANK. An unparseable
     * cell_ref, or the sheet not existing, returns a structured error.
     *
     * @param ax    The AxiomContext: logging, secrets, reflection, mutation.
     * @param input The decoded ReadCellInput for this invocation.
     */
    public static CellResult readCell(AxiomContext ax, ReadCellInput input) {
        ax.log().info("readCell handling", Map.of());
        try {
            if (input.getCellRef() == null || input.getCellRef().isEmpty()) {
                return CellResult.newBuilder().setError("cell_ref is required").build();
            }
            CellReference ref;
            try {
                ref = new CellReference(input.getCellRef());
            } catch (Exception e) {
                return CellResult.newBuilder().setError("invalid cell_ref: " + input.getCellRef()).build();
            }
            if (ref.getRow() < 0 || ref.getCol() < 0) {
                // CellReference parses a row-only or column-only reference (e.g.
                // "1234" or "AB") without throwing; a full cell lookup needs both.
                return CellResult.newBuilder()
                        .setError("invalid cell_ref (must name both a column and a row): " + input.getCellRef())
                        .build();
            }

            byte[] bytes = OfficeUtil.loadBytes(input.getFile());
            try (Workbook wb = OfficeUtil.openWorkbook(bytes, input.getFile().getPassword())) {
                Sheet sheet = OfficeUtil.resolveSheet(wb, input.getSheetName(), input.getSheetIndex());
                int r = ref.getRow();
                int c = ref.getCol();
                org.apache.poi.ss.usermodel.Row poiRow = sheet.getRow(r);
                org.apache.poi.ss.usermodel.Cell poiCell = poiRow == null ? null : poiRow.getCell(c);
                boolean found = poiCell != null;
                return CellResult.newBuilder()
                        .setCell(OfficeUtil.cellToProto(poiCell, new DataFormatter(), r, c))
                        .setFound(found)
                        .build();
            }
        } catch (OfficeUtil.OfficeError e) {
            return CellResult.newBuilder().setError(e.getMessage()).build();
        } catch (Exception e) {
            return CellResult.newBuilder().setError("could not read cell: " + e.getMessage()).build();
        }
    }
}
