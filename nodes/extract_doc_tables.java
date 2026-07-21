package nodes;

import axiom.AxiomContext;
import gen.Messages.DocTable;
import gen.Messages.OfficeFile;
import gen.Messages.TableRow;
import gen.Messages.TablesResult;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import java.io.ByteArrayInputStream;
import java.util.List;
import java.util.Map;

public class ExtractDocTables {

    /**
     * Extract every table embedded in a .docx Word document as structured
     * rows of cell text, in document order. .doc (legacy binary) is not
     * supported by this node — use ExtractDocText for legacy .doc text.
     * Bounded to 2,000 tables and 5,000 rows per table.
     *
     * @param ax    The AxiomContext: logging, secrets, reflection, mutation.
     * @param input The decoded OfficeFile for this invocation.
     */
    public static TablesResult extractDocTables(AxiomContext ax, OfficeFile input) {
        ax.log().info("extractDocTables handling", Map.of());
        try {
            byte[] bytes = OfficeUtil.loadBytes(input);
            OfficeUtil.FormatInfo fmt = OfficeUtil.detectFormat(bytes);
            if (!"docx".equals(fmt.format)) {
                return TablesResult.newBuilder()
                        .setError("ExtractDocTables supports .docx only (got: " + fmt.format + ")")
                        .build();
            }
            try (XWPFDocument doc = new XWPFDocument(new ByteArrayInputStream(bytes))) {
                List<XWPFTable> tables = doc.getTables();
                TablesResult.Builder result = TablesResult.newBuilder();
                int tableCount = 0;
                for (int ti = 0; ti < tables.size(); ti++) {
                    if (tableCount >= OfficeUtil.MAX_TABLES) break;
                    XWPFTable table = tables.get(ti);
                    List<XWPFTableRow> rows = table.getRows();
                    DocTable.Builder tb = DocTable.newBuilder().setIndex(ti);
                    int rowCount = 0;
                    for (XWPFTableRow row : rows) {
                        if (rowCount >= OfficeUtil.MAX_TABLE_ROWS) break;
                        TableRow.Builder rb = TableRow.newBuilder();
                        for (XWPFTableCell cell : row.getTableCells()) {
                            rb.addCells(OfficeUtil.orEmpty(cell.getText()));
                        }
                        tb.addRows(rb.build());
                        rowCount++;
                    }
                    tb.setRowCount(rowCount);
                    result.addTables(tb.build());
                    tableCount++;
                }
                result.setTableCount(tableCount);
                return result.build();
            }
        } catch (OfficeUtil.OfficeError e) {
            return TablesResult.newBuilder().setError(e.getMessage()).build();
        } catch (Exception e) {
            return TablesResult.newBuilder().setError("could not extract tables: " + e.getMessage()).build();
        }
    }
}
