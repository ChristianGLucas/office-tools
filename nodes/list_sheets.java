package nodes;

import axiom.AxiomContext;
import gen.Messages.OfficeFile;
import gen.Messages.SheetInfo;
import gen.Messages.SheetsResult;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.Map;

public class ListSheets {

    /**
     * List every sheet in a spreadsheet (.xlsx/.xls) — name, 0-based index,
     * used row/column extent, and whether it is hidden — plus the workbook's
     * active sheet index. Cheap: only reads sheet metadata, not cell data.
     *
     * @param ax    The AxiomContext: logging, secrets, reflection, mutation.
     * @param input The decoded OfficeFile for this invocation.
     */
    public static SheetsResult listSheets(AxiomContext ax, OfficeFile input) {
        ax.log().info("listSheets handling", Map.of());
        try {
            byte[] bytes = OfficeUtil.loadBytes(input);
            try (Workbook wb = OfficeUtil.openWorkbook(bytes, input.getPassword())) {
                SheetsResult.Builder result = SheetsResult.newBuilder()
                        .setSheetCount(wb.getNumberOfSheets())
                        .setActiveSheetIndex(wb.getActiveSheetIndex());
                for (int i = 0; i < wb.getNumberOfSheets(); i++) {
                    Sheet s = wb.getSheetAt(i);
                    int rowCount = s.getLastRowNum() >= 0 ? s.getLastRowNum() + 1 : 0;
                    int colCount = 0;
                    for (int r = s.getFirstRowNum(); r <= s.getLastRowNum(); r++) {
                        if (s.getRow(r) != null) {
                            colCount = Math.max(colCount, (int) s.getRow(r).getLastCellNum());
                        }
                    }
                    result.addSheets(SheetInfo.newBuilder()
                            .setIndex(i)
                            .setName(s.getSheetName())
                            .setRowCount(rowCount)
                            .setColCount(Math.max(colCount, 0))
                            .setIsHidden(wb.isSheetHidden(i))
                            .build());
                }
                return result.build();
            }
        } catch (OfficeUtil.OfficeError e) {
            return SheetsResult.newBuilder().setError(e.getMessage()).build();
        } catch (Exception e) {
            return SheetsResult.newBuilder().setError("could not list sheets: " + e.getMessage()).build();
        }
    }
}
