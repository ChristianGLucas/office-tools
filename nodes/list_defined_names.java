package nodes;

import axiom.AxiomContext;
import gen.Messages.DefinedNamesResult;
import gen.Messages.OfficeFile;

import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.List;
import java.util.Map;

public class ListDefinedNames {

    /**
     * List a workbook's defined (named) ranges and named constants — the
     * things Excel's Name Manager shows — each with its name, the formula
     * or reference it resolves to, whether it is workbook-scoped (-1) or
     * sheet-scoped, and whether it is a named function rather than a range.
     * A workbook with none returns an empty list, not an error.
     *
     * @param ax    The AxiomContext: logging, secrets, reflection, mutation.
     * @param input The decoded OfficeFile for this invocation.
     */
    public static DefinedNamesResult listDefinedNames(AxiomContext ax, OfficeFile input) {
        ax.log().info("listDefinedNames handling", Map.of());
        try {
            byte[] bytes = OfficeUtil.loadBytes(input);
            try (Workbook wb = OfficeUtil.openWorkbook(bytes, input.getPassword())) {
                List<? extends Name> names = wb.getAllNames();
                DefinedNamesResult.Builder result = DefinedNamesResult.newBuilder();
                int count = 0;
                for (Name n : names) {
                    result.addNames(gen.Messages.DefinedName.newBuilder()
                            .setName(OfficeUtil.orEmpty(n.getNameName()))
                            .setRefersTo(OfficeUtil.orEmpty(n.getRefersToFormula()))
                            .setSheetIndex(n.getSheetIndex())
                            .setIsFunctionName(n.isFunctionName())
                            .build());
                    count++;
                }
                result.setCount(count);
                return result.build();
            }
        } catch (OfficeUtil.OfficeError e) {
            return DefinedNamesResult.newBuilder().setError(e.getMessage()).build();
        } catch (Exception e) {
            return DefinedNamesResult.newBuilder().setError("could not list defined names: " + e.getMessage()).build();
        }
    }
}
