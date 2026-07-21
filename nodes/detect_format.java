package nodes;

import axiom.AxiomContext;
import gen.Messages.OfficeFile;
import gen.Messages.FormatResult;
import java.util.Map;

public class DetectFormat {

    /**
     * Cheaply identify which of the four Office formats this package handles
     * a file is — xlsx, xls, docx, or doc — without fully parsing it, using
     * Apache POI's magic-byte sniffing (OLE2 vs the OOXML zip container) plus
     * a zip-entry-name check for OOXML files. A file that is neither is not
     * an error: it returns format="unknown" with error left empty, useful
     * for triaging a document before choosing which node to call next.
     *
     * @param ax    The AxiomContext: logging, secrets, reflection, mutation.
     * @param input The decoded OfficeFile for this invocation.
     */
    public static FormatResult detectFormat(AxiomContext ax, OfficeFile input) {
        ax.log().info("detectFormat handling", Map.of());
        try {
            byte[] bytes = OfficeUtil.loadBytes(input);
            OfficeUtil.FormatInfo info = OfficeUtil.detectFormat(bytes);
            return FormatResult.newBuilder()
                    .setFormat(info.format)
                    .setMimeType(info.mimeType)
                    .setIsOoxml(info.isOoxml)
                    .setIsLegacyBinary(info.isLegacyBinary)
                    .build();
        } catch (OfficeUtil.OfficeError e) {
            return FormatResult.newBuilder().setError(e.getMessage()).build();
        } catch (Exception e) {
            return FormatResult.newBuilder().setError("unexpected failure: " + e.getMessage()).build();
        }
    }
}
