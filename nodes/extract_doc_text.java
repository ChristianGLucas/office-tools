package nodes;

import axiom.AxiomContext;
import gen.Messages.OfficeFile;
import gen.Messages.TextResult;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.ByteArrayInputStream;
import java.util.Map;

public class ExtractDocText {

    private static final long MAX_TEXT_CHARS = 20_000_000; // 20M chars

    /**
     * Extract the full plain text of a Word document — .docx (OOXML) via
     * XWPFWordExtractor, or legacy .doc (OLE2) via WordExtractor — the
     * paragraph text in document order, POI's own extractor doing the
     * unpacking. unit_count is the paragraph count for .docx (unavailable
     * for legacy .doc, reported as 0). Output is capped at ~20M characters
     * with truncated set if the real text was longer.
     *
     * @param ax    The AxiomContext: logging, secrets, reflection, mutation.
     * @param input The decoded OfficeFile for this invocation.
     */
    public static TextResult extractDocText(AxiomContext ax, OfficeFile input) {
        ax.log().info("extractDocText handling", Map.of());
        try {
            byte[] bytes = OfficeUtil.loadBytes(input);
            OfficeUtil.FormatInfo fmt = OfficeUtil.detectFormat(bytes);
            String text;
            int unitCount = 0;
            switch (fmt.format) {
                case "docx": {
                    try (XWPFDocument doc = new XWPFDocument(new ByteArrayInputStream(bytes));
                         XWPFWordExtractor ex = new XWPFWordExtractor(doc)) {
                        text = ex.getText();
                        unitCount = doc.getParagraphs().size();
                    }
                    break;
                }
                case "doc": {
                    try (POIFSFileSystem fs = new POIFSFileSystem(new ByteArrayInputStream(bytes));
                         HWPFDocument doc = new HWPFDocument(fs);
                         WordExtractor ex = new WordExtractor(doc)) {
                        text = ex.getText();
                    }
                    break;
                }
                default:
                    return TextResult.newBuilder().setError("not a Word document (.docx/.doc)").build();
            }
            boolean truncated = false;
            if (text.length() > MAX_TEXT_CHARS) {
                text = text.substring(0, (int) MAX_TEXT_CHARS);
                truncated = true;
            }
            return TextResult.newBuilder()
                    .setText(text)
                    .setUnitCount(unitCount)
                    .setTruncated(truncated)
                    .build();
        } catch (OfficeUtil.OfficeError e) {
            return TextResult.newBuilder().setError(e.getMessage()).build();
        } catch (Exception e) {
            return TextResult.newBuilder().setError("could not extract text: " + e.getMessage()).build();
        }
    }
}
