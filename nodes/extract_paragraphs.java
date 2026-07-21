package nodes;

import axiom.AxiomContext;
import gen.Messages.OfficeFile;
import gen.Messages.Paragraph;
import gen.Messages.ParagraphsResult;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

import java.io.ByteArrayInputStream;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class ExtractParagraphs {

    private static final Pattern HEADING_STYLE = Pattern.compile("^Heading(\\d+)$", Pattern.CASE_INSENSITIVE);

    /**
     * List every paragraph of a .docx Word document, in document order,
     * with its text, named style (Word's styleId, e.g. "Heading1",
     * "Normal"), and a derived is_heading/heading_level from the standard
     * "HeadingN" style-id convention Word's default template uses (a
     * document on a custom template with differently-named heading styles
     * will report is_heading=false — a disclosed limitation, not a crash).
     * .doc (legacy binary) is not supported by this node — use
     * ExtractDocText for legacy .doc text. Bounded to 20,000 paragraphs.
     *
     * @param ax    The AxiomContext: logging, secrets, reflection, mutation.
     * @param input The decoded OfficeFile for this invocation.
     */
    public static ParagraphsResult extractParagraphs(AxiomContext ax, OfficeFile input) {
        ax.log().info("extractParagraphs handling", Map.of());
        try {
            byte[] bytes = OfficeUtil.loadBytes(input);
            OfficeUtil.FormatInfo fmt = OfficeUtil.detectFormat(bytes);
            if (!"docx".equals(fmt.format)) {
                return ParagraphsResult.newBuilder()
                        .setError("ExtractParagraphs supports .docx only (got: " + fmt.format + ")")
                        .build();
            }
            try (XWPFDocument doc = new XWPFDocument(new ByteArrayInputStream(bytes))) {
                List<XWPFParagraph> paragraphs = doc.getParagraphs();
                ParagraphsResult.Builder result = ParagraphsResult.newBuilder();
                int count = 0;
                for (int i = 0; i < paragraphs.size(); i++) {
                    if (count >= OfficeUtil.MAX_PARAGRAPHS) break;
                    XWPFParagraph p = paragraphs.get(i);
                    String styleId = OfficeUtil.orEmpty(p.getStyleID());
                    Matcher m = HEADING_STYLE.matcher(styleId);
                    boolean isHeading = m.matches();
                    int level = isHeading ? Integer.parseInt(m.group(1)) : 0;
                    result.addParagraphs(Paragraph.newBuilder()
                            .setIndex(i)
                            .setText(OfficeUtil.orEmpty(p.getText()))
                            .setStyleName(styleId)
                            .setIsHeading(isHeading)
                            .setHeadingLevel(level)
                            .build());
                    count++;
                }
                result.setCount(count);
                return result.build();
            }
        } catch (OfficeUtil.OfficeError e) {
            return ParagraphsResult.newBuilder().setError(e.getMessage()).build();
        } catch (Exception e) {
            return ParagraphsResult.newBuilder().setError("could not extract paragraphs: " + e.getMessage()).build();
        }
    }
}
