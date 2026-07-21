package nodes;

import axiom.AxiomContext;
import gen.Messages.OfficeFile;
import gen.Messages.PropertiesResult;

import org.apache.poi.hpsf.SummaryInformation;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ooxml.POIXMLProperties;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.ByteArrayInputStream;
import java.util.Map;

public class GetProperties {

    /**
     * Read document-level metadata from an Office file — title, author,
     * subject, keywords, comments, category, created/modified timestamps,
     * last-modified-by, revision, application, and template — one shape
     * covering all four supported formats. OOXML files (.xlsx/.docx) read
     * this from their core + extended properties parts; legacy OLE2 files
     * (.xls/.doc) read it from the shared SummaryInformation stream POI
     * exposes identically for both. A well-formed file that simply has no
     * value set for a field returns that field empty, not an error.
     *
     * @param ax    The AxiomContext: logging, secrets, reflection, mutation.
     * @param input The decoded OfficeFile for this invocation.
     */
    public static PropertiesResult getProperties(AxiomContext ax, OfficeFile input) {
        ax.log().info("getProperties handling", Map.of());
        try {
            byte[] bytes = OfficeUtil.loadBytes(input);
            OfficeUtil.FormatInfo fmt = OfficeUtil.detectFormat(bytes);
            PropertiesResult.Builder b = PropertiesResult.newBuilder().setFormat(fmt.format);

            switch (fmt.format) {
                case "xlsx": {
                    try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(bytes))) {
                        applyCore(b, wb.getProperties());
                    }
                    break;
                }
                case "docx": {
                    try (XWPFDocument doc = new XWPFDocument(new ByteArrayInputStream(bytes))) {
                        applyCore(b, doc.getProperties());
                    }
                    break;
                }
                case "xls": {
                    try (HSSFWorkbook wb = new HSSFWorkbook(new ByteArrayInputStream(bytes))) {
                        applySummary(b, wb.getSummaryInformation());
                    }
                    break;
                }
                case "doc": {
                    try (POIFSFileSystem fs = new POIFSFileSystem(new ByteArrayInputStream(bytes));
                         HWPFDocument doc = new HWPFDocument(fs)) {
                        applySummary(b, doc.getSummaryInformation());
                    }
                    break;
                }
                default:
                    return PropertiesResult.newBuilder()
                            .setFormat(fmt.format)
                            .setError("unrecognized or unsupported Office format")
                            .build();
            }
            return b.build();
        } catch (OfficeUtil.OfficeError e) {
            return PropertiesResult.newBuilder().setError(e.getMessage()).build();
        } catch (Exception e) {
            return PropertiesResult.newBuilder().setError("could not read properties: " + e.getMessage()).build();
        }
    }

    private static void applyCore(PropertiesResult.Builder b, POIXMLProperties props) {
        POIXMLProperties.CoreProperties core = props.getCoreProperties();
        POIXMLProperties.ExtendedProperties ext = props.getExtendedProperties();
        b.setTitle(OfficeUtil.orEmpty(core.getTitle()));
        b.setSubject(OfficeUtil.orEmpty(core.getSubject()));
        b.setAuthor(OfficeUtil.orEmpty(core.getCreator()));
        b.setKeywords(OfficeUtil.orEmpty(core.getKeywords()));
        b.setComments(OfficeUtil.orEmpty(core.getDescription()));
        b.setCategory(OfficeUtil.orEmpty(core.getCategory()));
        b.setCreated(OfficeUtil.iso(core.getCreated()));
        b.setModified(OfficeUtil.iso(core.getModified()));
        b.setLastModifiedBy(OfficeUtil.orEmpty(core.getLastModifiedByUser()));
        b.setRevision(OfficeUtil.orEmpty(core.getRevision()));
        if (ext != null) {
            b.setApplication(OfficeUtil.orEmpty(ext.getApplication()));
            b.setTemplate(OfficeUtil.orEmpty(ext.getTemplate()));
        }
    }

    private static void applySummary(PropertiesResult.Builder b, SummaryInformation si) {
        if (si == null) return;
        b.setTitle(OfficeUtil.orEmpty(si.getTitle()));
        b.setSubject(OfficeUtil.orEmpty(si.getSubject()));
        b.setAuthor(OfficeUtil.orEmpty(si.getAuthor()));
        b.setKeywords(OfficeUtil.orEmpty(si.getKeywords()));
        b.setComments(OfficeUtil.orEmpty(si.getComments()));
        b.setCreated(OfficeUtil.iso(si.getCreateDateTime()));
        b.setModified(OfficeUtil.iso(si.getLastSaveDateTime()));
        b.setLastModifiedBy(OfficeUtil.orEmpty(si.getLastAuthor()));
        b.setRevision(OfficeUtil.orEmpty(si.getRevNumber()));
        b.setApplication(OfficeUtil.orEmpty(si.getApplicationName()));
        b.setTemplate(OfficeUtil.orEmpty(si.getTemplate()));
    }
}
