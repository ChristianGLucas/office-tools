package nodes;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;

import java.io.ByteArrayOutputStream;
import java.io.IOException;

/**
 * Shared fixture builders for node tests: minimal REAL .xlsx/.docx documents
 * built in-memory with Apache POI itself (the same library the nodes wrap),
 * so tests exercise the real Office parsing path deterministically without
 * checked-in binary fixture files. Not a node — no axiom.yaml entry.
 */
final class OfficeTestFixtures {
    private OfficeTestFixtures() {}

    /**
     * One-sheet workbook, "Sheet1":
     *   A1 = "Hello" (string)
     *   B1 = 42.0 (numeric)
     *   A2 = a FORMULA cell whose formula text is "1+2" but whose CACHED
     *        result is deliberately set to 999 — live-evaluating "1+2" would
     *        give 3, so any test asserting 999 proves the node reported the
     *        cached result rather than re-evaluating the formula.
     *   B2 = TRUE (boolean)
     * Plus one workbook-scoped defined name "Total" -> "Sheet1!$B$1".
     */
    static byte[] simpleWorkbook() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet("Sheet1");
            Row row0 = sheet.createRow(0);
            row0.createCell(0).setCellValue("Hello");
            row0.createCell(1).setCellValue(42.0);

            Row row1 = sheet.createRow(1);
            Cell formulaCell = row1.createCell(0);
            formulaCell.setCellFormula("1+2");
            formulaCell.setCellValue(999.0); // sets the CACHED result, does not re-evaluate
            row1.createCell(1).setCellValue(true);

            Name name = wb.createName();
            name.setNameName("Total");
            name.setRefersToFormula("Sheet1!$B$1");

            ByteArrayOutputStream out = new ByteArrayOutputStream();
            wb.write(out);
            return out.toByteArray();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * A .docx with one Heading1 paragraph "Title", one Normal paragraph
     * "Body text.", and one 2x2 table with cells "A1","B1"/"A2","B2".
     */
    static byte[] simpleDocx() {
        try (XWPFDocument doc = new XWPFDocument()) {
            XWPFParagraph heading = doc.createParagraph();
            heading.setStyle("Heading1");
            XWPFRun headingRun = heading.createRun();
            headingRun.setText("Title");

            XWPFParagraph body = doc.createParagraph();
            XWPFRun bodyRun = body.createRun();
            bodyRun.setText("Body text.");

            XWPFTable table = doc.createTable(2, 2);
            table.getRow(0).getCell(0).setText("A1");
            table.getRow(0).getCell(1).setText("B1");
            table.getRow(1).getCell(0).setText("A2");
            table.getRow(1).getCell(1).setText("B2");

            ByteArrayOutputStream out = new ByteArrayOutputStream();
            doc.write(out);
            return out.toByteArray();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    /** Not a valid Office file of any kind — neither OOXML zip nor OLE2. */
    static byte[] garbageBytes() {
        return new byte[] { 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08 };
    }
}
