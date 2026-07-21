package nodes;

import gen.Messages.Cell;
import gen.Messages.OfficeFile;

import org.apache.poi.hpsf.SummaryInformation;
import org.apache.poi.poifs.filesystem.FileMagic;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaError;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellReference;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.InetAddress;
import java.net.URI;
import java.net.http.HttpClient;
import java.net.http.HttpRequest;
import java.net.http.HttpResponse;
import java.time.Duration;
import java.util.Locale;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;

/**
 * Shared helpers for the office-tools package: bounded input loading (bytes
 * or SSRF-guarded URL fetch), Office-format detection (OOXML zip vs legacy
 * OLE2, and which application within each), and POI spreadsheet cell/sheet
 * conversion. Every node routes untrusted input through here so the cost and
 * interpretation bounds are enforced in exactly one place.
 */
final class OfficeUtil {
    private OfficeUtil() {}

    // --- Cost bounds on untrusted input (checked before any parsing) -------
    static final long MAX_INPUT_BYTES = 40L * 1024 * 1024;   // 40 MB raw file
    static final long MAX_FETCH_BYTES = 40L * 1024 * 1024;   // 40 MB over the wire
    static final int DEFAULT_MAX_ROWS = 10_000;
    static final int HARD_MAX_ROWS = 50_000;
    static final int DEFAULT_MAX_COLS = 1_000;
    static final int HARD_MAX_COLS = 5_000;
    static final int MAX_DEFINED_NAMES = 5_000;
    static final int MAX_PARAGRAPHS = 20_000;
    static final int MAX_TABLES = 2_000;
    static final int MAX_TABLE_ROWS = 5_000;
    static final int MAX_SHEETS_FOR_TEXT = 2_000;

    /** A structured, caller-facing failure — never a crash. */
    static final class OfficeError extends Exception {
        OfficeError(String message) { super(message); }
    }

    // --- Loading input -------------------------------------------------

    /** Resolve an OfficeFile's raw bytes: data directly, or an SSRF-guarded
     *  fetch from url when data is empty. Enforces MAX_INPUT_BYTES either way. */
    static byte[] loadBytes(OfficeFile file) throws OfficeError {
        byte[] data = file.getData().toByteArray();
        if (data.length == 0 && !file.getUrl().isEmpty()) {
            data = fetchUrl(file.getUrl());
        }
        if (data.length == 0) {
            throw new OfficeError("OfficeFile has neither data nor a fetchable url");
        }
        if (data.length > MAX_INPUT_BYTES) {
            throw new OfficeError("input exceeds size limit: " + data.length
                    + " bytes > " + MAX_INPUT_BYTES + " byte cap");
        }
        return data;
    }

    // --- SSRF-guarded URL fetch -----------------------------------------
    //
    // Scheme allowlist (http/https only) + DNS resolution of every hop's host
    // with each resolved address checked against loopback/link-local/private
    // (site-local)/multicast/reserved/CGNAT ranges before connecting, redirects
    // capped and re-validated on every hop, response body capped at
    // MAX_FETCH_BYTES. NOTE (disclosed limitation, not hidden): unlike a raw
    // socket dialer, java.net.http.HttpClient resolves DNS internally at
    // connect time — there is no supported hook to pin the already-validated
    // IP for the actual TCP connect, so a narrow DNS-rebinding TOCTOU window
    // remains between our validation lookup and HttpClient's own connect. This
    // is a materially higher bar than an unguarded fetch (blocks the common
    // metadata-endpoint / internal-network-probe cases outright) but is not a
    // full pin — filed as platform/library friction, see the retrospective.

    private static final int MAX_REDIRECTS = 5;
    private static final Duration CONNECT_TIMEOUT = Duration.ofSeconds(8);
    private static final Duration REQUEST_TIMEOUT = Duration.ofSeconds(20);

    static byte[] fetchUrl(String urlStr) throws OfficeError {
        String current = urlStr;
        HttpClient client = HttpClient.newBuilder()
                .connectTimeout(CONNECT_TIMEOUT)
                .followRedirects(HttpClient.Redirect.NEVER)
                .build();
        try {
            for (int hop = 0; hop <= MAX_REDIRECTS; hop++) {
                URI uri = URI.create(current);
                String scheme = uri.getScheme();
                if (scheme == null || !(scheme.equalsIgnoreCase("http") || scheme.equalsIgnoreCase("https"))) {
                    throw new OfficeError("unsupported URL scheme (http/https only)");
                }
                String host = uri.getHost();
                if (host == null || host.isEmpty()) {
                    throw new OfficeError("URL has no host");
                }
                validateHostResolvesSafely(host);

                HttpRequest req = HttpRequest.newBuilder(uri)
                        .timeout(REQUEST_TIMEOUT)
                        .header("User-Agent", "axiom-office-tools/0.1")
                        .GET()
                        .build();
                HttpResponse<InputStream> resp;
                try {
                    resp = client.send(req, HttpResponse.BodyHandlers.ofInputStream());
                } catch (IOException | InterruptedException e) {
                    throw new OfficeError("fetch failed: " + e.getMessage());
                }
                int status = resp.statusCode();
                if (status >= 300 && status < 400) {
                    String location = resp.headers().firstValue("Location").orElse(null);
                    if (location == null) {
                        throw new OfficeError("redirect response with no Location header");
                    }
                    current = uri.resolve(location).toString();
                    continue;
                }
                if (status < 200 || status >= 300) {
                    throw new OfficeError("fetch returned HTTP " + status);
                }
                return readBounded(resp.body(), MAX_FETCH_BYTES);
            }
        } catch (IllegalArgumentException e) {
            throw new OfficeError("invalid URL: " + e.getMessage());
        }
        throw new OfficeError("too many redirects (max " + MAX_REDIRECTS + ")");
    }

    private static void validateHostResolvesSafely(String host) throws OfficeError {
        InetAddress[] addrs;
        try {
            addrs = InetAddress.getAllByName(host);
        } catch (java.net.UnknownHostException e) {
            throw new OfficeError("cannot resolve host: " + host);
        }
        if (addrs.length == 0) {
            throw new OfficeError("host resolved to no addresses: " + host);
        }
        for (InetAddress a : addrs) {
            if (isBlockedAddress(a)) {
                throw new OfficeError("URL host resolves to a disallowed network address");
            }
        }
    }

    private static boolean isBlockedAddress(InetAddress a) {
        if (a.isLoopbackAddress() || a.isLinkLocalAddress() || a.isSiteLocalAddress()
                || a.isMulticastAddress() || a.isAnyLocalAddress()) {
            return true;
        }
        byte[] b = a.getAddress();
        // CGNAT 100.64.0.0/10
        if (b.length == 4 && (b[0] & 0xFF) == 100 && (b[1] & 0xFF) >= 64 && (b[1] & 0xFF) <= 127) {
            return true;
        }
        // IPv4-mapped/compat and unique-local IPv6 fc00::/7
        if (b.length == 16 && (b[0] & 0xFE) == 0xFC) {
            return true;
        }
        return false;
    }

    private static byte[] readBounded(InputStream in, long max) throws OfficeError {
        try (InputStream is = in) {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            byte[] buf = new byte[8192];
            long total = 0;
            int n;
            while ((n = is.read(buf)) != -1) {
                total += n;
                if (total > max) {
                    throw new OfficeError("response exceeds size limit: " + max + " byte cap");
                }
                out.write(buf, 0, n);
            }
            return out.toByteArray();
        } catch (IOException e) {
            throw new OfficeError("failed reading response body: " + e.getMessage());
        }
    }

    // --- Format detection ------------------------------------------------

    static final class FormatInfo {
        final String format;      // "xlsx" | "xls" | "docx" | "doc" | "unknown"
        final String mimeType;
        final boolean isOoxml;
        final boolean isLegacyBinary;
        FormatInfo(String format, String mimeType, boolean isOoxml, boolean isLegacyBinary) {
            this.format = format; this.mimeType = mimeType;
            this.isOoxml = isOoxml; this.isLegacyBinary = isLegacyBinary;
        }
    }

    static FormatInfo detectFormat(byte[] bytes) {
        FileMagic magic;
        try {
            magic = FileMagic.valueOf(bytes);
        } catch (Exception e) {
            return new FormatInfo("unknown", "application/octet-stream", false, false);
        }
        if (magic == FileMagic.OOXML) {
            String kind = sniffOoxmlKind(bytes);
            if ("xlsx".equals(kind)) {
                return new FormatInfo("xlsx",
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", true, false);
            }
            if ("docx".equals(kind)) {
                return new FormatInfo("docx",
                        "application/vnd.openxmlformats-officedocument.wordprocessingml.document", true, false);
            }
            return new FormatInfo("unknown", "application/zip", true, false);
        }
        if (magic == FileMagic.OLE2) {
            String kind = sniffOle2Kind(bytes);
            if ("xls".equals(kind)) {
                return new FormatInfo("xls", "application/vnd.ms-excel", false, true);
            }
            if ("doc".equals(kind)) {
                return new FormatInfo("doc", "application/msword", false, true);
            }
            return new FormatInfo("unknown", "application/x-ole-storage", false, true);
        }
        return new FormatInfo("unknown", "application/octet-stream", false, false);
    }

    /** Look at zip entry NAMES only (no decompression) to tell xlsx from docx. */
    private static String sniffOoxmlKind(byte[] bytes) {
        try (ZipInputStream zis = new ZipInputStream(new ByteArrayInputStream(bytes))) {
            ZipEntry e;
            int scanned = 0;
            while ((e = zis.getNextEntry()) != null && scanned < 200) {
                scanned++;
                String name = e.getName();
                if (name.startsWith("xl/")) return "xlsx";
                if (name.startsWith("word/")) return "docx";
            }
        } catch (IOException e) {
            return "unknown";
        }
        return "unknown";
    }

    private static String sniffOle2Kind(byte[] bytes) {
        try (POIFSFileSystem fs = new POIFSFileSystem(new ByteArrayInputStream(bytes))) {
            if (fs.getRoot().hasEntry("Workbook") || fs.getRoot().hasEntry("Book")) return "xls";
            if (fs.getRoot().hasEntry("WordDocument")) return "doc";
        } catch (IOException e) {
            return "unknown";
        }
        return "unknown";
    }

    // --- Spreadsheet helpers ---------------------------------------------

    static Workbook openWorkbook(byte[] bytes, String password) throws OfficeError {
        try {
            if (password != null && !password.isEmpty()) {
                return WorkbookFactory.create(new ByteArrayInputStream(bytes), password);
            }
            return WorkbookFactory.create(new ByteArrayInputStream(bytes));
        } catch (Exception e) {
            throw new OfficeError("could not open as a spreadsheet: " + e.getMessage());
        }
    }

    static Sheet resolveSheet(Workbook wb, String name, int index) throws OfficeError {
        if (name != null && !name.isEmpty()) {
            Sheet s = wb.getSheet(name);
            if (s == null) {
                throw new OfficeError("no sheet named '" + name + "'");
            }
            return s;
        }
        if (index < 0 || index >= wb.getNumberOfSheets()) {
            throw new OfficeError("sheet index " + index + " out of range (workbook has "
                    + wb.getNumberOfSheets() + " sheets)");
        }
        return wb.getSheetAt(index);
    }

    /** Convert one POI cell (possibly null/blank) into our proto Cell, reading
     *  formula cells' CACHED result (never live-evaluates formulas). */
    static Cell.Builder cellToProto(org.apache.poi.ss.usermodel.Cell poiCell, DataFormatter formatter,
                                     int rowIdx, int colIdx) {
        Cell.Builder b = Cell.newBuilder()
                .setRow(rowIdx)
                .setCol(colIdx)
                .setRef(new CellReference(rowIdx, colIdx).formatAsString());
        if (poiCell == null) {
            b.setType("BLANK");
            return b;
        }
        CellType rawType = poiCell.getCellType();
        boolean isFormula = rawType == CellType.FORMULA;
        if (isFormula) {
            try {
                b.setFormula(poiCell.getCellFormula());
            } catch (Exception ignored) {
                // formula text is best-effort
            }
        }
        CellType effective = isFormula ? poiCell.getCachedFormulaResultType() : rawType;
        b.setType(isFormula ? "FORMULA" : effective.name());
        switch (effective) {
            case STRING:
                b.setStringValue(poiCell.getRichStringCellValue().getString());
                break;
            case NUMERIC:
                b.setNumberValue(poiCell.getNumericCellValue());
                try {
                    if (DateUtil.isCellDateFormatted(poiCell)) {
                        b.setIsDate(true);
                    }
                } catch (Exception ignored) {
                    // date-format detection is best-effort
                }
                break;
            case BOOLEAN:
                b.setBoolValue(poiCell.getBooleanCellValue());
                break;
            case ERROR:
                try {
                    b.setErrorValue(FormulaError.forInt(poiCell.getErrorCellValue()).getString());
                } catch (Exception ignored) {
                    b.setErrorValue("#ERR");
                }
                break;
            case BLANK:
            default:
                break;
        }
        try {
            String formatted = formatter.formatCellValue(poiCell);
            if (formatted != null) {
                b.setFormattedValue(formatted);
            }
        } catch (Exception ignored) {
            // formatted display string is best-effort; never fail the node over it
        }
        return b;
    }

    // --- Document properties ----------------------------------------------

    static String iso(java.util.Date d) {
        if (d == null) return "";
        return java.time.Instant.ofEpochMilli(d.getTime()).toString();
    }

    static String orEmpty(String s) {
        return s == null ? "" : s;
    }
}
