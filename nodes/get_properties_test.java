package nodes;

import axiom.AxiomContext;
import com.google.protobuf.ByteString;
import gen.Messages.OfficeFile;
import gen.Messages.PropertiesResult;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import static org.junit.jupiter.api.Assertions.*;
import java.io.ByteArrayOutputStream;
import java.util.List;
import java.util.Map;
import java.util.Optional;

public class GetPropertiesTest {

    // A no-op AxiomContext a node author edits to drive a specific scenario.
    // Reflection exposes an empty graph, mutation is a sink. Implement only
    // what your assertions need.
    static final class TestContext implements AxiomContext {
        public Logger log() {
            return new Logger() {
                public void debug(String m, Map<String, String> a) {}
                public void info(String m, Map<String, String> a)  {}
                public void warn(String m, Map<String, String> a)  {}
                public void error(String m, Map<String, String> a) {}
            };
        }
        public Secrets secrets() { return name -> Optional.empty(); }
        public String executionId() { return "test-execution-id"; }
        public String flowId() { return "test-flow-id"; }
        public String tenantId() { return "test-tenant-id"; }
        public Reflection reflection() {
            return () -> new FlowReflection() {
                public List<ReflectionNode> nodes() { return List.of(); }
                public List<ReflectionEdge> edges() { return List.of(); }
                public List<ReflectionEdge> loopEdges() { return List.of(); }
                public FlowPosition position() { return new FlowPosition(0, 0, Map.of(), List.of()); }
                public String graphId() { return ""; }
            };
        }
        public Mutation mutation() {
            return () -> new FlowMutation() {
                public int addNode(String pkg, String ver, CanvasPosition pos) { return 0; }
                public void addEdge(int src, int dst, EdgeCondition cond) {}
            };
        }
    }

    @Test
    public void readsCorePropertiesFromXlsx() throws Exception {
        byte[] bytes;
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            wb.createSheet("Sheet1");
            wb.getProperties().getCoreProperties().setTitle("My Report");
            wb.getProperties().getCoreProperties().setCreator("Ada Lovelace");
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            wb.write(out);
            bytes = out.toByteArray();
        }
        AxiomContext ax = new TestContext();
        OfficeFile input = OfficeFile.newBuilder().setData(ByteString.copyFrom(bytes)).build();
        PropertiesResult result = GetProperties.getProperties(ax, input);
        assertEquals("", result.getError());
        assertEquals("xlsx", result.getFormat());
        assertEquals("My Report", result.getTitle());
        assertEquals("Ada Lovelace", result.getAuthor());
    }

    @Test
    public void unsetFieldsAreEmptyNotError() {
        AxiomContext ax = new TestContext();
        OfficeFile input = OfficeFile.newBuilder()
                .setData(ByteString.copyFrom(OfficeTestFixtures.simpleWorkbook()))
                .build();
        PropertiesResult result = GetProperties.getProperties(ax, input);
        assertEquals("", result.getError());
        assertEquals("", result.getTitle());
    }

    @Test
    public void malformedInputIsStructuredError() {
        AxiomContext ax = new TestContext();
        OfficeFile input = OfficeFile.newBuilder()
                .setData(ByteString.copyFrom(OfficeTestFixtures.garbageBytes()))
                .build();
        PropertiesResult result = GetProperties.getProperties(ax, input);
        assertNotEquals("", result.getError());
    }
}
