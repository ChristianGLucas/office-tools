package nodes;

import axiom.AxiomContext;
import com.google.protobuf.ByteString;
import gen.Messages.OfficeFile;
import gen.Messages.DefinedNamesResult;
import org.junit.jupiter.api.Test;
import static org.junit.jupiter.api.Assertions.*;
import java.util.List;
import java.util.Map;
import java.util.Optional;

public class ListDefinedNamesTest {

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
    public void listsTheOneDefinedName() {
        AxiomContext ax = new TestContext();
        OfficeFile input = OfficeFile.newBuilder()
                .setData(ByteString.copyFrom(OfficeTestFixtures.simpleWorkbook()))
                .build();
        DefinedNamesResult result = ListDefinedNames.listDefinedNames(ax, input);
        assertEquals("", result.getError());
        assertEquals(1, result.getCount());
        assertEquals("Total", result.getNames(0).getName());
        assertEquals("Sheet1!$B$1", result.getNames(0).getRefersTo());
        assertFalse(result.getNames(0).getIsFunctionName());
    }

    @Test
    public void noDefinedNamesIsEmptyNotError() throws Exception {
        byte[] bytes;
        try (org.apache.poi.xssf.usermodel.XSSFWorkbook wb = new org.apache.poi.xssf.usermodel.XSSFWorkbook()) {
            wb.createSheet("Sheet1");
            java.io.ByteArrayOutputStream out = new java.io.ByteArrayOutputStream();
            wb.write(out);
            bytes = out.toByteArray();
        }
        AxiomContext ax = new TestContext();
        OfficeFile input = OfficeFile.newBuilder().setData(ByteString.copyFrom(bytes)).build();
        DefinedNamesResult result = ListDefinedNames.listDefinedNames(ax, input);
        assertEquals("", result.getError());
        assertEquals(0, result.getCount());
    }
}
