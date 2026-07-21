package nodes;

import axiom.AxiomContext;
import com.google.protobuf.ByteString;
import gen.Messages.OfficeFile;
import gen.Messages.FormatResult;
import org.junit.jupiter.api.Test;
import static org.junit.jupiter.api.Assertions.*;
import java.util.List;
import java.util.Map;
import java.util.Optional;

public class DetectFormatTest {

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
    public void detectsXlsx() {
        AxiomContext ax = new TestContext();
        OfficeFile input = OfficeFile.newBuilder()
                .setData(ByteString.copyFrom(OfficeTestFixtures.simpleWorkbook()))
                .build();
        FormatResult result = DetectFormat.detectFormat(ax, input);
        assertEquals("", result.getError());
        assertEquals("xlsx", result.getFormat());
        assertTrue(result.getIsOoxml());
        assertFalse(result.getIsLegacyBinary());
        assertEquals("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", result.getMimeType());
    }

    @Test
    public void detectsDocx() {
        AxiomContext ax = new TestContext();
        OfficeFile input = OfficeFile.newBuilder()
                .setData(ByteString.copyFrom(OfficeTestFixtures.simpleDocx()))
                .build();
        FormatResult result = DetectFormat.detectFormat(ax, input);
        assertEquals("", result.getError());
        assertEquals("docx", result.getFormat());
        assertTrue(result.getIsOoxml());
    }

    @Test
    public void unrecognizedBytesAreUnknownNotError() {
        AxiomContext ax = new TestContext();
        OfficeFile input = OfficeFile.newBuilder()
                .setData(ByteString.copyFrom(OfficeTestFixtures.garbageBytes()))
                .build();
        FormatResult result = DetectFormat.detectFormat(ax, input);
        // Per the node's documented contract: unrecognized input is not an error.
        assertEquals("", result.getError());
        assertEquals("unknown", result.getFormat());
    }

    @Test
    public void emptyInputIsAStructuredError() {
        AxiomContext ax = new TestContext();
        FormatResult result = DetectFormat.detectFormat(ax, OfficeFile.newBuilder().build());
        assertNotEquals("", result.getError());
    }
}
