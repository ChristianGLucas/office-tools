package nodes;

import axiom.AxiomContext;
import com.google.protobuf.ByteString;
import gen.Messages.OfficeFile;
import gen.Messages.TextResult;
import org.junit.jupiter.api.Test;
import static org.junit.jupiter.api.Assertions.*;
import java.util.List;
import java.util.Map;
import java.util.Optional;

public class ExtractSpreadsheetTextTest {

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
    public void extractsCachedFormulaResultNotFormulaText() {
        // Regression: this must report the cached "999", never the live
        // re-evaluation of "1+2" (3), and never the raw formula source text.
        AxiomContext ax = new TestContext();
        OfficeFile input = OfficeFile.newBuilder()
                .setData(ByteString.copyFrom(OfficeTestFixtures.simpleWorkbook()))
                .build();
        TextResult result = ExtractSpreadsheetText.extractSpreadsheetText(ax, input);
        assertEquals("", result.getError());
        assertEquals(1, result.getUnitCount());
        assertFalse(result.getTruncated());
        assertTrue(result.getText().contains("Hello"));
        assertTrue(result.getText().contains("999"),
                "expected cached formula result 999 in: " + result.getText());
        assertFalse(result.getText().contains("1+2"),
                "must not leak raw formula text: " + result.getText());
    }

    @Test
    public void malformedInputIsStructuredError() {
        AxiomContext ax = new TestContext();
        OfficeFile input = OfficeFile.newBuilder()
                .setData(ByteString.copyFrom(OfficeTestFixtures.garbageBytes()))
                .build();
        TextResult result = ExtractSpreadsheetText.extractSpreadsheetText(ax, input);
        assertNotEquals("", result.getError());
    }
}
