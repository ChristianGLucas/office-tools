package nodes;

import axiom.AxiomContext;
import com.google.protobuf.ByteString;
import gen.Messages.OfficeFile;
import gen.Messages.TablesResult;
import org.junit.jupiter.api.Test;
import static org.junit.jupiter.api.Assertions.*;
import java.util.List;
import java.util.Map;
import java.util.Optional;

public class ExtractDocTablesTest {

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
    public void extractsTheOneTwoByTwoTable() {
        AxiomContext ax = new TestContext();
        OfficeFile input = OfficeFile.newBuilder()
                .setData(ByteString.copyFrom(OfficeTestFixtures.simpleDocx()))
                .build();
        TablesResult result = ExtractDocTables.extractDocTables(ax, input);
        assertEquals("", result.getError());
        assertEquals(1, result.getTableCount());
        var table = result.getTables(0);
        assertEquals(2, table.getRowCount());
        assertEquals("A1", table.getRows(0).getCells(0));
        assertEquals("B1", table.getRows(0).getCells(1));
        assertEquals("A2", table.getRows(1).getCells(0));
        assertEquals("B2", table.getRows(1).getCells(1));
    }

    @Test
    public void nonDocxIsStructuredError() {
        AxiomContext ax = new TestContext();
        OfficeFile input = OfficeFile.newBuilder()
                .setData(ByteString.copyFrom(OfficeTestFixtures.simpleWorkbook()))
                .build();
        TablesResult result = ExtractDocTables.extractDocTables(ax, input);
        assertNotEquals("", result.getError());
    }
}
