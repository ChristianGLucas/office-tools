package nodes;

import axiom.AxiomContext;
import com.google.protobuf.ByteString;
import gen.Messages.OfficeFile;
import gen.Messages.ReadSheetInput;
import gen.Messages.GridResult;
import org.junit.jupiter.api.Test;
import static org.junit.jupiter.api.Assertions.*;
import java.util.List;
import java.util.Map;
import java.util.Optional;

public class ReadSheetTest {

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
    public void readsTypedCellsAcrossTheUsedGrid() {
        AxiomContext ax = new TestContext();
        OfficeFile file = OfficeFile.newBuilder()
                .setData(ByteString.copyFrom(OfficeTestFixtures.simpleWorkbook()))
                .build();
        ReadSheetInput input = ReadSheetInput.newBuilder().setFile(file).setSheetName("Sheet1").build();
        GridResult result = ReadSheet.readSheet(ax, input);
        assertEquals("", result.getError());
        assertEquals("Sheet1", result.getSheetName());
        assertEquals(2, result.getRowCount());
        assertFalse(result.getTruncated());

        assertEquals("Hello", result.getRows(0).getCells(0).getStringValue());
        assertEquals("STRING", result.getRows(0).getCells(0).getType());
        assertEquals(42.0, result.getRows(0).getCells(1).getNumberValue());

        // A2 is a FORMULA cell: formula text "1+2" but a CACHED result of 999
        // (deliberately different from live-evaluating "1+2" == 3). Every
        // reported value must reflect the cache, never a live re-evaluation.
        var formulaCell = result.getRows(1).getCells(0);
        assertEquals("FORMULA", formulaCell.getType());
        assertEquals("1+2", formulaCell.getFormula());
        assertEquals(999.0, formulaCell.getNumberValue());
        assertEquals("999", formulaCell.getFormattedValue());

        assertTrue(result.getRows(1).getCells(1).getBoolValue());
    }

    @Test
    public void malformedInputIsStructuredError() {
        AxiomContext ax = new TestContext();
        OfficeFile file = OfficeFile.newBuilder()
                .setData(ByteString.copyFrom(OfficeTestFixtures.garbageBytes()))
                .build();
        GridResult result = ReadSheet.readSheet(ax, ReadSheetInput.newBuilder().setFile(file).build());
        assertNotEquals("", result.getError());
    }

    @Test
    public void unknownSheetNameIsStructuredError() {
        AxiomContext ax = new TestContext();
        OfficeFile file = OfficeFile.newBuilder()
                .setData(ByteString.copyFrom(OfficeTestFixtures.simpleWorkbook()))
                .build();
        GridResult result = ReadSheet.readSheet(ax,
                ReadSheetInput.newBuilder().setFile(file).setSheetName("NoSuchSheet").build());
        assertNotEquals("", result.getError());
    }
}
