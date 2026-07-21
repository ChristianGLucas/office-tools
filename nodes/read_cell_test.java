package nodes;

import axiom.AxiomContext;
import com.google.protobuf.ByteString;
import gen.Messages.OfficeFile;
import gen.Messages.ReadCellInput;
import gen.Messages.CellResult;
import org.junit.jupiter.api.Test;
import static org.junit.jupiter.api.Assertions.*;
import java.util.List;
import java.util.Map;
import java.util.Optional;

public class ReadCellTest {

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
    public void looksUpABooleanCellByReference() {
        AxiomContext ax = new TestContext();
        OfficeFile file = OfficeFile.newBuilder()
                .setData(ByteString.copyFrom(OfficeTestFixtures.simpleWorkbook()))
                .build();
        ReadCellInput input = ReadCellInput.newBuilder()
                .setFile(file).setSheetName("Sheet1").setCellRef("B2").build();
        CellResult result = ReadCell.readCell(ax, input);
        assertEquals("", result.getError());
        assertTrue(result.getFound());
        assertEquals("BOOLEAN", result.getCell().getType());
        assertTrue(result.getCell().getBoolValue());
    }

    @Test
    public void aReferencePastTheUsedRangeIsNotFoundNotAnError() {
        AxiomContext ax = new TestContext();
        OfficeFile file = OfficeFile.newBuilder()
                .setData(ByteString.copyFrom(OfficeTestFixtures.simpleWorkbook()))
                .build();
        CellResult result = ReadCell.readCell(ax,
                ReadCellInput.newBuilder().setFile(file).setSheetName("Sheet1").setCellRef("Z99").build());
        assertEquals("", result.getError());
        assertFalse(result.getFound());
        assertEquals("BLANK", result.getCell().getType());
    }

    @Test
    public void numericOnlyRefIsAStructuredErrorNotGarbage() {
        // Regression: CellReference("1234") parses as a row-only reference
        // (col unset, i.e. -1) without throwing. Before the fix this silently
        // returned a cell with col=-1 and no error instead of rejecting the
        // incomplete reference.
        AxiomContext ax = new TestContext();
        OfficeFile file = OfficeFile.newBuilder()
                .setData(ByteString.copyFrom(OfficeTestFixtures.simpleWorkbook()))
                .build();
        CellResult result = ReadCell.readCell(ax,
                ReadCellInput.newBuilder().setFile(file).setSheetName("Sheet1").setCellRef("1234").build());
        assertNotEquals("", result.getError());
    }

    @Test
    public void emptyCellRefIsAStructuredError() {
        AxiomContext ax = new TestContext();
        CellResult result = ReadCell.readCell(ax, ReadCellInput.newBuilder().build());
        assertNotEquals("", result.getError());
    }
}
