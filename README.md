# office-tools

An [Axiom](https://axiomide.com) node package that wraps [Apache POI](https://poi.apache.org/)
to extract structured data from Microsoft Office documents — spreadsheets
(`.xlsx`, `.xls`) and Word documents (`.docx`, `.doc`) — without opening
Office itself.

Built for the Axiom marketplace. Language: **Java**. License: **MIT**.

## Use it from your agent or app

Every node in this package is a **live, auto-scaling API endpoint** on the
[Axiom](https://axiomide.com) marketplace — call it from an AI agent or your own
code, with nothing to self-host.

**📦 See it on the marketplace:**
https://dev.axiomide.com/marketplace/christiangeorgelucas/office-tools@0.1.0

**Hook it up to an AI agent (MCP).** Add Axiom's hosted MCP server to any MCP
client and every node becomes a typed tool your agent can call — search the
catalog, inspect a schema, and invoke it directly.

```bash
# Claude Code
claude mcp add --transport http axiom https://api.axiomide.com/mcp \
  --header "Authorization: Bearer $AXIOM_API_KEY"
```

Claude Desktop, Cursor, or any config-based client:

```json
{
  "mcpServers": {
    "axiom": {
      "type": "http",
      "url": "https://api.axiomide.com/mcp",
      "headers": { "Authorization": "Bearer YOUR_AXIOM_API_KEY" }
    }
  }
}
```

**Call it from the CLI.**

```bash
axiom invoke christiangeorgelucas/office-tools/DetectFormat --input '{ ... }'
```

**Call it over HTTP.**

```bash
curl -X POST https://api.axiomide.com/invocations/v1/nodes/christiangeorgelucas/office-tools/0.1.0/DetectFormat \
  -H "Authorization: Bearer $AXIOM_API_KEY" \
  -H 'Content-Type: application/json' \
  -d '{ ... }'
```

> Input/output schema for each node is on the marketplace page above, or via
> `axiom inspect node christiangeorgelucas/office-tools/DetectFormat`.

### Get started free

Install the CLI:

```bash
# macOS / Linux — Homebrew
brew install axiomide/tap/axiom

# macOS / Linux — install script
curl -fsSL https://raw.githubusercontent.com/AxiomIDE/axiom-releases/main/install.sh | sh
```

**Windows:** download the `windows/amd64` `.zip` from the
[releases page](https://github.com/AxiomIDE/axiom-releases/releases), unzip it,
and put `axiom.exe` on your `PATH`.

Then `axiom version` to verify, `axiom login` (GitHub or Google) to authenticate,
and create an API key under **Console → API Keys**. Docs and sign-up at
**[axiomide.com](https://axiomide.com)**.

## What it does

Every node accepts (directly, or nested in an operation-specific input) an
`OfficeFile` envelope: raw file bytes, or a URL to fetch the file from, plus
an optional password for encrypted OOXML files. All input is bounded before
parsing (file size, row/column/paragraph/table counts) and malformed or
oversized input returns a structured `error` field rather than throwing.

### Nodes

| Node | Input | Output | What it does |
|---|---|---|---|
| `DetectFormat` | `OfficeFile` | `FormatResult` | Identify xlsx/xls/docx/doc from file magic bytes (no filename needed). |
| `GetProperties` | `OfficeFile` | `PropertiesResult` | Core document metadata: title, author, created/modified timestamps, etc. |
| `ListSheets` | `OfficeFile` | `SheetsResult` | List a workbook's sheets with row/column counts and visibility. |
| `ReadSheet` | `ReadSheetInput` | `GridResult` | Read an entire sheet's used cell grid, bounded by row/column caps. |
| `ReadRange` | `ReadRangeInput` | `GridResult` | Read one explicit A1-notation range (e.g. `B2:D5`). |
| `ReadCell` | `ReadCellInput` | `CellResult` | Look up a single cell by A1 reference. |
| `ListDefinedNames` | `OfficeFile` | `DefinedNamesResult` | List workbook-level named ranges/constants. |
| `ExtractSpreadsheetText` | `OfficeFile` | `TextResult` | Concatenate all cell text across a workbook's sheets. |
| `ExtractDocText` | `OfficeFile` | `TextResult` | Extract a Word document's plain-text paragraph content. |
| `ExtractParagraphs` | `OfficeFile` | `ParagraphsResult` | Extract `.docx` paragraphs with style/heading metadata. |
| `ExtractDocTables` | `OfficeFile` | `TablesResult` | Extract every table in a `.docx` as structured rows. |

Spreadsheet cells report their POI type (`STRING`/`NUMERIC`/`BOOLEAN`/
`FORMULA`/`BLANK`/`ERROR`), formula text where present (last-saved cached
result — formulas are never live-recomputed), and an Excel-formatted display
string.

## Safety bounds

- Max input file size: 40 MB (raw bytes or fetched over HTTP/HTTPS).
- Row/column and full-text-extraction sizes are capped, with a `truncated`
  flag on the result reporting when a cap was hit (`ReadSheet`, `ReadRange`,
  `ExtractSpreadsheetText`, `ExtractDocText`). Defined-name, paragraph, and
  table counts are also capped (5,000 / 20,000 / 2,000 respectively) but do
  not currently report a `truncated` signal — a doc this size is silently cut
  at the cap. Tracked for a future version.
- URL fetches are SSRF-guarded: scheme allowlist, DNS resolution of every
  redirect hop checked against loopback/link-local/private/CGNAT ranges,
  bounded redirects, and a bounded response size.
- `.doc`/`.xls` (legacy OLE2 binary) and `.docx`/`.xlsx` (OOXML) are both
  supported where POI supports them; `ExtractParagraphs` and
  `ExtractDocTables` are `.docx`-only (see each node's description).

## License

MIT — see [LICENSE](./LICENSE).
