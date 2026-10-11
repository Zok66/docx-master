# Tables (`{ "type": "table", ... }`)

A Block type used inside `edits[]` to insert a new table. For simple data tables a two-line declaration is enough; complex layouts (merged cells, custom borders, mixed cell content) build progressively from the same schema.

## Quick start

```jsonc
{ "type": "table",
  "headerRows": 1,
  "rows": [
    ["Model", "Accuracy", "F1"],
    ["LLM-A", "0.87",     "0.85"],
    ["LLM-B", "0.84",     "0.82"]
  ]}
```

Default: `borders: "all"` (thin black grid), `alignment` omitted (left), `layout: "autofit"`. `headerRows: 1` makes the first row repeat as a header on page breaks — does NOT auto-bold the text; bind a styled paragraph via `headerStyle` (below) or format header cells explicitly.

## Schema reference

```ts
TableBlock {
  type: "table"
  rows: TableCell[][]
  headerRows?: number          // top N rows get <w:tblHeader/>; default 0
  headerStyle?: string         // styleId applied to cell paragraphs in header rows (default; explicit Block[] styleId wins)
  cols?: { width: TableWidth }[]   // length = effective column count; if omitted, engine emits one <w:gridCol>
                                   // per effective column, each seeded with an equal share of the usable width
  borders?: BordersPreset | BordersCustom    // default "all"
  alignment?: "left" | "center" | "right"    // whole-table horizontal alignment on page
  vAlign?: "top" | "center" | "bottom"       // default cell vertical alignment; skill default "center"
  layout?: "fixed" | "autofit"               // default "autofit"
  bidiVisual?: boolean                       // RTL table — mirrors column order (tblPr/bidiVisual)
  padding?: Padding                          // CSS-shorthand cell padding — see Cell padding below
}

TableWidth = "auto" | Length

TableCell = string                                  // plain text, single paragraph
          | InlineNode[]                            // formatted text / cross-refs, single paragraph
          | Block[]                                 // multi-paragraph / images
              | { content: <string | InlineNode[] | Block[]>,
              colspan?: number,
              rowspan?: number,
              vAlign?: "top" | "center" | "bottom",
              borders?: BordersCustom,
              shading?: string,                     // hex RGB
              padding?: Padding,                    // overrides TableBlock.padding for this cell
              textDirection?: "lrTb" | "tbRl" | "btLr" | "lrTbV" | "tbLtlV"
                                                  // per-cell text flow (tcPr/textDirection;
                                                  // lrTbV / tbLtlV are legacy vertical variants)
              formula?: string                      // cell formula field, must start with "="
            }

BordersPreset = "all" | "none" | "outer" | "three-line"
BordersCustom = { top?, bottom?, left?, right?, insideH?, insideV?: BorderEdge }
BorderEdge    = "none" | "single" | "thick" | "double" | "dotted" | "dashed"
              | { style, size?: Length, color?: hex | "auto" }
Padding       = Length                                              // all four edges
              | [Length]                                            // [all]
              | [Length, Length]                                    // [vertical, horizontal]
              | [Length, Length, Length]                            // [top, horizontal, bottom]
              | [Length, Length, Length, Length]                    // [top, right, bottom, left]
```

## Cell content — four progressive forms

| Form | When to reach for it |
|---|---|
| `"text"` (plain string) | The 90% case. Single paragraph, no formatting, no spans. |
| `InlineNode[]` | Single paragraph with mixed run formatting (`{ text, format }`) or inline cross-refs (`{ refTo, display }`). |
| `Block[]` | Multiple paragraphs in one cell, images inside cells, or cell paragraphs needing their own `styleId`. |
| `{ content: ..., colspan, rowspan, vAlign, borders, shading }` | Anything above wrapped in an object that also carries spans or per-cell properties. |

Empty content (`""`, `[]`, or `{ content: [] }`) emits an empty `<w:p/>` — Word requires every cell to contain at least one block-level child.

## Cell merging

The agent declares only the "restart" cell (the one with content). The engine computes a grid occupancy map and injects vMerge continuation cells in subsequent rows at the same column. Continuation cells inherit gridSpan from the restart cell when both axes merge.

```jsonc
"rows": [
  [
    { "content": "Day", "rowspan": 2, "vAlign": "center" },
    { "content": "Performance", "colspan": 3 }
  ],
  ["Accuracy", "F1", "Latency"],
  [
    { "content": "Mon", "rowspan": 3 },
    "0.87", "0.85", "120ms"
  ],
  ["0.88", "0.86", "118ms"],
  ["0.86", "0.84", "122ms"]
]
```

Row 2 omits a cell at column 0 — claimed by `Day` from row 1. Rows 4 and 5 omit a cell at column 0 — claimed by `Mon`'s `rowspan: 3` from row 3.

Each row must total: `declared cells + cells claimed by ongoing rowspans + colspan-expanded width` = the table's column count. The engine throws with a clear error if a row's effective width disagrees.

## Borders

### Presets

| Preset | Effect |
|---|---|
| `"all"` | Thin single border on every edge (table-level + inside). Default. |
| `"none"` | All edges suppressed. |
| `"outer"` | Top / bottom / left / right thin; no inside lines. |
| `"three-line"` | Top thick + bottom thick + thin line under last header row. Sides + inside suppressed. Requires `headerRows ≥ 1` for the middle line; degrades silently to "top + bottom only" when `headerRows: 0`. Injects default top/bottom 4pt cell padding (see [Cell padding](#cell-padding)). |

### Custom

```jsonc
"borders": {
  "top":    { "style": "single", "size": 1.5 },
  "bottom": { "style": "single", "size": 1.5 },
  "insideH": "single",
  "left": "none",
  "right": "none",
  "insideV": "none"
}
```

Field `size` is a Length (`number` = pt, or `"Npt"` / `"Ncm"` / `"Nmm"` / `"Nin"`); engine converts to OOXML's 1/8 pt units (`w:sz`). Omitting a side defaults it to `"none"`. `"thick"` is a 1.5 pt shortcut. Full Length syntax: see [config-schema.md](config-schema.md) §"Length values".

### Cell-level override (escape hatch)

A cell in object form may carry its own `borders` overriding the table-level for that cell only. Used for "this row has a thick top line" or manual three-line construction beyond what the preset gives. Footgun: **adjacent cells do not auto-coordinate.** Setting cell A's right border to `"none"` does NOT remove the visible line between A and B — B is still drawing its own left border per the table-level cascade. To remove the line you must set both sides to `"none"`.

The `"three-line"` preset handles this coordination internally (engine injects per-cell `tcBorders` on the last header row). Outside of preset paths, agents writing manual coordination must do it explicitly.

## Headers

`headerRows: N` marks the top N rows as table headers — emits `<w:tblHeader/>` on each, making them repeat at the top of every page the table spans. Two important non-effects:

- **No auto-bold.** Header text renders identically to body cells unless agent specifies formatting.
- **No styling cascade.** Header rows don't get any default cell shading or border treatment beyond what the table-wide borders specify (except `"three-line"`'s middle line).

Bind a paragraph style to all header-row cells with `headerStyle`:

```jsonc
{ "type": "table",
  "headerRows": 1,
  "headerStyle": "TableHeader",
  "rows": [
    ["Header A", "Header B"],
    ["data",     "data"]
  ]}
```

`TableHeader` should be installed in `styles[]` with the desired font / weight / alignment. Cells with explicit `Block[]` containing their own `styleId` win over `headerStyle`.

## Column widths and layout

`cols` declares per-column widths; length must match the effective column count (see Cell merging). Mixing explicit Length values with `"auto"` works — auto columns split the remaining budget after explicit ones.

```jsonc
"cols": [{ "width": "auto" }, { "width": "3.5cm" }, { "width": "7cm" }]
```

Omit `cols` for autofit: engine seeds even-split widths from the host section's content width; Word reflows on open. Inside a `cell` locator the cell's actual width isn't auto-detected — declare `cols` matching the cell when it's narrow, especially under `layout: "fixed"`.

`layout`:
- `"autofit"` (default) — Word reflows columns to fit content.
- `"fixed"` — declared widths win; total exceeding the page width overflows the edge.

## Alignment axes

Three distinct alignment fields apply at different scopes:

| Field | Scope | Affects |
|---|---|---|
| `alignment` on TableBlock | whole table on page | Horizontal position (left / center / right). |
| `vAlign` on TableBlock | every cell without its own `vAlign` | Default `"center"`. Set `"top"` for form-style layouts where labels hug the top of each cell. |
| `vAlign` on cell object form | one cell | Overrides table-level default for that cell only. |
| `paraFormat.alignment` on a Paragraph inside a cell | one paragraph | Horizontal text alignment within the cell — separate from `vAlign`. |

Resolution for any cell: `cell.vAlign ?? block.vAlign ?? "center"`. Engine always emits `vAlign` on every cell — to restore Word's native top default, set `vAlign: "top"` at the table level.

**Caveat**: multiplier line spacing (e.g. 1.5x) puts the extra leading below the text, so single-line cells look top-aligned despite `vAlign: "center"`. Set the cell paragraph's `paraFormat.lineSpacing: 1`; multi-line content is unaffected.

## Cell padding

`padding` on TableBlock applies to every cell that doesn't carry its own; `padding` on a cell object overrides for that one cell. Resolution: `cell.padding ?? block.padding ?? <preset default>`.

Explicit `padding` always emits **all four** edges (CSS-faithful) — `padding: 0` flattens, including Word's TableNormal default of left/right 5.4pt. The `"three-line"` preset is the only path that emits a partial default (top/bottom 4pt; left/right omitted to inherit TableNormal); writing any `padding` explicitly fully replaces this.

## Row sort (`sort`)

Static equivalent of Word's 布局 → 排序, applied at write time:

```json
{ "type": "table", "headerRows": 1,
  "sort": { "column": 2, "order": "desc" },
  "rows": [ ["姓名", "分数"], ["张三", "88"], ["李四", "92"] ] }
```

- `column` is 1-based; `order` defaults to `"asc"`. Header rows are pinned; only body rows reorder.
- Key = target cell's plain text. Numeric on both sides → numeric compare (thousands separators stripped); otherwise zh-CN collation (Chinese sorts by pinyin).
- **Rowspan rejects `sort`** — reordering rows orphans vMerge continuations. Sort data before passing it to apply if merges are needed.

## Cell formulas (`formula`)

Word's table formula (Layout → 公式) as a static field:

```json
{ "formula": "=SUM(ABOVE)", "content": "255" }
```

- Must start with `=` — `=SUM(ABOVE)` / `=AVERAGE(LEFT)` / `=B2*C2` / `=SUM(A1:B3)`.
- Emitted as a 5-run complex field; Word re-evaluates on F9 / open-with-`updateFields`.
- The cell's plain-string `content` is the **cached** result Word shows before the first field update — precompute it (e.g. `255` for the example above). Leave `content: ""` for an empty placeholder.
- The skill does NOT evaluate formulas — it writes the field + cached value only.

## Edge cases the engine handles

- **Trailing paragraph after a table.** Word rejects a `<w:tbl>` as the last child of a body or cell. Engine appends `<w:p/>` after a trailing table in either container.
- **Adjacent tables.** Two `<w:tbl>` siblings need a `<w:p/>` between. Engine inserts one.
- **Empty cell.** `""` / `[]` / `{ content: [] }` becomes `<w:p/>`.
- **Mismatched row widths.** Engine throws at emit time with the row index and column counts.
- **Out-of-bounds spans.** `rowspan` past the last row or `colspan` past the row's remaining width throws at emit time.

## Caption placement

Place `TableCaption`-bound paragraphs **above** the table (mirror of figures which place captions below). Cross-refs from body text anchor on the caption paragraph, never on the table itself:

```jsonc
{ "type": "paragraph", "styleId": "TableCaption", "anchor": "tab-results", "text": "评测结果汇总" },
{ "type": "table", "headerRows": 1, "rows": [...] }
```

Then in body text:
```jsonc
[{ "text": "见 " }, { "refTo": { "type": "anchor", "name": "tab-results" }, "display": "label" }]
```

## Integration with the style system

| What | How |
|---|---|
| Cell paragraph styleId | `Block[]` form: `[{ "type": "paragraph", "styleId": "TableCellBody", "text": "..." }]`. Cascade resolves through styles.xml normally. |
| Auto-numbered captions inside cells | A cell paragraph bound to `FigureCaption` / `TableCaption` participates in the doc-wide counter. |
| Theme fonts inside cells | rPr cascade applies; no special wiring needed. |
| Bookmark anchors inside cells | Cell paragraph blocks can declare `anchor` — same BookmarkAllocator path as body paragraphs. Cross-refs from any position resolve. |
| `pattern_rules` / `bulk_rules` coverage | Body + layout-table cell paragraphs (indexed). Data-table cells are now reachable via per-cell-paragraph addressing in surgical edits, but **not** by whole-doc bulk rules. For predictable cell paragraph styling, bind via explicit `styleId` in inserts or use the `cell` locator. |
| MDF (match destination formatting) | TableBlock is non-paragraph; engine doesn't propagate anchor pPr into it. The cells' own pPr is what you declared. |
| Table-level styles (Word's `<w:tblStyle>`) | **Not supported** — see "What's not supported" below. |

## What's not supported (v1)

- **Nested tables via cell `Block[]`** — schema is acyclic. To put a table inside an existing cell, run a separate `apply` with a `cell` locator + insert op.
- **Deleting / replacing an existing table** — no locator selects a `<w:tbl>` directly. `paragraph` / `range` / `cell` / `heading` all target paragraphs.
- **trackChanges on TableBlock insertion** — engine throws at emit; OOXML has no clean "table inserted" tracked-change wrapper. Run table insertion in a separate apply without trackChanges, then enable trackChanges for cell-content edits afterward.
- **Table-level styles** (`<w:tblStyle>` reference + conditional region formatting like banded rows / firstRow auto-bold). OOXML's conditional regions are a separate sub-spec; this skill expects per-`TableBlock` declarations of borders / shading / alignment / headerStyle. Reuse across tables via JSON-fragment reuse.
- **Row no-break (`cantSplit`), zebra striping** — fall outside v1 scope. cantSplit workaround: user fixes in Word after open. Zebra striping: declare `shading` per cell (or use the JSON-fragment reuse pattern).
- **Cell text rotation (`textDirection`)** — **supported** per cell: `{ content: "...", textDirection: "tbRl" }` (`lrTb` / `tbRl` / `btLr` are the three Word UI exposes). Pair with `vAlign` for vertical-header cells.
- **Word built-in styles like "Grid Table 4 - Accent 1"** — not referenced; declare equivalent borders + shading manually.
- **Same-apply edit of a freshly inserted table's cells** — paragraph indices and cell locators reference pre-edit state. Insert the table in one apply, edit cells in a second.
- **Percentage column widths** — use fixed pt or `"auto"` instead.

## Classifier reason labels (overview output)

Overview tags each table with `reason:<label>` showing which signal triggered the layout/data classification. Use these to verify the classifier picked the right class, and to mentally override when a known edge case is misclassified.

| Reason | Trigger | Class |
|---|---|---|
| `1x1` | 1 row × 1 col table | layout |
| `singleTcStack` | Every row has exactly one `<w:tc>` and total paragraphs > 3 | layout |
| `outlineLvl` | Table contains any direct `<w:outlineLvl>` | layout |
| `bulkCell` | Some cell holds more than 5 paragraphs | layout |
| `multiColData` | Multi-row × multi-col, no layout signal fired | data |
| `fallback` | Degenerate (single-cell tables that aren't 1×1, etc.) | data |
