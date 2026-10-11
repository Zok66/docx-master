# Captions (`captions` table + `CaptionBlock` + `EquationBlock.captionId`)

Word-native caption-class numbering — SEQ + STYLEREF fields + bookmark + REF `\h`. Used for figure / table / equation / theorem / lemma / ... — anything that needs a chapter-prefixed enumerator with cross-references.

## Quick start

```jsonc
// Top-level apply config
{
  "captions": {
    "Equation": {
      "prefix": "(",
      "suffix": ")",
      "chapterPrefix": ["Heading1"],
      "styleId": "EquationNumber"
    },
    "Figure": {
      "prefix": "图",
      "chapterPrefix": ["Heading1"],
      "bodySeparator": " ",
      "styleId": "FigureCaption"
    },
    "Table": {
      "prefix": "表",
      "chapterPrefix": ["Heading1"],
      "bodySeparator": " ",
      "styleId": "TableCaption"
    }
  },
  "edits": [
    {
      "op": "insert-after",
      "at": { "type": "paragraph", "index": 10 },
      "content": [
        { "type": "equation", "latex": "a^2 + b^2 = c^2",
          "captionId": "Equation", "anchor": "eq-pythagoras" },
        { "type": "caption", "captionId": "Figure",
          "text": "系统架构示意图", "anchor": "fig-arch" }
      ]
    }
  ]
}
```

Field semantics for the `captions` config block are also in [`config-schema.md` § Captions](config-schema.md#captions). `chapterPrefix` is the SEQ-based mechanism for per-chapter caption numbering; it is distinct from `restart: "byHeading"` on a `numbering[]` scheme (which restarts list-class auto-numbering, not caption counters) — see [`config-schema.md` § Numbering](config-schema.md#numbering).

## Schema reference

```ts
captions: Record<string, CaptionEntry>

CaptionEntry {
  prefix?: string                   // literal before the counter (default "")
  suffix?: string                   // literal after the counter (default "")
  format?: "arabic" | "alphabetic" | "ALPHABETIC"
         | "roman" | "ROMAN" | "chinese" | "chinese-formal"
                                    // default "arabic"
  chapterPrefix?: Array<              // ordered, any depth; default []
    string                            //   bare styleId — use heading's native number rendering
    | { styleId: string;              //   force format (re-renders heading counter as Arabic /
        format?: SeqFormat }          //   alphabetic / roman / ... regardless of heading's native numFmt;
  >                                   //   the 中文 academic case: H1 displays "第一章", caption shows "1.1"
                                    // default [] (global, no restart)
  chapterSeparator?: string         // joins chapter levels + counter
                                    // (default ".")
  bodySeparator?: string            // between counter and CaptionBlock.text
                                    // (default " ")
  styleId: string                   // REQUIRED — caption paragraph's style
  subCounter?: {                    // enables subequations (1a)(1b)
    format?: "arabic" | "alphabetic" | ...   // default "alphabetic"
    prefix?: string                 // default ""
    suffix?: string                 // default ""
  }
}
```

Block-level:

```ts
CaptionBlock {
  type: "caption"
  captionId: string                 // references captions[<id>]
  text: string                      // may be ""
  anchor?: string                   // bookmark for cross-references
}

CaptionCounterReset {
  type: "caption-counter-reset"
  captionId: string
  newValue?: number                 // default 1
}
```

## Templates

### Chapter-prefixed (中文学术 / GB/T 7713)

H1 renders as `第一章 / 第二章 / 第三章` (chineseCounting numFmt with lvlText `第%1章`); captions read `图 1.1 / 表 2.1 / (3.1)` in Arabic. Full apply config — declare H1 as the chapter style, install multi-level numbering that pairs `chineseCounting` for H1 with Arabic for deeper headings, then mount captions with a `format: "arabic"` override so the chapter slot re-renders independently of H1's rendering:

```jsonc
{
  "styles": [
    { "id": "Heading1", "name": "heading 1", "fromParagraph": 12, "outlineLevel": 0 },
    { "id": "Heading2", "name": "heading 2", "fromParagraph": 20, "outlineLevel": 1 },
    { "id": "FigureCaption",   "name": "Figure Caption",   "fromParagraph": 40 },
    { "id": "TableCaption",    "name": "Table Caption",    "fromParagraph": 55 },
    { "id": "EquationNumber",  "name": "Equation Number",  "fromParagraph": 60 }
  ],
  "numbering": {
    "levels": [
      { "level": 0, "styleId": "Heading1", "numFmt": "chineseCounting",
        "lvlText": "第%1章" },
      { "level": 1, "styleId": "Heading2", "numFmt": "decimal",
        "lvlText": "%1.%2", "isLgl": true }
    ]
  },
  "captions": {
    "Figure":   { "prefix": "图",
                  "chapterPrefix": [{ "styleId": "Heading1", "format": "arabic" }],
                  "bodySeparator": " ", "styleId": "FigureCaption" },
    "Table":    { "prefix": "表",
                  "chapterPrefix": [{ "styleId": "Heading1", "format": "arabic" }],
                  "bodySeparator": " ", "styleId": "TableCaption" },
    "Equation": { "prefix": "(", "suffix": ")",
                  "chapterPrefix": [{ "styleId": "Heading1", "format": "arabic" }],
                  "styleId": "EquationNumber" }
  }
}
```

Predicted Word output for captions emitted under H1 = `第一章` / `第二章`:

```
图 1.1 系统架构          ← gap between 图 and 1 is autoSpace
表 1.1 评估指标
(1.1)                  ← inline equation number
图 2.1 实验流程
(2.1)
```

The `format: "arabic"` overrides H1's native rendering. Chinese theses typically style H1 as `chineseCounting` but want captions to read `(1.1) / 图 1.1`. Without the override, captions inherit H1's rendering → `(第一章.1)`. Drop the override (use bare string `"Heading1"`) when H1 is already Arabic.

**If your chapter prefix renders as `0`** check (in order): the heading style declares an outline level (set `outlineLevel` on the style entry); the `chapterPrefix.styleId` exactly matches a `styles[]` entry's `id`; the apply call carries the `captions` config (without it the cross-ref pipeline skips chapter-SEQ injection entirely).

**How `format` works**: at apply time, the engine injects a hidden `SEQ _chap_<styleId> \* <FORMAT>` field into each paragraph of the referenced style, wrapped in `<w:vanish/>` rPr so the counter advances without rendering (Word's SEQ `\h` switch is silently overridden by `\*` in the same field, so character-level vanish is the reliable hide mechanism). Captions read the counter via `SEQ _chap_<styleId> \c \* <FORMAT>` (`\c` = repeat current value, no increment). Word's F9 keeps both sides live, so adding / removing chapter headings in Word renumbers captions correctly.

Word's `STYLEREF "Heading 1" \n \* ARABIC` would seem simpler but doesn't reliably re-format non-Arabic source numFmts — `\n` returns the full lvlText ("第一章") and `\* ARABIC` doesn't extract the numeric portion (renders as "第一章.1"). The parallel hidden SEQ sidesteps that.

Identifier reservation: `_chap_` prefixed names are engine-reserved; schema rejects agent use as a captionId.

For English academic: same shape, drop the format override, replace `"图"` / `"表"` with `"Figure "` / `"Table "` (ASCII prefixes need a literal trailing space — `autoSpace` only fires at CJK boundaries) and `bodySeparator: " "` with `": "`.

For short papers / no chapters: drop `chapterPrefix` (global counter).

For theorem-class / lemmas / corollaries: same shape, custom identifier + prefix.

### Subequations

```jsonc
"captions": {
  "Equation": {
    "prefix": "(", "suffix": ")",
    "chapterPrefix": ["Heading1"],
    "styleId": "EquationNumber",
    "subCounter": { "format": "alphabetic" }
  }
}
```

On EquationBlock: omit `subGroup` for standalone `(1)` / `(2)` / `(3)`; `"subGroup": "start"` for `(2a)`; `"subGroup": "continue"` for `(2b)` / `(2c)` / ...

## How rendering works (under the hood)

Diagrammed: the bare-string `chapterPrefix` path. For the `{styleId, format}` override variant, replace the STYLEREF run with `SEQ _chap_<styleId> \c \* <FORMAT>` — see "How `format` works" above.

For a caption like `图 2.3` with `Figure` identifier + `chapterPrefix: ["Heading1"]`:

```
[bookmarkStart name="fig-arch"]
  "图"                                  ← prefix (no trailing space; autoSpace gaps 图↔2 at render)
  { STYLEREF "Heading1" \n }            ← chapter prefix → "2"
  "."                                   ← chapterSeparator
  { SEQ Figure \* ARABIC \s 1 }         ← counter → "3"
[bookmarkEnd]
" "                                     ← bodySeparator
"系统架构示意图"                          ← body text
```

REF `\h` targeting `fig-arch` returns `"图 2.3"` — the whole bookmark range, decoration included.

`\s 1` makes the SEQ counter reset at each Heading 1 boundary (counter sim mirrors this for the placeholder backfill). For two-level chapter prefix (`["Heading1", "Heading2"]`), `\s 2` instead.

## Cross-references to captions

```jsonc
{ "refTo": { "type": "anchor", "name": "fig-arch" }, "display": "label" }
```

For caption-class targets, `display: "label"` and `display: "number"` both return the SEQ-rendered text with full decoration (prefix + chapter + counter + suffix) — they collapse because the bookmark wraps just the number range, so REF `\h` returns the same text either way. `display: "full"` throws: caption-class anchors have no paragraph-wide secondary bookmark to source body text from, so the pre-F9 placeholder would diverge from Word's post-F9 render. Use `"label"` to cite captions. Full routing rules in [`cross-references.md`](cross-references.md).

## Editing existing captions

Use the `edit-caption` op:

```jsonc
{ "op": "edit-caption",
  "target": { "anchor": "fig-arch" },
  "text": "更新后的标题" }
```

Or target by identifier + body-order index:

```jsonc
{ "op": "edit-caption",
  "target": { "captionId": "Figure", "index": 3 },
  "text": "..." }
```

`edit-caption` replaces only the body text — SEQ / STYLEREF fields and bookmark stay intact so cross-references keep resolving. Throws on EquationBlock targets (no body to edit) — to change an equation, delete + re-emit. `replace_paragraph` / paragraph-level destructive ops on caption paragraphs are blocked by the field scan.

## Resetting a caption counter mid-document

```jsonc
{ "type": "caption-counter-reset", "captionId": "Equation", "newValue": 1 }
```

Standalone marker — the **next** caption of `captionId` renders with its counter at `newValue` (default 1). Useful for appendix sequences that don't align with outline-level restart, or multi-section docs where each major section gets its own caption counter.

## Conflicts and constraints

- A `styleId` referenced in `captions[<id>].styleId` MUST NOT also appear in `numbering[].levels[].styleId`. The two mechanisms would fight at render time. Engine throws at apply with a clear message.
- `chapterPrefix` references unknown `styleId` → schema throws.
- `chapterPrefix` references a style that isn't outline-numbered → pre-scan warns (Word's STYLEREF returns 0 at render time).
- `captionId` references undeclared identifier → schema throws.
- `subGroup` set on EquationBlock without matching `subCounter` config → schema throws.
- `anchor` set on EquationBlock without `captionId` → schema throws.

## Standardize re-emit (source-doc captions)

Existing SEQ captions in the source doc get re-rendered with the current `captions[<id>]` config on each apply. Bookmark + identifier + body text preserved. Identifier mismatch (SEQ exists for an unconfigured identifier) → pass through + warn.

## CJK prefix / separator spacing

**Don't pre-bake Pangu spacing in `prefix` / `suffix` / `bodySeparator` for CJK contexts.** Word's `autoSpace` automatically inserts the visual gap at CJK ↔ Latin/digit boundaries at render time; a typed ASCII space stacks on top and renders too wide. Use `"图"` / `"表"` / `"("` without trailing spaces and `" "` (single space) for `bodySeparator`; let Word do the CJK↔digit gap. ASCII-only prefixes (`"Figure "`, `"Table "`) keep their literal spaces — `autoSpace` only fires at CJK boundaries.

## Discovering existing captions

- `overview` shows a Captions section listing SEQ identifiers found in body (skip-if-empty).
- `inspect_caption <doc>` lists all identifiers with occurrence counts + referencing-REF counts.
- `inspect_caption <doc> <identifier>` dumps per-paragraph details.
- `migrate_captions <doc> [--style <styleId>]...` detects manually-numbered caption-shaped paragraphs (e.g. "图 2.1: ..." typed by hand, no SEQ field) and suggests identifiers — agent builds the apply config to convert. `--style` filters candidates to that paragraph styleId (repeatable for multiple styles); omit to scan the whole body.
