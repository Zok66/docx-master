# Cross-references (REF fields)

## When to use this

Never write `"如图 1 所示"` or `"见 [3]"` as plain runs — use `InlineRef` so Word resolves the cite at render time. Page-number citations (`PAGEREF`) are out of scope; surface to the user.

## Schema

`InlineRef` is an inline node alongside `InlineRun` inside a paragraph's `text` array:

```jsonc
{
  "type": "paragraph",
  "text": [
    { "text": "结果如" },
    { "refTo": { "type": "paragraph", "index": 42 }, "display": "label" },
    { "text": " 所示。" }
  ]
}
```

Fields:
- `refTo` — locator. Two forms:
  - `{ "type": "paragraph", "index": N }` — pre-edit 1-based paragraph index, same as every other `edits[]` locator. Target must exist in the source document.
  - `{ "type": "anchor", "name": "fig-arch" }` — named bookmark. Resolves against (a) any `ParagraphBlock.anchor` / `EquationBlock.anchor` declared anywhere in this `edits[]` — the engine pre-scans names before emit so refs can address anchors declared later in the array or later in the same op's Block list — or (b) bookmarks already in the source document that wrap a single paragraph. Name must match `^[A-Za-z_][A-Za-z0-9_-]{0,39}$`.
- `display` — what Word renders. `\h` is always added so the rendered text is clickable:
  - `"label"` (default) — for outline targets (Headings, lists) renders `lvlText` via REF `\n \h` (e.g. `第二章` / `1.2`); for caption-class targets (CaptionBlock + EquationBlock with `captionId`) renders the SEQ result with full decoration via REF `\h` (e.g. `图 2.1` / `(2.3)`).
  - `"number"` — for outline targets, paragraph number via `\r \h`. For caption-class targets **collapses to `"label"` semantics** (same bookmark wraps just the number range).
  - `"full"` — paragraph text content via REF `\h`. Works on outline targets and on non-caption paragraphs. **Throws on caption-class targets** — captions are cited by their decorated number, not body text. Use `"label"` instead.
- `format` — optional `RunFormat` (color, italic, size, …) applied to the rendered text. Format-bearing refs require Word round-trip to verify — see below.

## Target requirements

`label` / `number` require an auto-numbered target:
- Outline / list paragraph bound to a `numbering[]` level (REF `\n` /
  `\r` reads the numbering binding), or
- Caption-class target — `CaptionBlock` or `EquationBlock` with
  `captionId` (bookmark wraps the SEQ-decorated number range).

`full` works on any paragraph with text content. Caption-class
targets throw — use `"label"` to cite captions by their number.

When `label` / `number` hits an unbound target, apply refuses with a
message naming the fix (bind via `numbering[]`, add `captionId`, or
switch to `display: "full"` for a plain-text quote). A source paragraph
that `pattern_rules` / `bulk_rules` will bind counts as unbound here:
rules run after `edits[]`, so the binding isn't visible at ref
resolution — cite via `anchor` on a paragraph the same `edits[]` creates.

## Named anchors — ref paragraphs created in the same apply

When an `edits[]` insert creates a paragraph later refs will cite, give it an `anchor` so the refs can address it by name:

```jsonc
{
  "captions": {
    "Figure": { "prefix": "图", "chapterPrefix": ["Heading1"],
                "bodySeparator": " ", "styleId": "FigureCaption" }
  },
  "edits": [
    {
      "op": "insert-after",
      "at": { "type": "paragraph", "index": 50 },
      "content": [
        { "type": "image", "src": "diagrams/arch.png", "width": "12.7cm", "height": "8.5cm" },
        {
          "type": "caption",
          "captionId": "Figure",
          "anchor": "fig-architecture",
          "text": "系统总体架构"
        }
      ]
    },
    {
      "op": "insert-after",
      "at": { "type": "paragraph", "index": 60 },
      "content": [
        {
          "type": "paragraph",
          "text": [
            { "text": "本文系统的架构如" },
            { "refTo": { "type": "anchor", "name": "fig-architecture" }, "display": "label" },
            { "text": " 所示。" }
          ]
        }
      ]
    }
  ]
}
```

Anchor names live in a **flat namespace** across `edits[]` and source bookmarks. Names must be unique (duplicates throw at pre-scan). A ref can address any name that's declared somewhere in `edits[]` or exists as a source bookmark — declaration order doesn't matter, so forward refs ("如表 1 所示，下表展示..." with the table after the cite) are fine.

`anchor` on an insert with no current ref is still useful: the bookmark stays in the output for future apply runs to ref by name.

## First-open behavior and format verification

The placeholder text between `<w:fldChar fldCharType="separate"/>` and `<w:fldChar fldCharType="end"/>` is pre-computed at apply time from the simulated numbering counters, so the document reads correctly even if the user dismisses Word's field-update prompt. `<w:updateFields w:val="true"/>` is set in `settings.xml`, so the prompt fires automatically on next open; manual `Ctrl+A` → `F9` updates after that.

### Verifying format-bearing refs

Inspecting the apply-time XML alone is **not enough** for refs carrying a `format`. Word rewrites the field's result run on every update; a broken emit (missing rPr replication or `\* MERGEFORMAT`) still validates and reads correctly on first open, but the format disappears the moment Word updates.

Round-trip to verify:
1. Open the apply output in Word
2. Accept the "update fields" prompt and save
3. Unzip and grep `word/document.xml` — confirm the result run still carries the expected `<w:rPr>`

Schema validation and emit-time inspection both pass the broken case; only a real Word round-trip catches the regression.

## Reference list citations

Declare the bracket form **once** in `lvlText` and use `display: "label"` at the cite site. Never type literal `[` / `]` around an InlineRef whose target's lvlText already contains them — the brackets stack visibly.

Cite-render shape depends on the citation standard the doc follows. Read it from user spec, existing chrome in the doc, or ask. Two numeric shapes the InlineRef path covers:

- **Inline bracket** `[1]` — IEEE numeric, ACM, most CS journals, GB/T 7714 行内括注变体. Plain `InlineRef`, no `format`.
- **Superscript** `⁽¹⁾` — GB/T 7714 顺序编码制 (简中学术主流), Vancouver, Nature. `InlineRef` with `format: { "vertAlign": "superscript" }`. **Reference-list cites only** — caption-class refs (figure / table / equation) stay inline normal regardless of standard.

```jsonc
// numbering scheme for the reference list (shared by both shapes):
{ "levels": [
  { "level": 0, "numFmt": "decimal", "lvlText": "[%1]", "suff": "space",
    "styleId": "Reference" }
]}

// inline-bracket cite:
{ "type": "paragraph", "text": [
  { "text": "本文方法借鉴了" },                     // no Pangu space — autoSpace handles CJK↔[
  { "refTo": { "type": "anchor", "name": "ref-smith2024" }, "display": "label" },
  { "text": "中的思路。" }
]}
// → 本文方法借鉴了[3]中的思路。

// superscript cite (same anchor, same display, only format differs):
{ "refTo": { "type": "anchor", "name": "ref-smith2024" }, "display": "label",
  "format": { "vertAlign": "superscript" } }
// → same [3] text, rendered as a raised run via vertAlign — no character substitution
```

Author-year shapes (`(Smith, 2024)` — APA / MLA / GB/T 7714 著者-出版年制) aren't numeric and don't fit this path; out of scope for the auto-numbered InlineRef. Surface to the user when the doc requires this style.

Asymmetric variant (rare — list uses `1.`, cites use `[1]`): leave `lvlText: "%1"`, use `display: "number"`, wrap literal `[` / `]` at the cite site. Don't mix the two — either the brackets live in `lvlText` and nowhere else, or in the cite and nowhere else.

## Limitations

- **No PAGEREF**: page-number citations aren't supported (different field type, depends on Word's pagination layer).
- **Source bookmarks must wrap a single paragraph**: source bookmarks spanning multiple paragraphs or sitting at body level aren't resolvable via `refTo: { type: "anchor" }` — silently skipped at allocator construction.
- **Custom `lvlRestart` not honored in placeholder**: the counter simulator uses Word's default reset-on-higher-level. Custom restart points still produce correct text after Word updates the fields; the pre-update placeholder may diverge. Cosmetic only.
