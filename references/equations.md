# Equations (`{ "type": "equation", ... }` and `{ "math": "..." }`)

LaTeX → OMML rendering for display equations and inline math. Numbered equations use the captions pipeline (SEQ + STYLEREF + bookmark, Word's native caption mechanism) — see [`captions.md`](captions.md).

## Quick start

```jsonc
// Unnumbered display
{ "type": "equation", "latex": "E = mc^2" }

// Numbered display (chapter-prefixed when captions.Equation declares chapterPrefix)
{ "type": "equation", "latex": "a^2 + b^2 = c^2",
  "captionId": "Equation", "anchor": "eq-pythagoras" }

// Inline (inside a paragraph's text array)
{ "type": "paragraph", "text": [
  { "text": "By Pythagoras, " },
  { "math": "a^2 + b^2 = c^2" },
  { "text": "." }
]}
```

## Schema reference

```ts
EquationBlock {
  type: "equation"
  latex?: string              // LaTeX expression (Temml subset — see below)
  omml?: string               // raw OMML (escape hatch when temml fails)
                              // exactly one of latex / omml required
  styleId?: string            // paragraph style for the equation's <w:p>;
                              // defaults to "Equation"
  paraFormat?: ParagraphFormat
  captionId?: string          // declares a numbered equation; references
                              // captions[<id>]. Omit for unnumbered.
  subGroup?: "start" | "continue"  // subequations: (1a)(1b)(1c). Requires
                              // captionId + captions[<id>].subCounter.
  anchor?: string             // bookmark name; requires captionId
}

InlineEquation { math: string }   // inside Paragraph.text[] alongside { text, format } and { refTo }
```

## Numbered equations — layout

With `captionId` set, the engine emits a 3-column borderless table:

```
[ left spacer | centered OMML | (chapter.counter) ]
```

Right cell is a SEQ + STYLEREF caption with the prefix / suffix / chapter prefix declared in `captions[<id>]`. The whole "number + decoration" range carries a bookmark named after `anchor`; body-text InlineRefs resolve to the rendered caption text (e.g. `"(2.3)"`) via REF `\h`.

Without `captionId`, the engine emits a single centered paragraph with just the OMML — no table, no caption, no bookmark.

## LaTeX coverage

The renderer is [Temml](https://temml.org/) — supports most of LaTeX math mode. When a specific expression triggers temml's known issues, fall back to the `omml` escape hatch: pass pre-converted OMML directly and the engine embeds it as-is.

## Subequations — `subGroup`

```jsonc
// (1) standalone — followed by a sub-group
{ "type": "equation", "latex": "x = 1", "captionId": "Equation", "anchor": "eq-1" },

// (2a) subgroup start — parent counter advances, sub resets to a
{ "type": "equation", "latex": "y = 2", "captionId": "Equation",
  "subGroup": "start", "anchor": "eq-2a" },

// (2b) subgroup continue — parent unchanged, sub advances to b
{ "type": "equation", "latex": "z = 3", "captionId": "Equation",
  "subGroup": "continue", "anchor": "eq-2b" },

// (3) standalone after subgroup — parent advances normally
{ "type": "equation", "latex": "w = 4", "captionId": "Equation", "anchor": "eq-3" }
```

Requires `captions["Equation"].subCounter` to be declared (otherwise the schema rejects `subGroup`).

## Cross-references

`InlineRef` targets the equation's anchor:

```jsonc
{ "refTo": { "type": "anchor", "name": "eq-pythagoras" }, "display": "label" }
```

Returns `prefix + chapter + counter + suffix` from the SEQ result (e.g. `"(2.3)"`). Caption-class display routing detailed in [`cross-references.md`](cross-references.md).

## Integration with the style system

| What | How |
|---|---|
| Numbered vs unnumbered | Presence of `captionId` on the EquationBlock |
| Caption format | Declared once in `captions[<id>]` — prefix / suffix / format / chapterPrefix / styleId |
| Equation paragraph style | `EquationBlock.styleId` (default `"Equation"`); separate from the caption paragraph style |
| pattern_rules / bulk_rules | Apply to surrounding paragraphs; the equation paragraph contains only OMML so text-pattern matching skips it. |

## Edge cases

- **`captionId` references undeclared identifier.** Schema validation throws — declare the entry in `captions` first.
- **`anchor` without `captionId`.** Schema throws — without numbering, REF \h has no resolved target to return.
- **`display: "full"` on EquationBlock target.** Pre-scan throws — the equation paragraph has no body text.
- **trackChanges + equation insert.** Engine throws at emit. OOXML's tracked-change wrappers don't have a clean "equation inserted" shape. Run equation insertion without trackChanges; use trackChanges for subsequent surrounding edits.
- **Inline math inside an InlineRef's display text.** Not supported. Embed the InlineRef and the math as sibling InlineNodes.

## Known fragile LaTeX tokens

Use the `omml` escape hatch when:

**Throws at apply time** — error names `edits[N]` + LaTeX source:

- Strike-through macros (`\cancel`, `\bcancel`, `\xcancel`, `\sout`) — `<menclose>` strike notations have no OMML peer. `\fbox{}`, `\overline{}`, `\underline{}` work.

**Renders approximately** — no error, visual differs from LaTeX:

- `\big` / `\Big` / `\bigg` / `\Bigg` standalone delimiters — default size.
- Graded spacing (`\!`, `\,`, `\;`, `\quad`, `\hspace{}`) — collapsed to one space.
- Overlap macros (`\rlap`, `\llap`, `\mathrlap`, `\mathllap`) — overlap lost.
- Decorative wrappers (`\boxed`, `\colorbox`, `\mathbackground`) — border/background lost.

## What's not supported

- **LaTeX `\tag{}` / `\label{}` / `equation` environment.** Use the `captionId` + `anchor` pattern instead.
- **Editing an existing equation's LaTeX in place.** Locators target paragraphs, not OMML subtrees. To change an equation, `replace` the whole paragraph (or `delete` then `insert-after`).
