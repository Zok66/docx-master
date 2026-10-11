# `edits` block (used inside `apply`)

`edits` is a sub-block of `apply`'s config — surgical content + format changes at specific locations: replace / insert / delete paragraphs, swap a table cell, embed an image, restyle a paragraph or range. Optional Word tracked-changes mode.

Before writing `edits[]`, inventory the roles the inserted content will introduce (Heading 1..N, BodyText, ListNumber/Bullet, Caption, cross-refs) and install the matching `styles[]` / `numbering` / `captions` entries first, so each Block binds via `styleId` instead of ad-hoc per-op format. `pattern_rules` slots in alongside when the task spans new structure + chrome retags.

## Reconnaissance first

Always inspect before composing edits. Two `edits`-specific tools beyond the standard survey:

- `inspect_table` — top-level tables with `[row,col]` cell snippets (before composing a `cell` locator).
- `inspect_blockers` — paragraphs `apply` will refuse to touch in the edit phase (existing tracked changes / fields / SDT controls).

## Config shape (within an `apply` config)

```json
{
  "source": "input.docx",
  "output": "output.docx",
  // ... styles[], numbering, pattern_rules, etc. as needed ...
  "edits": [
    { "op": "...", "at": { "type": "..." } /* op-specific fields */ }
  ],
  "trackChanges": false
}
```

`edits[]` runs in array order during `apply`'s pipeline (after style/numbering install, before pattern_rules cleanup). Failures abort atomically; the original file is never modified.

### Locators (`at`)

All position indices — `index`, `from`/`to`, `table`/`row`/`col`, `paragraph`, `blank`, `runIndex` — are **1-based**, uniform across locator types and consistent with `overview` / `inspect_table` / `inspect_runs` / `find_text` display. Exception: `heading.level` is OOXML `outlineLevel` (0-based; 0=Heading 1).

| `type` | Selects |
|---|---|
| `paragraph` | The Nth indexed paragraph (`{ "type": "paragraph", "index": N }`). Matches `#NNN` in `overview`. |
| `range` | `{ ..., "from": A, "to": B }`, inclusive. Endpoints must share a container (body, or one specific layout-table cell). |
| `cell` | `{ ..., "table": T, "row": R, "col": C }` for the whole cell's paragraphs; add `"paragraph": K` to narrow to one paragraph (1-based within the cell), `"to": M` for a contiguous range. Cell coords match `inspect_table`'s `[r,c]` output; K range comes from the `paras:N` annotation shown per cell. |
| `heading` | `{ ..., "text": "...", "level"?: L }`. First paragraph whose rendered text matches and whose OOXML outline level is L (0-based). Disambiguate with `find_paragraphs` if multiple match, then switch to `paragraph` index. |
| `whole-body` | Every body paragraph. Pairs naturally with `format`; rarely with `replace`. |
| `run` | Two shapes: `{ "type": "run", "paragraph": N, ... }` for body or layout-cell paragraphs (global `#NNN`); `{ "type": "run", "table": T, "row": R, "col": C, "paragraph": K, ... }` for data-cell paragraphs (K is 1-based within the cell, matching `inspect_table` `paras:N`). Both accept `blank: K` (Kth whitespace-only underlined placeholder run) or `runIndex: M` (Mth run); default is `blank: 1`. Run indices match `inspect_runs` / `find_text` output. Pair only with `set-run`. |

### Ops

- **`replace`** — `{ ..., "with": [Block, ...] }`. Removes targets, inserts fragment in their place. Pair with `"overwriteFields": true` to allow the replace to absorb paragraphs that contain SEQ/REF/fldChar complex fields (typical caption-iteration scenario). Revisions and SDT remain blocking regardless.
- **`insert-before` / `insert-after`** — `{ ..., "content": [Block, ...] }`. Inserts fragment immediately before / after the target.
- **`delete`** — `{ ... }`. Removes the targeted paragraph(s).
- **`format`** — `{ ..., "styleId"?, "runFormat"?, "paraFormat"?, "clearDirect"? }`. Mutates existing paragraphs without changing their content. At least one of styleId / runFormat / paraFormat / clearDirect required. `clearDirect: ["pPr"] | ["rPr"] | ["pPr","rPr"] | "all"` — drop direct pPr/rPr before applying the new format. Keeps `<w:pStyle>` and `<w:numPr>` on pPr. Note: `["pPr"]` also strips the paragraph-mark rPr; pair with explicit `runFormat` to preserve it.
- **`set-run`** — `{ "at": <run-locator>, "with": "value text", "format"?: { ... } }`. Replaces the targeted run's text while preserving its rPr (font / underline / size carry through). `format` accepts the full `runFormat` field set (bold / italic / underline / strike / doubleStrike / charSpacing / emphasisMark / color / fontLatin / fontCJK / size / vertAlign / rtl); absent, the run's existing rPr stays verbatim.
- **`edit-caption`** — `{ "target": { "anchor": "<name>" } | { "captionId": "<id>", "index": N }, "text": "<string>" }`. Replaces the body text of an existing caption paragraph; preserves SEQ / STYLEREF fields and bookmark so cross-references keep resolving. Throws on EquationBlock targets (no body). See [`captions.md`](captions.md).
- **`merge`** — `{ ..., "keepPPr"?: "first" | "last" }`. Combine the targeted paragraphs (≥ 2) into one; all content children (runs, hyperlinks, bookmark marks, comment marks, etc.) moved into the survivor in document order; surviving `<w:pPr>` from the `keepPPr` end (default `"first"`). Non-survivor paragraphs are removed. Note: if a non-survivor paragraph's bookmark range opened in one paragraph and closed in another, the resulting bookmark span may differ — verify cross-references after merging bookmark-bearing paragraphs. Refused under `trackChanges: true` (run in a separate apply without trackChanges). Also refused if any target paragraph contains tracked changes, a complex field, or an SDT control — accept/reject revisions and resolve fields first.

### MDF (Match Destination Formatting) inheritance

`replace` / `insert-before` / `insert-after` make new `paragraph` blocks inherit the **anchor** paragraph's `<w:pPr>` and the anchor's first non-empty run's `<w:rPr>`. This matches Word's "Match Destination Formatting" paste mode. Anchor: first replaced (replace), first target (insert-before), last target (insert-after). Inheritance is additive at pPr-child granularity — explicit `styleId` / `paraFormat` on the Block always wins. Set `"styleId": "Normal"` to opt out. `image` / `page-break` / `horizontal-rule` blocks don't inherit.

**With explicit `styleId` on the Block**, the engine skips these anchor pPr fields so the style cascade reaches them: `outlineLvl`, `numPr`, `pageBreakBefore`, `widowControl`, `spacing`. `spacing` is included because heading / body styles typically declare paragraph spacing; if you need to preserve the anchor's spacing instead of the styleId's, set explicit `paraFormat.spaceBefore` / `spaceAfter` / `lineSpacing` on the Block. The full `<w:rPr>` from the anchor's pMark is also skipped (so the style's run rPr wins). Without this, a direct `<w:outlineLvl w:val="1"/>` on the anchor would override the heading level that `ProposalH1` declares, misclassifying the new paragraph in the TOC and navigation pane.

**With explicit `runFormat` on a run**, the engine does not inherit the anchor's run rPr for that run — the Block's explicit format is canonical.

### Blocks (in `with` / `content`)

```json
{ "type": "paragraph", "text": "...", "styleId"?, "paraFormat"?, "runFormat"?, "numbering"?, "anchor"? }
{ "type": "image", "src": "path", "width": <length>, "height": <length>, "alt"?, "styleId"?, "paraFormat"? }
{ "type": "table", "rows": [[...]], "headerRows"?, "headerStyle"?, "cols"?, "borders"?, "alignment"?, "vAlign"?, "layout"?, "bidiVisual"?, "padding"?, "sort"? }   ← full schema in [`tables.md`](tables.md)
{ "type": "equation", "latex"?: "...", "omml"?: "...", "styleId"?, "paraFormat"?,
  "captionId"?: "<id>", "subGroup"?: "start"|"continue", "anchor"? }
{ "type": "caption", "captionId": "<id>", "text": "...", "anchor"? }
{ "type": "caption-counter-reset", "captionId": "<id>", "newValue"?: N }
{ "type": "page-break" }
{ "type": "horizontal-rule" }
```

`equation` takes exactly one of `latex` / `omml` (omml is the escape
hatch when temml fails). `captionId` switches to the 3-col numbered
layout; `subGroup` only legal with captionId. See [`equations.md`](equations.md)
+ [`captions.md`](captions.md).

`caption` and `caption-counter-reset` are caption-pipeline blocks —
they need a `captions` table declared at the apply config root. See
[`captions.md`](captions.md).

`anchor` attaches a stable bookmark name (Word's `[A-Za-z_][\w-]{0,39}` rule) so later `InlineRef` nodes can target this new paragraph via `refTo: { "type": "anchor", "name": ... }`. The only way to ref a paragraph created in this same `apply` run — paragraph-index locators reference pre-edit state. See [`cross-references.md`](cross-references.md).

`text` is either a plain string (single run) or an array of inline nodes. Seven inline node types, discriminated by which key is present:

| Shape | Renders |
|---|---|
| `{ text, format? }` | Plain run with optional rPr |
| `{ refTo, display?, format? }` | Cross-reference (REF field) — see [`cross-references.md`](cross-references.md) |
| `{ math }` | Inline equation (OMML) — see [`equations.md`](equations.md) |
| `{ link, text, format? }` | Hyperlink (external URL or `#anchor` jump) |
| `{ field, format? }` | Dynamic field (`"page"` / `"numPages"` / `"date"`) |
| `{ styleRef, numberOnly?, format? }` | STYLEREF — render the nearest paragraph of a given style |
| `{ break }` | `<w:br/>` — `"line"` (soft line break), `"page"` (page break), `"column"` (column break). Use sparingly — prefer paragraph-level breaks over inline. |

**Hyperlinks** (`{ link, text }`) — `link` is any URI Word accepts (`https:` / `http:` / `mailto:` / `tel:` / `ftp:` / …); `#name` targets an internal bookmark — same anchor namespace as `refTo: { type: "anchor" }` (source bookmarks plus `anchor` on Blocks in this apply). The visible `text` is agent-supplied; the `Hyperlink` character style (Word's blue + underline) is applied automatically and injected into styles.xml on first use. Use `refTo` when the visible text should be the target's auto-numbered cite; use `link` when you supply the visible text and just want clickable navigation. Agent-supplied `format` merges on top of the Hyperlink character style — only the keys you set get overridden, the rest of the Hyperlink rPr (e.g. underline) survives.

**Dynamic fields** (`{ field }`) — closed set of three values: `"page"` (PAGE), `"numPages"` (NUMPAGES), `"date"` (DATE, Word's default format). Other field codes aren't supported via this node — author them as plain runs only if you need them. `updateFields` is set so values resolve on first open without manual F9.

**STYLEREF** (`{ styleRef, numberOnly? }`) — renders the nearest paragraph bound to `styleRef` (the styleId, not its display name; e.g. `"Heading1"` not `"Heading 1"`). Most common use: page header that follows the current chapter title. `numberOnly: true` restricts the output to the heading's auto-number, dropping body text — empty when the referenced style isn't bound to numbering. Unknown `styleRef` → Word renders `#ERROR`; verify the styleId is installed via `styles[]` or pre-existing in the source.

### Block-specific notes

Image dimensions are required. To cross-ref a figure, attach `anchor` to its caption paragraph — image paragraphs carry no numbering or text content for `display: "label"` / `"number"` / `"full"` to resolve against. **Table** block has its own schema for rows / cells / merging / borders / column widths / header repetition; details + four progressive cell-content forms in [`tables.md`](tables.md).

For display math, `equation` block carries LaTeX (rendered to native OMML via Temml). Inline math uses `{ "math": "..." }` inside Paragraph.text. See [`equations.md`](equations.md) — LaTeX coverage, caption + cross-ref pattern, and the v1 n-ary operand visual quirk.

**Express structure semantically.** Hierarchy and list shape bind via `styleId` and `numbering` — not by typing markers in `text`. Two paths to numbering:
- **styleId-bound** (preferred): if the styleId you set is bound to a numbering scheme via `numbering[].levels[].styleId`, the binding handles auto-numbering automatically — don't supply a `numbering` field on the block. Use for headings (`Heading1..N`) and list-bound styles (`ListNumber` / `ListBullet`).
- **ad-hoc** `numbering: { numId, level }`: for one-off paragraph-level numbering not tied to a style. Rare; used when you want a paragraph numbered without committing the style to a scheme.

Same applies to images in styled documents: bind `styleId` (typically `FigureImage`) so centering / spacing come from the cascade rather than per-call `paraFormat`. The bare-image emit (no pPr) is only correct for inline / unstyled context.

### Quote handling

`text` is emitted verbatim. Default to smart quotes in prose; ASCII `"` `'` only inside literal tokens (code, URLs, shell). Smart-quote shape is locale-dependent — read it from the document's prevailing language (existing chrome / user prompt / overview metadata), not from agent habit:

| Locale | Outer | Inner |
|---|---|---|
| Simplified Chinese (GB/T 15834) | `“”` | `‘’` |
| Traditional Chinese (台 / 港) | `「」` | `『』` |
| Japanese | `「」` | `『』` |
| English / Western | `“”` | `‘’` |

Picking the wrong locale's shape (e.g. `「」` in mainland 简中正文, or straight `"` anywhere in prose) is a real defect — agent must commit to one locale up front when starting the doc, not flip per paragraph.

### Format fields

`runFormat`: `bold` / `italic` / `underline` / `strike` (boolean, tri-state — `false` emits explicit off-toggle to override inherited true), `doubleStrike` (boolean — `w:dstrike` double strikethrough), `color` (`"RRGGBB"`), `fontLatin` / `fontCJK`, `size` (Length — see below), `charSpacing` (Length — `w:spacing` run spacing, e.g. `"2pt"` widens by 2pt), `emphasisMark` (`"dot"` / `"comma"` / `"circle"` / `"underDot"` / `"none"` — `w:em` 着重号), `vertAlign` (`"superscript"` / `"subscript"` / `"baseline"` — three-state; omit to inherit, use `"baseline"` only to opt out of a super/sub inherited from a character style), `rtl` (boolean — `w:rtl` RTL 文字，配合段落 `bidi` 用于阿拉伯文 / 希伯来文).

Inside a paragraph's `text` array, alongside `{ "text": ..., "format": ... }` runs, you can place `{ "refTo": ..., "display": ..., "format": ... }` cross-reference nodes. **Any cite to an auto-numbered target (figure / table / heading / equation / reference entry) must use these — never write the counter as literal text.** See [`cross-references.md`](cross-references.md) for the contract.

`paraFormat`: `alignment` (`"left" | "center" | "right" | "both"`), `spaceBefore` / `spaceAfter` (Length), `lineSpacing` (number multiplier / `"Npt"` exact / `{ atLeast: <length> }`), `firstLineIndent` / `hangingIndent` / `indentLeft` / `indentRight` (Length, plus `"Nchar"` for font-scaled indent; `0` / `"0pt"` explicitly overrides an inherited indent to zero, `null` leaves the field absent so the style cascade decides), `outlineLevel` (0–9), `bidi` (boolean — `w:bidi` 段落从右向左基准方向，RTL 段落必开).

**Length values** — anywhere a length is expected (`size`, `spaceBefore` / `spaceAfter`, `width` / `height`, border `size`, indent forms, table column widths) accept the same shape: a bare number (pt) or one of `"Npt"` / `"Ncm"` / `"Nmm"` / `"Nin"`. Use whatever matches the user's prompt — `"2.54cm"` is no extra cost over `"72pt"`. Indents additionally accept `"Nchar"` (Word's "首行缩进 N 字符" — auto-scales with font size; preferred for CJK body text where character-units survive size changes).

## Track-changes mode

`"trackChanges": true` emits edits as Word revision markup: text changes via `<w:ins>` / `<w:del>`, format changes via `<w:rPrChange>` / `<w:pPrChange>` snapshots. The user accepts / rejects in Word's Review tab. Paragraphs that *already* contain tracked changes are blocked; ask the user to accept / reject existing revisions first.

`"author"` (optional, only meaningful with `trackChanges: true`) — non-empty string written to the `w:author` attribute of every emitted `<w:ins>` / `<w:del>` / `*Change`. Omitted ⇒ empty value ⇒ Word displays "Unknown Author". Never default this to a tool brand; only set it when the user explicitly names who's authoring the review.

## Cell-fill strategy

Form cells often hold a label paragraph + several empty placeholder rows (the form designer pre-allocates blanks for handwriting). Two strategies:

- **`replace` a range** covering the empty placeholder rows: drops the placeholders, only the new content remains. Cleaner cell, no trailing blank rows.
- **`insert-after` the label**: leaves the empty placeholder rows below the inserted content. Word auto-grows the cell; visually fine for most forms but trailing blank rows are visible.

Default to `replace` for content-bearing fills; `insert-after` for non-destructive additions where the original layout is meant to be preserved.

For multi-paragraph content, pass multiple Blocks in one op rather than chaining many `insert-after`s — order is preserved and no stale-index issue.

`apply --dry-run` emits a `=== Paragraph index drift (pre-edit → post-edit) ===` section when the edits change body paragraph count. Use it to spot whether a downstream locator (`paragraph` index, cross-ref) you wrote against the pre-edit numbering still points where you intended.

## Edge cases

- **Cross-container range**: `range` cannot span body ↔ table-cell. Split into two edits.
- **Heading locator ambiguity**: returns first match. Use `find_paragraphs` to disambiguate, then refer by `paragraph` index.
- **Stale element**: if op A removes paragraphs and op B targets one of them, op B fails. Reorder so deletes / replaces happen last, or split into separate runs.
- **Empty body**: `whole-body` on a doc with zero paragraphs only accepts `insert-*` (appends before the trailing `<w:sectPr>`).

## Compose with other shapes

For messy templates, combine `styles[]` / `numbering` / `pattern_rules` (standardize-shape) with `edits[]` in one `apply` config — the engine installs structure first so `edits[]` references just-installed styleIds, and MDF doesn't propagate template chrome into your inserts. Role-based whole-doc reshape without per-locator edits → [standardize.md](standardize.md). Read-only check → `audit`.
