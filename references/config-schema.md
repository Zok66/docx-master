# `apply` Config Schema

Full field-by-field reference for the JSON config consumed by
`apply [--dry-run] <config.json>`.

## Length values

Anywhere a length is expected — `size`, `spaceBefore` / `spaceAfter`,
image `width` / `height`, border `size`, table column widths, paragraph
indents — the accepted shape is uniform:

| Form | Meaning | Example |
|---|---|---|
| `number` | pt (bare) | `12` = 12 pt |
| `"Npt"` | pt explicit | `"12pt"` |
| `"Ncm"` | centimeters | `"2.54cm"` |
| `"Nmm"` | millimeters | `"5mm"` |
| `"Nin"` | inches | `"1in"` |

Pick whichever matches the user's prompt verbatim — `"2.54cm"` and
`"72pt"` cost the agent the same to write; the engine converts to OOXML
internal units.

Indents (`firstLineIndent` / `hangingIndent` / `indentLeft` /
`indentRight`) additionally accept `"Nchar"` — Word's "首行缩进 N 字符",
emitted as `w:firstLineChars`, auto-scales with the run's font size.
Prefer `"2char"` for CJK body indents so the indent stays correct when
font size changes.

`lineSpacing` has three modes determined by input type:

- **`number`** → multiplier (auto rule). `1.5` = 1.5× line height.
- **`"Npt"` / `"Ncm"` / ...** → exact line height (exact rule).
- **`{ "atLeast": <Length> }`** → at-least line height (atLeast rule, rare).

Bare numbers are always multipliers — no magnitude heuristic. Use the
string form when you mean an exact line height.

## Top-level shape

```jsonc
{
  source: "/path/to/original.docx",  // optional. Path to the input file.
                                     // This file is NEVER modified — apply
                                     // copies it first, then writes the modified
                                     // copy. Omit to scaffold from the bundled
                                     // blank template (one empty Normal paragraph,
                                     // A4 portrait). Required when a `template`
                                     // block is declared.
  output: "/path/to/output.docx",    // REQUIRED. Path for the new file. Must differ from source.
  dryRun: false,                     // optional. true = in-memory pipeline + report, no file written.
                                     // Equivalent to the --dry-run CLI flag.
  allowValidationWarnings: false,    // optional. Keep the output docx even when this run
                                     // introduces new OOXML validation errors.
  mergeRuns: false,                  // optional. Coalesce adjacent identically-formatted
                                     // runs (strips rsid* / proofErr) before anything else
                                     // runs — for documents of unknown provenance.

  template:   { ... },               // optional. See § "Template import" below.
  theme:      { fonts: {...} },      // optional. See § "Theme" below.
  pageSetup:  { ... },               // optional. See § "Page setup" below.
  headerFooter: { header?, footer? }, // optional. Each surface declares one or
                                     //   more of `default` / `first` / `even`.
                                     //   See references/header-footer.md.
  styles:     [ ... ],               // required when the task reshapes typography;
                                     //   optional for edit-only configs. See § "Style entries".
  settings:   { elements: [ ... ] }, // optional. Raw <w:settings> children — updateFields,
                                     //   docVars, mailMerge, print / correction flags. See below.
  docProps:   { enabled, title, creator }, // optional, on by default — maintains
                                     //   docProps/core.xml + app.xml. See below.
  lint:       { panguSpacing: false }, // optional. Silences the advisory CJK↔Latin spacing report.
  numbering: { levels: [...] }       // optional. Single scheme, OR array of
            | [ {levels:[...]}, ... ], // schemes for parallel installations
                                     // (e.g. multi-level heading + single-level list).
  captions: { "<id>": { ... } },     // optional. Caption-class numbering (figure /
                                     // table / equation / theorem / ...). See
                                     // § "Captions" below; rendering pipeline in
                                     // captions.md.
  edits:    [ { op: "...", ... } ],  // optional. Location-based surgical edits
                                     // (replace / insert / delete / image / caption /
                                     // equation). See references/edit.md.
  trackChanges: false,               // optional. Wrap this run's edits in <w:ins>/<w:del>.
  author:   "...",                   // optional. Only with trackChanges: true — revision author.
  requirements: { id: "..." },       // optional. Annotation only — see § "Requirements" below.

  // Paragraph-to-style mapping, in resolution order:
  exclude:        [1, 2, 3],         // array of paragraph indices (numbers, not objects);
                                     //   see § "Paragraph mapping"
  assignments:    [ ... ],
  pattern_rules:  [ ... ],
  bulk_rules:     [ ... ],
}
```

## Top-level config fields

- **`source`** (string, optional) — path to the input docx. The original file is never modified. Omit to scaffold from a blank template (one empty Normal paragraph, A4 portrait). Required when a `template` block is declared.
- **`output`** (string, **required**) — path for the output docx. Must differ from `source`.
- **`dryRun`** (boolean, default `false`) — when `true`, run the pipeline in memory and produce a change report without writing a file. Equivalent to the `--dry-run` CLI flag.
- **`allowValidationWarnings`** (boolean, default `false`) — keep the output docx even when apply introduces new OOXML validation errors. Pre-existing source errors are always non-fatal (baseline-diff). CLI: `--allow-validation-warnings`. See [standardize.md](standardize.md) Validation behavior.
- **`mergeRuns`** (boolean, default `false`) — coalesce adjacent identically-formatted runs in `word/document.xml` before the pipeline starts, and strip `rsid*` attributes plus `<w:proofErr>` markers. Word splits a paragraph's text across many runs (revision ids, spell-check markers), so a phrase visible on the page is often not one contiguous string in the XML — text-matched anchors and `set-run` miss it. Because merging changes the run layout, the decision has to be made *before* any locator is written: enable it for documents of unknown provenance and address runs against the merged layout, rather than hand-addressing `runIndex` values on the original. Rendering is unchanged; `<w:ins>` / `<w:del>` wrappers are preserved. CLI: `--merge-runs`. Standalone equivalent: `node scripts/merge_runs.js`. See [SKILL.md → "Documents of unknown provenance"](../SKILL.md).
- **`settings`** (object, optional) — `{ elements: [{ name, attrs?, text?, children? }] }`. Raw `<w:settings>` children written through a name-keyed upsert: each entry merges into an existing element of the same name, or is inserted at its XSD-correct position in the 98-child order. Repeatable children (`docVar`, `rsid`, …) append one node per declaration. `name` and `attrs` keys are **bare** (`val`, `name`, `eastAsia`) — the `w:` prefix is implied; write `name: "docVars"`, not `"w:docVars"`. Covers `updateFields`, `displayBackgroundShape`, `proofState`, `hideSpellingErrors`, `mailMerge` (with `mainDocumentType` / `dataType` / `destination` children), `docVars` (with `docVar` children), `defaultTabStop`, print flags, `themeFontLang`, …
- **`docProps`** (object, optional, on by default) — `{ enabled?, title?, creator? }`. Maintains `docProps/core.xml` + `app.xml`: title / creator / revision, plus `Words` / `Characters` / `Paragraphs` counted the way Word counts Chinese (CJK per character, Latin per word). These parts are what `NUMPAGES` / `NUMWORDS` / `AUTHOR` fields read. `creator` is only applied when the part is new or still carries the tool's own previous value — a human-set author is never overwritten.
- **`lint`** (object, optional) — `{ panguSpacing?: boolean }`. The CJK↔Latin spacing audit is **advisory only** (reported in `lint.pangu`, never blocks the write); `panguSpacing: false` silences it for deliberate spacing.
- **`trackChanges`** (boolean, default `false`) — wrap this run's edits in `<w:ins>` / `<w:del>` markup. See [edit.md](edit.md) Track-changes mode.
- **`author`** (string, optional) — only meaningful with `trackChanges: true`; written to revision markup's `w:author` attribute. Omitted ⇒ Word displays "Unknown Author" (never defaulted to a tool brand).

## Style entries

`styles[]` holds one entry per semantic role. It is **required whenever the task reshapes typography** (the standardize shape) and may be omitted entirely for edit-only / header-only configs — the engine writes only what is declared. Each entry follows one of two shapes (mix freely across the array — pick per role).

### Mode A — extract from a representative paragraph (required when source has one)

```jsonc
{
  id:            "Heading1",   // REQUIRED. Style ID written into styles.xml.
                               // Use Word built-ins (Heading1, BodyText, Caption, ...)
                               // when the role matches — keeps TOC / nav pane / outline
                               // working without extra wiring.
  name:          "一级标题",    // REQUIRED. Display name shown in Word's style panel.
  basedOn:       "Normal",     // optional. Default "Normal".
  fromParagraph: 33,           // 1-based paragraph index. The tool extracts that
                               // paragraph's full computed rPr + pPr (using the
                               // dominant text run, skipping numbering-prefix-only
                               // runs) and uses them as the style definition.
                               //
                               // CORRECTNESS RULE: Required when the source already
                               // has any paragraph playing this role (any pre-existing
                               // styleId you are refining). Top-level typography fields
                               // (size, bold, alignment, lineSpacing, etc.) on a
                               // represented-role entry silently override the template's
                               // actual typography — use `overrides` for explicit
                               // adjustments instead. See standardize.md §1 for rationale.
  overrides: {                 // optional. Any field listed under Mode B can appear here.
    outlineLevel: 0,           // typical use: add structural fields the source lacks
    alignment:    "left",      // or override a specific value per user request
  },
}
```

### Mode B — define manually (empty-slot / user-spec-only fallback)

```jsonc
{
  id:              "Caption",  // REQUIRED.
  name:            "图表注",    // REQUIRED.
  basedOn:         "Normal",   // optional. Default "Normal".
  fontLatin:       "Arial",    // optional. Latin / Western text font.
  fontCJK:         "黑体",     // optional. CJK font. Most common field in
                               //   Chinese-academic configs.
  size:            10.5,       // optional. Length: bare = pt, or "Npt"/"Ncm"/"Nmm"/"Nin".
  bold:            false,      // optional. Default false.
  italic:          false,      // optional. Default false.
  color:           "auto",     // optional. Hex ("2E75B6") or "auto". Default "auto".
  vertAlign:       "superscript", // optional. "superscript" | "subscript" | "baseline".
                               //   Used on character styles (FootnoteReference,
                               //   EndnoteReference, or custom super/sub styles).
                               //   "baseline" is the explicit reset — distinct from
                               //   omitting the field (= inherit cascade).
  alignment:       "center",   // optional. "left" | "center" | "right" | "both".
  lineSpacing:     1.5,        // optional. Three forms, mode chosen by type:
                               //   number       → multiplier (auto), e.g. 1.5, 2
                               //   "Npt"/etc.   → exact line height
                               //   { atLeast }  → at-least line height
                               //   Bare numbers are ALWAYS multipliers — no magnitude
                               //   heuristic. Use "20pt" for exact 20-pt line height.
  spaceBefore:     12,         // optional. Length before paragraph.
  spaceAfter:      6,          // optional. Length after paragraph.
  firstLineIndent: "2char",    // optional. Length, "Nchar", or null.
                               //   "Nchar" → emitted as `w:firstLineChars` (1/100 char),
                               //     auto-scales with run font size — required for the
                               //     standard "首行缩进 2 字符" academic convention.
                               //   Other Length → emitted as `w:firstLine` (fixed
                               //     twips), does NOT scale with font.
                               //   `null` (or omitted) → no indent emitted; for
                               //     existing paragraphs the cascade decides.
                               //   `0` / `"0pt"` → explicitly emit zero-indent
                               //     (overrides an inherited indent to nothing).
                               //   Prefer "Nchar" for thesis/paper body text.
  hangingIndent:   null,       // optional. Same units as firstLineIndent. For
                               //   bibliography/reference entries use "2char" (or
                               //   pt/cm to match the leading "[N] " marker width).
  outlineLevel:    0,          // optional. 0–9: 0–8 are heading levels
                               //   (0 = H1, 1 = H2, …); 9 = body text. Set on
                               //   heading styles to enable TOC / outline view /
                               //   nav pane.
}
```

Modes mix in the same array. `overrides` companions `fromParagraph` — layer additions / replacements on top of the extracted shape.

## Numbering

```jsonc
numbering: {
  levels: [
    {
      level:   0,                        // REQUIRED. 0-8.
      numFmt:  "chineseCounting",        // REQUIRED. decimal / chineseCounting /
                                         //   bullet / lowerRoman / etc.
                                         //   See references/numbering-formats.md for table.
                                         //   Note: chineseCounting and chineseCountingThousand
                                         //   produce the same visible glyphs (一、二、三…) —
                                         //   prefer chineseCounting as the canonical form.
      lvlText: "第%1章",                 // REQUIRED. Display pattern (%N = level N counter).
      styleId: "Heading1",               // REQUIRED. Binds this level to a paragraph style.
                                         // PREREQUISITE: must either pre-exist in styles.xml
                                         // or be declared in styles[]. Otherwise apply throws
                                         // at install time with "style not found".
      start:   1,                        // optional. Starting number. Default 1.
      stripPrefixPatterns: ["%1.%2", "%1."],
                                         // optional. Alternative manual-prefix patterns
                                         //   to strip from paragraph text (tried in
                                         //   order, longest first). Use when the source
                                         //   mixes styles — some H2 written as "1.1 ...",
                                         //   others as "1. ...". Defaults to [lvlText].
      suff: "nothing",                   // "tab" | "space" | "nothing". Gap between
                                         //   marker and text. See numbering-formats.md
                                         //   for per-level guidance. If omitted, inferred
                                         //   from trailing whitespace in `lvlText` (0 →
                                         //   "nothing", 1 → "space", 2+ → "tab"); explicit
                                         //   is the intended path.
      numRPr: {                          // optional. rPr applied to the auto-generated
        color: "3370FF",                 //   number marker only — independent of the
        bold:  false,                    //   title text.
      },
      // Counter scope for single-level schemes. Default "continuous" — one
      // counter shared by every paragraph bound to this level.
      // Accepted values:
      //   "continuous": one numId, items continue regardless of intervening paragraphs.
      //   "perInstance": fork a fresh numId per contiguous run of same-styleId
      //       paragraphs so each list block restarts at 1 (procedural 1./2./3. lists).
      //   "byHeading": restart whenever the nearest preceding paragraph with
      //       outlineLvl changes.
      //   { "atStyleChange": "<styleId>" }: restart at every paragraph bound
      //       to the named styleId.
      //
      //   Block-level override: numbering: { numId, level, restart: true } forks a
      //   fresh numId with <w:startOverride val="1"/> at one paragraph — use when a
      //   single mid-list position needs a hard reset this field can't express.
      //
      //   For SEQ-based per-chapter caption numbering use captions.chapterPrefix
      //   instead — see § "Captions" below.
      //
      //   Ignored on multi-level schemes (which use lvlRestart on each level for resets).
      restart: "continuous",
    },
    ...
  ]
}
```

`restart` is a **level-entry** field — it lives inside `levels[i]`, not at the scheme root. For single-level schemes (the common case for `byHeading` / `atStyleChange`), this means setting it on the single level entry. The behavior is conceptually scheme-wide because single-level schemes have only one counter, but the JSON placement is inside `levels[i]`.

Omit `numbering` entirely if the document has no numbered headings/lists.

### Explicit `numId` on a scheme

By default the engine allocates fresh numIds for declared schemes. To pin a scheme to a specific id — so block-level `numbering: { numId }` references resolve predictably — set `"numId": N` on the scheme object (sibling of `levels`). The dry-run report includes a scheme → numId allocation table showing which id each scheme was assigned and whether it was pinned or allocated.

The engine always creates a fresh `<w:abstractNum>` + `<w:num>` pair. If the pinned `numId` is already present in the source `numbering.xml`, apply throws with a clear message naming the conflicting entry and instructing the agent to choose a different id or remove the explicit `numId` to let the engine allocate. The engine's auto-allocator also skips all source numIds, so allocated ids are always safe regardless of what the source contains.

Collision: two config-declared schemes requesting the same `numId`, or any config scheme requesting a `numId` already in the source, causes apply to throw, naming the conflicting entry.

Pattern templates and `numFmt` values: see [`numbering-formats.md`](numbering-formats.md).

## Captions

```jsonc
captions: {
  "<id>": {
    prefix?: string,          // literal before the counter (default "")
    suffix?: string,          // literal after the counter (default "")
    format?: "arabic" | "alphabetic" | "ALPHABETIC"
           | "roman" | "ROMAN" | "chinese" | "chinese-formal",
                              // default "arabic"
    chapterPrefix?: Array<    // ordered, any depth; default [] (global, no restart)
      string                  //   bare styleId — use heading's native number rendering
      | { styleId: string;    //   force format (re-renders as Arabic/alphabetic/roman/...
          format?: SeqFormat }//   regardless of heading's native numFmt)
    >,                        //   The 中文 academic case: H1 displays "第一章", captions
                              //   read "图 1.1" — use { styleId: "Heading1", format: "arabic" }
    chapterSeparator?: string,// joins chapter levels + counter (default ".")
    bodySeparator?: string,   // between counter and CaptionBlock.text (default " ")
    styleId: string,          // REQUIRED — caption paragraph's style
    subCounter?: {            // enables subequations (1a)(1b)
      format?: "arabic" | "alphabetic" | ...,  // default "alphabetic"
      prefix?: string,        // default ""
      suffix?: string         // default ""
    }
  }
}
```

`chapterPrefix` is the SEQ-based per-chapter caption mechanism. It is distinct from `restart: "byHeading"` on a numbering scheme: `chapterPrefix` drives caption counters (figures / tables / equations) via SEQ fields + a hidden parallel chapter SEQ injected into heading paragraphs; `byHeading` drives list-class auto-numbering restart inside `numbering[]`. They address different counter classes; do not substitute one for the other.

Block-level types used with captions — `CaptionBlock`, `EquationBlock.captionId`, `caption-counter-reset` — and the full rendering pipeline: see [`captions.md`](captions.md).

## Requirements (annotation only)

```jsonc
requirements: {
  BodyText: "正文请使用宋体小四号字，行距设为1.5倍，首行缩进两个字符",
  Heading1: "标题用黑体三号加粗居中显示",
}
```

The script does NOT parse this. It records each string and displays it side-by-side with the agent-resolved structured fields under "=== Style Resolution ===" in the change report. The agent (you) translates the natural-language spec into the structured `styles[i]` fields above — Chinese typography requires understanding negation ("不要加粗"), hierarchical references ("比一级小一号"), Chinese numerals ("两个字符"), and synonyms (思源黑体 / 苹方 / 方正小标宋), which a fixed regex parser silently fails on.

The Style Resolution display is the verification mechanism — agent / user / reviewer reads both lines and confirms the translation matches intent.

## Template import

```jsonc
template: {
  source: "/path/to/template.docx",  // path to the reference docx
  styles: ["BodyText", "Heading1", "Heading2", "Heading3", "Caption"],
                                     // styleIds to import. basedOn ancestors are
                                     // auto-pulled if they don't exist in source.
  importNumbering: true,             // default true. When an imported style references
                                     // a numId, the corresponding abstractNum + num
                                     // are migrated to fresh IDs in the source's
                                     // numbering.xml. The imported style's numPr is
                                     // rewritten to point at the new numId.
}
```

If a styleId already exists in source, the template's definition wins — the template is treated as the authoritative stylebook for the listed IDs. Other source styles are untouched.

## Theme

```jsonc
theme: {
  fonts: {
    majorLatin:    "Times New Roman",  // optional. Heading/Title Latin font (Office major scheme).
    majorEastAsia: "黑体",              // optional. Heading/Title CJK font.
    minorLatin:    "Times New Roman",  // optional. Body Latin font (Office minor scheme).
    minorEastAsia: "宋体",              // optional. Body CJK font.
  },
}
```

Modifies `word/theme/theme1.xml`. Any `docDefaults` / `styles[]` / direct rPr that references theme fonts (`<w:rFonts w:asciiTheme="majorHAnsi"/>` etc.) auto-resolves to the new values — the document-design layer instead of per-style declarations.

Use when the user wants the doc's underlying font scheme changed ("把这份文档的主题字体改成 X / Y"), not when they want one specific style restyled. Sparse: declare only the slots being changed; omitted slots keep the source's existing theme value.

## Page setup

Mutates `<w:sectPr>` children — paper size, orientation, margins, columns, page-number format. Sparse-by-design: only declared fields change; undeclared `<w:sectPr>` attributes (headerReference, type, docGrid, …) stay intact. Top-level fields apply to every section; `sections.<selector>` overrides specific sections.

```jsonc
pageSetup: {
  paperSize:   "A4",                  // or "A3"/"A5"/"Letter"/"Legal"/"B5"/"16K", or { width, height }
  orientation: "portrait",            // "portrait" | "landscape"
  margins: {                          // all Length, all optional, per-edge merge
    top: "2.54cm", bottom: "2.54cm", left: "3.17cm", right: "3.17cm",
    header: "1.5cm",                  // distance from page edge to header text
    footer: "1.75cm",
    gutter: "0cm",
  },
  columns: 2,                         // equal-width count; see column forms below
  pgNumType: {                        // page number numerals + restart per section
    fmt: "decimal",                   // "decimal" | "upperRoman" | "lowerRoman" | "upperLetter" | "lowerLetter"
    start: 1                          // optional; omit to continue from previous section
  },
  sections: {                         // only when some sections differ from defaults
    "1":   { pgNumType: { fmt: "lowerRoman" } },
    "2-3": { pgNumType: { fmt: "decimal", start: 1 } },
  },
}
```

### `columns` forms

- **`columns: N`** — equal-width count, default 0.5cm gap.
- **`columns: { count, space?, separator? }`** — equal-width with custom gap / vertical separator line.
- **`columns: { widths: [...], spaces?: [...], separator? }`** — unequal widths. `spaces.length` must equal `widths.length - 1`. Auto-sets OOXML `equalWidth="false"`.

`count` and `widths` are mutually exclusive — declare exactly one.

### `sections.<selector>` keys

- `"N"` — section N (1-based, matches `inspect_section <N>`).
- `"N-M"` — sections N through M inclusive.

Multiple selectors overlapping on the same section layer in object key order: later wins per field, and `margins` merges per-edge across layers (declaring `top` in one layer and `left` in a later layer leaves both set).

### Field semantics

- **margins** and **pgNumType** are per-field merged: declaring `top` (or `fmt`) doesn't touch the other fields. Omitted fields keep the source value (or the top-level default when overriding in a section).
- **orientation alone** reads current pgSz w/h and swaps as needed. Sections with no existing pgSz reject orientation-only — declare `paperSize` alongside.
- **columns** is replaced wholesale when declared. To leave columns untouched, omit the field.

The dry-run report includes per-section before → after — verify the selectors hit intended sections.

## Paragraph mapping

Paragraphs are 1-based, matching `#001`, `#002` labels in the overview skeleton. Paragraphs inside layout tables are included in the numbering; paragraphs inside data tables are not indexed.

```jsonc
exclude: [1, 2, 3],                    // never touch — overrides everything else
assignments: [                          // per-paragraph rules (highest precedence
  { para: 1,  action: "keep" },         //   among the non-exclude branches)
  { para: 33, action: "restyle", style: "Heading1" },
  { para: 47, action: "flag",    reason: "Ambiguous: looks like heading but no numbering" },
],
pattern_rules: [                       // regex-based by paragraph text
  { regex: "^图\\s*\\d+[-.]\\d+", style: "FigureCaption" },
  { regex: "^表\\s*\\d+[-.]\\d+", style: "TableCaption" },
  { regex: "^\\[\\d+\\]\\s+",     style: "Reference" },
  { regex: "^(关键词|Keywords?)\\s*[:：]", style: "Keywords", stripMatch: true },
  // First match wins (rules tried in order). The match must be anchored at the
  // start of paragraph text. `stripMatch: true` removes the matched leading text —
  // useful when the new style replaces the label via numbering or other mechanism.
],
bulk_rules: [                          // by visual fingerprint label from overview
  { fingerprint: "D", style: "BodyText" },
  { fingerprint: "E", style: "Code" },
],
```

### Resolution order

For each paragraph, the first matching branch wins:

1. `exclude` — paragraph is untouched. Full stop.
2. `assignments` — per-paragraph rule.
3. `pattern_rules` — first regex match anchored at start of text.
4. `bulk_rules` — match by fingerprint label.
5. No match — paragraph is left unchanged (implicit `keep`).

### Assignment actions

- `keep` — preserve original formatting exactly. `style` not used.
- `restyle` — apply the named style and strip conflicting direct formatting. `style` is required.
- `flag` — do not modify the paragraph; record it in the change report with the supplied `reason`. Use when the role assignment is genuinely ambiguous.

## Style-field resolution priority

For each entry's final field values, layered low → high (later wins):

1. Defaults.
2. `template` imported style (if same ID).
3. `styles[i].fromParagraph` extracted values.
4. `styles[i]` direct fields (fontLatin / fontCJK / size / bold / ...).
5. `styles[i].overrides` — deliberate per-style escape.

`requirements` is annotation only and does not participate in resolution.
