# Standardize-shape `apply` config

Standardize adds **structure** (semantic styleId, outlineLevel, auto-numbering, caption SEQ, cross-reference fields). Typography stays where the template put it — touch font / size / alignment only when (a) the user prompt explicitly names them, or (b) chrome is inconsistent enough that no exemplar can be extracted. Reading "标准化 / 排版" as "reset typography to canonical values" is the most common agent failure on this skill.

Operates by **pattern**: regex / fingerprint / styleId selects categories; engine applies uniformly. Per-paragraph enumeration is the last resort. Surgical edits at specific indices or table cells go through `edits[]` — see [`edit.md`](edit.md); both shapes combine in one `apply` config. Full schema: [`config-schema.md`](config-schema.md).

## How standardize thinks

Source paragraphs carry typography (font, size, indent, alignment) and structure (heading level, list membership, caption identity) interwoven through direct `pPr` / `rPr` and typed-prefix text. Standardize separates them:

- **Typography is preserved.** Where a paragraph already plays a role, `fromParagraph` extracts its typography into a fresh semantic style. Direct format on chrome paragraphs stays as-is.
- **Structure is re-bound.** Typed numbering prefixes (`一、`, `1.1`, `图 2.1`, `如图 3.2 所示`) come out of the text and into the right mechanism — auto-numbering scheme for heading/list prefixes, caption SEQ fields for figure/table/equation numbers, REF fields for in-prose cross-references.
- **Sparse by design.** Untouched paragraphs stay untouched; undeclared styles stay as they are. Declare only what's wrong, missing, or what the user explicitly asks to change.

## Workflow

1. **Survey.** `overview` first.
2. **Design ONE config** spanning the four decisions below. Reactive additions accrete debt.
3. **Dry-run** (`apply --dry-run`). Read each signal against intent — see §5.
4. **Apply.** Output is a fresh docx; the original is never modified.

---

## 1. Style roles & typography

Target: one style per semantic role the doc + content combined contain (`Heading1..N`, `BodyText`, `ListNumber`, `FigureImage` + `FigureCaption`, `TableCaption`, `EquationNumber`, `Reference`, etc.).

### `fromParagraph` extraction (default for represented roles)

If a role has any paragraph already playing it in the source, the entry **MUST** be `fromParagraph: N`. The engine extracts font / size / weight / indent / spacing from the actual paragraph; you add only fields the source can't carry (`outlineLevel`, `basedOn`) or user-spec overrides via `overrides: { ... }`. The template's values are the contract to preserve, not a starting point to redesign.

```jsonc
{ "id": "Heading1", "name": "heading 1", "fromParagraph": 33, "outlineLevel": 0 }
{ "id": "Heading2", "name": "heading 2", "fromParagraph": 47, "outlineLevel": 1,
  "overrides": { "fontCJK": "黑体" } }  // user-spec layered on source typography
```

`fromParagraph` reads the *dominant text run* — longest non-numbering-prefix run; `"1.1 研究方法"` gives the title's typography, not the prefix's.

**Top-level `size` / `bold` / `alignment` / `spaceBefore` / `lineSpacing` on a represented-role entry silently overrides the template** — the dominant over-declaration failure. Keep top-level fields for empty-slot or user-spec'd styles only.

**`fontLatin` vs `fontCJK`.** When user names only a CJK font ("正文宋体" / "标题黑体"), put it in `overrides.fontCJK` and leave `fontLatin` unset so the source's Latin font is preserved. Set both only when the user explicitly says the same font applies to Latin.

**Indent unit preservation.** Source's `w:firstLineChars` / `w:hangingChars` (Word's "首行缩进 N 字符") extracts as `"Nchar"` and auto-scales with font size; fixed twips give `"Npt"`. Don't manually convert "char" values to pt — locks indent to one font size.

**`FigureImage` line height.** `exact` lineRule clips the inline drawing to that height. If `fromParagraph` extracts one, override to a multiplier or `{ atLeast: <length> }`.

### Empty-slot styles

For roles with **no source paragraph** playing them yet — e.g., `ListNumber` in a doc that's never had a numbered list — explicit top-level fields are required. Declare only user-named attributes + locale defaults the empty slot needs:

```jsonc
{ "id": "ListNumber", "name": "List Number", "fontCJK": "宋体",
  "size": 12, "firstLineIndent": "2char" }
```

Undeclared fields cascade from `basedOn` / `Normal` / theme — leave them blank.

### Canonical ListBullet / ListNumber configuration

For Word-natural list rendering (bullet / number hangs in the margin, text aligns at a consistent indent), use:

```jsonc
{
  "id": "ListBullet",
  "fontCJK": "宋体",
  "size": "12pt",
  "paraFormat": {
    "indentLeft": "0.74cm",
    "hangingIndent": "0.74cm"
  }
}
```

Equivalent for `ListNumber` (substitute `ListNumber` for the `id`). Avoid `firstLineIndent` on list styles — it produces "first-line indent THEN bullet" layout (bullet appears N chars in from text margin), which is not the standard Word list appearance. The combination `indentLeft + hangingIndent` of the same value is what makes the bullet hang at the left margin while text aligns at `indentLeft`.

This recipe pairs with `numbering: [...]` declarations that bind the styleId via `pStyle` in their level definitions.

### Override existing > create new

Run `inspect_style_def` first — POI / WPS / school templates often play Normal / Heading 1 / etc. under auto-generated styleIds (`a`, `a1`, `2`, `10`, ...). Override by exact styleId; the engine mutates in place, preserving everything unspecified. Verify the style is actually used for its intended role first — overriding `Heading1` while it's misused as body would corrupt those paragraphs; reassign first.

### Source role ↔ styleId mismatches

Two directions, both common in POI / WPS / school templates:

- **One styleId, many roles** (headings / body / captions all bound to `a`). Override can't help — changing `a` shifts them all. Install fresh semantic styles in `styles[]` + route the right paragraphs to each via `pattern_rules` / `assignments`. Fresh styles are empty-slot in this design sense even if target paragraphs exist.
- **Many styleIds, one role** (`a` / `a1` / `style29` all rendering as body). Either pick the most-used as canonical and route the others to it via `bulk_rules` keyed on fingerprint, or install a fresh `BodyText` and route all paragraphs to it. The source styleId topology doesn't fix itself.

### Style identity (`<w:name>` aliasing)

`name` is optional. For an override on an existing styleId, omitted `name` preserves the source `<w:name>`. For a new style, omitted `name` defaults to `id`.

`name` must not alias any existing style's identity. Word treats `<w:name>` as the built-in style identity marker, including locale aliases ("Normal" ≡ "正文" ≡ "標準"; "Heading 1" ≡ "标题 1"; "Body Text" ≡ "正文文本"). When two different styleIds claim the same identity, Word silently drops the second style's `rPr` at render time. Three safe approaches:

- **Override existing by its styleId** so no new name enters the doc.
- **When creating new with a built-in styleId** (`Heading1`, `BodyText`, `Caption`, ...), use the canonical English built-in name matching the styleId — `name: "Body Text"` for `id: "BodyText"`, `name: "heading 1"` for `id: "Heading1"`. Word applies its own UI localization for display.
- **When creating new with a custom styleId** (e.g. `BodyEmphasis`), the simplest safe rule is `name = id`.

The engine catches direct string-equal collisions plus the major en/zh-CN aliases at preflight.

### Font scope layers

Three layers, narrowest to broadest:

- **Per-role override** (`styles[]` entries with `fontCJK` / `fontLatin` / `overrides`): only styled paragraphs change. Use when the user names specific roles ("标题黑体, 正文宋体", "Heading2 字号改小一号").
- **Whole-doc default** (declare a `Normal` entry; other styles' `basedOn: "Normal"` chain inherits): wide, covers most pStyle-bound paragraphs via cascade. Use for uniform defaults ("整篇统一用 Times New Roman").
- **Document-design font scheme** (`theme.fonts` block): widest, modifies theme1.xml so any docDefaults / styles / runs that reference theme fonts auto-resolve. Use for document-design language ("把这份文档的主题字体改成 X / Y").

Bias toward the narrowest layer that captures the intent — wider layers risk surprising the user with effects on chrome they didn't intend to touch.

Chinese font size names (初号 / 一号 / .../ 小六): see [`chinese-font-sizes.md`](chinese-font-sizes.md).

### Edge case: source base style violates the stated spec

If doc's `Normal` sets `bold: true` while the spec says 正文不加粗, override `Normal` in `styles[]` with the corrected fields — the base style's deviation cascades to every dependent.

---

## 2. Counter mechanisms

Three mechanisms for three classes of structural numbering. They don't substitute for each other; match the source's typed-counter shape to the right one.

### Auto-numbering (headings + lists)

When the document has typed heading prefixes (chrome or content prefixes), migrate to automatic numbering. Skip only when the user explicitly opts out, or no numbered headings exist.

If the manual scheme itself is inconsistent — e.g. H1 has numbers in chapter 1 but not chapter 2, or H2 uses chapter-prefixed `"1.1"` in some chapters and per-chapter-restart `"1."` in others — auto-migration is a normalization decision that may change author-intended semantics. Ask the user before applying.

Each level binds to a heading style via `styleId`; higher levels reset lower-level counters automatically:

```jsonc
"numbering": [
  {
    "levels": [
      { "level": 0, "numFmt": "chineseCounting", "lvlText": "%1、",  "suff": "nothing", "styleId": "Heading1" },
      { "level": 1, "numFmt": "chineseCounting", "lvlText": "（%2）", "suff": "nothing", "styleId": "Heading2" },
      { "level": 2, "numFmt": "decimal",         "lvlText": "%3.",   "suff": "space",   "styleId": "Heading3" }
    ]
  },
  {
    "levels": [
      { "level": 0, "numFmt": "decimal", "lvlText": "%1.", "suff": "space", "styleId": "ListNumber" }
    ]
  }
]
```

**`outlineLevel` is independent of numbering level.** They co-occur for headings but are different OOXML concepts: `outlineLevel` (on the paragraph style's pPr) controls TOC inclusion, navigation pane, outline view. Numbering level (the `level` field above) controls auto-number display. Set `outlineLevel` explicitly on heading styles via `overrides`; don't expect numbering to imply it. List styles (`ListBullet` / `ListNumber`) take numbering but **must not** carry `outlineLevel` — otherwise list items pollute the TOC.

**`stripPrefixPatterns` defaults to `[lvlText]`** — when the doc's manual prefix matches the new auto-numbering pattern, you don't need to write it. Specify it only when the source mixed prefix styles within one role. Patterns try in order, first match wins; **longer pattern must come first** or shorter ones swallow prefixes the longer ones wanted.

**`%N` placeholder matches Arabic digits or common Chinese numerals (一二三...百千).** It does NOT match Roman numerals, Latin letters, less-common Chinese forms (壹貳叄), circled digits (①), or other locale-specific shapes. The dry-run report's "Numbered-style paragraphs not matched by any stripPrefixPattern" section surfaces unstripped paragraphs with leading-text samples — read those samples and add a `pattern_rules` entry with `stripMatch: true` and an explicit regex covering the unrecognized shape.

**Heading vs list — different schemes, both in the same `numbering: [...]` array:**

- **Heading-class** (outline-bearing prefixes defining document structure): `Heading1..N` bound to **one unified multi-level scheme**.
- **List-class** (local enumerations inside body sections): `ListNumber` / `ListBullet` bound to a **separate single-level scheme** (restart per group).

Detection signals:

- **Position.** Heading at section start, followed by short title text; list item inside a section among other list items or body prose.
- **Surrounding format.** Heading already styled distinctly (bold, larger size); list items sit in body-style.
- **Pattern shape.** Headings tend to decimal hierarchy or chapter sentinels; lists tend to parenthesized or single-level digits.
- **Depth.** Headings nest 2–4 levels; lists usually 1.

Multi-level templates by document class (academic / technical / governmental / legal): see [`numbering-formats.md`](numbering-formats.md).

**Edge case: unnumbered special headings** (摘要 / Abstract / 目录 / 参考文献 / 致谢 / 附录) share Heading1's visual style but have no chapter number. Use a separate `HeadingNoNum` style, or bind these paragraphs to a styleId not in the heading numbering scheme.

**Edge case: appendix numbering** often restarts with a different scheme (附录A / A.1 / A.2) — a second `numbering` entry, bound to a separate set of appendix heading styles.

### Caption SEQ fields (figures / tables / equations)

Typed figure / table / equation numbers (`图 2.1` / `表 3-1` / `公式 (2.4)`) do **not** route through `pattern_rules`. They are SEQ-field-backed counters, not styleId-bound auto-numbering — a separate mechanism.

Migration path:

1. `migrate_captions` detects manually-numbered caption-shaped paragraphs and proposes identifiers.
2. Declare a `captions: { Figure: {...}, Table: {...}, Equation: {...} }` table at the apply config root (per-identifier prefix / suffix / chapter prefix / styleId).
3. In `edits[]`, convert each detected paragraph via a `replace` op with a `{ type: "caption", captionId: "Figure", text: "..." }` block, or for equations a `{ type: "equation", latex: "...", captionId: "Equation" }` block.

Full schema: [`captions.md`](captions.md), [`equations.md`](equations.md).

### Cross-references (typed in-prose counters)

Typed counters embedded in body prose ("如图 3.2 所示", "见第 2 章", "参考文献[5]") are not paragraph-leading prefixes — they sit mid-sentence. They route through `edits[]` `InlineRef` nodes, not `pattern_rules`.

Detection: `find_paragraphs --regex` with patterns matching the cite shape (e.g. `图\s*\d+[-．.]?\d*`, `第.*章`, `参考文献\[\d+\]`). Conversion: in `edits[]`, use a `replace` op that rebuilds the paragraph's `text` array with `{ refTo, display, format }` nodes replacing the typed counter substrings.

Full schema: [`cross-references.md`](cross-references.md). Page-number cites remain out of scope — surface to user.

### When NOT to migrate counters

- **Inconsistent manual schemes** across the document — ask the user before normalizing rather than picking a scheme silently.
- **Skipped heading levels** (H1 → H3 with no H2) — can't be synthesized; audit-only signal, surface to user.
- User explicitly opts out ("保留手动编号，只调字体").

---

## 3. Paragraph → style routing

Three targeting mechanisms, narrowest to broadest.

**No fingerprint coverage requirement.** Untargeted paragraphs simply stay as they are. Routing declares what *should change*, not what exists.

### Text-shape based: `pattern_rules`

One rule per chrome shape. `stripMatch: true` removes the manual prefix during restyle so auto-numbering takes over.

```jsonc
"pattern_rules": [
  { "regex": "^[一二三四五六七八九十百千]+、",     "style": "Heading1", "stripMatch": true },
  { "regex": "^（[一二三四五六七八九十百千]+）",   "style": "Heading2", "stripMatch": true },
  { "regex": "^\\d+\\.\\d+\\s",                "style": "Heading3", "stripMatch": true }
]
```

Real cases needing adaptation:

- **Inconsistent chrome**: section 1 uses bare colon-labels (`Name:`), sections 2+ use enumerator chrome (`1.`, `2.`). Convert what's structural; leave inline labels alone.
- **Unstable manual prefixes**: source has `1.1` in some chapters, `1.` in others. Use `stripPrefixPatterns: ["%1.%2", "%1."]` on the relevant level — longer first.
- **Chrome that shares text shape with body content**: body paragraphs cite `第三章` as a reference. Use `exclude` or refine the regex (anchor at paragraph start + require trailing context).

Validate coverage before applying with `find_paragraphs --regex <pat>`.

### Visual-shape based: `bulk_rules`

When chrome doesn't share a text pattern but does share a fingerprint:

```jsonc
"bulk_rules": [
  { "fingerprint": "B", "style": "BodyText" },
  { "fingerprint": "a8083d", "style": "Heading1" }
]
```

Fingerprint accepts either the letter label (in-session iteration) or the 6-char hash (stable across edits; use in persisted configs). Same hash accepted by `inspect_style`.

### Outliers / disambiguation: `assignments`

Per-paragraph, for two cases: genuine outliers no pattern catches, OR **text-shape collisions where two roles share the same prefix** (e.g. `Heading1` headings and an instruction list both start with `一、`). `assignments` overrides `pattern_rules` for the listed paragraphs, so listing a continuous block (say `#30–#35`) to route to a *different* style than the regex would pick is legitimate — not an outlier count violation.

```jsonc
"assignments": [{ "para": 142, "style": "BodyText" }]
```

If you're listing > 5 paragraphs that *aren't* a coherent same-role group (one-offs scattered across the doc), the pattern is wrong — refine the regex or add a missing pattern instead.

### Anti-targets: `exclude`

Paragraphs a regex caught wrongly:

```jsonc
"exclude": [173, 174]
```

Targeting precedence and assignment-action semantics: see [`config-schema.md`](config-schema.md).

---

## 4. What standardize leaves alone

The negative space matters as much as the positive. Standardize **does not** touch:

- **Section properties** — untouched unless declared via top-level `pageSetup`. See [`config-schema.md`](config-schema.md#page-setup).
- **Run-level direct format on untargeted paragraphs.** The uniform-strip rule fires only on paragraphs the engine restyles; untouched paragraphs keep their direct rPr verbatim.
- **Unreferenced style definitions.** Engine doesn't prune them (sparse-by-design). If the style picker matters to the user, prune in Word manually.
- **Tracked changes / SDT controls / complex fields.** The edit phase refuses these; `pattern_rules` won't reach inside them either. Run `inspect_blockers` if unsure.
- **Manual TOC body / page numbers / page-number cites.** See SKILL.md "Out of scope".
- **Footnote / endnote bodies.** They live in `word/footnotes.xml` / `endnotes.xml`, outside the `word/document.xml` index, so `apply` neither reindexes nor restyles them. Insert new notes with `extend` (`op: "footnote"` / `"endnote"`) — see [`extend.md`](extend.md).
- **Skipped heading levels.** Can't synthesize a missing level — audit flags, surface to user.

The flip side: **a narrow-scope request reduces to declaring less, not switching tools.** "Heading2 字号改小一号" → one `styles[]` entry with the size override; no `numbering`, no `pattern_rules`, no `edits`. "保留手动编号，只调字体" → `styles[]` with font overrides only; omit `numbering` and `stripMatch: true`. The engine's sparseness handles the rest.

---

## 5. Verifying via dry-run

`apply --dry-run` produces a change report. Four signals to read; each maps to the design decision that needs adjusting if the signal looks wrong.

- **Style Resolution.** Each installed style's final typography matches expectation. Mismatch usually means a top-level over-declaration on a represented role — revisit §1 `fromParagraph` and move declared fields into `overrides`.
- **Sample Affected Paragraphs.** `pattern_rules` hit the right targets. Wrong matches → refine regex or use `exclude`; missing matches → broaden regex or add a missing pattern. Pre-validate with `find_paragraphs --regex` to avoid surprises here.
- **Implicit-keep is a FAILURE signal.** Non-empty fingerprints not routed by any rule. Add `bulk_rules` / `assignments` until implicit-keep is empty, OR until the remaining ones are intentional non-content the user wants preserved (form labels, etc.).
- **"Numbered-style paragraphs not matched by any stripPrefixPattern."** Auto-numbering installed but the old typed prefix didn't strip. Read the leading-text samples and add an explicit `pattern_rules` entry with `stripMatch: true` and a regex covering the unrecognized shape.

### Reading `overview` skeleton conventions

- **Letter vs hash labels.** `[A]`, `[B]` sort by frequency this run (volatile across edits); the summary also shows a 6-char content hash next to each letter (`A [c4f9]: ...`). `bulk_rules.fingerprint` and `inspect_style` both accept either — letters for in-session iteration, hashes in configs that survive doc revisions.
- **Numbered ≠ unnumbered fingerprints.** Hash includes whether a paragraph carries a numbering reference; visually identical paragraphs split when one is auto-numbered and the other isn't.
- **Layout vs data tables.** Layout tables inline into the skeleton between `--- LAYOUT TABLE ---` markers; data tables summarize as one non-paragraph block.
- **Empty paragraph compression.** Consecutive empties compress (`--- empty ×N ---`); skeleton text truncates to ~40 chars — use `inspect_range` for full text.

---

## Cross-cutting edge cases

- **Empty paragraphs as spacing**: preserve them. Removing is structural, not stylistic, and risks breaking cover-page layout.
- **Table footnotes**: text right after a table starting with "注：" / "Note:" is a footnote, not body text — bind to `Reference` / `Footnote` rather than `BodyText`.
- **Pre-printed chrome** (forms, templates with printed labels and instructions): when the visual summary shows a long tail of low-occurrence fingerprints with short average text length (`avg ≤20ch`), those are usually printed labels / cover chrome, not author content — leave them untargeted.

---

## Putting it together

A complete standardize-shape config combines the four blocks. This is a reference shape; design the values from your survey, don't fill in a template.

```jsonc
{
  "source": "input.docx",
  "output": "output.docx",
  "styles": [
    { "id": "Heading1",   "name": "heading 1",  "fromParagraph": 33, "outlineLevel": 0 },
    { "id": "Heading2",   "name": "heading 2",  "fromParagraph": 47, "outlineLevel": 1 },
    { "id": "BodyText",   "name": "Body Text",  "fromParagraph": 12,
      "overrides": { "firstLineIndent": "2char" } },
    { "id": "ListNumber", "name": "List Number","fontCJK": "宋体","size": 12,"firstLineIndent": "2char" }
  ],
  "captions": {
    "Figure": { "prefix": "图", "bodySeparator": " ", "styleId": "FigureCaption" }
  },
  "numbering": [
    { "levels": [
      { "level": 0, "numFmt": "chineseCounting", "lvlText": "%1、", "suff": "nothing", "styleId": "Heading1" },
      { "level": 1, "numFmt": "chineseCounting", "lvlText": "（%2）","suff": "nothing", "styleId": "Heading2" }
    ]},
    { "levels": [
      { "level": 0, "numFmt": "decimal", "lvlText": "%1.", "suff": "space", "styleId": "ListNumber" }
    ]}
  ],
  "pattern_rules": [
    { "regex": "^[一二三四五六七八九十百千]+、",   "style": "Heading1", "stripMatch": true },
    { "regex": "^（[一二三四五六七八九十百千]+）", "style": "Heading2", "stripMatch": true }
  ],
  "edits": [
    /* CaptionBlock / InlineRef migration ops — see edit.md */
  ]
}
```

`exclude` and `assignments` are omitted above; declare them only when a rule catches wrongly or an outlier needs a per-paragraph correction.

---

## Compose with other shapes

- **Insertion + surgical edits**: add an `edits[]` block to the same `apply` config for content insertion or paragraph touch-ups the rules missed. See [`edit.md`](edit.md). The engine installs styles/numbering first, so `edits[]` references just-installed styleIds.
- **Read-only check before reshape**: `audit` (see [`audit.md`](audit.md)). Audit's violation list translates directly into standardize-shape blocks — style violations → `styles[]`, typed structural prefixes → `pattern_rules` + `numbering`, typed captions → `captions` + `edits[]` `CaptionBlock`.

---

## Validation behavior

`apply` validates the output docx after writing. Only errors **introduced by this apply run** are fatal; pre-existing errors in the source are warnings and do not block output. Override via [`allowValidationWarnings`](config-schema.md) when debugging.
