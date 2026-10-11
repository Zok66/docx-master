# `audit`

Read-only conformance check. Output is a violation list with `#NNN` references, never a new docx. Fixes happen in a separate `apply` call after the user reviews the report.

## Spec

Audit always runs against a **baseline**: Word canonical (SKILL.md "Target state") — mechanisms present in the doc must be canonical (one styleId per role, auto-numbering not typed, SEQ-backed captions, REF cross-refs, no orphan defs); absent mechanisms generate no signal. User-provided spec **layers on top** (adds / overrides specific items); never replaces the baseline. Sources of additional spec:

1. **The user's prompt** — rules typed in the request. Override / extend baseline.
2. **A reference docx** the user names as "the format we want to match" — reverse-engineer its styles + numbering and treat as additional spec items.
3. **An external standard** the user names (school template, journal guideline) — load the user-provided copy; don't fabricate from training memory.

Without an explicit user spec, baseline alone is the spec — the report is still an audit, not "describe the document".

## Workflow

1. Decompose the spec into checkable items.
2. Scan along the axes below; gather facts via the read-only tools listed.
3. Produce a structured violation list per item: spec requirement / document actual / affected paragraphs (letter labels + `#NNN` so the user can navigate).
4. Do not auto-fix. Translation to an `apply` config happens after the user reviews and confirms scope.

## Scanning axes

Three axes apply regardless of spec — what to *look at*, not what counts as a violation. Whether a signal *is* a violation is the spec's call.

### Axis 1 — does the style system express semantic roles?

Signals worth surfacing:
- Many styleIds playing the same visible role (`a` / `a1` / `style29` all rendering as body)
- One styleId playing many roles (every paragraph bound to `Normal` or `a`, role distinctions living entirely in direct `pPr`/`rPr`)
- Direct `pPr`/`rPr` on paragraphs that already carry a `pStyle` — distinguish content chrome's preserved typography (legitimate) from author drive-by overrides
- `<w:name>` collisions, including en/zh-CN aliases (`Normal` ≡ `正文`, `Heading 1` ≡ `标题 1`)
- Defined styles with zero usage

Tools: `overview` for the visual-style summary + direct-format-per-fingerprint, `inspect_style <fingerprint>` for which styleIds back a visible role, `inspect_style_def <styleId>` for the cascade and `basedOn` chain.

### Axis 2 — is structure carried by mechanisms or by typed text?

Anything that mimics a mechanism's output via literal characters is a candidate: numbering markers (`一、` / `1.1` / `第N章`), cross-reference counters in prose (`如图 3.2 所示`), caption numbers (`图 2-1`), TOC entries, page numbers in footers, footnote markers (`[1]`), multi-blank-paragraph spacing, tab-aligned column layout, underscore strings for fill-in blanks.

**Report every hit; group by context — never filter by intent.** Whether a typed prefix should be retagged is a *fix-time* decision belonging to the user / next stage, not an audit-time filter. Context dimensions worth grouping by so the consumer can decide:

- **Form-architecture chrome** — chrome paragraph followed by empty placeholder slots (a label heading a fill region). Often intentional stable label; standardize typically excludes via `exclude` or skips.
- **Instructional chrome** — typed-numbered paragraph whose own body IS the content (numbered instructions, notes). Form designer's content; usually preserved.
- **Author content with typed prefix** — chrome paragraph followed by its own body paragraphs (author typed `第三章 引言\n\n本章介绍...`). Default retag candidate per [standardize.md §3 chrome/content prefixes](standardize.md).

Tools: `overview` skeleton reveals typed prefix shapes; `find_paragraphs --regex` scans for typed counters (figure / table / chapter cites) — coverage view; `find_text` pinpoints the exact run / offset inside a hit; `inspect_neighbors` distinguishes the three context groups above; `inspect_caption` lists SEQ-backed identifiers (anything caption-shaped not listed there is typed); `migrate_captions` finds manually-numbered caption paragraphs explicitly.

### Axis 3 — is heading hierarchy and numbering correctly wired?

Signals:
- `outlineLevel` missing on heading styles (TOC / navigation pane breaks) or present on list styles (list items pollute TOC)
- Multiple parallel multi-level schemes for what should be one document outline; per-chapter independent counters
- Numbering bound paragraph-by-paragraph (`<w:numPr>` on each para) instead of style-bound via `numbering[].levels[].styleId`
- Paragraphs visually shaped like headings (bold + larger) but bound to a non-Heading styleId
- Skipped levels (H1 → H3 with no H2)

Tools: `overview` for heading distribution + numbering scheme cluster; `inspect_style_def` for per-style `outlineLevel` and `numId` wiring; `inspect_section` when the anomaly looks section-scoped.

## Spec-driven items

Beyond the three axes, the spec typically names specifics that tools can surface but can't judge alone. Audit these by reading the spec value against what `overview` / `inspect_range` / `inspect_section` show. Without a spec, none of these are violations on their own:

- Font (Latin / CJK separately), size per role
- Line spacing, space-before / space-after per role
- Page setup: margins, paper, orientation, header / footer distance (engine doesn't change `<w:sectPr>`, but audit reports)
- CJK-locale typography: first-line indent (`Nchar` vs `Npt`), quote style (`""` / `''` vs `「」`), CJK ↔ Latin / digit boundary spacing, `autoSpace`
- Caption position relative to figure / table (below vs above)
- Reference / citation format
- Chapter start-on-odd-page or new-section rules

## From violation list to `apply` config

After the user confirms scope, report items usually map cleanly:

- Heading typography inconsistent → `styles[Heading*]` with `fromParagraph` + `overrides`
- Typed structural prefixes → `pattern_rules` with `stripMatch: true` + a `numbering` scheme bound to heading styles
- Typed caption numbers → `captions` table + `edits[]` converting paragraphs to `CaptionBlock` (use `migrate_captions` output as the source list)
- One styleId overloaded with multiple roles → install fresh semantic styles + route via `pattern_rules` / `assignments` (the chaotic-source case in [standardize.md](standardize.md))
- Direct format on Heading-bound paragraphs → distinguish chrome typography (preserve) from drive-by overrides (only the latter goes into `overrides`)

## Notes

- Direct `pPr` / `rPr` is not a violation by virtue of being present — for content chrome it's the typographic source standardize preserves. It IS a violation when chrome's values conflict with the spec, or when same-role chrome is inconsistent across instances; the fix path is `overrides` on the installed style, not `fromParagraph` extraction (which would pull the violating values in as the new style's typography). See SKILL.md "How to think about formatting".
- Issues the engine's preflight catches (style.name collisions, etc.) still belong in the audit report; the report is for the user, not the engine.

## Compose with other shapes

- User confirms a fix scope → role-based reshape via [standardize.md](standardize.md), surgical changes via [edit.md](edit.md), both in one `apply` config.
- Schema-validate a file independent of any user spec → `validate`.
