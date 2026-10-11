# Header / Footer

Generates header and footer parts and binds them to every section. Sparse-by-design: omit the `headerFooter` block entirely to leave existing HF references untouched. When the block IS declared, every sectPr is rebound and any pre-existing HF parts become orphans in the archive (Word ignores them); new partNames allocate past the highest existing index — numbering does not compact.

## Shape

```jsonc
headerFooter: {
  header: {
    default: [ ...Block[] ],   // applies to pages not covered by first / even
    first:   [ ...Block[] ],   // optional. First page of each section.
                               //   Auto-sets <w:titlePg/> on every sectPr.
    even:    [ ...Block[] ],   // optional. Even-numbered pages.
                               //   Auto-sets <w:evenAndOddHeaders/> in settings.
    underline: true | BorderEdge,   // optional. Line below the header.
  },
  footer: {
    default, first?, even?,
    overline: true | BorderEdge,    // optional. Line above the footer.
  },
  sections: {                  // optional. Per-section override; see below.
    "1":   { header?, footer? },
    "3-5": { header?, footer? }
  }
}
```

At least one of `header` / `footer` (or a non-empty `sections`) is required; each declared surface needs at least one variant. Empty array `[]` is legal — means "this variant exists for the trigger flag but renders nothing" (used to blank a cover page's header while keeping `titlePg` set).

To blank only the first page, declare both `first: []` AND `default: [...]` — omitting `default` leaves non-first, non-even pages with no header/footer at all.

## Block subset

Inside header / footer content, only these block types are allowed (rejected at config parse, including any nesting depth — table cells too):

- `paragraph` — text, optionally styleId-bound. Inline nodes — `text`, `fields` (`page` / `numPages` / `date`), `styleRef`, `hyperlinks` — work the same as in body edits; shapes in [`edit.md`](edit.md) / [`cross-references.md`](cross-references.md).
- `image`
- `table` — common pattern for split layouts via a single-row borderless 3-column table.
- `horizontal-rule`

Disallowed: `caption`, `caption-counter-reset`, `equation`, `page-break`. The first three bind to body-side counters; emitting them inside HF double-increments and breaks the counter sim. Page breaks inside HF are no-ops (Word ignores).

## Paragraph styleId restrictions

HF paragraphs cannot use heading or body-text styleIds (`Heading1..9`, `Title`, `Subtitle`, `BodyText`) — also rejected at any nesting depth. These carry outline-level / numbering bindings that misbehave outside the body.

Use the built-in `Header` / `Footer` styleId (engine auto-injects them when missing) with a `paraFormat` / `runFormat` override for typography. Custom non-heading styleIds declared in top-level `styles[]` also work.

## Separator line (`underline` / `overline`)

`header.underline` draws a line below the header (attached to the last `<w:p>` of each variant); `footer.overline` draws a line above the footer (first `<w:p>`).

- `true` → thin black single line (0.5pt).
- `BorderEdge` → same shape as table cell borders, see [`tables.md`](tables.md#borders).

Skipped when the endpoint position has no `<w:p>` to attach to. Image endpoints attach normally (image renders inside `<w:p>`). If a variant ends with a `horizontal-rule`, the separator overwrites that rule's edge on the same side. When **every** variant of a surface skips, the dry-run report emits a `warning:` line so a declaration that fell through doesn't go unnoticed.

## Variant semantics

ECMA-376 sectPr supports three reference types per surface:

- **default** — every page not covered by `first` or `even`.
- **first** — section's first page only. `<w:titlePg/>` is set on each sectPr whose effective config declares this variant.
- **even** — even-numbered pages. `<w:evenAndOddHeaders/>` in settings.xml flips on whenever any section declares this variant (global flag).

HF config is the source of truth for both flags — re-applying with a smaller variant set cleans them up.

## Per-section overrides (`sections`)

Top-level `header` / `footer` is the default for every section. `sections` keys (`"N"` 1-based, or `"N-M"` inclusive — same shape as `pageSetup.sections`) override the surface **wholesale**: a section declaring `header` replaces the top-level header entirely (variants + `underline`) while inheriting the top-level `footer` if not redeclared.

Sections with identical effective config share parts — declaring the same header across 50 sections still emits one `headerN.xml`. Combine with `pageSetup.sections.<N>.pgNumType` for section-aware page numbering (roman on TOC, arabic restart on body) — see [`config-schema.md`](config-schema.md#page-setup).

## Combinations

- **HF + `edits[]`** — both share the body asset registry; image / hyperlink rIds allocate linearly across body and HF emissions. Registering the same `png` in both adds exactly one `<Default Extension="png">` entry to `[Content_Types].xml`.
- **HF + `trackChanges: true`** — track-changes wraps body edits only. HF mutation runs as a separate pass producing pristine `<w:hdr>` / `<w:ftr>` content (no `<w:ins>` / `<w:del>` wrappers). Use this combo when iterating on body content under change tracking but wanting header/footer rewritten cleanly.

## Not supported

- `ParagraphBlock.anchor` inside HF — bookmark semantics on a part rendered across multiple pages are ambiguous; the engine throws at emit. For page-number cross-references, use the `field: "page"` inline node — shapes in [`edit.md`](edit.md).
- HF cleanup (orphan strategy leaves old parts in archive; future phase may GC).
