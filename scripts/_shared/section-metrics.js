//#region lib/parse/section-metrics.ts
/** Content-area width in twips: `pgSz.w − pgMar.left − pgMar.right`.
*
* The result is the value LaTeX would call `\textwidth` — the horizontal
* box block-level content can occupy in this section before colliding
* with the page margin. Returns 0 when `pageSize.width` is unavailable
* (rare hand-built fixtures without `<w:pgSz>`).
*/
function sectionUsableWidthTwips(section) {
	return Math.max(0, section.pageSize.width - section.margins.left - section.margins.right);
}
/** Locate the section whose `paraRange` contains the given 1-based
* paragraph index. Returns `null` when no section matches (index out of
* any range, e.g. against a fresh paragraph index past the original doc
* length) or when `sections` is empty. Callers handling "insert at end"
* should pass the last existing paragraph index, or fall back to
* `sections[sections.length - 1]` directly when paragraphs is empty.
*/
function sectionForParagraph(sections, paragraphIndex) {
	for (const s of sections) if (paragraphIndex >= s.paraRange[0] && paragraphIndex <= s.paraRange[1]) return s;
	return null;
}

//#endregion
export { sectionUsableWidthTwips as n, sectionForParagraph as t };