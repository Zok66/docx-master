import { a as firstChildNS, c as paragraphRuns, h as NS, l as paragraphStyleId, m as walkBodyParagraphs, o as getChildren } from "./_shared/xml-utils.js";
import { t as loadDocx } from "./_shared/load.js";
import { t as parseFieldRuns } from "./_shared/field-parse.js";

//#region skill/tools/migrate-captions.ts
/**
* migrate-captions: detect manually-numbered caption-shaped paragraphs.
*
* Usage:
*   migrate_captions <docx-path> [--style <styleId> ...]
*
* Scans body for paragraphs whose pStyle matches one of the provided
* styleIds (or any style if --style not given) AND whose visible text
* starts with a caption-shaped prefix:
*   - "图 N", "图 N.M"
*   - "表 N", "表 N.M"
*   - "Figure N", "Figure N.M"
*   - "Table N", "Table N.M"
*   - "(N)", "(N.M)"
*   - "[N]", "[N.M]"
*
* Reports each candidate for the agent to review. Output is read-only —
* the agent then builds an apply config with the appropriate `captions`
* table + edit ops to delete / re-emit each candidate via CaptionBlock.
*
* This tool is intentionally read-only in v1. Automatic migration
* requires the agent to declare the captions table first (identifier
* names, chapter prefix style, etc.) — without that context, the tool
* can't know which identifier each pattern maps to. The agent uses this
* detection output to write the apply config.
*/
const w = NS.w;
const PATTERNS = [
	{
		regex: /^图\s*\d+(?:[.-]\d+)?/,
		identifier: "Figure"
	},
	{
		regex: /^表\s*\d+(?:[.-]\d+)?/,
		identifier: "Table"
	},
	{
		regex: /^Figure\s+\d+(?:[.-]\d+)?/i,
		identifier: "Figure"
	},
	{
		regex: /^Table\s+\d+(?:[.-]\d+)?/i,
		identifier: "Table"
	},
	{
		regex: /^Equation\s+\d+(?:[.-]\d+)?/i,
		identifier: "Equation"
	},
	{
		regex: /^\(\s*\d+(?:[.-]\d+)?\s*\)/,
		identifier: "Equation"
	},
	{
		regex: /^\[\s*\d+(?:[.-]\d+)?\s*\]/,
		identifier: "Equation"
	}
];
async function main() {
	const args = process.argv.slice(2);
	const file = args[0];
	if (!file) {
		console.error("Usage: migrate_captions <docx-path> [--style <styleId> ...]");
		process.exit(1);
	}
	const filterStyles = /* @__PURE__ */ new Set();
	for (let i = 1; i < args.length; i++) if (args[i] === "--style" && args[i + 1]) {
		filterStyles.add(args[i + 1]);
		i++;
	}
	const documentDoc = (await loadDocx(file)).documentDoc;
	if (!documentDoc) {
		console.error("migrate_captions: failed to read word/document.xml");
		process.exit(1);
	}
	const body = firstChildNS(documentDoc.documentElement, w, "body");
	if (!body) {
		console.error("migrate_captions: document has no body");
		process.exit(1);
	}
	const candidates = [];
	let paragraphIndex = 0;
	for (const para of walkBodyParagraphs(body)) {
		paragraphIndex++;
		const styleId = paragraphStyleId(para);
		if (filterStyles.size > 0 && (styleId === void 0 || !filterStyles.has(styleId))) continue;
		if (paragraphContainsSeq(para)) continue;
		const text = paragraphVisibleText(para);
		for (const { regex, identifier } of PATTERNS) {
			const m = text.match(regex);
			if (m) {
				candidates.push({
					paragraphIndex,
					styleId,
					matchedPrefix: m[0],
					text,
					suggestedIdentifier: identifier
				});
				break;
			}
		}
	}
	if (candidates.length === 0) {
		console.log("No manually-numbered caption-shaped paragraphs detected.");
		return;
	}
	console.log(`Manual caption candidates detected (${candidates.length}). Each is a paragraph whose text leads with a caption-shape and no SEQ field — the agent can convert via apply config (CaptionBlock + replace op).`);
	console.log("");
	for (const c of candidates) {
		const styleId = c.styleId ?? "(unstyled)";
		const preview = c.text.length > 60 ? c.text.slice(0, 57) + "..." : c.text;
		console.log(`  para ${String(c.paragraphIndex).padStart(4)}  style: ${styleId.padEnd(20)}  prefix: ${c.matchedPrefix.padEnd(12)}  suggestedIdentifier: ${c.suggestedIdentifier}`);
		console.log(`        text: ${preview}`);
	}
	console.log("");
	console.log("Next steps for the agent:");
	console.log("  1. Declare a captions[] table in your apply config matching the suggested");
	console.log("     identifiers (Figure / Table / Equation / ...).");
	console.log("  2. Add edit ops that `delete` each manual paragraph and `insert-after`");
	console.log("     a CaptionBlock { captionId, text, anchor? }. Strip the matched prefix");
	console.log("     from the text — the captions table re-renders prefix + counter.");
}
function paragraphContainsSeq(paragraph) {
	return parseFieldRuns(paragraphRuns(paragraph)).some((e) => e.kind === "field" && e.fieldType === "SEQ");
}
function paragraphVisibleText(paragraph) {
	let out = "";
	for (const r of getChildren(paragraph)) {
		if (r.namespaceURI !== w || r.localName !== "r") continue;
		for (const t of getChildren(r)) if (t.namespaceURI === w && t.localName === "t") out += t.textContent ?? "";
	}
	return out;
}
await main();

//#endregion
export {  };