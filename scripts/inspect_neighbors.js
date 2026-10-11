import { t as loadDocx } from "./_shared/load.js";

//#region skill/tools/inspect-neighbors.ts
/**
* inspect_neighbors <docx> <paraIndex> [--radius N]
*
* Returns up to `radius` adjacent elements on each side of the target
* paragraph in the document body's linear flow. Each entry tags its type
* (paragraph / image / table / equation / pageBreak / sectionBreak) and
* carries the type-specific structured payload (image dims, table shape,
* paragraph fingerprint + textPreview).
*
* Use cases:
*   - figure caption classification: image at distance 1 before
*   - table caption classification: table at distance 1 after
*   - first-paragraph-after-heading detection: heading-fingerprint paragraph
*     at distance 1 before
*   - layout sanity: are two images stacked with only an empty paragraph
*     between, suggesting one combined caption block?
*
* Default radius is 4 — enough to span "image, empty para, caption, empty
* para, next thing" without too much noise.
*/
async function main() {
	const argv = process.argv.slice(2);
	let radius = 4;
	const positional = [];
	for (let i = 0; i < argv.length; i++) {
		const a = argv[i];
		if (a === "--radius") {
			const next = argv[++i];
			if (!next || isNaN(parseInt(next, 10))) {
				console.error("--radius requires a positive integer");
				process.exit(1);
			}
			radius = parseInt(next, 10);
			if (radius < 1) {
				console.error("--radius must be ≥ 1");
				process.exit(1);
			}
		} else if (a.startsWith("--")) {
			console.error(`unknown flag: ${a}`);
			process.exit(1);
		} else positional.push(a);
	}
	const file = positional[0];
	const idxArg = positional[1];
	if (!file || !idxArg) {
		console.error("Usage: node scripts/inspect_neighbors.js <docx-path> <paragraph-index> [--radius N]");
		process.exit(1);
	}
	const targetIdx = parseInt(idxArg, 10);
	if (isNaN(targetIdx) || targetIdx < 1) {
		console.error("Invalid paragraph index (must be 1-based positive integer)");
		process.exit(1);
	}
	try {
		const doc = await loadDocx(file);
		const paraInList = doc.neighborItems.findIndex((it) => it.kind === "paragraph" && it.paraIndex === targetIdx);
		if (paraInList < 0) {
			const paragraphs = doc.paragraphs;
			const max = paragraphs.length;
			const range = max > 0 ? `#${paragraphs[0].index}–#${paragraphs[max - 1].index}` : "(none)";
			const closest = paragraphs.reduce((best, p) => Math.abs(p.index - targetIdx) < Math.abs(best.index - targetIdx) ? p : best, paragraphs[0]);
			console.error(`Paragraph #${targetIdx} not found. Document has ${max} indexed paragraphs (${range}). Closest: #${closest.index}.`);
			console.error("Note: paragraphs inside data tables are not indexed and cannot be referenced.");
			process.exit(1);
		}
		console.log(renderReport(doc, paraInList, radius));
	} catch (err) {
		console.error(`Error: ${err.message}`);
		process.exit(1);
	}
}
function renderReport(doc, anchorIdx, radius) {
	const items = doc.neighborItems;
	const paraByIdx = new Map(doc.paragraphs.map((p) => [p.index, p]));
	const anchor = items[anchorIdx];
	const anchorPara = anchor.kind === "paragraph" ? paraByIdx.get(anchor.paraIndex) : null;
	const lines = [];
	lines.push(`#${anchorPara?.index ?? "?"} — neighbors within ${radius} hops${anchor.kind === "paragraph" && anchor.isEmpty ? " (anchor is empty paragraph)" : ""}`);
	if (anchorPara) {
		const txt = anchorPara.text.length > 80 ? anchorPara.text.slice(0, 77) + "…" : anchorPara.text;
		lines.push(`  anchor: [${anchorPara.fingerprint}] ${JSON.stringify(txt)}`);
	}
	lines.push("");
	const beforeEntries = [];
	for (let j = anchorIdx - 1, d = 1; j >= 0 && d <= radius; j--, d++) beforeEntries.push({
		distance: d,
		item: items[j]
	});
	const afterEntries = [];
	for (let j = anchorIdx + 1, d = 1; j < items.length && d <= radius; j++, d++) afterEntries.push({
		distance: d,
		item: items[j]
	});
	lines.push("before:");
	if (beforeEntries.length === 0) lines.push("  (start of document)");
	else for (const e of beforeEntries) lines.push(`  [${e.distance}] ${formatItem(e.item, paraByIdx)}`);
	lines.push("");
	lines.push("after:");
	if (afterEntries.length === 0) lines.push("  (end of document)");
	else for (const e of afterEntries) lines.push(`  [${e.distance}] ${formatItem(e.item, paraByIdx)}`);
	return lines.join("\n");
}
function formatItem(item, paraByIdx) {
	switch (item.kind) {
		case "paragraph": {
			const p = paraByIdx.get(item.paraIndex);
			const fp = p?.fingerprint ?? "?";
			const tag = item.isEmpty ? " (empty)" : "";
			const txt = p ? p.text.length > 60 ? `"${p.text.slice(0, 57)}…"` : `"${p.text}"` : `"<paragraph #${item.paraIndex}>"`;
			return `paragraph #${item.paraIndex} [${fp}]${tag} ${txt}`;
		}
		case "image": return `IMAGE ${item.widthCm.toFixed(1)}×${item.heightCm.toFixed(1)}cm`;
		case "table": return `TABLE ${item.rows}×${item.cols} ${item.classification}`;
		case "equation": return `EQUATION`;
		case "pageBreak": return `pageBreak`;
		case "sectionBreak": return `sectionBreak`;
	}
}
main();

//#endregion
export {  };