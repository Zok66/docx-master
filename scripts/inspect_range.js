import { t as loadDocx } from "./_shared/load.js";
import { n as formatComputedRPrParts, r as pad, t as formatComputedPPrParts } from "./_shared/format.js";

//#region skill/tools/inspect-range.ts
async function main() {
	const file = process.argv[2];
	const fromArg = process.argv[3];
	const toArg = process.argv[4];
	if (!file || !fromArg || !toArg) {
		console.error("Usage: node scripts/inspect_range.js <docx-path> <from> <to>");
		process.exit(1);
	}
	const from = parseInt(fromArg, 10);
	const to = parseInt(toArg, 10);
	if (isNaN(from) || isNaN(to) || from < 1 || to < from) {
		console.error("Invalid range");
		process.exit(1);
	}
	try {
		const doc = await loadDocx(file);
		const out = [];
		for (const p of doc.paragraphs) {
			if (p.index < from || p.index > to) continue;
			out.push(...renderPara(p));
			out.push("");
		}
		if (out.length === 0) out.push(`No paragraphs in range ${from}-${to}`);
		console.log(out.join("\n"));
	} catch (err) {
		console.error(`Error: ${err.message}`);
		process.exit(1);
	}
}
function renderPara(p) {
	const lines = [];
	lines.push(`#${pad(p.index)} [${p.fingerprint}]`);
	lines.push(`  text: ${JSON.stringify(p.text)}`);
	lines.push(`  style: "${p.styleId}" (${p.styleName})  insideTable: ${p.context.insideTable ?? "(none)"}`);
	lines.push(`  computed rPr: ${formatRPr(p.rPr)}`);
	lines.push(`  computed pPr: { ${formatComputedPPrParts(p.pPr, {
		includeIndentSides: true,
		includeNumLevel: true
	}).join(", ")} }`);
	lines.push(`  section: ${p.context.sectionIndex}`);
	return lines;
}
function formatRPr(r, opts) {
	const parts = formatComputedRPrParts(r, opts);
	if (r.underline) parts.push(`underline: ${r.underline}`);
	if (r.highlight) parts.push(`highlight: ${r.highlight}`);
	if (r.strike) parts.push(`strike: ${r.strike}`);
	if (r.caps) parts.push(`caps: ${r.caps}`);
	return `{ ${parts.join(", ")} }`;
}
main();

//#endregion
export {  };