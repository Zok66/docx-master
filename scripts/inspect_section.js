import { t as loadDocx } from "./_shared/load.js";
import { i as paperName, o as tw2mm } from "./_shared/format.js";

//#region skill/tools/inspect-section.ts
async function main() {
	const file = process.argv[2];
	const idxArg = process.argv[3];
	if (!file || idxArg === void 0) {
		console.error("Usage: node scripts/inspect_section.js <docx-path> <section-index>  (1-based; matches `Section N` in overview)");
		process.exit(1);
	}
	const idx = parseInt(idxArg, 10);
	try {
		const doc = await loadDocx(file);
		if (idx < 1 || idx > doc.sections.length) {
			console.error(`Section ${idx} not found. Document has ${doc.sections.length} section(s); valid 1..${doc.sections.length}.`);
			process.exit(1);
		}
		const sec = doc.sections[idx - 1];
		const out = [];
		out.push(`Section ${idx} (paragraphs #${sec.paraRange[0]}-#${sec.paraRange[1]})`);
		const paper = paperName(sec.pageSize.width, sec.pageSize.height);
		out.push(`  Paper:       ${paper} (${tw2mm(sec.pageSize.width)} × ${tw2mm(sec.pageSize.height)} mm)`);
		out.push(`  Orientation: ${sec.orientation}`);
		out.push(`  Margins:     top=${tw2mm(sec.margins.top)} bottom=${tw2mm(sec.margins.bottom)} left=${tw2mm(sec.margins.left)} right=${tw2mm(sec.margins.right)} mm`);
		out.push(`  Header: ${sec.header ?? "(none)"}`);
		if (sec.headerHasImage) out.push(`    contains image`);
		out.push(`  Footer: ${sec.footer ?? "(none)"}`);
		if (sec.footerPageNumFormat) out.push(`  Footer page format: ${sec.footerPageNumFormat}`);
		if (idx > 1) {
			const prev = doc.sections[idx - 2];
			out.push("");
			out.push(`  Differs from Section ${idx - 1}:`);
			const diffs = diffSections(prev, sec);
			if (diffs.length === 0) out.push(`    (no differences)`);
			for (const d of diffs) out.push(`    ${d}`);
		}
		console.log(out.join("\n"));
	} catch (err) {
		console.error(`Error: ${err.message}`);
		process.exit(1);
	}
}
function diffSections(a, b) {
	const out = [];
	if (a.pageSize.width !== b.pageSize.width || a.pageSize.height !== b.pageSize.height) out.push(`paper changed: ${tw2mm(a.pageSize.width)}×${tw2mm(a.pageSize.height)}mm → ${tw2mm(b.pageSize.width)}×${tw2mm(b.pageSize.height)}mm`);
	if (a.orientation !== b.orientation) out.push(`orientation: ${a.orientation} → ${b.orientation}`);
	for (const k of [
		"top",
		"bottom",
		"left",
		"right"
	]) if (a.margins[k] !== b.margins[k]) out.push(`margin.${k}: ${tw2mm(a.margins[k])}mm → ${tw2mm(b.margins[k])}mm`);
	if ((a.header ?? "") !== (b.header ?? "")) out.push(a.header === null ? `+ Header added: "${b.header}"` : b.header === null ? `- Header removed` : `Header changed`);
	if ((a.footer ?? "") !== (b.footer ?? "")) out.push(a.footer === null ? `+ Footer added` : b.footer === null ? `- Footer removed` : `Footer changed`);
	if ((a.footerPageNumFormat ?? "") !== (b.footerPageNumFormat ?? "")) out.push(`Footer page format: ${a.footerPageNumFormat ?? "(none)"} → ${b.footerPageNumFormat ?? "(none)"}`);
	return out;
}
main();

//#endregion
export {  };