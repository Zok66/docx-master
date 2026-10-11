import { a as firstChildNS, d as textContent, h as NS, o as getChildren, p as wVal, s as getChildrenNS } from "./_shared/xml-utils.js";
import { t as summarizeTable } from "./_shared/table-classifier.js";
import { n as computeRawFingerprint } from "./_shared/fingerprint.js";
import { t as loadDocx } from "./_shared/load.js";
import { a as truncate, n as formatComputedRPrParts, r as pad, t as formatComputedPPrParts } from "./_shared/format.js";

//#region skill/tools/inspect-style.ts
async function main() {
	const file = process.argv[2];
	const handle = process.argv[3];
	if (!file || !handle) {
		console.error("Usage: node scripts/inspect_style.js <docx-path> <fingerprint-label-or-hash>");
		process.exit(1);
	}
	try {
		const doc = await loadDocx(file);
		const sum = doc.summary.find((s) => s.label === handle || s.hash === handle);
		if (!sum) {
			const letters = doc.summary.map((s) => s.label).join(", ");
			const hashes = doc.summary.map((s) => s.hash).join(", ");
			console.error(`Fingerprint "${handle}" not found.\n  Available letters: [${letters}]\n  Available hashes:  [${hashes}]`);
			process.exit(1);
		}
		const label = sum.label;
		const matches = doc.paragraphs.filter((p) => p.fingerprint === label);
		const cellMatches = findDataCellMatches(doc, sum.rawFingerprint);
		const styleIds = Array.from(new Set([...matches.map((p) => p.styleId), ...cellMatches.map((c) => c.styleId)]));
		const totalCount = matches.length + cellMatches.length;
		const out = [];
		out.push(`Fingerprint ${label}: ${sum.description}`);
		if (matches.length > 0) {
			const first = matches[0];
			out.push(`Computed rPr: ${formatRPr(first)}`);
			out.push(`Computed pPr: ${formatPPr(first, matches)}`);
		} else if (cellMatches.length > 0) {
			const first = cellMatches[0];
			out.push(`Computed rPr: ${formatRPrRaw(first.rPr)}`);
			out.push(`Computed pPr: ${formatPPrRaw(first.pPr, cellMatches)}`);
		}
		out.push(`Referenced pStyles: [${styleIds.map((s) => `"${s}"`).join(", ")}]`);
		if (cellMatches.length > 0) out.push(`Occurrences: ${totalCount} (of which ${cellMatches.length} inside data-table cells)`);
		else out.push(`Occurrences: ${totalCount}`);
		out.push("");
		const indexedLimit = Math.min(20, matches.length);
		for (let i = 0; i < indexedLimit; i++) {
			const p = matches[i];
			out.push(`  #${pad(p.index)} [${p.fingerprint}]  "${truncate(p.text, 40)}"`);
		}
		if (matches.length > indexedLimit) out.push(`  ... (showing first ${indexedLimit} indexed, total ${matches.length})`);
		const cellLimit = Math.min(20, cellMatches.length);
		for (let i = 0; i < cellLimit; i++) {
			const c = cellMatches[i];
			const { table, row, col, paragraph } = c.coords;
			const loc = `T${table}R${row}C${col} K${paragraph}`;
			out.push(`  ${loc.padEnd(14)} [${label}]  "${truncate(c.text, 40)}"`);
		}
		if (cellMatches.length > cellLimit) out.push(`  ... (showing first ${cellLimit} cell entries, total ${cellMatches.length})`);
		console.log(out.join("\n"));
	} catch (err) {
		console.error(`Error: ${err.message}`);
		process.exit(1);
	}
}
/** Walk all data-table cells and return paragraphs whose fingerprint matches
* `rawFingerprint`. Uses `StyleResolver` directly to compute rPr/pPr for
* each cell paragraph (same cascade logic as DocumentParser). */
function findDataCellMatches(doc, rawFingerprint) {
	const out = [];
	const body = firstChildNS(doc.documentDoc.documentElement, NS.w, "body");
	if (!body) return out;
	const defaultStyleId = doc.resolver.getDefaultParagraphStyleId() || "Normal";
	let tableIdx = 0;
	for (const child of getChildren(body)) {
		if (child.namespaceURI !== NS.w || child.localName !== "tbl") continue;
		tableIdx++;
		if (summarizeTable(child).classification !== "data") continue;
		const rows = getChildrenNS(child, NS.w, "tr");
		for (let ri = 0; ri < rows.length; ri++) {
			const cells = getChildrenNS(rows[ri], NS.w, "tc");
			for (let ci = 0; ci < cells.length; ci++) {
				const paras = getChildrenNS(cells[ci], NS.w, "p");
				for (let pi = 0; pi < paras.length; pi++) {
					const pEl = paras[pi];
					const pPrEl = firstChildNS(pEl, NS.w, "pPr");
					const pStyleEl = pPrEl ? firstChildNS(pPrEl, NS.w, "pStyle") : null;
					const literalStyleId = pStyleEl ? wVal(pStyleEl) : null;
					if (literalStyleId?.startsWith("_")) continue;
					const styleId = literalStyleId || defaultStyleId;
					const runs = getChildrenNS(pEl, NS.w, "r");
					let chosenRPrEl = null;
					let bestLen = -1;
					for (const r of runs) {
						const rPrEl = firstChildNS(r, NS.w, "rPr");
						let txt = "";
						for (const c of getChildren(r)) if (c.namespaceURI === NS.w && c.localName === "t") txt += textContent(c);
						const len = txt.length > 0 && /^[\d一二三四五六七八九十百千零０-９.()（）【】［］〔〕[\]•·○●◆◇■□★☆※\-、，,\s]+$/.test(txt) ? 0 : txt.length;
						if (len >= bestLen) {
							bestLen = len;
							chosenRPrEl = rPrEl;
						}
					}
					const paraMarkRPrEl = pPrEl ? firstChildNS(pPrEl, NS.w, "rPr") : null;
					const computed = doc.resolver.computeRunStyle(styleId, chosenRPrEl || paraMarkRPrEl);
					const directPPr = doc.resolver.parsePPr(pPrEl);
					const finalPPr = { ...computed.pPr };
					for (const k of Object.keys(directPPr)) {
						const v = directPPr[k];
						if (v !== void 0) Object.assign(finalPPr, { [k]: v });
					}
					if (computeRawFingerprint(computed.rPr, finalPPr) !== rawFingerprint) continue;
					let text = "";
					for (const r of runs) for (const c of getChildren(r)) if (c.namespaceURI === NS.w && c.localName === "t") text += textContent(c);
					out.push({
						coords: {
							table: tableIdx,
							row: ri + 1,
							col: ci + 1,
							paragraph: pi + 1
						},
						rPr: computed.rPr,
						pPr: finalPPr,
						styleId,
						text
					});
				}
			}
		}
	}
	return out;
}
function formatRPr(p) {
	return formatRPrRaw(p.rPr);
}
function formatRPrRaw(r) {
	const parts = formatComputedRPrParts(r, {
		truthyToggles: true,
		filterAutoColor: true
	});
	if (r.underline) parts.push(`underline: ${r.underline}`);
	return parts.length === 0 ? "{}" : `{ ${parts.join(", ")} }`;
}
function formatPPr(p, all) {
	const numIdDisplay = resolveNumIdFromList(all.flatMap((q) => q.pPr.numId !== void 0 ? [q.pPr.numId] : []), all.some((q) => q.pPr.numId === void 0));
	const parts = formatComputedPPrParts(p.pPr, { numIdDisplay });
	return parts.length === 0 ? "{}" : `{ ${parts.join(", ")} }`;
}
function formatPPrRaw(pPr, all) {
	const parts = formatComputedPPrParts(pPr, { numIdDisplay: resolveNumIdFromList(all.flatMap((c) => c.pPr.numId !== void 0 ? [c.pPr.numId] : []), all.some((c) => c.pPr.numId === void 0)) });
	return parts.length === 0 ? "{}" : `{ ${parts.join(", ")} }`;
}
function resolveNumIdFromList(numIds, hasUnset) {
	const seen = new Set(numIds);
	if (seen.size === 0) return null;
	if (seen.size === 1 && !hasUnset) return Array.from(seen)[0];
	return "mixed";
}
main();

//#endregion
export {  };