import { g as DocxReader, h as NS, s as getChildrenNS } from "./_shared/xml-utils.js";
import { t as summarizeTable } from "./_shared/table-classifier.js";
import { c as walkTopLevelTables, r as paragraphText, s as walkIndexedParagraphs } from "./_shared/locator.js";

//#region skill/tools/inspect-table.ts
/**
* `inspect_table` — list top-level tables with cell text snippets.
*
* Output is the cell coordinate space the `cell` locator addresses:
*   table T row R col C  → text snippet (first 40 chars)
*
* Used before composing a `cell` locator. Lists every top-level table —
* data, layout — because cell locators can address any of them
* (paragraph indices skip data tables, but cell locators don't).
*/
async function main() {
	const file = process.argv[2];
	if (!file) {
		console.error("Usage: node scripts/inspect_table.js <docx-path>");
		process.exit(1);
	}
	try {
		const documentDoc = await (await DocxReader.open(file)).readXml("word/document.xml");
		if (!documentDoc) {
			console.error("word/document.xml not found");
			process.exit(1);
		}
		const tables = walkTopLevelTables(documentDoc);
		if (tables.length === 0) {
			console.log("(no top-level tables)");
			return;
		}
		const indexByElement = /* @__PURE__ */ new Map();
		for (const p of walkIndexedParagraphs(documentDoc)) indexByElement.set(p.element, p.index);
		const out = [];
		for (const t of tables) {
			const summary = summarizeTable(t.element);
			const rows = getChildrenNS(t.element, NS.w, "tr");
			out.push(`=== Table ${t.tableIndex} (${summary.classification}) ${rows.length}×${summary.cols} ===`);
			for (let r = 0; r < rows.length; r++) {
				const cells = getChildrenNS(rows[r], NS.w, "tc");
				for (let c = 0; c < cells.length; c++) {
					const paras = getChildrenNS(cells[c], NS.w, "p");
					const paraSpan = formatParaSpan(cells[c], indexByElement);
					if (paras.length <= 1) {
						const text = paras[0] ? paragraphText(paras[0]).trim() : "";
						const snippet = text.length > 40 ? text.slice(0, 40) + "…" : text;
						out.push(`  [${r + 1},${c + 1}] ${JSON.stringify(snippet)}${paraSpan}`);
					} else {
						out.push(`  [${r + 1},${c + 1}] paras:${paras.length}${paraSpan}`);
						for (let k = 0; k < paras.length; k++) {
							const pText = paragraphText(paras[k]).trim();
							const snippet = pText.length > 40 ? pText.slice(0, 40) + "…" : pText;
							out.push(`         K${k + 1}: ${JSON.stringify(snippet)}`);
						}
					}
				}
			}
			out.push("");
		}
		console.log(out.join("\n").trimEnd());
	} catch (err) {
		console.error(`Error: ${err.message}`);
		process.exit(1);
	}
}
/** "  paras: 58–89" / "  paras: 60" / "" (empty when cell paragraphs are
* unindexed, i.e. inside a data table — cell locator addresses them
* by [r,c] anyway). Layout-table paragraphs each have a #NNN; surfacing
* the span lets agents pick a non-cross-cell range locator. */
function formatParaSpan(tc, indexByElement) {
	const indices = [];
	for (const p of getChildrenNS(tc, NS.w, "p")) {
		const idx = indexByElement.get(p);
		if (idx !== void 0) indices.push(idx);
	}
	if (indices.length === 0) return "";
	if (indices.length === 1) return `  paras: ${indices[0]}`;
	const first = indices[0];
	const last = indices[indices.length - 1];
	return last - first + 1 === indices.length ? `  paras: ${first}–${last}` : `  paras: ${indices.join(", ")}`;
}
main();

//#endregion
export {  };