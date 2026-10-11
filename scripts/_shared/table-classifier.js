import { a as firstChildNS, d as textContent, h as NS, i as descendantsNS, p as wVal, s as getChildrenNS } from "./xml-utils.js";

//#region lib/parse/table-classifier.ts
/** Single cell holding more paragraphs than this is treated as a body
* container and forces the table into the `layout` bucket. Calibrated
* against survey of real fixtures: true data cells never exceed 3
* paragraphs; multi-line form cells (申报书 checkbox lists, proposal
* evaluation rubrics) sit at 5; layout content containers start at 4
* but typically reach 10+. Threshold 5 keeps form / data cells stable
* while catching content-heavy multi-tc layout cases that S1 misses. */
const BULK_CELL_PARA_THRESHOLD = 5;
function summarizeTable(tbl) {
	const rows = getChildrenNS(tbl, NS.w, "tr");
	const rowCount = rows.length;
	let maxCols = 0;
	let maxCellParas = 0;
	const rowTcCounts = [];
	for (const tr of rows) {
		const tcs = getChildrenNS(tr, NS.w, "tc");
		let cells = 0;
		for (const tc of tcs) {
			const tcPr = firstChildNS(tc, NS.w, "tcPr");
			const gridSpan = tcPr ? firstChildNS(tcPr, NS.w, "gridSpan") : null;
			const span = gridSpan ? parseInt(wVal(gridSpan) || "1", 10) : 1;
			cells += span;
			const cellParas = descendantsNS(tc, NS.w, "p").length;
			if (cellParas > maxCellParas) maxCellParas = cellParas;
		}
		rowTcCounts.push(tcs.length);
		if (cells > maxCols) maxCols = cells;
	}
	const row1Texts = [];
	if (rows.length > 0) for (const tc of getChildrenNS(rows[0], NS.w, "tc")) row1Texts.push(collectCellText(tc).trim());
	const totalParas = descendantsNS(tbl, NS.w, "p").length;
	const allSingleTcPerRow = rowTcCounts.every((c) => c === 1);
	const hasOutlineHeading = descendantsNS(tbl, NS.w, "outlineLvl").length > 0;
	if (rowCount === 1 && maxCols === 1) return {
		classification: "layout",
		rows: rowCount,
		cols: maxCols,
		row1Texts,
		classificationReason: "1x1"
	};
	let classification = "data";
	let classificationReason = "fallback";
	if (allSingleTcPerRow && totalParas > 3) {
		classification = "layout";
		classificationReason = "singleTcStack";
	} else if (hasOutlineHeading) {
		classification = "layout";
		classificationReason = "outlineLvl";
	} else if (maxCellParas > BULK_CELL_PARA_THRESHOLD) {
		classification = "layout";
		classificationReason = "bulkCell";
	} else if (rowCount > 1 && maxCols > 1) {
		classification = "data";
		classificationReason = "multiColData";
	}
	return {
		classification,
		rows: rowCount,
		cols: maxCols,
		row1Texts,
		classificationReason
	};
}
function collectCellText(tc) {
	return descendantsNS(tc, NS.w, "t").map((t) => textContent(t)).join("");
}

//#endregion
export { summarizeTable as t };