//#region lib/parse/format.ts
/** Pad a 1-based paragraph index to 3 digits — "001", "042", "123". */
function pad(n) {
	return n.toString().padStart(3, "0");
}
/** Collapse internal whitespace and clip with an ellipsis. */
function truncate(s, n) {
	const collapsed = s.replace(/\s+/g, " ").trim();
	return collapsed.length <= n ? collapsed : collapsed.slice(0, n) + "…";
}
/** Recognize standard paper sizes by their twip dimensions. */
function paperName(width, height) {
	for (const [name, w, h] of [
		[
			"A4",
			11906,
			16838
		],
		[
			"A4",
			16838,
			11906
		],
		[
			"A3",
			16838,
			23811
		],
		[
			"A5",
			8392,
			11906
		],
		[
			"Letter",
			12240,
			15840
		],
		[
			"Legal",
			12240,
			20160
		]
	]) if (Math.abs(width - w) < 50 && Math.abs(height - h) < 50) return name;
	return "Custom";
}
/** Render Word's spacing/@line + spacing/@lineRule pair as human text. */
function formatLineSpacing(line, rule) {
	const r = rule || "auto";
	if (r === "exact") return `${line / 20}pt fixed`;
	if (r === "atLeast") return `${line / 20}pt atLeast`;
	return `${parseFloat((line / 240).toFixed(2))}×`;
}
/** Twips → millimeters, rounded to 1dp (Word page-setup-friendly). */
function tw2mm(t) {
	return +(t / 56.6929).toFixed(1);
}
/** Render the common subset of `ComputedRunStyle` as a parts array — fonts,
* size, bold, italic, color, vertAlign. Caller wraps with `{ ... }` and
* appends tool-specific extras (underline / highlight / strike / caps). */
function formatComputedRPrParts(r, opts = {}) {
	const parts = [];
	const truthy = opts.truthyToggles ?? false;
	const pushToggle = (name, v) => {
		if (v === void 0) return;
		if (truthy && !v) return;
		parts.push(truthy ? `${name}: true` : `${name}: ${v}`);
	};
	const latin = r.fontAscii ?? r.fontHAnsi;
	const setSlots = [latin, r.fontEastAsia].filter((v) => v !== void 0);
	if (setSlots.length >= 2 && setSlots.every((v) => v === setSlots[0])) parts.push(`font: "${setSlots[0]}"`);
	else {
		if (r.fontEastAsia) parts.push(`fontCJK: "${r.fontEastAsia}"`);
		if (latin) parts.push(`fontLatin: "${latin}"`);
	}
	if (r.size !== void 0) parts.push(`size: ${r.size / 2}pt`);
	pushToggle("bold", r.bold);
	pushToggle("italic", r.italic);
	if (r.color && (!opts.filterAutoColor || r.color !== "auto")) parts.push(`color: ${r.color}`);
	if (r.vertAlign) parts.push(`vertAlign: ${r.vertAlign}`);
	return parts;
}
/** Render the common subset of `ComputedParaStyle` as a parts array. */
function formatComputedPPrParts(pp, opts = {}) {
	const parts = [];
	if (pp.alignment) parts.push(`alignment: ${pp.alignment}`);
	if (pp.spaceBefore !== void 0) parts.push(`spaceBefore: ${pp.spaceBefore / 20}pt`);
	if (pp.spaceAfter !== void 0) parts.push(`spaceAfter: ${pp.spaceAfter / 20}pt`);
	if (pp.lineSpacing !== void 0) parts.push(`lineSpacing: ${formatLineSpacing(pp.lineSpacing, pp.lineRule)}`);
	if (opts.includeIndentSides) {
		if (pp.indentLeft !== void 0) parts.push(`indentLeft: ${pp.indentLeft / 20}pt`);
		if (pp.indentRight !== void 0) parts.push(`indentRight: ${pp.indentRight / 20}pt`);
	}
	if (pp.firstLineIndentChars !== void 0) parts.push(`firstLineIndent: ${pp.firstLineIndentChars / 100}char`);
	else if (pp.firstLineIndent !== void 0) parts.push(`firstLineIndent: ${pp.firstLineIndent / 20}pt`);
	if (pp.hangingIndentChars !== void 0) parts.push(`hangingIndent: ${pp.hangingIndentChars / 100}char`);
	else if (pp.hangingIndent !== void 0) parts.push(`hangingIndent: ${pp.hangingIndent / 20}pt`);
	if (pp.outlineLevel !== void 0) parts.push(`outlineLevel: ${pp.outlineLevel}`);
	const numId = opts.numIdDisplay === void 0 ? pp.numId ? String(pp.numId) : null : opts.numIdDisplay;
	if (numId) parts.push(`numId: ${numId}`);
	if (opts.includeNumLevel && pp.numLevel !== void 0) parts.push(`numLevel: ${pp.numLevel}`);
	return parts;
}

//#endregion
export { truncate as a, paperName as i, formatComputedRPrParts as n, tw2mm as o, pad as r, formatComputedPPrParts as t };