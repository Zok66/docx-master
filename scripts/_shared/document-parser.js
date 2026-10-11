import { a as firstChildNS, d as textContent, f as wAttr, h as NS, i as descendantsNS, n as attr, o as getChildren, p as wVal, s as getChildrenNS, u as parseToggle } from "./xml-utils.js";
import { t as summarizeTable } from "./table-classifier.js";

//#region lib/parse/style-resolver.ts
var StyleResolver = class {
	docDefaultsRPr = {};
	docDefaultsPPr = {};
	styles = /* @__PURE__ */ new Map();
	themeFonts = {};
	themeColors = {};
	rawStyles = /* @__PURE__ */ new Map();
	resolvedCache = /* @__PURE__ */ new Map();
	constructor(stylesDoc, themeDoc) {
		if (themeDoc) this.parseTheme(themeDoc);
		if (stylesDoc) this.parseStyles(stylesDoc);
	}
	parseTheme(theme) {
		const root = theme.documentElement;
		if (!root) return;
		const themeElements = root.getElementsByTagNameNS(NS.a, "themeElements")[0];
		if (!themeElements) return;
		const fontScheme = themeElements.getElementsByTagNameNS(NS.a, "fontScheme")[0];
		if (fontScheme) {
			const major = fontScheme.getElementsByTagNameNS(NS.a, "majorFont")[0];
			const minor = fontScheme.getElementsByTagNameNS(NS.a, "minorFont")[0];
			if (major) {
				const latin = major.getElementsByTagNameNS(NS.a, "latin")[0];
				const ea = major.getElementsByTagNameNS(NS.a, "ea")[0];
				if (latin) this.themeFonts.majorLatin = attr(latin, "", "typeface") || void 0;
				if (ea) this.themeFonts.majorEastAsia = attr(ea, "", "typeface") || void 0;
			}
			if (minor) {
				const latin = minor.getElementsByTagNameNS(NS.a, "latin")[0];
				const ea = minor.getElementsByTagNameNS(NS.a, "ea")[0];
				if (latin) this.themeFonts.minorLatin = attr(latin, "", "typeface") || void 0;
				if (ea) this.themeFonts.minorEastAsia = attr(ea, "", "typeface") || void 0;
			}
		}
		const clrScheme = themeElements.getElementsByTagNameNS(NS.a, "clrScheme")[0];
		if (clrScheme) for (const slot of [
			"dk1",
			"lt1",
			"dk2",
			"lt2",
			"accent1",
			"accent2",
			"accent3",
			"accent4",
			"accent5",
			"accent6",
			"hlink",
			"folHlink"
		]) {
			const el = clrScheme.getElementsByTagNameNS(NS.a, slot)[0];
			if (!el) continue;
			const srgb = el.getElementsByTagNameNS(NS.a, "srgbClr")[0];
			const sys = el.getElementsByTagNameNS(NS.a, "sysClr")[0];
			if (srgb) {
				const v = attr(srgb, "", "val");
				if (v) this.themeColors[slot] = v.toUpperCase();
			} else if (sys) {
				const v = attr(sys, "", "lastClr") || attr(sys, "", "val");
				if (v) this.themeColors[slot] = v.toUpperCase();
			}
		}
	}
	parseStyles(stylesDoc) {
		const root = stylesDoc.documentElement;
		if (!root) return;
		const docDefaults = firstChildNS(root, NS.w, "docDefaults");
		if (docDefaults) {
			const rPrDefault = firstChildNS(docDefaults, NS.w, "rPrDefault");
			if (rPrDefault) {
				const rPr = firstChildNS(rPrDefault, NS.w, "rPr");
				this.docDefaultsRPr = this.parseRPr(rPr);
			}
			const pPrDefault = firstChildNS(docDefaults, NS.w, "pPrDefault");
			if (pPrDefault) {
				const pPr = firstChildNS(pPrDefault, NS.w, "pPr");
				this.docDefaultsPPr = this.parsePPr(pPr);
			}
		}
		for (const styleEl of getChildrenNS(root, NS.w, "style")) {
			const id = wAttr(styleEl, "styleId") || "";
			if (!id) continue;
			const type = wAttr(styleEl, "type") || "paragraph";
			const isDefault = wAttr(styleEl, "default") === "1";
			const name = wVal(firstChildNS(styleEl, NS.w, "name")) || id;
			const basedOnEl = firstChildNS(styleEl, NS.w, "basedOn");
			const basedOn = basedOnEl ? wVal(basedOnEl) : null;
			const rPr = this.parseRPr(firstChildNS(styleEl, NS.w, "rPr"));
			const pPr = this.parsePPr(firstChildNS(styleEl, NS.w, "pPr"));
			this.rawStyles.set(id, {
				id,
				name,
				type,
				basedOn,
				rPr,
				pPr,
				isDefault
			});
		}
		for (const [id, raw] of this.rawStyles) this.styles.set(id, {
			id,
			name: raw.name,
			type: raw.type,
			basedOn: raw.basedOn,
			rPr: raw.rPr,
			pPr: raw.pPr,
			isDefault: raw.isDefault,
			usageCount: 0
		});
	}
	parseRPr(rPrEl) {
		const out = {};
		if (!rPrEl) return out;
		const rFonts = firstChildNS(rPrEl, NS.w, "rFonts");
		if (rFonts) {
			const ascii = wAttr(rFonts, "ascii");
			const hAnsi = wAttr(rFonts, "hAnsi");
			const eastAsia = wAttr(rFonts, "eastAsia");
			const asciiTheme = wAttr(rFonts, "asciiTheme");
			const hAnsiTheme = wAttr(rFonts, "hAnsiTheme");
			const eastAsiaTheme = wAttr(rFonts, "eastAsiaTheme");
			let rAscii = asciiTheme ? this.resolveThemeFont(asciiTheme, false) : ascii || void 0;
			let rHAnsi = hAnsiTheme ? this.resolveThemeFont(hAnsiTheme, false) : hAnsi || void 0;
			let rEastAsia = eastAsiaTheme ? this.resolveThemeFont(eastAsiaTheme, true) : eastAsia || void 0;
			if (rAscii && hasCJK(rAscii)) {
				rEastAsia ??= rAscii;
				rAscii = void 0;
			}
			if (rHAnsi && hasCJK(rHAnsi)) {
				rEastAsia ??= rHAnsi;
				rHAnsi = void 0;
			}
			if (rAscii) out.fontAscii = rAscii;
			if (rHAnsi) out.fontHAnsi = rHAnsi;
			if (rEastAsia) out.fontEastAsia = rEastAsia;
		}
		const sz = firstChildNS(rPrEl, NS.w, "sz");
		if (sz) {
			const v = wVal(sz);
			if (v) out.size = parseInt(v, 10);
		}
		const szCs = firstChildNS(rPrEl, NS.w, "szCs");
		if (szCs && out.size === void 0) {
			const v = wVal(szCs);
			if (v) out.size = parseInt(v, 10);
		}
		const b = firstChildNS(rPrEl, NS.w, "b");
		const bCs = firstChildNS(rPrEl, NS.w, "bCs");
		const bToggle = parseToggle(b) ?? parseToggle(bCs);
		if (bToggle !== void 0) out.bold = bToggle;
		const i = firstChildNS(rPrEl, NS.w, "i");
		const iCs = firstChildNS(rPrEl, NS.w, "iCs");
		const iToggle = parseToggle(i) ?? parseToggle(iCs);
		if (iToggle !== void 0) out.italic = iToggle;
		const u = firstChildNS(rPrEl, NS.w, "u");
		if (u) out.underline = wVal(u) || "single";
		const color = firstChildNS(rPrEl, NS.w, "color");
		if (color) {
			const themeColor = wAttr(color, "themeColor");
			if (themeColor) out.color = this.resolveThemeColor(themeColor) ?? wVal(color) ?? "auto";
			else {
				const v = wVal(color);
				if (v) out.color = v;
			}
		}
		const highlight = firstChildNS(rPrEl, NS.w, "highlight");
		if (highlight) out.highlight = wVal(highlight) || void 0;
		const strikeToggle = parseToggle(firstChildNS(rPrEl, NS.w, "strike"));
		if (strikeToggle !== void 0) out.strike = strikeToggle;
		const capsToggle = parseToggle(firstChildNS(rPrEl, NS.w, "caps"));
		if (capsToggle !== void 0) out.caps = capsToggle;
		const vertAlign = firstChildNS(rPrEl, NS.w, "vertAlign");
		if (vertAlign) {
			const v = wVal(vertAlign);
			if (v === "superscript" || v === "subscript" || v === "baseline") out.vertAlign = v;
		}
		return out;
	}
	parsePPr(pPrEl) {
		const out = {};
		if (!pPrEl) return out;
		const pStyle = firstChildNS(pPrEl, NS.w, "pStyle");
		if (pStyle) {
			const v = wVal(pStyle);
			if (v) out.pStyle = v;
		}
		const jc = firstChildNS(pPrEl, NS.w, "jc");
		if (jc) {
			const v = wVal(jc);
			if (v) out.alignment = v;
		}
		const spacing = firstChildNS(pPrEl, NS.w, "spacing");
		if (spacing) {
			const before = wAttr(spacing, "before");
			const after = wAttr(spacing, "after");
			const line = wAttr(spacing, "line");
			const lineRule = wAttr(spacing, "lineRule");
			if (before) out.spaceBefore = parseInt(before, 10);
			if (after) out.spaceAfter = parseInt(after, 10);
			if (line) out.lineSpacing = parseInt(line, 10);
			if (lineRule) out.lineRule = lineRule;
		}
		const ind = firstChildNS(pPrEl, NS.w, "ind");
		if (ind) {
			const left = wAttr(ind, "left") ?? wAttr(ind, "start");
			const right = wAttr(ind, "right") ?? wAttr(ind, "end");
			const firstLine = wAttr(ind, "firstLine");
			const hanging = wAttr(ind, "hanging");
			const firstLineChars = wAttr(ind, "firstLineChars");
			const hangingChars = wAttr(ind, "hangingChars");
			if (left) out.indentLeft = parseInt(left, 10);
			if (right) out.indentRight = parseInt(right, 10);
			if (firstLine) out.firstLineIndent = parseInt(firstLine, 10);
			if (firstLineChars) out.firstLineIndentChars = parseInt(firstLineChars, 10);
			if (hanging) out.hangingIndent = parseInt(hanging, 10);
			if (hangingChars) out.hangingIndentChars = parseInt(hangingChars, 10);
		}
		const outlineLvl = firstChildNS(pPrEl, NS.w, "outlineLvl");
		if (outlineLvl) {
			const v = wVal(outlineLvl);
			if (v !== null) out.outlineLevel = parseInt(v, 10);
		}
		const numPr = firstChildNS(pPrEl, NS.w, "numPr");
		if (numPr) {
			const numId = firstChildNS(numPr, NS.w, "numId");
			const ilvl = firstChildNS(numPr, NS.w, "ilvl");
			if (numId) {
				const v = wVal(numId);
				if (v) out.numId = v;
			}
			if (ilvl) {
				const v = wVal(ilvl);
				if (v !== null) out.numLevel = parseInt(v, 10);
			}
		}
		return out;
	}
	/**
	* Mutate the parsed theme fonts — call before parseTheme dependents read
	* `themeFonts`. Used by `applyThemeFontOverrides` to keep the resolver's
	* cache in sync after the theme1.xml DOM has been updated.
	*/
	setThemeFontOverrides(spec) {
		if (spec.majorLatin) this.themeFonts.majorLatin = spec.majorLatin;
		if (spec.majorEastAsia) this.themeFonts.majorEastAsia = spec.majorEastAsia;
		if (spec.minorLatin) this.themeFonts.minorLatin = spec.minorLatin;
		if (spec.minorEastAsia) this.themeFonts.minorEastAsia = spec.minorEastAsia;
	}
	/**
	* Walk a stylesDoc and expand every rFonts element's themed font attrs
	* (asciiTheme / hAnsiTheme / eastAsiaTheme / cstheme) to literal attrs
	* (ascii / hAnsi / eastAsia / cs) by resolving against the parsed theme.
	*
	* Why this matters: OOXML §17.3.2.27 prescribes "themed wins over literal"
	* when both are specified on the same rFonts element. That rule fires
	* during the cascade merge too — if the doc's <w:docDefaults> uses themed
	* fonts (very common in Office / WPS / 国产 docx templates) and an agent
	* injects a style with literal fonts via this skill, Word merges them and
	* the doc-defaults theme reference silently overrides the agent's literal
	* value at render time. The agent's pre-flight inspect_* output and our
	* Style Resolution report would say "宋体" while Word renders 等线.
	*
	* Pre-expanding stylesDoc to literal-only attrs eliminates the cascade
	* conflict: every later rFonts merge sees only literal-vs-literal, and
	* child-overrides-parent string semantics apply as expected.
	*
	* Safe to call before any other styles.xml mutations. Mutates stylesDoc
	* in place. Idempotent (a second call finds no themed attrs to expand).
	*/
	expandThemedFontsInStyles(stylesDoc) {
		const root = stylesDoc.documentElement;
		if (!root) return;
		const visit = (el) => {
			if (el.namespaceURI === NS.w && el.localName === "rFonts") this.expandRFontsThemedAttrs(el);
			const childNodes = el.childNodes;
			for (let i = 0; i < childNodes.length; i++) {
				const c = childNodes[i];
				if (c.nodeType === 1) visit(c);
			}
		};
		visit(root);
	}
	expandRFontsThemedAttrs(rFonts) {
		const w = NS.w;
		for (const [themedAttr, literalAttr, isEastAsia] of [
			[
				"asciiTheme",
				"ascii",
				false
			],
			[
				"hAnsiTheme",
				"hAnsi",
				false
			],
			[
				"eastAsiaTheme",
				"eastAsia",
				true
			],
			[
				"cstheme",
				"cs",
				false
			]
		]) {
			const themed = wAttr(rFonts, themedAttr);
			if (!themed) continue;
			const resolved = this.resolveThemeFont(themed, isEastAsia);
			if (resolved) rFonts.setAttributeNS(w, `w:${literalAttr}`, resolved);
			rFonts.removeAttributeNS(w, themedAttr);
		}
	}
	resolveThemeFont(themeRef, isEastAsia) {
		switch (themeRef) {
			case "majorAscii":
			case "majorHAnsi":
			case "majorBidi": return this.themeFonts.majorLatin;
			case "minorAscii":
			case "minorHAnsi":
			case "minorBidi": return this.themeFonts.minorLatin;
			case "majorEastAsia": return this.themeFonts.majorEastAsia;
			case "minorEastAsia": return this.themeFonts.minorEastAsia;
		}
		return isEastAsia ? this.themeFonts.minorEastAsia : this.themeFonts.minorLatin;
	}
	resolveThemeColor(themeRef) {
		const slot = {
			dark1: "dk1",
			light1: "lt1",
			dark2: "dk2",
			light2: "lt2",
			hyperlink: "hlink",
			followedHyperlink: "folHlink",
			background1: "lt1",
			background2: "lt2",
			text1: "dk1",
			text2: "dk2"
		}[themeRef] || themeRef;
		return this.themeColors[slot];
	}
	/** Compute final style for a paragraph + optional inline run rPr. */
	computeRunStyle(paraStyleId, runRPrElement) {
		const styleResolved = this.resolveStyleChain(paraStyleId);
		let rPr = { ...styleResolved.rPr };
		let pPr = { ...styleResolved.pPr };
		if (runRPrElement) {
			const direct = this.parseRPr(runRPrElement);
			const rStyle = firstChildNS(runRPrElement, NS.w, "rStyle");
			if (rStyle) {
				const csId = wVal(rStyle);
				if (csId) {
					const cs = this.resolveStyleChain(csId);
					rPr = mergeRPr(rPr, cs.rPr);
				}
			}
			rPr = mergeRPr(rPr, direct);
		}
		return {
			rPr,
			pPr,
			styleName: styleResolved.styleName,
			styleChain: styleResolved.chain
		};
	}
	/** Resolve a style by id, walking basedOn chain from root downwards. */
	resolveStyleChain(styleId) {
		const cached = this.resolvedCache.get(styleId);
		if (cached) return cached;
		const chain = this.buildChain(styleId);
		let rPr = { ...this.docDefaultsRPr };
		let pPr = { ...this.docDefaultsPPr };
		let styleName = styleId;
		for (const id of chain) {
			const raw = this.rawStyles.get(id);
			if (!raw) continue;
			pPr = mergePPr(pPr, raw.pPr);
			rPr = mergeRPr(rPr, raw.rPr);
			styleName = raw.name;
		}
		if (chain.length === 0) styleName = styleId || "Normal";
		const result = {
			rPr,
			pPr,
			styleName,
			chain
		};
		this.resolvedCache.set(styleId, result);
		return result;
	}
	buildChain(styleId) {
		const visited = /* @__PURE__ */ new Set();
		const path = [];
		let cur = styleId;
		while (cur && !visited.has(cur)) {
			visited.add(cur);
			const raw = this.rawStyles.get(cur);
			if (!raw) break;
			path.push(cur);
			cur = raw.basedOn;
		}
		return path.reverse();
	}
	getStyleDefinition(styleId) {
		return this.styles.get(styleId) || null;
	}
	getAllStyles() {
		return Array.from(this.styles.values());
	}
	getDocDefaults() {
		return {
			rPr: this.docDefaultsRPr,
			pPr: this.docDefaultsPPr
		};
	}
	getThemeFonts() {
		return this.themeFonts;
	}
	getThemeColors() {
		return this.themeColors;
	}
	incrementUsage(styleId) {
		const def = this.styles.get(styleId);
		if (def) def.usageCount++;
	}
	/** styleId of the paragraph style flagged `w:default="1"`, or null when
	* the document has none. Parsers use this as the fallback for paragraphs
	* that omit `<w:pStyle>` so usage attribution lands on the actual
	* default style (POI/WPS auto-generate ids like "a") instead of the
	* literal name "Normal" which usually doesn't match any styleId. */
	getDefaultParagraphStyleId() {
		for (const s of this.styles.values()) if (s.isDefault && s.type === "paragraph") return s.id;
		return null;
	}
};
/**
* Overlay only the defined fields of `overlay` onto a copy of `base`. Used
* for merging both rPr and pPr inheritance frames; the field semantics are
* the same (later wins iff non-undefined) so one helper covers both.
*/
/** Does `s` contain a CJK Unified Ideograph (basic block or extension A)?
* Font names are short ASCII or short CJK; one CJK char is enough signal
* that the value belongs in the eastAsia slot, not ascii/hAnsi. */
function hasCJK(s) {
	for (let i = 0; i < s.length; i++) {
		const c = s.charCodeAt(i);
		if (c >= 19968 && c <= 40959) return true;
		if (c >= 13312 && c <= 19903) return true;
	}
	return false;
}
function mergeStyle(base, overlay) {
	const out = { ...base };
	for (const k of Object.keys(overlay)) {
		const v = overlay[k];
		if (v !== void 0) Object.assign(out, { [k]: v });
	}
	return out;
}
function mergeRPr(base, overlay) {
	return mergeStyle(base, overlay);
}
function mergePPr(base, overlay) {
	return mergeStyle(base, overlay);
}
/**
* Mutate `themeDoc` (theme1.xml) in place to override the typeface attrs of
* the theme font scheme. Call BEFORE constructing StyleResolver so the
* resolver picks up the new values from a fresh parseTheme; if the resolver
* already exists, follow up with `resolver.setThemeFontOverrides(spec)` to
* keep its cache in sync.
*
* Only the slots specified in `spec` are touched. `spec` keys map to:
*   majorLatin    → fontScheme/majorFont/latin@typeface
*   majorEastAsia → fontScheme/majorFont/ea@typeface
*   minorLatin    → fontScheme/minorFont/latin@typeface
*   minorEastAsia → fontScheme/minorFont/ea@typeface
*
* If theme1.xml has the empty `<a:ea typeface=""/>` pattern (very common in
* zh-CN templates that rely on system fallback), this writes a real typeface
* value into the empty slot. Combined with the stylesDoc themed-attr
* expansion that runs later, the new theme font ends up being the literal
* value used by every downstream cascade.
*/
function applyThemeFontOverrides(themeDoc, spec) {
	const root = themeDoc.documentElement;
	if (!root) return;
	const themeElements = root.getElementsByTagNameNS(NS.a, "themeElements")[0];
	if (!themeElements) return;
	const fontScheme = themeElements.getElementsByTagNameNS(NS.a, "fontScheme")[0];
	if (!fontScheme) return;
	const setTypeface = (fontEl, childLocalName, value) => {
		if (!fontEl) return;
		let target = fontEl.getElementsByTagNameNS(NS.a, childLocalName)[0];
		if (!target) {
			target = themeDoc.createElementNS(NS.a, `a:${childLocalName}`);
			fontEl.appendChild(target);
		}
		target.setAttribute("typeface", value);
	};
	const majorFont = fontScheme.getElementsByTagNameNS(NS.a, "majorFont")[0];
	const minorFont = fontScheme.getElementsByTagNameNS(NS.a, "minorFont")[0];
	if (spec.majorLatin) setTypeface(majorFont, "latin", spec.majorLatin);
	if (spec.majorEastAsia) setTypeface(majorFont, "ea", spec.majorEastAsia);
	if (spec.minorLatin) setTypeface(minorFont, "latin", spec.minorLatin);
	if (spec.minorEastAsia) setTypeface(minorFont, "ea", spec.minorEastAsia);
}

//#endregion
//#region lib/parse/document-parser.ts
var DocumentParser = class {
	docXml;
	resolver;
	numberingDoc;
	headerDocs;
	footerDocs;
	rels;
	nextParaIndex = 1;
	currentSection = 0;
	constructor(docXml, resolver, numberingDoc, options) {
		this.docXml = docXml;
		this.resolver = resolver;
		this.numberingDoc = numberingDoc;
		this.headerDocs = options?.headerDocs ?? /* @__PURE__ */ new Map();
		this.footerDocs = options?.footerDocs ?? /* @__PURE__ */ new Map();
		this.rels = options?.rels ?? /* @__PURE__ */ new Map();
	}
	parse() {
		const root = this.docXml.documentElement;
		const body = firstChildNS(root, NS.w, "body");
		if (!body) return {
			paragraphs: [],
			elements: [],
			sections: [],
			neighborItems: []
		};
		const allParagraphs = [];
		const flatItems = [];
		const sections = [];
		let sectionStartIndex = 1;
		const children = getChildren(body);
		for (let idx = 0; idx < children.length; idx++) {
			const child = children[idx];
			if (child.namespaceURI !== NS.w) continue;
			if (child.localName === "p") {
				this.handleParagraph(child, allParagraphs, flatItems, null);
				const pPr = firstChildNS(child, NS.w, "pPr");
				const sectPr = pPr ? firstChildNS(pPr, NS.w, "sectPr") : null;
				if (sectPr) {
					const endIndex = this.nextParaIndex - 1;
					sections.push(this.buildSectionInfo(sectPr, sections.length, sectionStartIndex, endIndex));
					sectionStartIndex = this.nextParaIndex;
					flatItems.push({
						kind: "sectionBreak",
						sectionIndex: this.currentSection
					});
					this.currentSection++;
				}
			} else if (child.localName === "tbl") {
				const summary = summarizeTable(child);
				if (summary.classification === "layout") {
					const innerParas = [];
					flatItems.push({
						kind: "table",
						classification: "layout",
						rows: summary.rows,
						cols: summary.cols,
						row1Texts: summary.row1Texts,
						classificationReason: summary.classificationReason,
						sectionIndex: this.currentSection,
						innerParagraphs: innerParas
					});
					for (const tr of getChildrenNS(child, NS.w, "tr")) for (const tc of getChildrenNS(tr, NS.w, "tc")) this.walkLayoutCell(tc, allParagraphs, flatItems, innerParas);
				} else flatItems.push({
					kind: "table",
					classification: summary.classification,
					rows: summary.rows,
					cols: summary.cols,
					row1Texts: summary.row1Texts,
					classificationReason: summary.classificationReason,
					sectionIndex: this.currentSection
				});
			} else if (child.localName === "sectPr") {
				const endIndex = this.nextParaIndex - 1;
				sections.push(this.buildSectionInfo(child, sections.length, sectionStartIndex, endIndex));
				sectionStartIndex = this.nextParaIndex;
			}
		}
		if (sections.length === 0 && allParagraphs.length > 0) sections.push({
			index: 0,
			paraRange: [1, allParagraphs.length],
			pageSize: {
				width: 0,
				height: 0
			},
			margins: {
				top: 0,
				bottom: 0,
				left: 0,
				right: 0
			},
			orientation: "portrait",
			header: null,
			footer: null,
			headerHasImage: false
		});
		const neighborItems = this.buildNeighborItems(flatItems);
		return {
			paragraphs: allParagraphs,
			elements: this.compressElements(flatItems),
			sections,
			neighborItems
		};
	}
	walkLayoutCell(tc, all, flat, innerParas) {
		for (const c of getChildren(tc)) {
			if (c.namespaceURI !== NS.w) continue;
			if (c.localName === "p") this.handleParagraph(c, all, flat, "layout", innerParas);
			else if (c.localName === "tbl") {
				const summary = summarizeTable(c);
				if (summary.classification === "layout") {
					flat.push({
						kind: "table",
						classification: "layout",
						rows: summary.rows,
						cols: summary.cols,
						row1Texts: summary.row1Texts,
						classificationReason: summary.classificationReason,
						sectionIndex: this.currentSection,
						innerParagraphs: innerParas
					});
					for (const tr of getChildrenNS(c, NS.w, "tr")) for (const tc2 of getChildrenNS(tr, NS.w, "tc")) this.walkLayoutCell(tc2, all, flat, innerParas);
				} else flat.push({
					kind: "table",
					classification: summary.classification,
					rows: summary.rows,
					cols: summary.cols,
					row1Texts: summary.row1Texts,
					classificationReason: summary.classificationReason,
					sectionIndex: this.currentSection
				});
			}
		}
	}
	handleParagraph(pEl, all, flat, insideTable, innerCollect) {
		const pPr = firstChildNS(pEl, NS.w, "pPr");
		const pStyleEl = pPr ? firstChildNS(pPr, NS.w, "pStyle") : null;
		const literalStyleId = pStyleEl ? wVal(pStyleEl) : null;
		if (literalStyleId && literalStyleId.startsWith("_")) return;
		const styleId = literalStyleId || this.resolver.getDefaultParagraphStyleId() || "Normal";
		const runs = getChildrenNS(pEl, NS.w, "r");
		let dominantRun = null;
		let dominantLen = -1;
		let firstNonEmptyRun = null;
		let firstSubstantiveRun = null;
		for (const r of runs) {
			let text = "";
			for (const c of getChildren(r)) if (c.namespaceURI === NS.w && c.localName === "t") text += textContent(c);
			if (text.length > 0 && firstNonEmptyRun === null) firstNonEmptyRun = r;
			const isNumberingPrefix = text.length > 0 && /^[\d一二三四五六七八九十百千零０-９.()（）【】［］〔〕[\]•·○●◆◇■□★☆※\-、，,\s]+$/.test(text);
			if (!isNumberingPrefix && text.length > 0 && firstSubstantiveRun === null) firstSubstantiveRun = r;
			const len = isNumberingPrefix ? 0 : text.length;
			if (len >= dominantLen) {
				dominantLen = len;
				dominantRun = r;
			}
		}
		const chosenRun = dominantLen > 0 ? dominantRun : firstSubstantiveRun || firstNonEmptyRun || runs[0] || null;
		const chosenRPr = chosenRun ? firstChildNS(chosenRun, NS.w, "rPr") : null;
		const paraMarkRPr = pPr ? firstChildNS(pPr, NS.w, "rPr") : null;
		const computed = this.resolver.computeRunStyle(styleId, chosenRPr || paraMarkRPr);
		const directPPr = this.resolver.parsePPr(pPr);
		const finalPPr = { ...computed.pPr };
		for (const k of Object.keys(directPPr)) {
			const v = directPPr[k];
			if (v !== void 0) Object.assign(finalPPr, { [k]: v });
		}
		let text = "";
		let hasDrawing = false;
		let hasEquation = false;
		let imageDims = null;
		let hasPageBreak = false;
		for (const r of runs) for (const c of getChildren(r)) if (c.namespaceURI === NS.w) {
			if (c.localName === "t") text += textContent(c);
			else if (c.localName === "tab") text += "	";
			else if (c.localName === "br") if (wAttr(c, "type") === "page") hasPageBreak = true;
			else text += "\n";
			else if (c.localName === "drawing") {
				hasDrawing = true;
				const dims = extractDrawingDims(c);
				if (dims) imageDims = dims;
			}
		} else if (c.namespaceURI === NS.m) {
			if (c.localName === "oMath" || c.localName === "oMathPara") hasEquation = true;
		}
		for (const c of getChildren(pEl)) if (c.namespaceURI === NS.m) {
			if (c.localName === "oMath" || c.localName === "oMathPara") hasEquation = true;
		}
		this.resolver.incrementUsage(styleId);
		const styleName = this.resolver.getStyleDefinition(styleId)?.name || computed.styleName || styleId;
		const directRPr = this.resolver.parseRPr(chosenRPr || paraMarkRPr);
		const para = {
			index: this.nextParaIndex++,
			text,
			rPr: computed.rPr,
			pPr: finalPPr,
			directRPr,
			directPPr,
			styleId,
			styleName,
			fingerprint: "",
			context: {
				insideTable,
				sectionIndex: this.currentSection
			}
		};
		all.push(para);
		if (innerCollect) innerCollect.push(para);
		const isImageOnly = hasDrawing && text.trim().length === 0;
		const isEquationOnly = hasEquation && text.trim().length === 0;
		if (hasDrawing && imageDims) flat.push({
			kind: "image",
			widthCm: imageDims.widthCm,
			heightCm: imageDims.heightCm,
			sectionIndex: this.currentSection,
			paraIndex: para.index
		});
		if (hasEquation) flat.push({
			kind: "equation",
			sectionIndex: this.currentSection,
			paraIndex: para.index
		});
		if (hasPageBreak) flat.push({
			kind: "pageBreak",
			sectionIndex: this.currentSection
		});
		flat.push({
			kind: "para",
			paragraph: para,
			isImageOnly,
			isEquationOnly
		});
	}
	buildSectionInfo(sectPr, index, startIdx, endIdx) {
		const pgSz = firstChildNS(sectPr, NS.w, "pgSz");
		const pgMar = firstChildNS(sectPr, NS.w, "pgMar");
		const width = pgSz ? parseInt(wAttr(pgSz, "w") || "0", 10) : 0;
		const height = pgSz ? parseInt(wAttr(pgSz, "h") || "0", 10) : 0;
		const orientation = (pgSz ? wAttr(pgSz, "orient") : null) === "landscape" ? "landscape" : "portrait";
		const margins = {
			top: pgMar ? parseInt(wAttr(pgMar, "top") || "0", 10) : 0,
			bottom: pgMar ? parseInt(wAttr(pgMar, "bottom") || "0", 10) : 0,
			left: pgMar ? parseInt(wAttr(pgMar, "left") || "0", 10) : 0,
			right: pgMar ? parseInt(wAttr(pgMar, "right") || "0", 10) : 0
		};
		let header = null;
		let footer = null;
		let headerHasImage = false;
		let footerPageNumFormat;
		for (const ref of getChildrenNS(sectPr, NS.w, "headerReference")) {
			const rId = wAttr(ref, "id");
			const type = wAttr(ref, "type");
			if (rId && (type === "default" || type === null)) {
				const rel = this.rels.get(rId);
				const doc = rel ? this.headerDocs.get(rel.target) : null;
				if (doc) {
					const text = collectAllText(doc.documentElement).trim();
					header = text.length > 0 ? text : "(empty)";
					headerHasImage = descendantsNS(doc.documentElement, NS.w, "drawing").length > 0;
				}
			}
		}
		for (const ref of getChildrenNS(sectPr, NS.w, "footerReference")) {
			const rId = wAttr(ref, "id");
			const type = wAttr(ref, "type");
			if (rId && (type === "default" || type === null)) {
				const rel = this.rels.get(rId);
				const doc = rel ? this.footerDocs.get(rel.target) : null;
				if (doc) {
					const text = collectAllText(doc.documentElement).trim();
					footer = text.length > 0 ? text : "(empty)";
					const instrTexts = descendantsNS(doc.documentElement, NS.w, "instrText").map((e) => textContent(e)).join(" ");
					if (/PAGE/i.test(instrTexts)) {
						const m = instrTexts.match(/\\\*\s*([A-Za-z]+)/);
						footerPageNumFormat = m ? m[1] : "decimal";
					}
				}
			}
		}
		const pgNumType = firstChildNS(sectPr, NS.w, "pgNumType");
		if (pgNumType) {
			const fmt = wAttr(pgNumType, "fmt");
			if (fmt) footerPageNumFormat = fmt;
		}
		return {
			index,
			paraRange: [startIdx, endIdx],
			pageSize: {
				width,
				height
			},
			margins,
			orientation,
			header,
			footer,
			headerHasImage,
			footerPageNumFormat
		};
	}
	/**
	* Flatten the internal RawItem list into the public NeighborItem stream.
	* Each empty paragraph stays as its own entry (unlike compressElements).
	* inspect_neighbors uses this to compute on-demand neighbor windows.
	*/
	buildNeighborItems(items) {
		const out = [];
		for (const it of items) switch (it.kind) {
			case "para": {
				const isEmpty = it.paragraph.text.trim().length === 0 && !it.isImageOnly && !it.isEquationOnly;
				out.push({
					kind: "paragraph",
					paraIndex: it.paragraph.index,
					isEmpty,
					sectionIndex: it.paragraph.context.sectionIndex
				});
				break;
			}
			case "image":
				out.push({
					kind: "image",
					widthCm: it.widthCm,
					heightCm: it.heightCm,
					sectionIndex: it.sectionIndex
				});
				break;
			case "table":
				out.push({
					kind: "table",
					classification: it.classification,
					rows: it.rows,
					cols: it.cols,
					sectionIndex: it.sectionIndex
				});
				break;
			case "equation":
				out.push({
					kind: "equation",
					sectionIndex: it.sectionIndex
				});
				break;
			case "pageBreak":
				out.push({
					kind: "pageBreak",
					sectionIndex: it.sectionIndex
				});
				break;
			case "sectionBreak":
				out.push({
					kind: "sectionBreak",
					sectionIndex: it.sectionIndex
				});
				break;
		}
		return out;
	}
	compressElements(items) {
		const out = [];
		let emptyRunCount = 0;
		let emptyRunFirstIdx = 0;
		let emptyRunSection = 0;
		const flushEmpty = () => {
			if (emptyRunCount > 0) {
				out.push({
					kind: "emptyRun",
					count: emptyRunCount,
					firstIndex: emptyRunFirstIdx,
					sectionIndex: emptyRunSection
				});
				emptyRunCount = 0;
			}
		};
		for (const it of items) if (it.kind === "para") {
			if (it.paragraph.context.insideTable === "layout") continue;
			if (it.paragraph.text.trim().length === 0 && !it.isImageOnly && !it.isEquationOnly) {
				if (emptyRunCount === 0) {
					emptyRunFirstIdx = it.paragraph.index;
					emptyRunSection = it.paragraph.context.sectionIndex;
				}
				emptyRunCount++;
				continue;
			}
			flushEmpty();
			out.push({
				kind: "paragraph",
				paragraph: it.paragraph
			});
		} else if (it.kind === "table") {
			flushEmpty();
			out.push({
				kind: "table",
				classification: it.classification,
				rows: it.rows,
				cols: it.cols,
				row1Texts: it.row1Texts,
				classificationReason: it.classificationReason,
				sectionIndex: it.sectionIndex,
				paragraphs: it.innerParagraphs ?? []
			});
		} else if (it.kind === "image") {
			flushEmpty();
			out.push({
				kind: "image",
				widthCm: it.widthCm,
				heightCm: it.heightCm,
				sectionIndex: it.sectionIndex
			});
		} else if (it.kind === "equation") {
			flushEmpty();
			out.push({
				kind: "equation",
				sectionIndex: it.sectionIndex
			});
		} else if (it.kind === "pageBreak") {
			flushEmpty();
			out.push({
				kind: "pageBreak",
				sectionIndex: it.sectionIndex
			});
		} else if (it.kind === "sectionBreak") {
			flushEmpty();
			out.push({
				kind: "sectionBreak",
				sectionIndex: it.sectionIndex
			});
		}
		flushEmpty();
		return out;
	}
};
function extractDrawingDims(drawingEl) {
	const extents = descendantsNS(drawingEl, NS.wp, "extent");
	for (const e of extents) {
		const cx = parseInt(attrLocal(e, "cx") || "0", 10);
		const cy = parseInt(attrLocal(e, "cy") || "0", 10);
		if (cx > 0 && cy > 0) return {
			widthCm: cx / 36e4,
			heightCm: cy / 36e4
		};
	}
	return null;
}
function attrLocal(el, name) {
	return el.getAttribute(name);
}
function collectAllText(el) {
	return descendantsNS(el, NS.w, "t").map((t) => textContent(t)).join("");
}

//#endregion
export { StyleResolver as n, applyThemeFontOverrides as r, DocumentParser as t };