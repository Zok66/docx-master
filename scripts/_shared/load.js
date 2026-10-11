import { a as firstChildNS, d as textContent, f as wAttr, g as DocxReader, h as NS, p as wVal, s as getChildrenNS } from "./xml-utils.js";
import { n as StyleResolver, t as DocumentParser } from "./document-parser.js";
import { t as Fingerprinter } from "./fingerprint.js";
import { basename } from "node:path";
import { statSync } from "node:fs";

//#region lib/xml/load.ts
async function loadDocx(filePath) {
	const reader = await DocxReader.open(filePath);
	const stylesDoc = await reader.readXml("word/styles.xml");
	const themeDoc = await reader.readXml("word/theme/theme1.xml");
	const numberingDoc = await reader.readXml("word/numbering.xml");
	const documentDoc = await reader.readXml("word/document.xml");
	if (!documentDoc) throw new Error("word/document.xml not found");
	const relsDoc = await reader.readXml("word/_rels/document.xml.rels");
	const rels = /* @__PURE__ */ new Map();
	if (relsDoc) {
		const root = relsDoc.documentElement;
		if (root) for (const r of getChildrenNS(root, "http://schemas.openxmlformats.org/package/2006/relationships", "Relationship")) {
			const id = r.getAttribute("Id") || "";
			const type = r.getAttribute("Type") || "";
			const target = r.getAttribute("Target") || "";
			if (id) rels.set(id, {
				type,
				target
			});
		}
	}
	const headerDocs = /* @__PURE__ */ new Map();
	const footerDocs = /* @__PURE__ */ new Map();
	for (const [, rel] of rels) if (rel.type.endsWith("/header")) {
		const path = `word/${rel.target}`;
		const d = await reader.readXml(path);
		if (d) headerDocs.set(rel.target, d);
	} else if (rel.type.endsWith("/footer")) {
		const path = `word/${rel.target}`;
		const d = await reader.readXml(path);
		if (d) footerDocs.set(rel.target, d);
	}
	const coreDoc = await reader.readXml("docProps/core.xml");
	let author;
	let title;
	if (coreDoc && coreDoc.documentElement) {
		const dcCreator = coreDoc.documentElement.getElementsByTagNameNS(NS.dc, "creator")[0];
		const dcTitle = coreDoc.documentElement.getElementsByTagNameNS(NS.dc, "title")[0];
		if (dcCreator) author = textContent(dcCreator) || void 0;
		if (dcTitle) title = textContent(dcTitle) || void 0;
	}
	const resolver = new StyleResolver(stylesDoc, themeDoc);
	if (stylesDoc) resolver.expandThemedFontsInStyles(stylesDoc);
	const parsed = new DocumentParser(documentDoc, resolver, numberingDoc, {
		headerDocs,
		footerDocs,
		rels
	}).parse();
	const fpResult = new Fingerprinter().assign(parsed.paragraphs, resolver);
	const stat = statSync(filePath);
	return {
		reader,
		resolver,
		paragraphs: parsed.paragraphs,
		elements: parsed.elements,
		neighborItems: parsed.neighborItems,
		sections: parsed.sections,
		summary: fpResult.summary,
		metadata: {
			fileName: basename(filePath),
			fileSize: stat.size,
			author,
			title
		},
		numberingDoc,
		themeDoc,
		stylesDoc,
		documentDoc
	};
}
function parseNumbering(numberingDoc) {
	if (!numberingDoc) return [];
	const root = numberingDoc.documentElement;
	if (!root) return [];
	const abstractMap = /* @__PURE__ */ new Map();
	for (const a of getChildrenNS(root, NS.w, "abstractNum")) {
		const id = wAttr(a, "abstractNumId") || "";
		const lvls = [];
		for (const lvl of getChildrenNS(a, NS.w, "lvl")) {
			const ilvl = parseInt(wAttr(lvl, "ilvl") || "0", 10);
			const numFmt = firstChildNS(lvl, NS.w, "numFmt");
			const lvlText = firstChildNS(lvl, NS.w, "lvlText");
			const start = firstChildNS(lvl, NS.w, "start");
			const pStyle = firstChildNS(lvl, NS.w, "pStyle");
			lvls.push({
				level: ilvl,
				format: numFmt && wVal(numFmt) || "decimal",
				text: lvlText && wVal(lvlText) || "",
				pStyle: pStyle ? wVal(pStyle) || void 0 : void 0,
				start: start ? parseInt(wVal(start) || "1", 10) : 1
			});
		}
		abstractMap.set(id, lvls);
	}
	const out = [];
	for (const n of getChildrenNS(root, NS.w, "num")) {
		const id = wAttr(n, "numId") || "";
		const absRef = firstChildNS(n, NS.w, "abstractNumId");
		const absId = absRef && wVal(absRef) || "";
		const lvls = abstractMap.get(absId) || [];
		out.push({
			numId: id,
			abstractNumId: absId,
			levels: lvls
		});
	}
	return out;
}

//#endregion
export { parseNumbering as n, loadDocx as t };