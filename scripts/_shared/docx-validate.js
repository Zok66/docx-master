import { _ as parseXml, a as firstChildNS, f as wAttr, g as DocxReader, h as NS, o as getChildren, s as getChildrenNS } from "./xml-utils.js";
import { dirname, join } from "node:path";
import { existsSync, readFileSync, readdirSync } from "node:fs";
import { fileURLToPath, pathToFileURL } from "node:url";

//#region lib/shared/docx-validate.ts
/**
* OOXML validator backed by xmllint-wasm (libxml2 compiled to WebAssembly,
* the same engine the Anthropic docx skill uses via Python lxml). Provides
* full ECMA-376 XSD validation — element ordering, types, enumerations,
* minOccurs/maxOccurs, attribute constraints — without Python or native
* deps. Schemas are bundled under `<scripts>/_shared/schemas/` at build
* time.
*
* Two layers:
*   - **XSD validation** (`validateXsd`): per-part schema check via
*     xmllint-wasm. Catches everything the schema covers.
*   - **Cross-part / package-level checks** (`checkCrossPartRefs`,
*     `checkContentTypes`, `checkRelationships`, `checkMultiLevelType`):
*     reference integrity and Word-specific quirks XSD doesn't enforce.
*     Examples: a `<w:numId>` referenced from document.xml must exist in
*     numbering.xml; a multi-level abstractNum without `<w:multiLevelType>`
*     silently fails to render in Word even though XSD permits its absence.
*
* Public API:
*   - `validateDocxFile(path)` — async, opens a docx and validates every
*     part. Used by `apply` post-write and the standalone `validate` CLI.
*/
const W = NS.w;
const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types";
const MODULE_DIR = dirname(fileURLToPath(import.meta.url));
function findSchemasDir() {
	const candidates = [join(MODULE_DIR, "schemas"), join(MODULE_DIR, "..", "..", "vendor", "ooxml-schemas")];
	for (const p of candidates) if (existsSync(join(p, "ISO-IEC29500-4_2016", "wml.xsd"))) return p;
	throw new Error(`OOXML XSD schemas not found. Looked in: ${candidates.join(", ")}. Run \`bun run build:skill\` to stage them under dist/.`);
}
/** xmllint-wasm can't be statically bundled by tsdown — its package entry
* does runtime `require("./xmllint-node.js")` and references `xmllint.wasm`
* by relative path. We mark it `external` in tsdown.config and copy the
* package's runtime files into the bundled output under `_shared/
* xmllint-wasm/`. At runtime we try the bare specifier first (works in dev
* via node_modules) and fall back to the colocated path (built bundle). */
let cachedValidateXML = null;
async function getValidateXML() {
	if (cachedValidateXML) return cachedValidateXML;
	const candidates = ["xmllint-wasm", pathToFileURL(join(MODULE_DIR, "xmllint-wasm", "index-node.js")).href];
	let lastErr = null;
	for (const spec of candidates) try {
		cachedValidateXML = (await import(spec)).validateXML;
		return cachedValidateXML;
	} catch (err) {
		lastErr = err;
	}
	throw new Error(`xmllint-wasm runtime not found. Looked for: ${candidates.join(", ")}. Last error: ${lastErr?.message ?? "(unknown)"}.`);
}
let cachedSchemas = null;
/**
* Loaded XSDs are placed at flat paths (basename only) in xmllint-wasm's
* in-memory FS, because Emscripten's MEMFS doesn't auto-create intermediate
* directories on file write — nested paths like `ecma/fouth-edition/x.xsd`
* fail with `ErrnoError 44 (ENOENT)`. Each schema's `xs:import` /
* `xs:include` `schemaLocation` is rewritten to its basename to keep the
* cross-schema reference graph intact under flattening.
*
* All 39 schemas have unique basenames so flattening is collision-free.
*/
function getSchemas() {
	if (cachedSchemas) return cachedSchemas;
	const root = findSchemasDir();
	const out = [];
	const walk = (dir) => {
		for (const ent of readdirSync(dir, { withFileTypes: true })) {
			const p = join(dir, ent.name);
			if (ent.isDirectory()) walk(p);
			else if (ent.name.endsWith(".xsd")) {
				let contents = readFileSync(p, "utf-8");
				contents = contents.replace(/schemaLocation="([^"]+)"/g, (_, loc) => {
					return `schemaLocation="${loc.split("/").pop() ?? loc}"`;
				});
				contents = contents.replace(/(<xs(?:d)?:import\b[^>]*?\bnamespace="http:\/\/www\.w3\.org\/XML\/1998\/namespace")(?![^>]*\bschemaLocation=)([^>]*?\/>)/g, `$1 schemaLocation="xml.xsd"$2`);
				out.push({
					fileName: ent.name,
					contents
				});
			}
		}
	};
	walk(root);
	cachedSchemas = out;
	return out;
}
/**
* Pick the main XSD to validate a given part against. Schemas are flattened
* to basename in MEMFS, so the picker returns basenames too.
*/
function pickSchema(partPath) {
	if (partPath === "[Content_Types].xml") return "opc-contentTypes.xsd";
	if (partPath.endsWith(".rels")) return "opc-relationships.xsd";
	if (partPath.startsWith("docProps/")) {
		if (partPath.endsWith("/app.xml")) return "shared-documentPropertiesExtended.xsd";
		if (partPath.endsWith("/core.xml")) return "opc-coreProperties.xsd";
		if (partPath.endsWith("/custom.xml")) return "shared-documentPropertiesCustom.xsd";
		return null;
	}
	if (partPath.startsWith("word/")) {
		if (partPath.includes("/theme/")) return "dml-main.xsd";
		if (partPath.endsWith("/sources.xml")) return "shared-bibliography.xsd";
		if (partPath.endsWith("/people.xml") || /\/comments(Extended|Ids|Extensible)\.xml$/.test(partPath)) return null;
		return "wml.xsd";
	}
	return null;
}
/**
* Errors libxml2 emits on every OOXML doc because the schemas don't cover
* a few late-added Microsoft extensions and Dublin Core terms. The official
* Anthropic validator filters the same list — see scripts/office/validators/
* base.py:IGNORED_VALIDATION_ERRORS.
*/
const IGNORED_XSD_ERROR_PATTERNS = [/hyphenationZone/i, /purl\.org\/dc\//i];
/**
* Namespaces the ECMA-376 core schemas know about. Attributes / elements in
* any other namespace (Microsoft extensions w14 / w15 / w16, mc:Ignorable,
* etc.) are stripped before validation — same approach as the official
* Anthropic validator's `_clean_ignorable_namespaces`. Without this, a
* vanilla Word doc emits hundreds of "attribute not allowed" errors against
* the core schemas because Word adds the extension attrs by default.
*/
const OOXML_NAMESPACES = new Set([
	"http://schemas.openxmlformats.org/officeDocument/2006/math",
	"http://schemas.openxmlformats.org/officeDocument/2006/relationships",
	"http://schemas.openxmlformats.org/schemaLibrary/2006/main",
	"http://schemas.openxmlformats.org/drawingml/2006/main",
	"http://schemas.openxmlformats.org/drawingml/2006/chart",
	"http://schemas.openxmlformats.org/drawingml/2006/chartDrawing",
	"http://schemas.openxmlformats.org/drawingml/2006/diagram",
	"http://schemas.openxmlformats.org/drawingml/2006/picture",
	"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
	"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
	"http://schemas.openxmlformats.org/wordprocessingml/2006/main",
	"http://schemas.openxmlformats.org/presentationml/2006/main",
	"http://schemas.openxmlformats.org/spreadsheetml/2006/main",
	"http://schemas.openxmlformats.org/officeDocument/2006/sharedTypes",
	"http://www.w3.org/XML/1998/namespace",
	"urn:schemas-microsoft-com:vml",
	"urn:schemas-microsoft-com:office:office",
	"http://schemas.openxmlformats.org/officeDocument/2006/bibliography",
	"http://schemas.openxmlformats.org/officeDocument/2006/customXml",
	"http://schemas.openxmlformats.org/officeDocument/2006/customXmlProps"
]);
/**
* Strip MC-extension attributes and elements (anything whose namespace
* isn't in OOXML_NAMESPACES) from the XML before validating. Returns the
* cleaned string. Operates via regex on the serialized XML — full DOM
* round-trip is too slow when called per validation across many parts.
*/
/**
* Strip MC-extension attributes and elements (anything whose namespace
* isn't in OOXML_NAMESPACES) from the XML before validating. Returns the
* cleaned string.
*
* Implementation notes (rewritten — the previous regex version corrupted
* any document containing `wps:` / `wpg:` / `wpc:` / VML content):
*
*   - Prefixes are collected once into a Set. The old version kept a list
*     with one entry per `xmlns:` occurrence (dozens of duplicates), so the
*     same destructive replace ran again and again over already-mangled XML.
*   - Elements are removed by **tag scan with a nesting counter**, not by
*     regex. The old `<p:x>[\s\S]*?</p:y>` pattern used a greedy body and
*     didn't require the closing tag to match the opening one, so a
*     `<wps:wsp>…</wps:wsp>` run could swallow siblings and leave orphan
*     end tags — the XML then failed to parse at all ("Opening and ending
*     tag mismatch"), which is what surfaced as 12 bogus errors on any
*     document carrying a VML watermark.
*   - Text inside a removed element is dropped with it (that's the point:
*     validate the core-schema view of the part).
*   - Attribute stripping is quote-aware so `>` inside an attribute value
*     can't terminate the tag early.
*
* VML (`urn:schemas-microsoft-com:vml`) and the `office` namespace are in
* OOXML_NAMESPACES because `wml.xsd` actually imports them — watermarks,
* OLE previews and legacy shapes validate against the real schema instead
* of being deleted first.
*/
function stripMcExtensions(xml) {
	const extPrefixes = /* @__PURE__ */ new Set();
	for (const m of xml.matchAll(/xmlns:([\w.-]+)="([^"]+)"/g)) if (!OOXML_NAMESPACES.has(m[2] ?? "")) extPrefixes.add(m[1]);
	if (extPrefixes.size === 0) return xml;
	const isExtension = (name) => {
		const colon = name.indexOf(":");
		return colon > 0 && extPrefixes.has(name.slice(0, colon));
	};
	/** Drop `prefix:attr="…"` where prefix is an extension prefix. `xmlns:*`
	*  declarations survive (their prefix is `xmlns`, never an ext prefix). */
	const stripExtAttrs = (attrs) => extPrefixes.size === 0 || !attrs ? attrs : attrs.replace(/([\s])?([\w.-]+):([\w.-]+)\s*=\s*(?:"[^"]*"|'[^']*')/g, (full, lead, prefix) => extPrefixes.has(prefix) ? "" : full);
	const tagRe = /<(\/?)([\w.:-]+)((?:"[^"]*"|'[^']*'|[^>"'])*?)(\/?)>/g;
	let out = "";
	let cursor = 0;
	let skip = 0;
	let m;
	while ((m = tagRe.exec(xml)) !== null) {
		const [full, slash, name, attrs, selfClose] = m;
		const isClose = slash === "/";
		const ext = isExtension(name);
		if (skip > 0) {
			if (ext && !isClose && !selfClose) skip += 1;
			else if (ext && isClose) {
				skip -= 1;
				if (skip === 0) cursor = m.index + full.length;
			}
			continue;
		}
		if (ext) {
			if (isClose) continue;
			if (selfClose) {
				out += xml.slice(cursor, m.index);
				cursor = m.index + full.length;
				continue;
			}
			skip = 1;
			continue;
		}
		out += xml.slice(cursor, m.index) + "<" + slash + name + stripExtAttrs(attrs) + selfClose + ">";
		cursor = m.index + full.length;
	}
	out += xml.slice(cursor);
	return out.replace(/\s[\w.-]+:Ignorable\s*=\s*"[^"]*"/g, "");
}
async function validateXsd(parts) {
	const errors = [];
	const schemas = getSchemas();
	const schemaMap = new Map(schemas.map((s) => [s.fileName, s]));
	const groups = /* @__PURE__ */ new Map();
	for (const [partPath, contents] of parts) {
		const schemaPath = pickSchema(partPath);
		if (!schemaPath) continue;
		if (!schemaMap.has(schemaPath)) continue;
		if (!groups.has(schemaPath)) groups.set(schemaPath, []);
		groups.get(schemaPath).push({
			partPath,
			contents
		});
	}
	for (const [schemaPath, files] of groups) {
		const mainSchema = schemaMap.get(schemaPath);
		const preload = schemas.filter((s) => s.fileName !== schemaPath);
		const inputMap = /* @__PURE__ */ new Map();
		const xmlInputs = files.map((f, i) => {
			const token = `input_${i}.xml`;
			inputMap.set(token, f.partPath);
			return {
				fileName: token,
				contents: /^(word|ppt|xl)\//.test(f.partPath) ? stripMcExtensions(f.contents) : f.contents
			};
		});
		let result;
		try {
			result = await (await getValidateXML())({
				xml: xmlInputs,
				schema: [mainSchema],
				preload
			});
		} catch (err) {
			const detail = err instanceof Error ? `${err.message}${err.stack ? `\n${err.stack.split("\n").slice(0, 3).join("\n")}` : ""}` : `(non-Error throw: ${typeof err}) ${JSON.stringify(err)}`;
			errors.push({
				part: `(xmllint)`,
				message: `validator crashed for schema ${schemaPath}: ${detail}`
			});
			continue;
		}
		if (result.valid) continue;
		for (const e of result.errors) {
			const text = e.message;
			if (IGNORED_XSD_ERROR_PATTERNS.some((re) => re.test(text))) continue;
			const reportedFile = e.loc?.fileName ?? "";
			const part = inputMap.get(reportedFile) ?? reportedFile ?? schemaPath;
			errors.push({
				part,
				message: text
			});
		}
	}
	return errors;
}
/** Word built-in styleIds that resolve via docDefaults / latentStyles even
* without an explicit `<w:style>` entry. */
const WORD_BUILTIN_STYLE_IDS = new Set([
	"Normal",
	"NoList",
	"TableNormal",
	"DefaultParagraphFont"
]);
/**
* `<w:multiLevelType>` is optional per CT_AbstractNum schema, but Word
* silently fails to render numbering on multi-level schemes that omit it.
* XSD won't flag this — we have to.
*/
function checkMultiLevelType(numberingXml) {
	if (!numberingXml) return [];
	const doc = parseXmlOrNull(numberingXml);
	if (!doc?.documentElement) return [];
	const errors = [];
	for (const abs of getChildrenNS(doc.documentElement, W, "abstractNum")) {
		const aid = wAttr(abs, "abstractNumId") ?? "(unknown)";
		const levels = getChildrenNS(abs, W, "lvl");
		if (!firstChildNS(abs, W, "multiLevelType") && levels.length > 1) errors.push({
			part: "word/numbering.xml",
			message: `abstractNum ${aid}: missing <w:multiLevelType> on a ${levels.length}-level scheme — Word silently fails to render numbering without it.`
		});
	}
	return errors;
}
/**
* Every `<w:numId>` referenced from document.xml must exist in
* numbering.xml; every `<w:pStyle>` that numbering.xml binds must exist in
* styles.xml (excepting Word built-ins). pStyle references in document.xml
* itself are not flagged — Word falls back to Normal gracefully.
*/
function checkCrossPartRefs(parts) {
	const errors = [];
	const stylesDoc = parseXmlOrNull(parts.get("word/styles.xml") ?? null);
	const numberingDoc = parseXmlOrNull(parts.get("word/numbering.xml") ?? null);
	const documentDoc = parseXmlOrNull(parts.get("word/document.xml") ?? null);
	const knownStyleIds = /* @__PURE__ */ new Set();
	if (stylesDoc?.documentElement) for (const s of getChildrenNS(stylesDoc.documentElement, W, "style")) {
		const sid = wAttr(s, "styleId");
		if (sid) knownStyleIds.add(sid);
	}
	const knownNumIds = /* @__PURE__ */ new Set();
	if (numberingDoc?.documentElement) for (const n of getChildrenNS(numberingDoc.documentElement, W, "num")) {
		const nid = wAttr(n, "numId");
		if (nid) knownNumIds.add(nid);
	}
	if (documentDoc?.documentElement && knownNumIds.size > 0) {
		let pIdx = 0;
		const visit = (el) => {
			for (const child of getChildren(el)) {
				if (child.namespaceURI === W && child.localName === "p") pIdx += 1;
				if (child.namespaceURI === W && child.localName === "numId") {
					const nid = wAttr(child, "val");
					if (nid && !knownNumIds.has(nid)) errors.push({
						part: "word/document.xml",
						message: `paragraph #${pIdx}: numId="${nid}" doesn't exist in numbering.xml.`
					});
				}
				visit(child);
			}
		};
		visit(documentDoc.documentElement);
	}
	if (numberingDoc?.documentElement && knownStyleIds.size > 0) for (const abs of getChildrenNS(numberingDoc.documentElement, W, "abstractNum")) {
		const aid = wAttr(abs, "abstractNumId");
		for (const lvl of getChildrenNS(abs, W, "lvl")) {
			const pStyle = firstChildNS(lvl, W, "pStyle");
			if (!pStyle) continue;
			const sid = wAttr(pStyle, "val");
			if (sid && !knownStyleIds.has(sid) && !WORD_BUILTIN_STYLE_IDS.has(sid)) errors.push({
				part: "word/numbering.xml",
				message: `abstractNum ${aid} lvl ${wAttr(lvl, "ilvl") ?? "?"}: pStyle references "${sid}", which doesn't exist in styles.xml — the style→numbering binding will silently fail.`
			});
		}
	}
	return errors;
}
/** Every part in the package must have either an Override or a Default
* content type entry; every Override must point at an existing part. */
function checkContentTypes(parts, packageEntries) {
	const errors = [];
	const ct = parts.get("[Content_Types].xml");
	if (!ct) return [{
		part: "[Content_Types].xml",
		message: "missing"
	}];
	const ctDoc = parseXmlOrNull(ct);
	if (!ctDoc?.documentElement) return [{
		part: "[Content_Types].xml",
		message: "unparseable"
	}];
	const overrides = /* @__PURE__ */ new Set();
	for (const ov of getChildrenNS(ctDoc.documentElement, CT_NS, "Override")) {
		const partName = ov.getAttribute("PartName");
		if (partName) overrides.add(partName.replace(/^\//, ""));
	}
	const defaults = /* @__PURE__ */ new Set();
	for (const def of getChildrenNS(ctDoc.documentElement, CT_NS, "Default")) {
		const ext = def.getAttribute("Extension");
		if (ext) defaults.add(ext.toLowerCase());
	}
	const declared = (path) => {
		if (overrides.has(path)) return true;
		const ext = path.split(".").pop()?.toLowerCase();
		return ext ? defaults.has(ext) : false;
	};
	for (const path of packageEntries) {
		if (path.endsWith("/")) continue;
		if (path === "[Content_Types].xml") continue;
		if (!declared(path)) errors.push({
			part: "[Content_Types].xml",
			message: `part "${path}" has neither Override nor Default content type — Word will reject the package.`
		});
	}
	for (const ov of overrides) if (!packageEntries.includes(ov)) errors.push({
		part: "[Content_Types].xml",
		message: `Override targets "/${ov}" but no such part exists in the package.`
	});
	return errors;
}
/** Every `r:id`/`r:embed`/`r:link` attribute used in an XML part must
* resolve to a Relationship in the corresponding .rels file. */
function checkRelationships(parts, packageEntries) {
	const errors = [];
	const entrySet = new Set(packageEntries);
	for (const relsPath of packageEntries) {
		if (!relsPath.endsWith(".rels")) continue;
		const text = parts.get(relsPath);
		if (!text) continue;
		const relsDoc = parseXmlOrNull(text);
		if (!relsDoc?.documentElement) continue;
		const ridSeen = /* @__PURE__ */ new Set();
		const ridToTargetMode = /* @__PURE__ */ new Map();
		const ridToTarget = /* @__PURE__ */ new Map();
		for (const rel of getChildrenNS(relsDoc.documentElement, PKG_REL_NS, "Relationship")) {
			const rid = rel.getAttribute("Id");
			const target = rel.getAttribute("Target") ?? "";
			const mode = rel.getAttribute("TargetMode") ?? "Internal";
			if (!rid) continue;
			if (ridSeen.has(rid)) errors.push({
				part: relsPath,
				message: `duplicate relationship Id "${rid}".`
			});
			ridSeen.add(rid);
			ridToTarget.set(rid, target);
			ridToTargetMode.set(rid, mode);
		}
		const baseDir = relsBaseDir(relsPath);
		for (const [rid, target] of ridToTarget) {
			if (ridToTargetMode.get(rid) === "External") continue;
			const resolved = resolveRelativePath(baseDir, target);
			if (!entrySet.has(resolved)) errors.push({
				part: relsPath,
				message: `relationship "${rid}" Target "${target}" → "${resolved}" doesn't exist in the package.`
			});
		}
	}
	for (const xmlPath of packageEntries) {
		if (!xmlPath.endsWith(".xml")) continue;
		const relsPath = relsForPart(xmlPath);
		if (!relsPath || !entrySet.has(relsPath)) continue;
		const relsDoc = parseXmlOrNull(parts.get(relsPath) ?? null);
		if (!relsDoc?.documentElement) continue;
		const knownRids = /* @__PURE__ */ new Set();
		for (const rel of getChildrenNS(relsDoc.documentElement, PKG_REL_NS, "Relationship")) {
			const rid = rel.getAttribute("Id");
			if (rid) knownRids.add(rid);
		}
		const xmlDoc = parseXmlOrNull(parts.get(xmlPath) ?? null);
		if (!xmlDoc?.documentElement) continue;
		const visit = (el) => {
			for (const attrName of [
				"id",
				"embed",
				"link"
			]) {
				const ridAttr = el.getAttributeNS?.(R_NS, attrName) ?? null;
				if (ridAttr && !knownRids.has(ridAttr)) errors.push({
					part: xmlPath,
					message: `<${el.localName}> r:${attrName}="${ridAttr}" references a non-existent relationship in ${relsPath}.`
				});
			}
			for (const child of getChildren(el)) visit(child);
		};
		visit(xmlDoc.documentElement);
	}
	return errors;
}
async function validateOoxmlParts(parts, packageEntries) {
	const errors = [];
	errors.push(...await validateXsd(parts));
	errors.push(...checkMultiLevelType(parts.get("word/numbering.xml") ?? null));
	errors.push(...checkCrossPartRefs(parts));
	if (packageEntries) {
		errors.push(...checkContentTypes(parts, packageEntries));
		errors.push(...checkRelationships(parts, packageEntries));
	}
	return errors;
}
async function validateDocxFile(filePath) {
	const reader = await DocxReader.open(filePath);
	const entries = reader.listEntries();
	const parts = /* @__PURE__ */ new Map();
	for (const e of entries) {
		if (!e.endsWith(".xml") && !e.endsWith(".rels")) continue;
		const text = await reader.readText(e);
		if (text !== null) parts.set(e, text);
	}
	return validateOoxmlParts(parts, entries);
}
/** Validate an `<m:oMath>...</m:oMath>` fragment against shared-math.xsd.
*  Defense in depth against schema-invalid converter output before it
*  reaches document.xml. Returns one error per validator complaint;
*  empty array = valid. */
async function validateOMath(omml) {
	const schemas = getSchemas();
	const main = schemas.find((s) => s.fileName === "shared-math.xsd");
	if (!main) return [`shared-math.xsd not found in bundled schemas`];
	const preload = schemas.filter((s) => s.fileName !== "shared-math.xsd");
	let result;
	try {
		result = await (await getValidateXML())({
			xml: [{
				fileName: "math.xml",
				contents: omml
			}],
			schema: [main],
			preload
		});
	} catch (err) {
		return [`validator crashed: ${err?.message ?? String(err)}`];
	}
	if (result.valid) return [];
	const out = [];
	for (const e of result.errors) {
		if (IGNORED_XSD_ERROR_PATTERNS.some((re) => re.test(e.message))) continue;
		out.push(e.message);
	}
	return out;
}
function relsForPart(partPath) {
	if (partPath.endsWith(".rels")) return null;
	const slash = partPath.lastIndexOf("/");
	if (slash < 0) return `_rels/${partPath}.rels`;
	return `${partPath.slice(0, slash)}/_rels/${partPath.slice(slash + 1)}.rels`;
}
/**
* Directory that a part's Relationships resolve against.
* OPC resolves a Target against the SOURCE PART's directory — never against the
* `.rels` file's own directory. Getting this wrong breaks every side-car rels
* part (`customXml/item1.xml.rels`), where the rels file and its source part
* live in the same folder.
*   "_rels/.rels"                  -> ""       (package root)
*   "word/_rels/document.xml.rels" -> "word"
*   "customXml/item1.xml.rels"     -> "customXml"
*   "customXml/_rels/item1.xml.rels" -> "customXml"
*/
function relsBaseDir(relsPath) {
	const slash = relsPath.lastIndexOf("/");
	const dir = slash < 0 ? "" : relsPath.slice(0, slash);
	if (dir === "_rels") return "";
	if (dir.endsWith("/_rels")) return dir.slice(0, -"/_rels".length);
	return dir;
}
function resolveRelativePath(baseDir, target) {
	if (target.startsWith("/")) return target.slice(1);
	const segs = (baseDir ? `${baseDir}/${target}` : target).split("/");
	const out = [];
	for (const s of segs) {
		if (s === "." || s === "") continue;
		if (s === "..") out.pop();
		else out.push(s);
	}
	return out.join("/");
}
function parseXmlOrNull(text) {
	if (!text) return null;
	try {
		return parseXml(text);
	} catch {
		return null;
	}
}

//#endregion
export { validateOMath as n, validateDocxFile as t };