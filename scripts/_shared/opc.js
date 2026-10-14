/**
 * Minimal OPC (zip) reader — no dependencies beyond Node's own zlib.
 *
 * An OOXML file is a zip whose entries are mostly UTF-8 XML. Every diagnostic
 * tool in this skill needs exactly that much: enumerate the parts, pull one out
 * as text. Reading the central directory by hand avoids a dependency on a zip
 * library (a skill should not require an install to inspect a file), and keeps
 * the behaviour identical across platforms and Node versions.
 *
 * Read-only by design. Nothing here writes, and `apply` / `extend` keep using
 * their own writers.
 */
import { readFileSync } from "node:fs";
import { inflateRawSync } from "node:zlib";

const EOCD = 0x06054b50;
const CEN = 0x02014b50;
const LOC = 0x04034b50;

/**
 * Open a package and return a reader:
 *   { buf, names, has(name), readText(name) }
 * `readText` returns null for a missing part (callers branch on that) and
 * throws on a corrupt local header, which means the file itself is broken.
 */
export function openZip(path) {
	const buf = readFileSync(path);
	// The end-of-central-directory record sits at the end, after a comment of
	// unknown length (up to 64 KiB) — scan backwards for its signature.
	let eocd = -1;
	const floor = Math.max(0, buf.length - 66000);
	for (let i = buf.length - 22; i >= floor; i--) {
		if (buf.readUInt32LE(i) === EOCD) {
			eocd = i;
			break;
		}
	}
	if (eocd < 0) throw new Error("not a zip/OPC package (no end-of-central-directory record)");
	const count = buf.readUInt16LE(eocd + 10);
	let off = buf.readUInt32LE(eocd + 16);
	const entries = new Map();
	for (let i = 0; i < count; i++) {
		if (off + 46 > buf.length || buf.readUInt32LE(off) !== CEN) break;
		const method = buf.readUInt16LE(off + 10);
		const compSize = buf.readUInt32LE(off + 20);
		const nameLen = buf.readUInt16LE(off + 28);
		const extraLen = buf.readUInt16LE(off + 30);
		const commentLen = buf.readUInt16LE(off + 32);
		const localOff = buf.readUInt32LE(off + 42);
		const name = buf.toString("utf8", off + 46, off + 46 + nameLen);
		entries.set(name, { method, compSize, localOff });
		off += 46 + nameLen + extraLen + commentLen;
	}
	return {
		buf,
		names: [...entries.keys()],
		has: (n) => entries.has(n),
		readText(name) {
			const e = entries.get(name);
			if (!e) return null;
			if (buf.readUInt32LE(e.localOff) !== LOC) throw new Error(`corrupt local header for ${name}`);
			const nLen = buf.readUInt16LE(e.localOff + 26);
			const eLen = buf.readUInt16LE(e.localOff + 28);
			const start = e.localOff + 30 + nLen + eLen;
			const raw = buf.subarray(start, start + e.compSize);
			const out = e.method === 0 ? raw : inflateRawSync(raw);
			return out.toString("utf8");
		}
	};
}

/**
 * Every part that can legally carry a bookmark or a reference to one.
 *
 * Comments and textboxes live inside `document.xml` (comments can also sit in
 * `comments.xml`), so they need no separate handling. Headers / footers are
 * resolved through the document's rels, with a conventional-name sweep as
 * fallback — a bookmark defined in a footer is still referenceable from the
 * body, and a reference can just as easily live in a header.
 */
export function referenceBearingParts(zip) {
	const names = ["word/document.xml"];
	const seen = new Set(names);
	const add = (n) => {
		if (!n || seen.has(n)) return;
		seen.add(n);
		names.push(n);
	};

	const rels = zip.readText("word/_rels/document.xml.rels") || "";
	for (const m of rels.matchAll(/<Relationship\b[^>]*>/g)) {
		const type = (m[0].match(/Type="([^"]*)"/) || [])[1] || "";
		const target = (m[0].match(/Target="([^"]*)"/) || [])[1] || "";
		if (!target || /^https?:/.test(target)) continue;
		if (/\/(header|footer)$/.test(type)) add(`word/${target.replace(/^\/?word\//, "")}`);
	}
	for (const n of zip.names) {
		if (/^word\/(header|footer)\d*\.xml$/.test(n)) add(n);
	}
	for (const n of ["word/footnotes.xml", "word/endnotes.xml", "word/comments.xml"]) add(n);

	const parts = [];
	for (const name of names) {
		const xml = zip.readText(name);
		if (xml) parts.push({ name, xml });
	}
	return parts;
}
