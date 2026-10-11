import { t as validateDocxFile } from "./_shared/docx-validate.js";

//#region skill/tools/validate.ts
/**
* `validate` — comprehensive OOXML validation of a .docx file.
*
* Modeled after the official Anthropic docx skill's validator: schema
* element-ordering, whitespace preservation, cross-part style/numbering
* references, content-types completeness, relationship Id integrity, ID
* uniqueness. The XSD-based check the official skill runs is the only thing
* this can't replicate (no good pure-TS XSD engine for OOXML's multi-schema
* imports); the CT_* ordering rules cover the most common XSD violations
* Word actually rejects.
*
* `apply` runs this automatically on every write. This standalone CLI is for
* spot-checking arbitrary .docx files (a template the user gave you, an
* output you suspect is corrupt, etc.).
*
* Exit code 0 on clean, 1 on errors.
*/
async function main() {
	const file = process.argv[2];
	if (!file) {
		console.error("Usage: node scripts/validate.js <docx-path>");
		process.exit(1);
	}
	try {
		const errors = await validateDocxFile(file);
		if (errors.length === 0) {
			console.log("All validations PASSED");
			return;
		}
		console.error(`Found ${errors.length} validation error(s):`);
		for (const e of errors) console.error(`  ${e.part}: ${e.message}`);
		process.exit(1);
	} catch (err) {
		console.error(`Error: ${err.message}`);
		process.exit(1);
	}
}
main();

//#endregion
export {  };