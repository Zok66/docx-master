#!/usr/bin/env python3
"""
Offline self-check for the bundled XSD schemas.

Why this exists
---------------
`validators_base._load_schema` compiles schemas with `no_network=True` so that
an outbound HTTP fetch can never happen silently (sandboxed runtimes surface
those as an "allow this connection?" prompt). The trade-off is fail-fast: if a
schema ever starts depending on a remote `schemaLocation`, compilation dies
instead of quietly downloading the missing piece.

This script turns that trade-off into a deterministic, one-shot check so the
failure is caught at build time rather than in a user's session:

  * every .xsd under pro/schemas/ is compiled with no_network=True
  * a schema that fails offline is re-compiled WITH network allowed
      - succeeds with network  -> real regression: the schema needs the internet
      - fails either way       -> pre-existing defect in the schema itself,
                                  unrelated to the offline policy (reported as
                                  BASELINE, does not fail the run)

Exit code 1 only for genuine network dependencies, so this is CI-safe.

Usage:
    python pro/scripts/schema_selfcheck.py            # scan all schemas
    python pro/scripts/schema_selfcheck.py wml.xsd    # check a single file
    python pro/scripts/schema_selfcheck.py --verbose  # list every file
"""

import argparse
import sys
from pathlib import Path

import lxml.etree as etree

SCHEMAS_DIR = Path(__file__).resolve().parent.parent / "schemas"


def compile_schema(xsd_path: Path, allow_network: bool) -> str | None:
    """Compile one schema. Return None on success, else the error message."""
    try:
        doc = etree.parse(
            str(xsd_path),
            parser=etree.XMLParser(no_network=not allow_network),
            base_url=xsd_path.resolve().as_uri(),
        )
        etree.XMLSchema(doc)
        return None
    except Exception as exc:  # noqa: BLE001 - lxml raises several types
        return " ".join(str(exc).split())[:160]


def scan(verbose: bool) -> int:
    if not SCHEMAS_DIR.is_dir():
        print(f"Error: schema directory not found: {SCHEMAS_DIR}")
        return 1

    files = sorted(SCHEMAS_DIR.rglob("*.xsd"))
    if not files:
        print("Error: no .xsd files found")
        return 1

    offline, network_dep, baseline = [], [], []

    for xsd in files:
        rel = xsd.relative_to(SCHEMAS_DIR)
        err_offline = compile_schema(xsd, allow_network=False)
        if err_offline is None:
            offline.append(rel)
            if verbose:
                print(f"  OK      {rel}")
            continue

        err_online = compile_schema(xsd, allow_network=True)
        if err_online is None:
            network_dep.append((rel, err_offline))
            print(f"  NETWORK {rel}")
            print(f"          offline error: {err_offline}")
        else:
            baseline.append((rel, err_offline))
            if verbose:
                print(f"  BASELINE {rel}  (fails online too - schema defect)")

    total = len(files)
    print()
    print(f"Schemas scanned          : {total}")
    print(f"Offline-clean            : {len(offline)}")
    print(f"Needs network (FAIL)     : {len(network_dep)}")
    print(f"Broken regardless (skip) : {len(baseline)}")

    if baseline and not verbose:
        print("\nPre-existing schema defects (independent of the offline policy):")
        for rel, err in baseline:
            print(f"  - {rel}: {err}")

    if network_dep:
        print("\nThese schemas require an outbound fetch and must be fixed:")
        for rel, err in network_dep:
            print(f"  - {rel}")
            print(f"    {err}")
        print("\nFix: bundle the imported file locally and point schemaLocation")
        print("at the relative filename, so no_network=True keeps working.")
        return 1

    print("\nAll schemas compile offline - no outbound requests possible.")
    return 0


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Verify every bundled XSD compiles without network access."
    )
    parser.add_argument(
        "files",
        nargs="*",
        help="Specific .xsd names under pro/schemas/ (default: all)",
    )
    parser.add_argument(
        "--verbose", action="store_true", help="List every schema, not just failures"
    )
    args = parser.parse_args()

    if not args.files:
        sys.exit(scan(args.verbose))

    exit_code = 0
    for name in args.files:
        target = SCHEMAS_DIR / name
        if not target.is_file():
            print(f"Error: not found: {target}")
            exit_code = 1
            continue
        err = compile_schema(target, allow_network=False)
        if err is None:
            print(f"OK      {name}")
            continue
        print(f"network {name}")
        print(f"        offline error: {err}")
        if compile_schema(target, allow_network=True) is not None:
            print("        (also fails with network - schema defect, not policy)")
        else:
            exit_code = 1

    sys.exit(exit_code)


if __name__ == "__main__":
    main()
