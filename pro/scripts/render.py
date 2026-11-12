"""Portable document -> PDF -> image rendering, for the visual self-check loop.

Why this exists
---------------
`soffice.py` covers LibreOffice only. Plenty of machines have no LibreOffice but
do have Word or WPS, and plenty have neither. This module probes what is actually
present and reports the renderer's TRUE identity, then degrades explicitly rather
than silently.

Renderer chain (first that works wins)
-------------------------------------
1. ``soffice``        LibreOffice, cross-platform, reproducible, no GUI side effects.
2. ``word`` / ``wps`` Windows COM. Identity is decided by ``Application.Path`` --
                      WPS registers itself under ``Word.Application`` and returns
                      ``Name == "Microsoft Word"``, so ``Name`` cannot be trusted.
3. ``word-mac``       Word for macOS via AppleScript.
4. none               Exit 3 with install guidance. Never produce a fake "pass".

Image chain (first that works wins)
-----------------------------------
1. ``pymupdf``  (pip, cross-platform)
2. ``pdftoppm`` (poppler)
3. ``magick``   (ImageMagick + Ghostscript)
4. none        Keep the PDF only; the caller can still open it by hand.

The renderer-difference caveat is not decoration: a layout that passes under WPS
is not guaranteed to pass under Microsoft Word, or the reverse. Every report this
module emits carries ``renderer`` and ``renderer_identity`` so a conclusion can be
traced back to the engine that produced it.

Usage
-----
    python render.py REPORT.docx                     # PDF + PNG for every page
    python render.py REPORT.docx --pages 1,6-8       # selected pages only
    python render.py REPORT.docx --dpi 200 --out ./prev
    python render.py --probe                         # availability report, JSON
    python render.py --probe --human                 # same, as a readable table

Exit codes: 0 ok · 2 bad input · 3 no PDF renderer · 4 PDF produced, no image engine.
"""

from __future__ import annotations

import argparse
import base64
import json
import os
import platform
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

# Same-directory import: this module lives beside soffice.py.
sys.path.insert(0, str(Path(__file__).resolve().parent))
from soffice import find_soffice, run_soffice  # noqa: E402

IS_WINDOWS = os.name == "nt"
IS_MACOS = sys.platform == "darwin"

RENDER_WARNING = (
    "Visual conclusions hold only for the renderer that produced them. "
    "A layout that passes under WPS is not guaranteed to pass under Microsoft Word, "
    "or the reverse. Verify in the software your reader actually uses."
)


# --------------------------------------------------------------------------- #
# PDF renderer
# --------------------------------------------------------------------------- #

# The COM helper. Kept as one string and passed via -EncodedCommand (UTF-16LE
# base64) so quotes, spaces and non-ASCII paths never need shell escaping.
_PS_RENDER = r"""
$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}
$src = $env:DOCX_RENDER_SRC
$out = $env:DOCX_RENDER_OUT
$result = [ordered]@{
  ok = $false; progId = $null; name = $null; version = $null;
  build = $null; path = $null; kind = $null; error = $null
}

# No interactive desktop -> COM may hang instead of failing. Bail out cleanly.
if ([Environment]::UserInteractive -eq $false) {
  $result.error = 'no interactive desktop session; COM rendering skipped'
  $result | ConvertTo-Json -Compress
  exit 0
}

function Classify($p, $name, $build) {
  if ($p -and ($p -match 'Kingsoft|WPS|wps')) { return 'wps' }
  if ($p -and ($p -match 'Microsoft Office|Office\d+|WINWORD')) { return 'word' }
  # Path absent or unrecognised: fall back to build numbering.
  # Word is 16.x (Office 2016+); WPS reports 12.x or 11.x.
  if ($build -match '^(1[0-9])\.') { if ([int]$Matches[1] -ge 16) { return 'word' } else { return 'wps' } }
  if ($name -match 'WPS') { return 'wps' }
  return 'unknown'
}

foreach ($progId in @('Word.Application', 'KWPS.Application')) {
  if ($result.ok) { break }
  $app = $null; $doc = $null
  try {
    $app = New-Object -ComObject $progId -ErrorAction Stop
    try { $app.Visible = $false } catch {}
    try { $app.DisplayAlerts = 0 } catch {}
    $doc = $app.Documents.Open($src, $false, $true)   # ConfirmConversions=false, ReadOnly=true
    $doc.ExportAsFixedFormat($out, 17)                # wdExportFormatPDF
    $result.ok = $true
    $result.progId = $progId
    $result.name = "$($app.Name)"
    try { $result.version = "$($app.Version)" } catch {}
    try { $result.build = "$($app.Build)" } catch {}
    try { $result.path = "$($app.Path)" } catch {}
    $result.kind = Classify $result.path $result.name $result.build
  } catch {
    $result.error = "$progId : $($_.Exception.Message)"
  } finally {
    if ($doc) { try { $doc.Close(0) } catch {} }
    if ($app) { try { $app.Quit(0) } catch {} }
    if ($app) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app) } catch {} }
  }
}
$result | ConvertTo-Json -Compress
"""


def _ps_run(script: str, env_extra: dict, timeout: int) -> subprocess.CompletedProcess:
    """Run PowerShell. Output is read as bytes and decoded leniently on purpose:
    a zh-CN console emits GBK (cp936) for some cmdlets even when the script sets
    [Console]::OutputEncoding, and strict text-mode decoding crashes the reader
    thread before we ever see the result."""
    encoded = base64.b64encode(script.encode("utf-16-le")).decode("ascii")
    env = os.environ.copy()
    env.update(env_extra)
    proc = subprocess.run(
        ["powershell", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-EncodedCommand", encoded],
        capture_output=True,
        timeout=timeout,
        env=env,
    )
    for stream in ("stdout", "stderr"):
        raw = getattr(proc, stream) or b""
        if isinstance(raw, bytes):
            for codec in ("utf-8", "cp936", "latin-1"):
                try:
                    setattr(proc, stream, raw.decode(codec))
                    break
                except UnicodeDecodeError:
                    continue
    return proc


def _office_pids() -> set[int]:
    """PIDs of Word / WPS already running, so a timeout can reap only our own."""
    if not IS_WINDOWS:
        return set()
    try:
        r = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq winword.exe", "/FI", "IMAGENAME eq wps.exe", "/NH", "/FO", "CSV"],
            capture_output=True, timeout=20,
        )
        text = (r.stdout or b"").decode("cp936", errors="replace")
        pids = set()
        for line in text.splitlines():
            parts = [p.strip('"') for p in line.split('","')]
            if len(parts) >= 2 and parts[1].isdigit():
                pids.add(int(parts[1]))
        return pids
    except Exception:
        return set()


def _reap_new_office(before: set[int]) -> None:
    """Kill Word/WPS processes spawned by this run only — never the user's session."""
    if not IS_WINDOWS:
        return
    for pid in _office_pids() - before:
        try:
            subprocess.run(["taskkill", "/PID", str(pid), "/T", "/F"], capture_output=True, timeout=15)
        except Exception:
            pass


def _render_windows_com(src: Path, out: Path, timeout: int) -> dict:
    before = _office_pids()
    try:
        r = _ps_run(_PS_RENDER, {"DOCX_RENDER_SRC": str(src), "DOCX_RENDER_OUT": str(out)}, timeout)
    except subprocess.TimeoutExpired:
        _reap_new_office(before)
        return {
            "ok": False,
            "kind": "timeout",
            "error": f"COM rendering exceeded {timeout}s and was reaped "
                     "(a WPS/Office dialog may be blocking automation)",
        }
    except FileNotFoundError:
        return {"ok": False, "kind": "missing", "error": "powershell not found"}

    for line in reversed((r.stdout or "").splitlines()):
        line = line.strip()
        if not line.startswith("{"):
            continue
        try:
            data = json.loads(line)
        except json.JSONDecodeError:
            continue
        if data.get("ok") and out.exists() and out.stat().st_size > 0:
            return data
        return {
            "ok": False,
            "kind": data.get("kind") or "com",
            "error": data.get("error") or "COM export produced no file",
        }
    return {"ok": False, "kind": "com", "error": (r.stderr or "").strip()[:300] or "no output from COM helper"}


_AS_RENDER = """
on run argv
    set srcPath to item 1 of argv
    set outPath to item 2 of argv
    tell application "Microsoft Word"
        set theDoc to open file name srcPath with read only
        save as theDoc file name outPath file format format PDF
        close theDoc saving no
    end tell
    return "ok"
end run
"""


def _render_macos_word(src: Path, out: Path, timeout: int) -> dict:
    if not Path("/Applications/Microsoft Word.app").exists():
        return {"ok": False, "kind": "missing", "error": "Microsoft Word.app not installed"}
    try:
        r = subprocess.run(
            ["osascript", "-e", _AS_RENDER, str(src), str(out)],
            capture_output=True, text=True, timeout=timeout,
        )
    except subprocess.TimeoutExpired:
        return {"ok": False, "kind": "timeout", "error": f"AppleScript exceeded {timeout}s"}
    if r.returncode == 0 and out.exists() and out.stat().st_size > 0:
        return {"ok": True, "kind": "word-mac", "name": "Microsoft Word for Mac",
                "path": "/Applications/Microsoft Word.app", "progId": None, "error": None}
    return {"ok": False, "kind": "word-mac", "error": (r.stderr or "").strip()[:300]}


def _render_soffice(src: Path, out_dir: Path, timeout: int) -> dict:
    exe = find_soffice()
    if not exe:
        return {"ok": False, "kind": "missing", "error": "soffice not found"}
    try:
        r = run_soffice(
            ["--headless", "--norestore", "--convert-to", "pdf", "--outdir", str(out_dir), str(src)],
            capture_output=True, text=True, timeout=timeout,
        )
    except subprocess.TimeoutExpired:
        return {"ok": False, "kind": "timeout", "error": f"soffice exceeded {timeout}s"}
    produced = out_dir / (src.stem + ".pdf")
    if r.returncode == 0 and produced.exists() and produced.stat().st_size > 0:
        version = ""
        try:
            v = subprocess.run([exe, "--version"], capture_output=True, text=True, timeout=20)
            version = (v.stdout or v.stderr or "").strip().splitlines()[0]
        except Exception:
            pass
        return {"ok": True, "kind": "soffice", "name": version or "LibreOffice", "path": exe,
                "progId": None, "error": None}
    return {"ok": False, "kind": "soffice", "error": (r.stderr or r.stdout or "").strip()[:300]}


def render_pdf(src: Path, out_dir: Path, timeout: int = 120, prefer: str | None = None) -> dict:
    """Produce out_dir/<stem>.pdf. Returns a report dict; .ok is the gate."""
    out_dir.mkdir(parents=True, exist_ok=True)
    pdf = out_dir / (src.stem + ".pdf")

    chain: list[str] = []
    if prefer:
        chain.append(prefer)
    chain += [c for c in ("soffice", "com", "word-mac") if c != prefer]

    attempts: list[dict] = []
    for step in chain:
        if step == "soffice":
            res = _render_soffice(src, out_dir, timeout)
        elif step == "com":
            if not IS_WINDOWS:
                continue
            res = _render_windows_com(src, pdf, timeout)
        elif step == "word-mac":
            if not IS_MACOS:
                continue
            res = _render_macos_word(src, pdf, timeout)
        else:
            continue

        attempts.append({"step": step, "error": res.get("error")})
        if res.get("ok"):
            return {
                "ok": True,
                "pdf": str(pdf),
                "renderer": res.get("kind"),
                "renderer_identity": {
                    "name": res.get("name"),
                    "path": res.get("path"),
                    "version": res.get("version"),
                    "build": res.get("build"),
                    "progId": res.get("progId"),
                },
                "attempts": attempts,
                "warning": RENDER_WARNING,
            }

    return {
        "ok": False,
        "renderer": None,
        "attempts": attempts,
        "error": "no PDF renderer available",
        "hint": _install_hint(),
    }


# --------------------------------------------------------------------------- #
# PDF -> image
# --------------------------------------------------------------------------- #

def _parse_pages(spec: str | None, total: int) -> list[int]:
    if not spec:
        return list(range(total))
    picked: list[int] = []
    for chunk in spec.split(","):
        chunk = chunk.strip()
        if not chunk:
            continue
        if "-" in chunk:
            a, _, b = chunk.partition("-")
            lo = int(a) if a.strip() else 1
            hi = int(b) if b.strip() else total
            picked.extend(range(lo, hi + 1))
        else:
            picked.append(int(chunk))
    return sorted({p - 1 for p in picked if 1 <= p <= total})


def _find_magick() -> str | None:
    """ImageMagick only.

    On Windows ``shutil.which("convert")`` resolves to ``system32\\convert.exe``
    — the FAT->NTFS utility — which happily answers ``--version`` and would be
    reported as an image engine that cannot render a single page. Verify the
    banner, not just the name.
    """
    for name in ("magick", "convert"):
        exe = shutil.which(name)
        if not exe:
            continue
        if IS_WINDOWS and exe.lower().replace("/", "\\").endswith("system32\\convert.exe"):
            continue
        try:
            r = subprocess.run([exe, "--version"], capture_output=True, timeout=20)
            out = (r.stdout or b"").decode("utf-8", errors="replace")
            if "ImageMagick" in out:
                return exe
        except Exception:
            continue
    return None


def _image_engine() -> str | None:
    try:
        import pymupdf  # noqa: F401

        return "pymupdf"
    except ImportError:
        pass
    try:
        import fitz  # noqa: F401

        return "pymupdf"
    except ImportError:
        pass
    if shutil.which("pdftoppm"):
        return "pdftoppm"
    if _find_magick():
        return "magick"
    return None


def pdf_to_images(pdf: Path, out_dir: Path, pages: str | None, dpi: int) -> dict:
    engine = _image_engine()
    if not engine:
        return {"ok": False, "engine": None, "images": [],
                "error": "no PDF->image engine (pip install pymupdf | poppler pdftoppm | ImageMagick)"}

    out_dir.mkdir(parents=True, exist_ok=True)
    images: list[dict] = []

    if engine == "pymupdf":
        try:
            import pymupdf
        except ImportError:
            import fitz as pymupdf
        doc = pymupdf.open(str(pdf))
        total = doc.page_count
        for idx in _parse_pages(pages, total):
            pix = doc[idx].get_pixmap(dpi=dpi)
            path = out_dir / f"page-{idx + 1:02d}.png"
            pix.save(str(path))
            images.append({"page": idx + 1, "file": str(path),
                           "width": pix.width, "height": pix.height,
                           "kb": round(path.stat().st_size / 1024)})
        doc.close()
        return {"ok": True, "engine": engine, "images": images, "pages": total}

    if engine == "pdftoppm":
        args = ["pdftoppm", "-png", "-r", str(dpi)]
        if pages:
            total = _pdf_page_count(pdf)
            sel = _parse_pages(pages, total) if total else []
            if sel:
                args += ["-f", str(sel[0] + 1), "-l", str(sel[-1] + 1)]
        stem = out_dir / "page"
        args += [str(pdf), str(stem)]
        r = subprocess.run(args, capture_output=True, text=True, timeout=180)
        if r.returncode != 0:
            return {"ok": False, "engine": engine, "images": [],
                    "error": (r.stderr or "").strip()[:300]}
        for f in sorted(out_dir.glob("page-*.png")):
            images.append({"page": int(f.stem.split("-")[-1]), "file": str(f),
                           "kb": round(f.stat().st_size / 1024)})
        return {"ok": True, "engine": engine, "images": images}

    # magick
    exe = _find_magick()
    if not exe:
        return {"ok": False, "engine": "magick", "images": [],
                "error": "ImageMagick not found (system32\\convert.exe is not it)"}
    total = _pdf_page_count(pdf) or 1
    chosen = _parse_pages(pages, total)
    for idx in chosen:
        path = out_dir / f"page-{idx + 1:02d}.png"
        r = subprocess.run(
            [exe, "-density", str(dpi), f"{pdf}[{idx}]", str(path)],
            capture_output=True, text=True, timeout=180,
        )
        if r.returncode == 0 and path.exists():
            images.append({"page": idx + 1, "file": str(path),
                           "kb": round(path.stat().st_size / 1024)})
    return {"ok": bool(images), "engine": engine, "images": images,
            "error": None if images else "magick produced no output (Ghostscript missing?)"}


def _pdf_page_count(pdf: Path) -> int:
    try:
        import pymupdf
    except ImportError:
        try:
            import fitz as pymupdf
        except ImportError:
            return 0
    try:
        doc = pymupdf.open(str(pdf))
        n = doc.page_count
        doc.close()
        return n
    except Exception:
        return 0


# --------------------------------------------------------------------------- #
# Probe
# --------------------------------------------------------------------------- #

def _install_hint() -> str:
    if IS_WINDOWS:
        return ("Install one: `winget install TheDocumentFoundation.LibreOffice`, "
                "or any Microsoft Office / WPS Office build.")
    if IS_MACOS:
        return "Install one: `brew install --cask libreoffice`, or Microsoft Word for Mac."
    return "Install one: `apt install libreoffice` / `dnf install libreoffice` / `pacman -S libreoffice-fresh`."


def probe() -> dict:
    report: dict = {
        "platform": f"{platform.system()} {platform.release()} ({platform.machine()})",
        "python": platform.python_version(),
        "pdf_renderers": {},
        "image_engines": {},
        "recommended_pdf": None,
        "recommended_image": None,
    }

    exe = find_soffice()
    report["pdf_renderers"]["soffice"] = {"ok": bool(exe), "path": exe}

    if IS_WINDOWS:
        before = _office_pids()
        com = {"ok": False, "detail": "not probed"}
        try:
            ps = (
                "$r=@();foreach($p in @('Word.Application','KWPS.Application')){"
                "try{$a=New-Object -ComObject $p;[void][Runtime.InteropServices.Marshal]::ReleaseComObject($a);"
                "$r+=$p+'|ok'}catch{$r+=$p+'|no'}};$r -join ';'"
            )
            out = _ps_run(ps, {}, 60).stdout or ""
            found = [seg for seg in out.strip().split(";") if seg.endswith("|ok")]
            com = {"ok": bool(found), "detail": ", ".join(found) or "neither Word nor WPS registered"}
        except subprocess.TimeoutExpired:
            _reap_new_office(before)
            com = {"ok": False, "detail": "COM probe timed out"}
        except Exception as exc:  # noqa: BLE001
            com = {"ok": False, "detail": str(exc)[:160]}
        report["pdf_renderers"]["windows-com"] = com
        report["pdf_renderers"]["windows-com"]["note"] = (
            "WPS registers under Word.Application and reports Name='Microsoft Word'; "
            "identity is decided at render time from Application.Path."
        )

    if IS_MACOS:
        has_word = Path("/Applications/Microsoft Word.app").exists()
        report["pdf_renderers"]["word-mac"] = {"ok": has_word,
                                               "path": "/Applications/Microsoft Word.app" if has_word else None}

    for name in ("pymupdf", "pdftoppm", "magick"):
        if name == "pymupdf":
            ok = _image_engine() == "pymupdf"
            report["image_engines"][name] = {"ok": ok}
        elif name == "pdftoppm":
            p = shutil.which("pdftoppm")
            report["image_engines"][name] = {"ok": bool(p), "path": p}
        else:
            p = _find_magick()
            report["image_engines"][name] = {"ok": bool(p), "path": p}

    for chain in ("soffice", "windows-com", "word-mac"):
        if report["pdf_renderers"].get(chain, {}).get("ok"):
            report["recommended_pdf"] = chain
            break
    for chain in ("pymupdf", "pdftoppm", "magick"):
        if report["image_engines"].get(chain, {}).get("ok"):
            report["recommended_image"] = chain
            break
    return report


def _print_probe(report: dict) -> None:
    line = "-" * 66
    print(f"docx-master render probe\n{line}")
    print(f"platform  {report['platform']}")
    print(f"python    {report['python']}")
    print("\nPDF renderers")
    for name, info in report["pdf_renderers"].items():
        tag = "OK  " if info.get("ok") else "MISS"
        detail = info.get("path") or info.get("detail") or ""
        print(f"  {tag} {name:<14} {detail}")
    print("\nPDF -> image")
    for name, info in report["image_engines"].items():
        tag = "OK  " if info.get("ok") else "MISS"
        print(f"  {tag} {name:<14} {info.get('path') or ''}")
    print(line)
    rec_pdf = report["recommended_pdf"]
    rec_img = report["recommended_image"]
    if rec_pdf and rec_img:
        print(f"Visual self-check: available ({rec_pdf} -> {rec_img})")
    elif rec_pdf:
        print(f"Visual self-check: partial — {rec_pdf} makes the PDF, no image engine "
              f"to turn it into a picture.")
    else:
        print("Visual self-check: unavailable. " + _install_hint())
    print("\n" + RENDER_WARNING)


# --------------------------------------------------------------------------- #
# CLI
# --------------------------------------------------------------------------- #

def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(
        description="Render a .docx/.doc to PDF + PNG for visual self-check (portable)."
    )
    ap.add_argument("source", nargs="?", help="input .docx / .doc")
    ap.add_argument("--out", default=None, help="output directory (default: <source>_render)")
    ap.add_argument("--pages", default=None, help="pages to rasterise, e.g. 1,6-8 (default: all)")
    ap.add_argument("--dpi", type=int, default=150, help="raster DPI (default 150)")
    ap.add_argument("--timeout", type=int, default=120, help="per-renderer timeout in seconds")
    ap.add_argument("--prefer", choices=["soffice", "com", "word-mac"], default=None,
                    help="try this renderer first (others still used as fallback)")
    ap.add_argument("--pdf-only", action="store_true", help="stop after the PDF")
    ap.add_argument("--probe", action="store_true", help="report available renderers and exit")
    ap.add_argument("--human", action="store_true", help="with --probe: readable table instead of JSON")
    args = ap.parse_args(argv)

    if args.probe:
        report = probe()
        if args.human:
            _print_probe(report)
            return 0
        print(json.dumps(report, indent=2, ensure_ascii=False))
        return 0

    if not args.source:
        ap.error("source is required unless --probe is given")

    src = Path(args.source).expanduser().resolve()
    if not src.exists():
        print(json.dumps({"ok": False, "error": f"no such file: {src}"}, ensure_ascii=False))
        return 2

    out_dir = Path(args.out).expanduser().resolve() if args.out else src.with_name(src.stem + "_render")

    with tempfile.TemporaryDirectory(prefix="docx_render_") as _scratch:
        pdf_report = render_pdf(src, out_dir, timeout=args.timeout, prefer=args.prefer)
        if not pdf_report.get("ok"):
            print(json.dumps(pdf_report, indent=2, ensure_ascii=False))
            return 3

        if args.pdf_only:
            print(json.dumps(pdf_report, indent=2, ensure_ascii=False))
            return 0

        img_report = pdf_to_images(Path(pdf_report["pdf"]), out_dir, args.pages, args.dpi)
        merged = {
            "ok": img_report.get("ok", False),
            "source": str(src),
            "pdf": pdf_report["pdf"],
            "renderer": pdf_report["renderer"],
            "renderer_identity": pdf_report["renderer_identity"],
            "image_engine": img_report.get("engine"),
            "dpi": args.dpi,
            "image_count": len(img_report.get("images") or []),
            "images": img_report.get("images") or [],
            "attempts": pdf_report.get("attempts"),
            "warning": RENDER_WARNING,
        }
        if not img_report.get("ok"):
            merged["image_error"] = img_report.get("error")
        print(json.dumps(merged, indent=2, ensure_ascii=False))
        return 0 if merged["ok"] else 4


if __name__ == "__main__":
    raise SystemExit(main())
