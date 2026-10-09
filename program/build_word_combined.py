#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
LaTeX -> Word one-step builder for this workflow.

Pipeline:
  1) Find PDF figures used by \\includegraphics and rasterize them to temporary PNG files.
  2) Build DOCX with Pandoc + pandoc-tex-numbering + reference.docx.
  3) Post-process the DOCX so numbered display equations are centered and equation numbers are right-aligned.

Usage:
  py build_word_combined.py test.tex my-reference.docx final.docx

Optional:
  py build_word_combined.py test.tex my-reference.docx final.docx --dpi 600 --keep-temp

Dependencies:
  - pandoc
  - pandoc-tex-numbering
  - lxml
  - PyMuPDF 1.24.11 (recommended for Python 3.8)
"""

import argparse
import hashlib
import os
import re
import shutil
import subprocess
import sys
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

try:
    from lxml import etree
except ImportError:
    print("ERROR: Missing lxml. Install it with: py -m pip install lxml", file=sys.stderr)
    sys.exit(2)

try:
    import pymupdf as fitz
except ImportError:
    try:
        import fitz
    except ImportError:
        print(
            "ERROR: Missing PyMuPDF. For Python 3.8 install: "
            "py -m pip install PyMuPDF==1.24.11",
            file=sys.stderr,
        )
        sys.exit(2)

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
}
W = "{%s}" % NS["w"]
M = "{%s}" % NS["m"]

# Matches the common form used in the thesis source:
#   \\includegraphics[width=0.70\\textwidth]{figures/foo.pdf}
INCLUDEGRAPHICS_RE = re.compile(
    r"(?P<prefix>\\includegraphics\s*(?:\[[^\]]*\])?\s*\{)"
    r"(?P<path>[^{}]+?\.pdf)"
    r"(?P<suffix>\})",
    re.IGNORECASE,
)


def find_program(name: str) -> str:
    found = shutil.which(name)
    if found:
        return found

    # pip-installed scripts may not be on PATH on Windows.
    scripts_dir = Path(sys.executable).resolve().parent / "Scripts"
    candidates = [scripts_dir / name, scripts_dir / (name + ".exe")]
    for candidate in candidates:
        if candidate.exists():
            return str(candidate)

    raise FileNotFoundError(name)


def rasterize_pdf(pdf_path: Path, png_path: Path, dpi: int) -> None:
    doc = fitz.open(str(pdf_path))
    try:
        if doc.page_count < 1:
            raise RuntimeError(f"PDF has no pages: {pdf_path}")
        page = doc.load_page(0)
        pix = page.get_pixmap(dpi=dpi, alpha=False)
        png_path.parent.mkdir(parents=True, exist_ok=True)
        pix.save(str(png_path))
    finally:
        doc.close()


def preprocess_pdf_figures(tex_path: Path, temp_tex_path: Path, media_dir: Path, dpi: int):
    text = tex_path.read_text(encoding="utf-8-sig")
    # English caption text is inserted as a separate styled Word paragraph.
    text = re.sub(r"\\(?:Figure|Table)Caption\{([^{}]+)\}\{([^{}]+)\}", lambda m: r"\caption{" + m[1] + "}", text)
    # Pandoc cannot resolve custom tabularx column types.  The Word table
    # postprocessor supplies centring, so use native X columns in its input.
    text = re.sub(r"(\\begin\{tabularx\}\{\\textwidth\})\{(Y+)\}", lambda m: m[1] + "{" + "X" * len(m[2]) + "}", text)
    source_dir = tex_path.parent.resolve()
    converted = []

    def replace(match: re.Match) -> str:
        raw = match.group("path").strip()
        # Skip URLs or macro-heavy paths; this builder targets normal local figure paths.
        if "://" in raw or "\\" in raw and not raw.startswith("."):
            # A backslash in a path often indicates a TeX macro rather than a plain filename.
            # Leave it unchanged instead of guessing.
            return match.group(0)

        pdf_path = Path(raw)
        if not pdf_path.is_absolute():
            pdf_path = (source_dir / pdf_path).resolve()

        if not pdf_path.exists():
            raise FileNotFoundError(f"Figure PDF not found: {raw} -> {pdf_path}")

        digest = hashlib.sha1(str(pdf_path).encode("utf-8")).hexdigest()[:10]
        safe_stem = re.sub(r"[^A-Za-z0-9._-]+", "_", pdf_path.stem)
        png_name = f"{safe_stem}_{digest}.png"
        png_path = media_dir / png_name
        rasterize_pdf(pdf_path, png_path, dpi)
        converted.append((pdf_path, png_path))

        # temp .tex is written beside the original .tex, so use a path relative to source_dir.
        rel_png = os.path.relpath(str(png_path), str(source_dir)).replace(os.sep, "/")
        return f"{match.group('prefix')}{rel_png}{match.group('suffix')}"

    new_text = INCLUDEGRAPHICS_RE.sub(replace, text)
    temp_tex_path.write_text(new_text, encoding="utf-8")
    return converted


def patch_equation_layout(src_docx: Path, out_docx: Path) -> int:
    """Center display equations and right-align appended numbers such as (1.1)."""
    with ZipFile(str(src_docx), "r") as zin:
        xml = zin.read("word/document.xml")
        root = etree.fromstring(xml)

        # Determine text area width from section properties.
        sects = root.findall(".//w:sectPr", NS)
        sect = sects[-1] if sects else None
        pg_sz = sect.find("w:pgSz", NS) if sect is not None else None
        page_w = int(pg_sz.get(W + "w")) if pg_sz is not None else 12240
        mar = sect.find("w:pgMar", NS) if sect is not None else None
        left = int(mar.get(W + "left")) if mar is not None else 1800
        right = int(mar.get(W + "right")) if mar is not None else 1800
        text_w = page_w - left - right
        center_pos = text_w // 2
        right_pos = text_w

        changed = 0
        for p in root.findall(".//w:p", NS):
            omp = p.find("m:oMathPara", NS)
            if omp is None:
                continue
            om = omp.find("m:oMath", NS)
            if om is None:
                continue

            children = list(om)
            direct_r = [
                (i, ch, "".join(ch.xpath(".//m:t/text()", namespaces=NS)))
                for i, ch in enumerate(children)
                if ch.tag == M + "r"
            ]
            if len(direct_r) < 3:
                continue

            # pandoc-tex-numbering appends (...) as direct math runs.
            i1, c1, t1 = direct_r[-3]
            i2, c2, t2 = direct_r[-2]
            i3, c3, t3 = direct_r[-1]
            if t1 != "(" or t3 != ")" or not re.fullmatch(r"\d+(?:\.\d+)*", t2 or ""):
                continue

            number = f"({t2})"
            remove_indices = {i1, i2, i3}

            # Remove the spacing run inserted immediately before the number, if present.
            prev_candidates = [x for x in direct_r if x[0] < i1]
            if prev_candidates:
                ip, cp, tp = prev_candidates[-1]
                if tp and all(
                    ch.isspace()
                    or ch
                    in "\u2000\u2001\u2002\u2003\u2004\u2005\u2006\u2007\u2008\u2009\u200a\u200b\u202f\u205f\u3000"
                    for ch in tp
                ):
                    remove_indices.add(ip)

            for idx in sorted(remove_indices, reverse=True):
                om.remove(children[idx])

            # Move the equation out of oMathPara into a normal paragraph with two tab stops.
            p.remove(omp)
            pPr = p.find("w:pPr", NS)
            if pPr is None:
                pPr = etree.Element(W + "pPr")
                p.insert(0, pPr)

            jc = pPr.find("w:jc", NS)
            if jc is not None:
                pPr.remove(jc)
            tabs_old = pPr.find("w:tabs", NS)
            if tabs_old is not None:
                pPr.remove(tabs_old)

            tabs = etree.SubElement(pPr, W + "tabs")
            tcenter = etree.SubElement(tabs, W + "tab")
            tcenter.set(W + "val", "center")
            tcenter.set(W + "pos", str(center_pos))
            tright = etree.SubElement(tabs, W + "tab")
            tright.set(W + "val", "right")
            tright.set(W + "pos", str(right_pos))

            r1 = etree.SubElement(p, W + "r")
            etree.SubElement(r1, W + "tab")
            p.append(om)
            r2 = etree.SubElement(p, W + "r")
            etree.SubElement(r2, W + "tab")
            rn = etree.SubElement(p, W + "r")
            tn = etree.SubElement(rn, W + "t")
            tn.text = number
            changed += 1

        new_xml = etree.tostring(
            root, xml_declaration=True, encoding="UTF-8", standalone=False
        )

        with ZipFile(str(out_docx), "w", ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = new_xml if item.filename == "word/document.xml" else zin.read(item.filename)
                zout.writestr(item, data)

    return changed


def run(args):
    tex = Path(args.tex).resolve()
    reference = Path(args.reference).resolve()
    output = Path(args.output).resolve()

    if not tex.exists():
        raise FileNotFoundError(f"Input TeX not found: {tex}")
    if not reference.exists():
        raise FileNotFoundError(f"Reference DOCX not found: {reference}")

    source_dir = tex.parent
    build_dir = source_dir / "._word_build"
    media_dir = build_dir / "media"
    build_dir.mkdir(parents=True, exist_ok=True)
    media_dir.mkdir(parents=True, exist_ok=True)

    temp_tex = source_dir / (tex.stem + ".__word_temp__.tex")
    temp_docx = build_dir / (output.stem + ".pandoc.docx")

    print(f"[1/4] Preparing PDF figures at {args.dpi} dpi ...")
    converted = preprocess_pdf_figures(tex, temp_tex, media_dir, args.dpi)
    if converted:
        for src, dst in converted:
            print(f"      PDF -> PNG: {src.name} -> {dst.name}")
    else:
        print("      No PDF figures found in \\includegraphics.")

    try:
        pandoc = find_program("pandoc")
    except FileNotFoundError:
        raise RuntimeError("Pandoc was not found on PATH.")

    try:
        numbering_filter = find_program("pandoc-tex-numbering")
    except FileNotFoundError:
        raise RuntimeError(
            "pandoc-tex-numbering was not found. Install it with: "
            "py -m pip install pandoc-tex-numbering"
        )

    print("[2/4] Running Pandoc + pandoc-tex-numbering ...")
    resource_path = os.pathsep.join([str(source_dir), str(media_dir)])
    cmd = [
        pandoc,
        str(temp_tex),
        "--filter",
        numbering_filter,
        "--reference-doc",
        str(reference),
        "--resource-path",
        resource_path,
        "-M",
        "number-theorems=false",
        "-o",
        str(temp_docx),
    ]
    subprocess.run(cmd, cwd=str(source_dir), check=True)

    print("[3/4] Centering equations and right-aligning equation numbers ...")
    output.parent.mkdir(parents=True, exist_ok=True)
    changed = patch_equation_layout(temp_docx, output)
    print(f"      Patched {changed} numbered equation(s).")

    print("[4/4] Done.")
    print(f"      Output: {output}")

    if not args.keep_temp:
        try:
            if temp_tex.exists():
                temp_tex.unlink()
            shutil.rmtree(build_dir, ignore_errors=True)
        except Exception as exc:
            print(f"WARNING: Could not fully remove temporary files: {exc}", file=sys.stderr)
    else:
        print(f"      Temporary files kept in: {build_dir}")
        print(f"      Temporary TeX: {temp_tex}")


def main():
    parser = argparse.ArgumentParser(
        description="Build a DOCX from LaTeX, converting PDF figures to PNG and fixing equation-number alignment."
    )
    parser.add_argument("tex", help="Input .tex file")
    parser.add_argument("reference", help="Pandoc reference .docx")
    parser.add_argument("output", help="Final output .docx")
    parser.add_argument(
        "--dpi",
        type=int,
        default=450,
        help="Rasterization DPI for PDF figures (default: 450)",
    )
    parser.add_argument(
        "--keep-temp",
        action="store_true",
        help="Keep temporary TeX/PNG/DOCX files for debugging",
    )
    args = parser.parse_args()

    try:
        run(args)
    except subprocess.CalledProcessError as exc:
        print(f"ERROR: Pandoc failed with exit code {exc.returncode}.", file=sys.stderr)
        sys.exit(exc.returncode or 1)
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
