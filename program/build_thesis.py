#!/usr/bin/env python3
"""Build thesis-template.tex to PDF and style-matched DOCX."""

from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
from pathlib import Path

from build_word_combined import patch_equation_layout, preprocess_pdf_figures


CORE = Path(__file__).resolve().parent
PACKAGE = CORE.parent
ROOT = PACKAGE
DEFAULT_DOCX_PYTHON = Path(sys.executable)


def find(name):
    if name == "pandoc-tex-numbering":
        return str(CORE / "numbering_filter.py")
    bundled = PACKAGE / "tools" / (name + ".exe")
    if bundled.exists():
        return str(bundled)
    value = shutil.which(name)
    if value:
        return value
    scripts = Path(sys.executable).resolve().parent / "Scripts"
    for candidate in (scripts / name, scripts / f"{name}.exe"):
        if candidate.exists():
            return str(candidate)
    raise RuntimeError(f"Required program not found: {name}")


def run(command, cwd=None):
    print("+", " ".join(map(str, command)))
    subprocess.run([str(x) for x in command], cwd=str(cwd or ROOT), check=True)


def build_pdf(tex):
    latexmk = find("latexmk")
    run([latexmk, "-xelatex", "-interaction=nonstopmode", "-halt-on-error", "-file-line-error", tex.name])
    return tex.with_suffix(".pdf")


def build_word(tex, reference, output, dpi):
    import fitz
    build_dir = ROOT / "._thesis_build"
    media_dir = build_dir / "media"
    build_dir.mkdir(exist_ok=True)
    media_dir.mkdir(exist_ok=True)
    # Record native text and line geometry, not screenshots, for front matter.
    cover_pdf = tex.with_suffix(".pdf")
    if not cover_pdf.exists() or tex.stat().st_mtime > cover_pdf.stat().st_mtime:
        print("Refreshing PDF geometry for editable front matter")
        cover_pdf = build_pdf(tex)
    if cover_pdf.exists():
        with fitz.open(str(cover_pdf)) as pdf:
            front_pages = []
            for page in pdf:
                raw = page.get_text("rawdict")
                if not page.get_text().strip():
                    if len(front_pages) == 3:
                        front_pages.append({"spans": [], "lines": []})
                        break # explicit blank immediately after declarations
                    continue # intentional TeX recto/verso blank pages
                spans = []
                for block in raw["blocks"]:
                    for line in block.get("lines", []):
                        for span in line["spans"]:
                            groups = []
                            for char in span["chars"]:
                                if not groups or char["bbox"][0] - groups[-1][-1]["bbox"][2] > 2:
                                    groups.append([])
                                groups[-1].append(char)
                            for group in groups:
                                spans.append({"text": "".join(c["c"] for c in group), "font": span["font"], "size": span["size"], "x": group[0]["origin"][0], "baseline": group[0]["origin"][1], "width": group[-1]["bbox"][2] - group[0]["bbox"][0]})
                front_pages.append({"spans": spans, "lines": [list(d["rect"]) for d in page.get_drawings() if abs(d["rect"].height) < 1]})
                if len(front_pages) == 4:
                    break
            (build_dir / "front-layout.json").write_text(json.dumps(front_pages, ensure_ascii=False), encoding="utf-8")
    temp_tex = build_dir / "thesis-word-source.tex"
    converted = preprocess_pdf_figures(tex, temp_tex, media_dir, dpi)
    # Pandoc does not reliably resolve LaTeX textwidth for raster images.
    # Use the actual PDF body width so every format receives the same size.
    import re
    with fitz.open(str(cover_pdf)) as pdf:
        body_width_cm = pdf[0].rect.width * 2.54 / 72 - 2 * 3.17
    source = temp_tex.read_text(encoding="utf-8")
    source = re.sub(r"width\s*=\s*(\d*(?:\.\d+)?)\\(?:textwidth|linewidth)",
                    lambda m: "width=%.5fcm" % (float(m.group(1) or 1) * body_width_cm), source)
    temp_tex.write_text(source, encoding="utf-8")
    for src, dst in converted:
        print(f"PDF figure: {src.name} -> {dst.name}")

    pandoc_docx = build_dir / "pandoc.docx"
    equation_docx = build_dir / "equations.docx"
    resource_path = os.pathsep.join([str(ROOT), str(media_dir), str(build_dir)])
    cmd = [
        find("pandoc"), str(temp_tex), "--from=latex", "--to=docx",
        "--reference-doc", str(reference), "--resource-path", resource_path,
        "--citeproc", "-M", "link-citations=true", "--bibliography", str(ROOT / "references.bib"),
        "--csl", str(PACKAGE / "numeric-citation.csl"),
        "--top-level-division=chapter", "-o", str(pandoc_docx),
    ]
    numbering = None
    try:
        numbering = find("pandoc-tex-numbering")
    except RuntimeError:
        print("NOTE: pandoc-tex-numbering not found; Pandoc native equation handling is used.")
    if numbering:
        cmd[3:3] = ["--filter", numbering, "-M", "number-theorems=false"]
    if os.environ.get("THESIS_MATH_MANIFEST"):
        cmd.extend(["--filter", str(CORE / "capture_math.py")])
    run(cmd)
    patch_equation_layout(pandoc_docx, equation_docx)
    docx_python = os.environ.get("DOCX_PYTHON")
    if docx_python:
        docx_python = Path(docx_python)
    elif DEFAULT_DOCX_PYTHON.exists():
        docx_python = DEFAULT_DOCX_PYTHON
    else:
        docx_python = Path(sys.executable)
    run([docx_python, CORE / "postprocess_thesis_docx.py", equation_docx, output])
    return output


def main():
    global ROOT
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--tex", default="thesis-template.tex")
    parser.add_argument("--reference", default="thesis-reference.docx")
    parser.add_argument("--word-output", default="thesis-template.docx")
    parser.add_argument("--dpi", type=int, default=450)
    parser.add_argument("--pdf-only", action="store_true")
    parser.add_argument("--word-only", action="store_true")
    parser.add_argument("--math-mode", choices=["native", "mathtype"], default="native")
    args = parser.parse_args()
    tex = (ROOT / args.tex).resolve()
    ROOT = tex.parent
    if args.math_mode == "mathtype" and not args.pdf_only:
        (ROOT / "._thesis_build").mkdir(exist_ok=True)
        os.environ["THESIS_MATH_MANIFEST"] = str(ROOT / "._thesis_build" / "math-manifest.json")
    else:
        os.environ.pop("THESIS_MATH_MANIFEST", None)
    os.environ["THESIS_TEX"] = str(tex)
    os.environ["THESIS_FRONT_LAYOUT"] = str(ROOT / "._thesis_build" / "front-layout.json")
    os.environ["DOCX_PYTHON"] = sys.executable
    os.environ["PATH"] = str(Path(sys.executable).parent) + os.pathsep + os.environ.get("PATH", "")
    if (ROOT / "figures" / "测试多图.pdf").exists():
        run([sys.executable, CORE / "compose_subfigures.py"])
    if not args.word_only:
        print("PDF:", build_pdf(tex))
    if not args.pdf_only:
        reference = (PACKAGE / args.reference).resolve()
        if not reference.exists():
            raise RuntimeError("Run make_thesis_reference.py first or use build_all.bat")
        docx_python = os.environ.get("DOCX_PYTHON") or (
            str(DEFAULT_DOCX_PYTHON) if DEFAULT_DOCX_PYTHON.exists() else sys.executable
        )
        run([docx_python, CORE / "calibrate_reference_docx.py", reference])
        print("DOCX:", build_word(tex, reference, (ROOT / args.word_output).resolve(), args.dpi))
        run([docx_python, CORE / "finalize_word_layout.py", (ROOT / args.word_output).resolve()])
        run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", CORE / "update_word_fields.ps1", "-Path", (ROOT / args.word_output).resolve()])
        run([docx_python, "-c", "from docx import Document; from docx.oxml.ns import qn; from finalize_word_layout import format_cached_toc; from pathlib import Path; import sys; format_cached_toc(Path(sys.argv[1])); d=Document(sys.argv[1]); u=d.settings.element.find(qn('w:updateFields')); u.set(qn('w:val'),'false') if u is not None else None; d.save(sys.argv[1])", (ROOT / args.word_output).resolve()], cwd=CORE)
        if args.math_mode == "mathtype":
            run([docx_python, CORE / "mathtype_convert.py", (ROOT / args.word_output).resolve(), os.environ["THESIS_MATH_MANIFEST"]])
            # Refresh pagination without asking Word to update embedded equations.
            run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", CORE / "update_word_fields.ps1", "-Path", (ROOT / args.word_output).resolve(), "-PageNumbersOnly"])
            run([docx_python, "-c", "from finalize_word_layout import format_cached_toc; from pathlib import Path; import sys; format_cached_toc(Path(sys.argv[1]))", (ROOT / args.word_output).resolve()], cwd=CORE)
        # Apply after Word saves so Show All exposes no conversion placeholders.
        run([docx_python, CORE / "clean_equation_markers.py", (ROOT / args.word_output).resolve()])


if __name__ == "__main__":
    main()
