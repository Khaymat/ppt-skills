#!/usr/bin/env python3
"""Render a .pptx to PNG files (one per slide) so the result can be looked at before delivery.

    python scripts/render_check.py deck.pptx -o preview/

Uses PowerPoint (COM) on Windows when available, otherwise LibreOffice (soffice) + PyMuPDF.
Also writes preview/contact_sheet.png when Pillow is present.
"""
from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path


def via_powerpoint(pptx: Path, out: Path) -> bool:
    if sys.platform != "win32":
        return False
    try:
        import win32com.client  # type: ignore
    except ImportError:
        script = (f"$p=New-Object -ComObject PowerPoint.Application;"
                  f"$d=$p.Presentations.Open('{pptx.resolve()}',1,0,0);"
                  f"$d.Export('{out.resolve()}','PNG',1280,720);$d.Close();$p.Quit()")
        return subprocess.run(["powershell", "-NoProfile", "-Command", script]).returncode == 0
    app = win32com.client.Dispatch("PowerPoint.Application")
    d = app.Presentations.Open(str(pptx.resolve()), True, False, False)
    d.Export(str(out.resolve()), "PNG", 1280, 720)
    d.Close()
    app.Quit()
    return True


def via_libreoffice(pptx: Path, out: Path) -> bool:
    exe = shutil.which("soffice") or shutil.which("libreoffice")
    if not exe:
        return False
    import fitz  # PyMuPDF
    with tempfile.TemporaryDirectory() as tmp:
        subprocess.run([exe, "--headless", "--convert-to", "pdf", "--outdir", tmp, str(pptx)], check=True)
        pdf = next(Path(tmp).glob("*.pdf"))
        for i, page in enumerate(fitz.open(pdf), 1):
            page.get_pixmap(dpi=96).save(out / f"Slide{i}.PNG")
    return True


def contact_sheet(out: Path, cols: int = 3):
    from PIL import Image
    files = sorted(out.glob("Slide*.PNG"), key=lambda p: int("".join(c for c in p.stem if c.isdigit())))
    if not files:
        return
    tw, th = 640, 360
    rows = -(-len(files) // cols)
    sheet = Image.new("RGB", (cols * tw, rows * th), "white")
    for k, f in enumerate(files):
        sheet.paste(Image.open(f).convert("RGB").resize((tw, th)), ((k % cols) * tw, (k // cols) * th))
    sheet.save(out / "contact_sheet.png")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx", type=Path)
    ap.add_argument("-o", "--out", type=Path, default=Path("preview"))
    a = ap.parse_args()
    a.out.mkdir(parents=True, exist_ok=True)
    if not (via_powerpoint(a.pptx, a.out) or via_libreoffice(a.pptx, a.out)):
        raise SystemExit("No renderer found. Install PowerPoint (Windows) or LibreOffice plus PyMuPDF.")
    contact_sheet(a.out)
    print(f"rendered to {a.out}")


if __name__ == "__main__":
    main()
