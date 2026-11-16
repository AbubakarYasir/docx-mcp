"""PDF export: print-ready and digital-ready PDFs of the open document.

Engines
  word         Microsoft Word through COM (Windows). The reference renderer:
               fields, indexes and tables of contents are refreshed first, and
               fonts, tagged structure and heading bookmarks come out as Word
               draws them.
  libreoffice  LibreOffice headless, for machines without Word.
  auto         Word when it can be reached, otherwise LibreOffice.

Profiles
  digital  screen-optimised, heading bookmarks shown, hyperlinks kept, book
           spreads (cover alone on the right), right-to-left page order for
           Arabic/Urdu books, title shown in the viewer.
  print    print-optimised, TrimBox at the page edge, BleedBox and crop marks
           around it (bleed and marks are optional), viewer scaling off so a
           printer never shrinks the page.

Numerals: Word draws automatic numbers (note marks, page numbers, list numbers)
with the application's "Numeral" option, which is a setting of the computer,
not of the file. The export sets it per book for the run and restores it:
Hindi (١٢٣) for Arabic/Urdu books, Context for English books. A preset or the
caller can choose with `numerals` (hindi, arabic, context, system) or the
custom property docx-mcp.numerals.
"""

from __future__ import annotations

import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

MM = 72 / 25.4

_NUMERALS = {"arabic": 0, "hindi": 1, "context": 2, "system": 3}

_WORD_PS1 = r"""
param([string]$In, [string]$Out, [int]$Optimize = 1, [int]$Bookmarks = 1,
      [int]$PdfA = 0, [int]$Numerals = -1, [int]$Update = 1, [int]$Diacritics = -1)
$ErrorActionPreference = "Stop"
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
$old = $word.Options.ArabicNumeral
$oldUse = $word.Options.UseDiffDiacColor
$oldDiac = $word.Options.DiacriticColorVal
try {
    if ($Numerals -ge 0) { $word.Options.ArabicNumeral = $Numerals }
    if ($Diacritics -ge 0) {
        $word.Options.UseDiffDiacColor = $true
        $word.Options.DiacriticColorVal = $Diacritics
    }
    $doc = $word.Documents.Open($In, $false, $true)
    if ($Update -eq 1) {
        $doc.Fields.Update() | Out-Null
        foreach ($i in $doc.Indexes) { $i.Update() }
        foreach ($t in $doc.TablesOfContents) { $t.Update() }
        $doc.Repaginate()
        foreach ($i in $doc.Indexes) { $i.Update() }
        foreach ($t in $doc.TablesOfContents) { $t.UpdatePageNumbers() }
    }
    $doc.ExportAsFixedFormat($Out, 17, $false, $Optimize, 0, 1, 1, 0, $true, $true,
                             $Bookmarks, $true, $true, [bool]$PdfA)
    Write-Output ("OK pages=" + $doc.ComputeStatistics(2))
    $doc.Close(0)
} finally {
    $word.Options.ArabicNumeral = $old
    $word.Options.UseDiffDiacColor = $oldUse
    $word.Options.DiacriticColorVal = $oldDiac
    $word.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}
"""


def _bgr(hex_color: str | None) -> int:
    """Word's colour value (0xBBGGRR) for an RRGGBB hex; -1 for none."""
    if not hex_color:
        return -1
    h = hex_color.lstrip("#")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return r + (g << 8) + (b << 16)


def _powershell() -> str | None:
    if sys.platform != "win32":
        return None
    return shutil.which("powershell") or shutil.which("pwsh")


def _word_export(
    docx: Path,
    pdf: Path,
    *,
    optimize_print: bool,
    bookmarks: bool,
    pdfa: bool,
    numerals: str | None,
    update_fields: bool,
    diacritic_color: str | None = None,
) -> int:
    ps = _powershell()
    if ps is None:
        raise RuntimeError("Microsoft Word export needs Windows with Word installed.")
    with tempfile.TemporaryDirectory() as td:
        script = Path(td) / "export.ps1"
        script.write_text(_WORD_PS1, encoding="utf-8-sig")
        args = [
            ps, "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(script),
            "-In", str(docx), "-Out", str(pdf),
            "-Optimize", "0" if optimize_print else "1",
            "-Bookmarks", "1" if bookmarks else "0",
            "-PdfA", "1" if pdfa else "0",
            "-Numerals", str(_NUMERALS.get(numerals, -1)),
            "-Update", "1" if update_fields else "0",
            "-Diacritics", str(_bgr(diacritic_color)),
        ]  # fmt: skip
        r = subprocess.run(args, capture_output=True, text=True, timeout=900)
    if r.returncode != 0 or "OK pages=" not in r.stdout:
        raise RuntimeError(f"Word export failed: {(r.stderr or r.stdout).strip()[-800:]}")
    return int(r.stdout.split("OK pages=")[1].split()[0])


def _libreoffice_export(docx: Path, pdf: Path, *, bookmarks: bool, pdfa: bool) -> None:
    lo = shutil.which("libreoffice") or shutil.which("soffice")
    if lo is None:
        raise RuntimeError(
            "Neither Microsoft Word nor LibreOffice was found. Install one of them "
            "('libreoffice' or 'soffice' on PATH)."
        )
    bm = "true" if bookmarks else "false"
    opts = (
        f'{{"ExportBookmarks":{{"type":"boolean","value":"{bm}"}},'
        f'"SelectPdfVersion":{{"type":"long","value":"{2 if pdfa else 0}"}},'
        '"UseTaggedPDF":{"type":"boolean","value":"true"}}'
    )
    with tempfile.TemporaryDirectory() as td:
        r = subprocess.run(
            [lo, "--headless", f"-env:UserInstallation=file:///{td}/lo".replace("\\", "/"),
             "--convert-to", f"pdf:writer_pdf_Export:{opts}", "--outdir", td, str(docx)],
            capture_output=True, text=True, timeout=900,
        )  # fmt: skip
        made = Path(td) / (docx.stem + ".pdf")
        if r.returncode != 0 or not made.exists():
            raise RuntimeError(f"LibreOffice export failed: {(r.stderr or r.stdout).strip()}")
        pdf.parent.mkdir(parents=True, exist_ok=True)
        shutil.move(str(made), pdf)


def finish_pdf(
    path: str | os.PathLike,
    *,
    profile: str = "digital",
    rtl: bool = False,
    bleed_mm: float = 0.0,
    crop_marks: bool = False,
    title: str | None = None,
    author: str | None = None,
    two_page: bool = True,
) -> dict:
    """Post-process an exported PDF in place: page boxes, crop marks, viewer
    settings and metadata. Bookmarks, links and tags are kept."""
    from pypdf import PdfReader, PdfWriter
    from pypdf.generic import (
        ArrayObject,
        BooleanObject,
        DecodedStreamObject,
        DictionaryObject,
        NameObject,
        RectangleObject,
    )

    path = Path(path)
    writer = PdfWriter(clone_from=PdfReader(str(path)))
    bleed = max(0.0, bleed_mm) * MM
    slug = (bleed + 10 * MM) if crop_marks else bleed
    marked = 0
    if profile == "print":
        for page in writer.pages:
            x0, y0, x1, y1 = (float(v) for v in page.mediabox)
            page.trimbox = RectangleObject([x0, y0, x1, y1])
            page.bleedbox = RectangleObject([x0 - bleed, y0 - bleed, x1 + bleed, y1 + bleed])
            if slug:
                box = RectangleObject([x0 - slug, y0 - slug, x1 + slug, y1 + slug])
                page.mediabox = box
                page.cropbox = box
            if crop_marks:
                off, ln = bleed + 3 * MM, 6 * MM
                ops = ["Q", "q", "0.25 w", "1 1 1 1 K", "[] 0 d"]
                for cx, sx in ((x0, -1), (x1, 1)):
                    for cy, sy in ((y0, -1), (y1, 1)):
                        ha, hb = cx + sx * off, cx + sx * (off + ln)
                        va, vb = cy + sy * off, cy + sy * (off + ln)
                        ops.append(f"{ha:.2f} {cy:.2f} m {hb:.2f} {cy:.2f} l S")
                        ops.append(f"{cx:.2f} {va:.2f} m {cx:.2f} {vb:.2f} l S")
                ops.append("Q")
                pre, post = DecodedStreamObject(), DecodedStreamObject()
                pre.set_data(b"q\n")
                post.set_data(("\n" + "\n".join(ops) + "\n").encode())
                contents = page.get("/Contents")
                arr = ArrayObject([writer._add_object(pre)])
                if contents is not None:
                    c = contents.get_object() if hasattr(contents, "get_object") else contents
                    if isinstance(c, ArrayObject):
                        arr.extend(c)
                    else:
                        arr.append(contents)
                arr.append(writer._add_object(post))
                page[NameObject("/Contents")] = arr
                marked += 1
    vp = DictionaryObject()
    vp[NameObject("/DisplayDocTitle")] = BooleanObject(True)
    if rtl:
        vp[NameObject("/Direction")] = NameObject("/R2L")
    if profile == "print":
        vp[NameObject("/PrintScaling")] = NameObject("/None")
    writer._root_object[NameObject("/ViewerPreferences")] = vp
    if profile == "digital":
        writer._root_object[NameObject("/PageMode")] = NameObject("/UseOutlines")
        writer._root_object[NameObject("/PageLayout")] = NameObject(
            "/TwoPageRight" if two_page else "/SinglePage"
        )
    meta = {}
    if title:
        meta["/Title"] = title
    if author:
        meta["/Author"] = author
    if meta:
        writer.add_metadata(meta)
    with open(path, "wb") as fh:
        writer.write(fh)
    return {
        "pages": len(writer.pages),
        "trim_box": profile == "print",
        "bleed_mm": bleed_mm if profile == "print" else 0,
        "crop_marks": marked,
    }


_PREF = "docx-mcp.pdf."


class PdfExportMixin:
    def set_pdf_preferences(
        self,
        *,
        bleed_mm: float | None = None,
        crop_marks: bool | None = None,
        pdfa: bool | None = None,
        two_page: bool | None = None,
        numerals: str | None = None,
    ) -> dict:
        """Store a client's PDF wishes in the document (custom properties), so
        every later export_pdf follows them unless the call says otherwise.
        bleed_mm and crop_marks apply to print exports."""
        if numerals is not None and numerals not in _NUMERALS:
            raise ValueError(f"numerals must be one of {sorted(_NUMERALS)}")
        if bleed_mm is not None and not 0 <= bleed_mm <= 20:
            raise ValueError("bleed_mm must be between 0 and 20")
        for key, val in (("bleed-mm", bleed_mm), ("crop-marks", crop_marks),
                         ("pdfa", pdfa), ("two-page", two_page)):  # fmt: skip
            if val is not None:
                self.set_custom_property(_PREF + key, str(val).lower())
        if numerals is not None:
            self.set_custom_property("docx-mcp.numerals", numerals)
        return self.get_pdf_preferences()

    def get_pdf_preferences(self) -> dict:
        try:
            props = self.get_custom_properties()
        except Exception:  # noqa: BLE001
            props = {}

        def flag(key):
            v = props.get(_PREF + key)
            return None if v is None else str(v).lower() == "true"

        bleed = props.get(_PREF + "bleed-mm")
        return {
            "bleed_mm": float(bleed) if bleed not in (None, "") else None,
            "crop_marks": flag("crop-marks"),
            "pdfa": flag("pdfa"),
            "two_page": flag("two-page"),
            "numerals": self._pdf_numerals(),
        }

    def _pdf_numerals(self) -> str:
        try:
            pref = self.get_custom_properties().get("docx-mcp.numerals")
        except Exception:  # noqa: BLE001
            pref = None
        if pref in _NUMERALS:
            return pref
        return "hindi" if self._is_rtl() else "context"

    def export_pdf(
        self,
        output_path: str,
        profile: str = "digital",
        *,
        engine: str = "auto",
        numerals: str | None = None,
        bleed_mm: float | None = None,
        crop_marks: bool | None = None,
        pdfa: bool | None = None,
        two_page: bool | None = None,
        update_fields: bool = True,
    ) -> dict:
        """Export the document as a print-ready or digital-ready PDF.

        Args:
            output_path: where to write the PDF.
            profile: digital (bookmarks, spreads, right-to-left page order for
                Arabic books) or print (TrimBox, optional bleed and crop marks,
                no viewer scaling).
            engine: auto, word or libreoffice.
            numerals: hindi «١٢٣», arabic «123», context, system — how Word
                draws note and page numbers for this export. Default: the
                document's docx-mcp.numerals property, else hindi for
                right-to-left books and context for English.
            bleed_mm: print bleed around the trim (default: the document's
                set_pdf_preferences, else 3 for print).
            crop_marks: crop marks outside the bleed (default: preference,
                else on for print).
            pdfa: archival PDF/A (Word: PDF/A-1b, LibreOffice: PDF/A-2b).
            two_page: digital view as book spreads (default on).
            update_fields: refresh fields, indexes and contents first (Word).
        """
        if self.workdir is None:
            raise RuntimeError("No document is open.")
        if profile not in ("digital", "print"):
            raise ValueError("profile must be digital or print")
        if engine not in ("auto", "word", "libreoffice"):
            raise ValueError("engine must be auto, word or libreoffice")
        if numerals is not None and numerals not in _NUMERALS:
            raise ValueError(f"numerals must be one of {sorted(_NUMERALS)}")
        printing = profile == "print"
        pref = self.get_pdf_preferences()
        if bleed_mm is None:
            bleed_mm = pref["bleed_mm"] if pref["bleed_mm"] is not None else 3.0
        if not printing:
            bleed_mm = 0.0
        if crop_marks is None:
            crop_marks = printing if pref["crop_marks"] is None else pref["crop_marks"]
        crop_marks = bool(crop_marks) and printing
        if pdfa is None:
            pdfa = bool(pref["pdfa"])
        if two_page is None:
            two_page = pref["two_page"] is not False
        numerals = numerals or self._pdf_numerals()
        rtl = self._is_rtl()
        try:
            diacritic = self.get_custom_properties().get("docx-mcp.diacritic-color")
        except Exception:  # noqa: BLE001
            diacritic = None

        self.save(self.source_path, backup=False)
        out = Path(output_path).resolve()
        out.parent.mkdir(parents=True, exist_ok=True)
        src = Path(self.source_path).resolve()

        used, err = None, None
        if engine in ("auto", "word") and _powershell():
            try:
                _word_export(
                    src, out, optimize_print=printing, bookmarks=True, pdfa=pdfa,
                    numerals=numerals, update_fields=update_fields,
                    diacritic_color=diacritic,
                )  # fmt: skip
                used = "word"
            except Exception as e:  # noqa: BLE001
                if engine == "word":
                    raise
                err = str(e)
        elif engine == "word":
            raise RuntimeError("Microsoft Word export needs Windows with Word installed.")
        if used is None:
            _libreoffice_export(src, out, bookmarks=True, pdfa=pdfa)
            used = "libreoffice"

        try:
            props = self.get_properties()
        except Exception:  # noqa: BLE001
            props = {}
        info = finish_pdf(
            out,
            profile=profile,
            rtl=rtl,
            bleed_mm=bleed_mm,
            crop_marks=crop_marks,
            title=props.get("title") or None,
            author=props.get("creator") or props.get("author") or None,
            two_page=two_page,
        )
        result = {
            "pdf_path": str(out),
            "engine": used,
            "profile": profile,
            "numerals": numerals if used == "word" else "renderer default",
            **info,
        }
        if err:
            result["word_error"] = err
        return result

    def convert_to_pdf(self, output_path: str) -> dict:
        """Plain PDF of the document (digital profile, best available engine)."""
        r = self.export_pdf(output_path, "digital")
        return {"pdf_path": r["pdf_path"], "engine": r["engine"]}
