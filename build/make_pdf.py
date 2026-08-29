#!/usr/bin/env python3
"""Convert the Paramount initiating-coverage note to a typeset PDF."""
import subprocess, pathlib, re

SRC = "/mnt/user-data/outputs/Paramount_Resources_Initiating_Coverage.md"
OUT = "/mnt/user-data/outputs/Paramount_Resources_Initiating_Coverage.pdf"
WORK = pathlib.Path("/home/claude/pou")

body = subprocess.run(
    ["pandoc", SRC, "-f", "markdown-tex_math_dollars-tex_math_single_backslash",
     "-t", "html5"],
    capture_output=True, text=True, check=True).stdout

# The first table is the key-statistics block; tag it so it can be styled apart.
body = body.replace("<table>", '<table class="stats">', 1)

# Pull the title block out of the flow so it can be laid out as a masthead.
body = re.sub(r"<h1[^>]*>.*?</h1>\s*", "", body, count=1, flags=re.S)
body = re.sub(r"<h2[^>]*>.*?</h2>\s*", "", body, count=1, flags=re.S)
body = re.sub(r"<p><strong>25 August 2026</strong></p>\s*", "", body, count=1)
body = body.replace("<hr />", "", 1)

CSS = """
@page { size: A4; margin: 20mm 18mm 20mm 18mm; }
body { font-family: "Caladea", "Bitstream Charter", serif; font-size: 10pt;
       line-height: 1.45; color: #1a1a1a; margin: 0; }
.masthead { border-bottom: 2.5pt solid #1F3864; padding-bottom: 8pt; margin-bottom: 14pt; }
.ticker { font-family: "Carlito", sans-serif; font-size: 9pt; letter-spacing: 1.4pt;
          color: #6b6b6b; text-transform: uppercase; margin: 0 0 3pt 0; }
.coname { font-family: "Carlito", sans-serif; font-size: 20pt; font-weight: bold;
          color: #1F3864; margin: 0 0 5pt 0; line-height: 1.15; }
.reco { font-family: "Carlito", sans-serif; font-size: 12.5pt; color: #1a1a1a;
        margin: 0 0 4pt 0; }
.reco strong { color: #1F3864; }
.dateline { font-family: "Carlito", sans-serif; font-size: 8.5pt; color: #6b6b6b; margin: 0; }
h1, h2 { font-family: "Carlito", sans-serif; color: #1F3864; font-weight: bold;
         margin: 16pt 0 6pt 0; font-size: 12.5pt; line-height: 1.25;
         border-bottom: 0.6pt solid #c9d2e4; padding-bottom: 3pt; }
h3 { font-family: "Carlito", sans-serif; color: #2c3e5c; font-weight: bold;
     font-size: 10.5pt; margin: 12pt 0 4pt 0; }
p { margin: 0 0 7pt 0; text-align: justify; }
strong { color: #10233f; }
em { color: #3a3a3a; }
hr { border: 0; border-top: 0.6pt solid #d5d5d5; margin: 12pt 0; }
ul { margin: 0 0 8pt 0; padding-left: 16pt; }
li { margin-bottom: 3pt; }
table { border-collapse: collapse; width: 100%; margin: 8pt 0 12pt 0;
        font-family: "Carlito", sans-serif; font-size: 8.8pt; }
th { background: #1F3864; color: #ffffff; font-weight: bold; text-align: right;
     padding: 4.5pt 6pt; border: 0; }
th:first-child { text-align: left; }
td { padding: 3.6pt 6pt; border-bottom: 0.5pt solid #dfe3ea; text-align: right; }
td:first-child { text-align: left; }
tr:nth-child(even) td { background: #f6f8fb; }
table.stats { width: 62%; font-size: 9.2pt; margin: 0 0 14pt 0; }
table.stats td { border-bottom: 0.5pt solid #e3e7ee; padding: 3.4pt 6pt; }
table.stats td:first-child { color: #4a4a4a; }
table.stats tr:nth-child(even) td { background: #f6f8fb; }
table.stats tr:first-child td { background: #e8edf6; font-weight: bold; }
table.stats tr:nth-child(2) td { background: #e8edf6; font-weight: bold; }
.disclaimer { font-size: 8pt; color: #6b6b6b; font-style: italic;
              border-left: 2.2pt solid #c9d2e4; padding: 5pt 9pt; margin: 10pt 0 14pt 0;
              background: #fafbfc; text-align: left; }
h2 { page-break-after: avoid; }
h3 { page-break-after: avoid; }
table { page-break-inside: avoid; }
"""

MAST = """<div class="masthead">
<p class="ticker">Energy &middot; Canadian Exploration &amp; Production &middot; Initiating Coverage</p>
<p class="coname">Paramount Resources Ltd.</p>
<p class="reco"><strong>HOLD</strong> &nbsp;|&nbsp; Target C$33.00 &nbsp;|&nbsp; TSX: POU &nbsp;|&nbsp; Last C$32.65</p>
<p class="dateline">25 August 2026</p>
</div>"""

# Style the two italic disclaimer paragraphs distinctly.
body = re.sub(
    r"<p><em>(This note is a hypothetical.*?)</em></p>",
    r'<p class="disclaimer">\1</p>', body, flags=re.S)
body = re.sub(
    r"<p><em>(Prepared as an independent.*?)</em></p>",
    r'<p class="disclaimer">\1</p>', body, flags=re.S)

html = (f"<!DOCTYPE html><html><head><meta charset='utf-8'>"
        f"<style>{CSS}</style></head><body>{MAST}{body}</body></html>")
(WORK / "note.html").write_text(html, encoding="utf-8")

subprocess.run([
    "wkhtmltopdf",
    "--enable-local-file-access",
    "--page-size", "A4",
    "--margin-top", "18mm", "--margin-bottom", "18mm",
    "--margin-left", "16mm", "--margin-right", "16mm",
    "--quiet",
    str(WORK / "note.html"), str(WORK / "raw.pdf")
], check=True)

# wkhtmltopdf here is built against unpatched Qt, so its footer switches are
# ignored. Stamp the footer on afterwards instead.
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.colors import HexColor
from reportlab.lib.units import mm
import io

reader = PdfReader(str(WORK / "raw.pdf"))
n = len(reader.pages)
writer = PdfWriter()
W_, H_ = A4
for i, page in enumerate(reader.pages, start=1):
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    c.setStrokeColor(HexColor("#C9D2E4")); c.setLineWidth(0.5)
    c.line(16*mm, 13*mm, W_ - 16*mm, 13*mm)
    c.setFont("Helvetica", 7); c.setFillColor(HexColor("#6B6B6B"))
    c.drawString(16*mm, 9.5*mm, "Paramount Resources Ltd. (TSX: POU) — Initiating Coverage")
    c.drawRightString(W_ - 16*mm, 9.5*mm, f"Page {i} of {n}")
    c.save(); buf.seek(0)
    page.merge_page(PdfReader(buf).pages[0])
    writer.add_page(page)
with open(OUT, "wb") as f:
    writer.write(f)
print(f"PDF written: {OUT}  ({n} pages)")
