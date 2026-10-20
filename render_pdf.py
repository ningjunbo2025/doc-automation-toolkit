# -*- coding: utf-8 -*-
"""pptx -> pdf -> png：用真 LibreOffice 渲染，再做 PDF->PNG。
用法: python render_pdf.py deck.pptx outdir [dpi]
"""
import os, sys, subprocess, glob, shutil
import tempfile

SOFFICE = (os.environ.get("SOFFICE") or shutil.which("soffice") or os.path.join("C:" + os.sep, "Program Files", "LibreOffice", "program", "soffice.exe"))
PROFILE = os.environ.get("SOFFICE_PROFILE") or os.path.join(tempfile.gettempdir(), "md2docx-lo-profile")


def pptx_to_pdf(pptx, outdir):
    os.makedirs(outdir, exist_ok=True)
    os.makedirs(PROFILE, exist_ok=True)
    subprocess.run(
        [SOFFICE, "--headless", "--norestore", "--nolockcheck", "--nodefault",
         "-env:UserInstallation=file:///" + PROFILE.replace("\\", "/"),
         "--convert-to", "pdf", "--outdir", outdir, pptx],
        check=True, timeout=300,
        stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
    )
    pdf = os.path.join(outdir, os.path.splitext(os.path.basename(pptx))[0] + ".pdf")
    if not os.path.exists(pdf):
        raise SystemExit("转换失败，未生成 PDF: " + pdf)
    return pdf


def pdf_to_png(pdf, outdir, dpi=150):
    import fitz
    doc = fitz.open(pdf)
    paths = []
    for i, page in enumerate(doc, 1):
        pix = page.get_pixmap(dpi=dpi)
        p = os.path.join(outdir, "slide-%02d.png" % i)
        pix.save(p)
        paths.append(p)
    doc.close()
    return paths


if __name__ == "__main__":
    pptx = sys.argv[1]
    outdir = sys.argv[2]
    dpi = int(sys.argv[3]) if len(sys.argv) > 3 else 150
    pdf = pptx_to_pdf(pptx, outdir)
    print("PDF:", pdf)
    for p in pdf_to_png(pdf, outdir, dpi):
        print("  ", p)
