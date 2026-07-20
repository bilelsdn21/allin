"""
PDF splitter - the review step for PDF passports
------------------------------------------------
Turns every PDF in the 'images' folder into normal picture files you can
rename, so they flow back into the usual process.

For each PDF page:
  - a normal (portrait) page  -> one image  : "<pdf> - p1.png"
  - a wide page (two passports side by side) -> LEFT + RIGHT halves + the FULL
    page, so you can keep whichever is correct.

All results go into the 'to_review' folder.

WHAT YOU DO AFTER RUNNING THIS:
  1. Open the 'to_review' folder and look at the images.
  2. For each real passport, RENAME it to the client's full name
     (exactly like the Excel: "NAME SURNAME").
  3. MOVE it into the 'images' folder.
  4. Delete the leftovers you don't need (blank pages, the half you didn't use).
  5. Run run.bat as usual to stamp them.

Your original PDFs are moved into 'pdfs_split' (kept safe - never deleted),
so the next normal run won't list them as skipped anymore.

Run by double-clicking split_pdfs.bat, or:  python split_pdfs.py
"""

import shutil
import sys
from pathlib import Path

try:
    import fitz  # PyMuPDF
except ImportError:
    sys.exit("This needs PyMuPDF. Install it with:  pip install pymupdf")

from PIL import Image

HERE = Path(__file__).parent
DPI = 200                 # render quality (good for printing)
WIDE_RATIO = 1.15         # width > height * this  -> treat as two side by side


def render_page(page):
    pix = page.get_pixmap(dpi=DPI)
    return Image.frombytes("RGB", (pix.width, pix.height), pix.samples)


def split_all(log=print):
    """Render every PDF in images/ into to_review/. Returns candidate count."""
    images_dir = HERE / "images"
    review = HERE / "to_review"
    done = HERE / "pdfs_split"
    review.mkdir(exist_ok=True)
    done.mkdir(exist_ok=True)

    # dedupe: on Windows glob is case-insensitive, so match by suffix instead
    pdfs = sorted(p for p in images_dir.iterdir()
                  if p.is_file() and p.suffix.lower() == ".pdf")
    if not pdfs:
        log("No PDF files in the 'images' folder. Nothing to split.")
        return 0

    made = 0
    failed = []   # (pdf name, plain-words problem)
    for pdf in pdfs:
        # one bad PDF must never stop the whole batch
        try:
            doc = fitz.open(pdf)
            if doc.needs_pass:
                doc.close()
                raise RuntimeError("password-protected")
            stem = pdf.stem
            log(f"\n{pdf.name}  ({doc.page_count} page(s))")

            for i in range(doc.page_count):
                img = render_page(doc.load_page(i))
                w, h = img.size
                base = f"{stem} - p{i + 1}"

                if w > h * WIDE_RATIO:
                    # wide page -> probably two passports side by side
                    img.crop((0, 0, w // 2, h)).save(review / f"{base} - LEFT.png")
                    img.crop((w // 2, 0, w, h)).save(review / f"{base} - RIGHT.png")
                    img.save(review / f"{base} - FULL.png")
                    log(f"   page {i + 1}: wide -> LEFT + RIGHT (+ FULL)")
                    made += 2
                else:
                    img.save(review / f"{base}.png")
                    log(f"   page {i + 1}: saved")
                    made += 1

            doc.close()
            shutil.move(str(pdf), str(done / pdf.name))
        except Exception as e:
            if "password" in str(e).lower():
                why = ("This PDF has a password. Ask for a copy without "
                       "a password. / Ce PDF a un mot de passe. Demandez "
                       "une copie sans mot de passe.")
            else:
                why = ("This PDF is damaged and cannot be opened. Ask for "
                       "the file again. / Ce PDF est endommagé. Redemandez "
                       "le fichier.")
            log(f"\n{pdf.name}  -> PROBLEM: {why}")
            failed.append((pdf.name, why))

    log(f"\nDone. {made} passport image(s) waiting in the 'to_review' folder.")
    if failed:
        log(f"{len(failed)} PDF(s) could not be read - they stay in 'images'.")
    log("Next: rename the real passports to the client name, move them into")
    log("'images', delete the leftovers, then run the check.")
    return made


def main():
    split_all(log=print)


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\n[ERROR] {e}")
    input("\nPress Enter to close...")
