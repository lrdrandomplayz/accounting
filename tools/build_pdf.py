"""Convert the Markdown notes and model questions to printable PDFs.

Run from the repository root:  python3 tools/build_pdf.py
Needs the `markdown` package and a Chromium or Chrome browser.
Set CHROME=/path/to/chrome if it is not found automatically.
"""
import os
import shutil
import subprocess
import tempfile
from pathlib import Path

import markdown

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "pdf"
SOURCES = sorted((ROOT / "notes").glob("*.md")) + [
    ROOT / "model-questions" / "Model_Questions.md",
    ROOT / "model-questions" / "Answer_Key.md",
    ROOT / "model-questions" / "Mock_Test.md",
    ROOT / "model-questions" / "Mock_Test_Answers.md",
]

CSS = """
@page { size: A4; margin: 16mm 14mm; }
body { font-family: Arial, Helvetica, 'WenQuanYi Zen Hei', 'Noto Sans CJK SC', sans-serif; font-size: 10.5pt; line-height: 1.45; color: #1a1a1a; }
h1 { font-size: 20pt; color: #1f3864; border-bottom: 3px solid #1f3864; padding-bottom: 4px; }
h2 { font-size: 14pt; color: #1f3864; margin-top: 18px; border-bottom: 1px solid #c9d3e6; padding-bottom: 2px; }
h3 { font-size: 12pt; color: #2f5496; margin-top: 14px; }
h2, h3 { break-after: avoid; }
table { border-collapse: collapse; width: 100%; margin: 8px 0 12px; font-size: 9.5pt; break-inside: avoid; }
th, td { border: 1px solid #b4bfd4; padding: 4px 6px; vertical-align: top; }
th { background: #d9e1f2; text-align: left; }
tr:nth-child(even) td { background: #f5f7fb; }
blockquote { margin: 8px 0; padding: 6px 12px; background: #fff8dc; border-left: 4px solid #e0b300; }
code { font-family: Consolas, monospace; font-size: 9pt; background: #eef1f6; padding: 0 3px; }
hr { border: none; border-top: 1px dashed #999; margin: 18px 0; }
.foot { margin-top: 24px; font-size: 8.5pt; color: #666; text-align: center; }
"""


def find_chrome():
    if os.environ.get("CHROME"):
        return os.environ["CHROME"]
    for name in ("chromium", "chromium-browser", "google-chrome", "chrome"):
        if shutil.which(name):
            return shutil.which(name)
    for p in Path("/opt/pw-browsers").glob("chromium-*/chrome-linux/chrome"):
        return str(p)
    raise SystemExit("Chromium not found. Set CHROME=/path/to/chrome")


def main():
    OUT.mkdir(exist_ok=True)
    chrome = find_chrome()
    with tempfile.TemporaryDirectory() as tmp:
        for src in SOURCES:
            body = markdown.markdown(src.read_text(encoding="utf-8"), extensions=["tables"])
            html = (
                f"<!doctype html><html><head><meta charset='utf-8'><title>{src.stem}</title>"
                f"<style>{CSS}</style></head><body>{body}"
                "<p class='foot'>Senior 1 Accounting notes. Free to use and share.</p></body></html>"
            )
            page = Path(tmp) / f"{src.stem}.html"
            page.write_text(html, encoding="utf-8")
            target = OUT / f"{src.stem}.pdf"
            subprocess.run(
                [chrome, "--headless", "--no-sandbox", "--disable-gpu", "--no-pdf-header-footer",
                 f"--print-to-pdf={target}", page.as_uri()],
                check=True, capture_output=True,
            )
            print("wrote", target.relative_to(ROOT))


if __name__ == "__main__":
    main()
