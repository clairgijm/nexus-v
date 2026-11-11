"""LISEZMOI pages in HTML, from the simple Markdown the generators write.

Handles what those pages use: titles (#, ##), paragraphs, bullet lists (-), tables (| a | b |), **bold** and `code`.
The page is self-contained, in the style of the check reports.
Usage: write_html(path, lines) from a generator, or python lisezmoi.py <file.md> to convert a page by hand.
"""
import html
import re
import sys
from pathlib import Path

STYLE = """
body { font-family: "Segoe UI", system-ui, sans-serif; color: #1c2330; background: #ffffff; margin: 24px auto;
  max-width: 1100px; padding: 0 16px; font-size: 14px; line-height: 1.5; }
h1 { font-size: 1.6rem; margin: 0 0 12px; color: #0569a3; padding-bottom: 6px; border-bottom: 3px solid #10aed6; }
h2 { font-size: 1.15rem; margin: 28px 0 8px; padding: 7px 12px; color: #0569a3; background: #e8f4fb;
  border-left: 5px solid #0569a3; border-radius: 4px; }
table { border-collapse: collapse; width: 100%; margin: 8px 0 14px; }
th, td { border: 1px solid #d8dce3; padding: 6px 8px; text-align: left; vertical-align: top; }
th { background: #eef0f4; font-weight: 600; }
code { font-family: Consolas, monospace; font-size: .92em; background: #eef0f4; padding: 1px 4px; border-radius: 3px; }
p.file { margin: 18px 0 4px; }
@media print { body { margin: 0; max-width: none; } tr { break-inside: avoid; } }
"""


def inline(text):
    """Escaped text with **bold** and `code`."""
    out = html.escape(text)
    out = re.sub(r"`([^`]+)`", r"<code>\1</code>", out)
    return re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", out)


def to_html(lines):
    """The HTML page of Markdown lines."""
    body, title, i = [], "LISEZMOI", 0
    while i < len(lines):
        line = lines[i].rstrip()
        if not line.strip():
            i += 1
        elif line.startswith("# "):
            title = line[2:].strip()
            body.append(f"<h1>{inline(title)}</h1>")
            i += 1
        elif line.startswith("## "):
            body.append(f"<h2>{inline(line[3:].strip())}</h2>")
            i += 1
        elif line.startswith("|"):
            rows = []
            while i < len(lines) and lines[i].startswith("|"):
                cells = [c.strip() for c in lines[i].strip().strip("|").split("|")]
                if not all(re.fullmatch(r":?-{3,}:?", c) for c in cells):
                    rows.append(cells)
                i += 1
            head, rest = rows[0], rows[1:]
            body.append("<table><thead><tr>" + "".join(f"<th>{inline(c)}</th>" for c in head) + "</tr></thead><tbody>"
                        + "".join("<tr>" + "".join(f"<td>{inline(c)}</td>" for c in r) + "</tr>" for r in rest)
                        + "</tbody></table>")
        elif line.startswith("- "):
            items = []
            while i < len(lines) and lines[i].startswith("- "):
                items.append(lines[i][2:].strip())
                i += 1
            body.append("<ul>" + "".join(f"<li>{inline(t)}</li>" for t in items) + "</ul>")
        else:
            text = []
            while i < len(lines) and lines[i].strip() and not lines[i].startswith(("#", "|", "- ")):
                text.append(lines[i].strip())
                i += 1
            css = ' class="file"' if len(text) == 1 and text[0].startswith("**") else ""
            body.append(f"<p{css}>{inline(' '.join(text))}</p>")
    return (f'<!doctype html><html lang="fr"><head><meta charset="utf-8">'
            f'<meta name="viewport" content="width=device-width, initial-scale=1"><title>{html.escape(title)}</title>'
            f"<style>{STYLE}</style></head><body>\n" + "\n".join(body) + "\n</body></html>\n")


def write_html(path, lines):
    """Write the HTML page of Markdown `lines` (a list of lines, or one text) at `path` (.html)."""
    if isinstance(lines, str):
        lines = lines.splitlines()
    Path(path).write_text(to_html(lines), encoding="utf-8")


if __name__ == "__main__":
    source = Path(sys.argv[1])
    write_html(source.with_suffix(".html"), source.read_text(encoding="utf-8"))
    print("écrit :", source.with_suffix(".html"))
