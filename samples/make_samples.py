"""Generate the sample NDA files used for demos and tests.

  nda_v1_ours.docx                 our last version (clean)
  nda_v2_counterparty_tracked.docx counterparty markup as Word tracked changes
  nda_v2_counterparty_clean.docx   same changes, sent as a clean copy (for compare mode)

Each paragraph is a list of segments: ("=", text) unchanged, ("-", text)
deleted by the counterparty, ("+", text) inserted by the counterparty.
A paragraph whose segments are all "+" is a new paragraph.
"""
import sys
from datetime import datetime, timezone
from pathlib import Path
from xml.sax.saxutils import escape

from docx import Document
from docx.oxml import parse_xml
from docx.shared import Pt

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
AUTHOR = "Northwind Legal"
DATE = datetime(2026, 11, 12, 9, 30, tzinfo=timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

H = "heading"
P = [
    (H, [("=", "Mutual Non-Disclosure Agreement")]),
    ("", [("=", 'between Lumen Analytics GmbH, Berlin ("Lumen") and Northwind Systems Ltd., London ("Northwind")')]),
    (H, [("=", "1. Purpose")]),
    ("", [("=", "1.1 The parties "), ("-", "wish"), ("+", "intend"),
          ("=", ' to evaluate a potential collaboration on workforce analytics software (the "Purpose").')]),
    (H, [("=", "2. Confidential Information")]),
    ("", [("=", '2.1 "Confidential Information" means all information disclosed by a party, whether in writing'),
          ("-", ", orally or in any other form,"),
          ("=", " that is marked as confidential"),
          ("-", " or would reasonably be understood to be confidential"),
          ("=", ".")]),
    (H, [("=", "3. Exclusions")]),
    ("", [("=", "3.1 Confidential Information does not include information that (a) is or becomes public without breach of this Agreement, (b) was lawfully known to the recipient before disclosure, (c) is lawfully received from a third party without restriction, "),
          ("-", "or "),
          ("=", "(d) is independently developed without use of the Confidential Information"),
          ("+", ", or (e) is retained in the unaided memory of the recipient's personnel (residuals)"),
          ("=", ".")]),
    (H, [("=", "4. Obligations")]),
    ("", [("=", "4.1 The recipient shall use Confidential Information solely for the Purpose and protect it with at least reasonable care.")]),
    (H, [("=", "5. Term")]),
    ("", [("=", "5.1 The obligations under this Agreement continue for "), ("-", "five (5)"), ("+", "two (2)"),
          ("=", " years after the last disclosure"),
          ("-", "; for trade secrets, for as long as they remain trade secrets"),
          ("=", ".")]),
    (H, [("=", "6. Return of Information")]),
    ("", [("=", "6.1 Upon request, the recipient shall return or destroy all Confidential Information within "),
          ("-", "fourteen (14)"), ("+", "thirty (30)"),
          ("=", " days and confirm this in writing"),
          ("+", ", save for copies retained in automatic electronic backups, which remain subject to this Agreement until deleted"),
          ("=", ".")]),
    (H, [("=", "7. Remedies")]),
    ("", [("=", "7.1 Each party acknowledges that damages may be an insufficient remedy and that the disclosing party may seek injunctive relief.")]),
    ("", [("+", "7.2 In no event shall either party's aggregate liability under this Agreement exceed EUR 10,000.")]),
    (H, [("=", "8. Governing Law")]),
    ("", [("=", "8.1 This Agreement is governed by "), ("-", "German law"), ("+", "the laws of England and Wales"),
          ("=", ". The courts of "), ("-", "Berlin"), ("+", "London"), ("=", " have exclusive jurisdiction.")]),
    (H, [("+", "9. Non-Solicitation")]),
    ("", [("+", "9.1 For twelve (12) months after the last disclosure, neither party shall solicit or employ any employee of the other party.")]),
]

_rev_id = [100]


def _next_id():
    _rev_id[0] += 1
    return _rev_id[0]


def run_xml(text, deleted=False):
    tag = "w:delText" if deleted else "w:t"
    return f'<w:r {W}><{tag} xml:space="preserve">{escape(text)}</{tag}></w:r>'


def build(mode, out_path):
    """mode: 'v1' (ours), 'tracked', 'v2clean'."""
    doc = Document()
    doc.styles["Normal"].font.name = "Calibri"
    doc.styles["Normal"].font.size = Pt(11)
    # remove the default empty first paragraph
    body = doc.element.body
    for p in list(body.iterchildren()):
        if p.tag.endswith("}p"):
            body.remove(p)

    for style, segs in P:
        all_new = all(s[0] == "+" for s in segs)
        if mode == "v1" and all_new:
            continue
        para = doc.add_paragraph(style="Heading 2" if style == H else None)
        if mode == "tracked" and all_new:
            # mark the paragraph mark itself as inserted
            ppr = para._p.get_or_add_pPr()
            ppr.append(parse_xml(
                f'<w:rPr {W}><w:ins w:id="{_next_id()}" w:author="{AUTHOR}" w:date="{DATE}"/></w:rPr>'))
        for kind, text in segs:
            if kind == "=":
                para._p.append(parse_xml(run_xml(text)))
            elif kind == "-":
                if mode == "v1":
                    para._p.append(parse_xml(run_xml(text)))
                elif mode == "tracked":
                    para._p.append(parse_xml(
                        f'<w:del {W} w:id="{_next_id()}" w:author="{AUTHOR}" w:date="{DATE}">'
                        f'<w:r><w:delText xml:space="preserve">{escape(text)}</w:delText></w:r></w:del>'))
            elif kind == "+":
                if mode == "v2clean":
                    para._p.append(parse_xml(run_xml(text)))
                elif mode == "tracked":
                    para._p.append(parse_xml(
                        f'<w:ins {W} w:id="{_next_id()}" w:author="{AUTHOR}" w:date="{DATE}">'
                        f'<w:r><w:t xml:space="preserve">{escape(text)}</w:t></w:r></w:ins>'))
    doc.save(out_path)


if __name__ == "__main__":
    out = Path(sys.argv[1] if len(sys.argv) > 1 else Path(__file__).parent)
    build("v1", out / "nda_v1_ours.docx")
    build("tracked", out / "nda_v2_counterparty_tracked.docx")
    build("v2clean", out / "nda_v2_counterparty_clean.docx")
    print("wrote samples to", out)
