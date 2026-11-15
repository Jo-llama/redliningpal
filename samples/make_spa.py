"""Generate the SPA negotiation scenario (Buyer = our client).

  spa/spa_v1_ours.docx              our draft, stamped rc-ver = V1_ID
  spa/spa_v2_theirs_tracked.docx    Seller's markup (tracked), stamp kept
  spa/spa_v2-1_theirs_corrected.docx Seller's "corrected" markup: two more
                                    tracked changes + one UNTRACKED edit
                                    (fraud carve-out silently removed)
  spa/spa_v2_theirs_clean.docx      v2 with changes accepted, no stamp
                                    ("clean copy for your convenience")
"""
import sys
import zipfile
from datetime import datetime, timezone
from pathlib import Path
from xml.sax.saxutils import escape

from docx import Document
from docx.oxml import parse_xml
from docx.shared import Pt

V1_ID = "7f3a9c2e-5b1d-4c8a-9e6f-0a1b2c3d4e5f"
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
AUTHOR = "Northwind Legal"
H = "heading"

# --- our draft (v1) -------------------------------------------------------
V1 = [
    (H, "Share Purchase Agreement"),
    ("", 'between Northwind Holdings Ltd. (the "Seller") and Halden Capital Partners GmbH (the "Buyer") regarding Nordlicht Software GmbH (the "Company")'),
    (H, "1. Sale and Purchase"),
    ("", '1.1 The Seller sells to the Buyer all shares in the Company (the "Shares") with economic effect as of the Locked Box Date.'),
    (H, "2. Purchase Price"),
    ("", '2.1 The purchase price is EUR 48,000,000 (the "Purchase Price"), payable on the Closing Date.'),
    ("", "2.2 The Seller shall pay to the Buyer an amount equal to any Leakage occurring between the Locked Box Date and the Closing Date."),
    (H, "3. Warranties"),
    ("", "3.1 The Seller warrants to the Buyer that the statements set out in Annex 3.1 are true and accurate on the Signing Date and on the Closing Date."),
    ("", "3.2 Where a warranty is qualified by the Seller's knowledge, the Seller's knowledge includes the actual knowledge of the managing directors and what they would have known after due enquiry."),
    (H, "4. Limitation of Liability"),
    ("", "4.1 The Seller is not liable for individual claims below EUR 25,000 (de minimis)."),
    ("", "4.2 The Seller is liable only if the aggregate of all claims exceeds EUR 240,000 (basket), in which case the Seller is liable from the first euro."),
    ("", "4.3 The Seller's aggregate liability for breach of warranties is limited to 30% of the Purchase Price."),
    ("", "4.4 None of the limitations in this Clause 4 apply in case of intent or fraud."),
    (H, "5. Limitation Period"),
    ("", "5.1 Claims for breach of warranties become time-barred twenty-four (24) months after the Closing Date; tax claims become time-barred six (6) months after the final tax assessment."),
    (H, "6. Tax Indemnity"),
    ("", "6.1 The Seller shall indemnify the Buyer against all Taxes of the Company relating to periods ending on or before the Locked Box Date."),
    (H, "7. Material Adverse Change"),
    ("", "7.1 The Buyer may withdraw from this Agreement if a Material Adverse Change occurs before Closing."),
    (H, "8. Non-Compete"),
    ("", "8.1 For three (3) years after the Closing Date, the Seller shall not compete with the business of the Company."),
    (H, "9. Employees"),
    ("", "9.1 The Seller shall procure that the managing directors listed in Annex 9.1 remain employed until the Closing Date."),
    (H, "10. Governing Law and Disputes"),
    ("", "10.1 This Agreement is governed by German law. Disputes are finally settled under the DIS Arbitration Rules; the seat of arbitration is Frankfurt am Main."),
]

# --- Seller's markup (v2): segments per paragraph, keyed by clause number --
V2 = {
    "2.2": [("=", "2.2 The Seller shall pay to the Buyer an amount equal to any Leakage occurring between the Locked Box Date and the Closing Date"),
            ("+", ", provided that the Buyer notifies the Seller of such Leakage within three (3) months after the Closing Date"),
            ("=", ".")],
    "3.1": [("=", "3.1 The Seller warrants to the Buyer that the statements set out in Annex 3.1 are true and accurate on the Signing Date"),
            ("-", " and on the Closing Date"), ("=", ".")],
    "3.2": [("=", "3.2 Where a warranty is qualified by the Seller's knowledge, the Seller's knowledge includes the actual knowledge of the managing directors"),
            ("-", " and what they would have known after due enquiry"), ("=", ".")],
    "4.1": [("=", "4.1 The Seller is not liable for individual claims below EUR "), ("-", "25,000"), ("+", "100,000"), ("=", " (de minimis).")],
    "4.2": [("=", "4.2 The Seller is liable only if the aggregate of all claims exceeds EUR 240,000 (basket), in which case the Seller is liable "),
            ("-", "from the first euro"), ("+", "only for the excess"), ("=", ".")],
    "4.3": [("=", "4.3 The Seller's aggregate liability for breach of warranties is limited to "), ("-", "30%"), ("+", "10%"), ("=", " of the Purchase Price.")],
    "5.1": [("=", "5.1 Claims for breach of warranties become time-barred "), ("-", "twenty-four (24)"), ("+", "twelve (12)"),
            ("=", " months after the Closing Date; tax claims become time-barred six (6) months after the final tax assessment.")],
    "6.1": [("=", "6.1 The Seller shall indemnify the Buyer against all Taxes of the Company relating to periods ending on or before the Locked Box Date"),
            ("+", ", to the extent such Taxes exceed the provisions in the Locked Box Accounts"), ("=", ".")],
    "7.1": [("-", "7.1 The Buyer may withdraw from this Agreement if a Material Adverse Change occurs before Closing.")],
    "8.1": [("=", "8.1 For "), ("-", "three (3) years"), ("+", "one (1) year"), ("=", " after the Closing Date, the Seller shall not compete with the business of the Company.")],
    "10.1": [("=", "10.1 This Agreement is governed by German law. Disputes are finally settled under the "), ("-", "DIS Arbitration Rules"), ("+", "LCIA Rules"),
             ("=", "; the seat of arbitration is "), ("-", "Frankfurt am Main"), ("+", "London"), ("=", ".")],
}
# new paragraph inserted after 9.1
V2_INSERT_AFTER = {"9.1": [("+", "9.2 The Buyer shall not terminate employment relationships of employees of the Company for operational reasons (betriebsbedingte Kündigung) for twenty-four (24) months after the Closing Date.")]}

# --- Seller's "corrected" markup (v2.1) = v2 with overrides ---------------
V2_1 = dict(V2)
V2_1["8.1"] = [("=", "8.1 For "), ("-", "three (3) years"), ("+", "two (2) years"), ("=", " after the Closing Date, the Seller shall not compete with the business of the Company.")]
V2_1["4.3"] = [("=", "4.3 The Seller's aggregate liability for breach of warranties is limited to "), ("-", "30%"), ("+", "10%"), ("=", " of the Purchase Price"),
               ("+", "; claims covered by the W&I insurance policy shall be pursued exclusively against the insurer"), ("=", ".")]
# UNTRACKED edit: "intent or fraud" -> "intent", typed with track changes off
V2_1["4.4"] = [("=", "4.4 None of the limitations in this Clause 4 apply in case of intent.")]


def clause_no(text):
    first = text.split(" ", 1)[0]
    return first if first[0].isdigit() and "." in first[:-1] else None


def version_paras(spec, inserts):
    out = []
    for style, text in V1:
        key = clause_no(text) if style != H else None
        segs = spec.get(key, [("=", text)]) if key else [("=", text)]
        out.append((style, segs))
        if key in inserts:
            out.append(("", inserts[key]))
    return out


_rev = [100]


def _rid():
    _rev[0] += 1
    return _rev[0]


def write_docx(paras, mode, path, stamp=None):
    """mode: 'tracked' | 'clean' (theirs accepted) | 'plain'."""
    date = datetime(2026, 11, 12, 9, 30, tzinfo=timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    doc = Document()
    doc.styles["Normal"].font.name = "Calibri"
    doc.styles["Normal"].font.size = Pt(11)
    body = doc.element.body
    for p in list(body.iterchildren()):
        if p.tag.endswith("}p"):
            body.remove(p)
    for style, segs in paras:
        if mode == "clean" and all(k == "-" for k, _ in segs):
            continue
        para = doc.add_paragraph(style="Heading 2" if style == H else None)
        whole_new = all(k == "+" for k, _ in segs)
        whole_del = all(k == "-" for k, _ in segs)
        if mode == "tracked" and (whole_new or whole_del):
            tag = "ins" if whole_new else "del"
            para._p.get_or_add_pPr().append(parse_xml(
                f'<w:rPr {W}><w:{tag} w:id="{_rid()}" w:author="{AUTHOR}" w:date="{date}"/></w:rPr>'))
        for kind, text in segs:
            t = escape(text)
            if kind == "=" or (mode == "clean" and kind == "+") or (mode == "plain" and kind != "+"):
                para._p.append(parse_xml(f'<w:r {W}><w:t xml:space="preserve">{t}</w:t></w:r>'))
            elif mode == "tracked" and kind == "-":
                para._p.append(parse_xml(f'<w:del {W} w:id="{_rid()}" w:author="{AUTHOR}" w:date="{date}">'
                                         f'<w:r><w:delText xml:space="preserve">{t}</w:delText></w:r></w:del>'))
            elif mode == "tracked" and kind == "+":
                para._p.append(parse_xml(f'<w:ins {W} w:id="{_rid()}" w:author="{AUTHOR}" w:date="{date}">'
                                         f'<w:r><w:t xml:space="preserve">{t}</w:t></w:r></w:ins>'))
    doc.save(path)
    if stamp:
        add_stamp(path, stamp)


def add_stamp(path, value):
    """Write custom document property rc-ver (what the add-in / stamp CLI does)."""
    custom = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
              '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" '
              'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
              '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="rc-ver">'
              f'<vt:lpwstr>{value}</vt:lpwstr></property></Properties>')
    src = Path(path)
    tmp = src.with_suffix(".tmp")
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "[Content_Types].xml":
                data = data.replace(b"</Types>", b'<Override PartName="/docProps/custom.xml" '
                                    b'ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/></Types>')
            elif item.filename == "_rels/.rels":
                data = data.replace(b"</Relationships>", b'<Relationship Id="rIdCustom" '
                                    b'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" '
                                    b'Target="docProps/custom.xml"/></Relationships>')
            zout.writestr(item, data)
        zout.writestr("docProps/custom.xml", custom)
    tmp.replace(src)


if __name__ == "__main__":
    out = Path(sys.argv[1] if len(sys.argv) > 1 else Path(__file__).parent / "spa")
    out.mkdir(parents=True, exist_ok=True)
    v1 = [(s, [("=", t)]) for s, t in V1]
    write_docx(v1, "plain", out / "spa_v1_ours.docx", stamp=V1_ID)
    write_docx(version_paras(V2, V2_INSERT_AFTER), "tracked", out / "spa_v2_theirs_tracked.docx", stamp=V1_ID)
    write_docx(version_paras(V2_1, V2_INSERT_AFTER), "tracked", out / "spa_v2-1_theirs_corrected.docx", stamp=V1_ID)
    write_docx(version_paras(V2, V2_INSERT_AFTER), "clean", out / "spa_v2_theirs_clean.docx")
    print("wrote SPA samples to", out)
