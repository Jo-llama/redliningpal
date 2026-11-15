"""Write the email scenarios used by the intake tests (RFC 822 .eml files).

Addresses match config/intake.example.json. For a live Gmail demo you send
the same kind of mails yourself; see README "Gmail demo".
"""
import sys
from email.message import EmailMessage
from email.utils import format_datetime
from datetime import datetime, timezone, timedelta
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
SPA = ROOT / "samples" / "spa"
OURS = "Jo Mackiewicz <buyer.counsel@halden-legal.example>"
THEIRS = "Mark Hughes <m.hughes@northwind-legal.example>"
THEIRS2 = "Amy Chen <a.chen@northwind-legal.example>"
INTAKE = "redline-intake@halden-legal.example"
DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
T0 = datetime(2026, 11, 10, 9, 0, tzinfo=timezone.utc)


def mail(path, *, frm, to, subject, body, msg_id, when, cc=INTAKE, reply_to=None, refs=None, attachments=()):
    m = EmailMessage()
    m["From"], m["To"], m["Subject"] = frm, to, subject
    if cc:
        m["Cc"] = cc
    m["Date"] = format_datetime(when)
    m["Message-ID"] = msg_id
    if reply_to:
        m["In-Reply-To"] = reply_to
    if refs:
        m["References"] = " ".join(refs)
    m.set_content(body)
    for fpath, name, ctype in attachments:
        maintype, subtype = ctype.split("/", 1)
        m.add_attachment(Path(fpath).read_bytes(), maintype=maintype, subtype=subtype, filename=name)
    Path(path).write_bytes(bytes(m))


def main(out):
    out = Path(out)
    out.mkdir(parents=True, exist_ok=True)
    subj = "[SPA-Nordlicht] Share Purchase Agreement - Buyer draft"

    # 01 our first draft goes out (intake in CC) -> registers v1 (ours)
    mail(out / "01_ours_v1.eml", frm=OURS, to=THEIRS, subject=subj, msg_id="<e01@halden-legal.example>", when=T0,
         body="Dear Mark,\n\nplease find attached our first draft of the SPA.\n\nBest regards\nJo",
         attachments=[(SPA / "spa_v1_ours.docx", "SPA_Nordlicht_v1.docx", DOCX)])

    # 02 Seller's markup in the same thread -> v2 (theirs), base v1
    mail(out / "02_theirs_v2.eml", frm=THEIRS, to=OURS, subject="RE: " + subj, msg_id="<e02@northwind-legal.example>",
         when=T0 + timedelta(days=2), reply_to="<e01@halden-legal.example>", refs=["<e01@halden-legal.example>"],
         body="Dear Jo,\n\nattached our comments on the draft in track changes.\n\nKind regards\nMark",
         attachments=[(SPA / "spa_v2_theirs_tracked.docx", "SPA_Nordlicht_v1_NW comments.docx", DOCX)])

    # 03 "corrected" version next morning -> v2.1 supersedes v2, same base
    mail(out / "03_theirs_v2-1_corrected.eml", frm=THEIRS, to=OURS, subject="RE: " + subj,
         msg_id="<e03@northwind-legal.example>", when=T0 + timedelta(days=3),
         reply_to="<e02@northwind-legal.example>", refs=["<e01@halden-legal.example>", "<e02@northwind-legal.example>"],
         body="Dear Jo,\n\nplease disregard our draft of yesterday - the attached version replaces it. "
              "We had missed a point on the cap.\n\nKind regards\nMark",
         attachments=[(SPA / "spa_v2-1_theirs_corrected.docx", "SPA_Nordlicht_NW comments_corrected.docx", DOCX)])

    # 04 duplicate delivery of 02 (same Message-ID)
    (out / "04_duplicate_of_02.eml").write_bytes((out / "02_theirs_v2.eml").read_bytes())

    # 05 colleague sends a clean copy + PDF in a NEW thread, no subject tag, no stamp
    pdf = out / "_comparison.pdf"
    pdf.write_bytes(b"%PDF-1.4\n% placeholder comparison\n")
    mail(out / "05_theirs_clean_copy_new_thread.eml", frm=THEIRS2, to=OURS, subject="Nordlicht - clean version for your convenience",
         msg_id="<e05@northwind-legal.example>", when=T0 + timedelta(days=2, hours=3),
         body="Hi Jo,\n\nfor convenience, a clean version of Mark's comments and a comparison PDF.\n\nAmy",
         attachments=[(SPA / "spa_v2_theirs_clean.docx", "SPA_Nordlicht_clean.docx", DOCX),
                      (pdf, "SPA_Nordlicht_comparison.pdf", "application/pdf")])

    # 06 no attachment
    mail(out / "06_no_attachment.eml", frm=THEIRS, to=OURS, subject="RE: " + subj, msg_id="<e06@northwind-legal.example>",
         when=T0 + timedelta(days=4), reply_to="<e03@northwind-legal.example>",
         body="Jo, are you free for a call on Thursday?\n\nMark")

    # 07 unknown sender, unrelated contract -> triage
    mail(out / "07_unknown_sender.eml", frm="Someone <legal@unknown-co.example>", to=OURS, subject="NDA for review",
         msg_id="<e07@unknown-co.example>", when=T0 + timedelta(days=5),
         body="Please review the attached NDA.",
         attachments=[(ROOT / "samples" / "nda_v2_counterparty_tracked.docx", "NDA.docx", DOCX)])
    print("wrote fixtures to", out)


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else Path(__file__).parent / "emails")
