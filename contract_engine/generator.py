import os, re, shutil, zipfile, uuid, tempfile, io
from datetime import date
from lxml import etree
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders

BASE_DIR     = os.path.dirname(__file__)
UNPACKED_DIR = os.path.join(BASE_DIR, "unpacked")

SERVICES_ALL = [f"SVC_{str(i).zfill(2)}" for i in range(1, 29)]
OPTIONAL_FIELDS = [
    "TOTAL_FEE","INITIAL_PAYMENT","SECOND_PAYMENT","SECOND_PAYMENT_DATE",
    "TRUST_PAYMENT","TRUST_PAYMENT_DATE","FINAL_PAYMENT","PGY2_FEE",
    "EXTERNSHIP_MONTHS","ADDITIONAL_NOTES",
]

_jobs: dict = {}


def _pack_docx(work_dir: str) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(work_dir):
            dirs[:] = sorted([d for d in dirs if not d.startswith(".")])
            for fname in sorted(files):
                full = os.path.join(root, fname)
                zf.write(full, os.path.relpath(full, work_dir))
    return buf.getvalue()


def send_email(to_addr: str, candidate_name: str, docx_bytes: bytes, rep_mode=False) -> tuple:
    host = os.environ.get("SMTP_HOST", "smtp.gmail.com")
    user = os.environ.get("SMTP_USER", "wcanormg@gmail.com")
    pw   = os.environ.get("SMTP_PASS", "")
    port = int(os.environ.get("SMTP_PORT", "587"))
    frm  = os.environ.get("SMTP_FROM", user)
    if not pw:
        return False, "SMTP password not configured"

    msg = MIMEMultipart()
    msg["From"]    = frm
    msg["To"]      = to_addr
    msg["Subject"] = f"{'[REP SUBMISSION] ' if rep_mode else ''}PSA Contract — {candidate_name}"
    body = (
        f"A sales rep has submitted contract details for {candidate_name}.\n\n"
        "Please find the generated PSA contract attached for your review.\n\n"
        "Residents Medical Group"
        if rep_mode else
        f"Please find the generated PSA contract for {candidate_name} attached.\n\n"
        "Review before sending to the candidate.\n\n"
        "Residents Medical Group"
    )
    msg.attach(MIMEText(body, "plain"))
    safe = candidate_name.replace(" ", "_")
    part = MIMEBase("application", "octet-stream")
    part.set_payload(docx_bytes)
    encoders.encode_base64(part)
    part.add_header("Content-Disposition", "attachment", filename=f"PSA_{safe}.docx")
    msg.attach(part)
    try:
        with smtplib.SMTP(host, port, timeout=15) as s:
            s.ehlo(); s.starttls(); s.login(user, pw)
            s.sendmail(frm, to_addr, msg.as_string())
        return True, "sent"
    except Exception as e:
        return False, str(e)


def generate(field_values: dict, selected_ids: list,
             send_to: str = "", rep_mode: bool = False) -> dict:
    selected_ids = set(selected_ids)
    if not field_values.get("EXECUTION_DATE", "").strip():
        field_values["EXECUTION_DATE"] = date.today().strftime("%B %d, %Y")

    job_id  = uuid.uuid4().hex
    tmp_dir = tempfile.mkdtemp(prefix=f"psa_{job_id}_")
    work_dir = os.path.join(tmp_dir, "work")

    try:
        shutil.copytree(UNPACKED_DIR, work_dir)
        doc_path = os.path.join(work_dir, "word", "document.xml")
        with open(doc_path, "r", encoding="utf-8") as f:
            xml = f.read()

        for sid in SERVICES_ALL:
            if sid not in selected_ids:
                xml = re.sub(rf'<!--BEGIN_{sid}-->.*?<!--END_{sid}-->', '', xml, flags=re.DOTALL)
        xml = re.sub(r'<!--BEGIN_ADDITIONAL_NOTES-->.*?<!--END_ADDITIONAL_NOTES-->\s*', '', xml, flags=re.DOTALL)

        for key, value in field_values.items():
            if value and str(value).strip():
                safe = str(value).replace('&','&amp;').replace('<','&lt;').replace('>','&gt;')
                xml = xml.replace('{{'+key+'}}', safe)

        empty = [k for k in OPTIONAL_FIELDS if not field_values.get(k, "").strip()]
        if empty:
            W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
            tree = etree.fromstring(xml.encode('utf-8'))
            to_remove = []
            for para in tree.iter(f'{{{W}}}p'):
                para_text = ''.join((t.text or '') for t in para.iter(f'{{{W}}}t'))
                for key in empty:
                    if ('{{'+key+'}}') in para_text:
                        to_remove.append(para); break
            for para in to_remove:
                p = para.getparent()
                if p is not None: p.remove(para)
            xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                   + etree.tostring(tree, encoding='unicode', xml_declaration=False))

        xml = re.sub(r'<!--.*?-->', '', xml, flags=re.DOTALL)
        with open(doc_path, "w", encoding="utf-8") as f:
            f.write(xml)

        docx_bytes = _pack_docx(work_dir)
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)

    cname = field_values.get("CANDIDATE_FULL_NAME", "Contract")
    _jobs[job_id] = (docx_bytes, cname)
    old = list(_jobs.keys())[:-50]
    for k in old: del _jobs[k]

    email_ok, email_msg = False, "no address provided"
    if send_to:
        email_ok, email_msg = send_email(send_to, cname, docx_bytes, rep_mode=rep_mode)

    return {"job_id": job_id, "email_sent": email_ok, "email_msg": email_msg,
            "execution_date": field_values.get("EXECUTION_DATE", "")}


def get_docx(job_id: str):
    return _jobs.get(job_id, (None, None))
