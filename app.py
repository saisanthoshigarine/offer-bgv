import os, io, json, uuid, base64, hashlib, re, threading, shutil, logging
from datetime import datetime, timedelta

# ── Load .env FIRST before any os.environ.get() calls ────────────────────────
from dotenv import load_dotenv
load_dotenv()

import pandas as pd
from flask import (Flask, request, jsonify, session, redirect,
                   render_template, send_file)
from werkzeug.utils import secure_filename

# reportlab
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table,
                                 TableStyle, HRFlowable)
from reportlab.pdfgen import canvas as rlc

# Brevo / email
import sib_api_v3_sdk
from sib_api_v3_sdk.rest import ApiException

# MongoDB
from pymongo import MongoClient

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY')
app.config['PERMANENT_SESSION_LIFETIME'] = timedelta(minutes=30)

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[logging.StreamHandler()]
)
logger = logging.getLogger(__name__)


# ── MongoDB setup ─────────────────────────────────────────────────────────────
MONGO_URI    = os.environ.get('MONGODB_URI')
_mongo_client = None

def get_db():
    global _mongo_client
    if _mongo_client is None:
        _mongo_client = MongoClient(MONGO_URI)
    return _mongo_client['offerflow']


# ── MongoDB CRUD helpers (drop-in replacements for JSON helpers) ───────────────
def get_users():
    return {u['id']: u for u in get_db().users.find({}, {'_id': 0})}

def get_cands():
    return {c['id']: c for c in get_db().candidates.find({}, {'_id': 0})}

def get_verifs():
    return {v['id']: v for v in get_db().verifications.find({}, {'_id': 0})}

def get_tokens():
    return {t['token']: t for t in get_db().tokens.find({}, {'_id': 0})}

def get_email_history():
    return {e['id']: e for e in get_db().email_history.find({}, {'_id': 0})}

def save_users(d):
    db = get_db()
    for uid, u in d.items():
        db.users.replace_one({'id': uid}, u, upsert=True)

def save_cands(d):
    db = get_db()
    for cid, c in d.items():
        db.candidates.replace_one({'id': cid}, c, upsert=True)

def save_verifs(d):
    db = get_db()
    for vid, v in d.items():
        db.verifications.replace_one({'id': vid}, v, upsert=True)

def save_tokens(d):
    db = get_db()
    for token, t in d.items():
        db.tokens.replace_one({'token': token}, {**t, 'token': token}, upsert=True)

def save_email_history(d):
    db = get_db()
    for eid, e in d.items():
        db.email_history.replace_one({'id': eid}, e, upsert=True)


# ── Validation Helpers ────────────────────────────────────────────────────────
def validate_email(email):
    pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return re.match(pattern, email) is not None

def validate_required_fields(data, required_fields):
    missing = [field for field in required_fields if not data.get(field, '').strip()]
    return missing

def validate_file_size(file, max_size_mb=5):
    if file:
        file.seek(0, os.SEEK_END)
        size = file.tell()
        file.seek(0)
        return size <= max_size_mb * 1024 * 1024
    return False

def sanitize_string(s):
    if not s:
        return ''
    return re.sub(r'[<>"\']', '', str(s).strip())


BASE_URL      = os.environ.get('BASE_URL', 'http://localhost:5000')
BREVO_API_KEY = os.environ.get('BREVO_API_KEY')
SENDER_EMAIL  = os.environ.get('SENDER_EMAIL')
SENDER_NAME   = os.environ.get('SENDER_NAME')

BASE_DIR   = os.path.dirname(os.path.abspath(__file__))

# ── Use /tmp on Vercel (read-only FS), local dirs otherwise ──────────────────
IS_VERCEL  = bool(os.environ.get('VERCEL') or os.environ.get('VERCEL_ENV'))
UPLOAD_DIR = "/tmp/uploads"           if IS_VERCEL else os.path.join(BASE_DIR, 'uploads')
LETTER_DIR = "/tmp/generated_letters" if IS_VERCEL else os.path.join(BASE_DIR, 'generated_letters')

for d in [UPLOAD_DIR+'/letterheads', UPLOAD_DIR+'/excel',
          UPLOAD_DIR+'/documents', LETTER_DIR]:
    os.makedirs(d, exist_ok=True)


# ── Letterhead path resolver ──────────────────────────────────────────────────
def _resolve_lh_path(stored: str) -> str:
    if not stored:
        return ''
    normalised = stored.replace('\\', '/').strip()
    if os.path.isabs(normalised) and os.path.exists(normalised):
        return normalised
    fname  = os.path.basename(normalised)
    in_tmp = os.path.join(UPLOAD_DIR, 'letterheads', fname)
    if os.path.exists(in_tmp):
        return in_tmp
    base_rel = os.path.join(BASE_DIR, *normalised.split('/'))
    if os.path.exists(base_rel):
        return base_rel
    return ''


# ── Seed letterheads into /tmp on Vercel ─────────────────────────────────────
def init_data():
    """Copy letterhead files from project dir into /tmp (Vercel only)."""
    project_lh = os.path.join(BASE_DIR, 'uploads', 'letterheads')
    tmp_lh     = os.path.join(UPLOAD_DIR, 'letterheads')
    if os.path.isdir(project_lh):
        for f in os.listdir(project_lh):
            src = os.path.join(project_lh, f)
            dst = os.path.join(tmp_lh, f)
            if not os.path.exists(dst):
                shutil.copy2(src, dst)

init_data()


def hash_pw(pw): return hashlib.sha256(pw.encode()).hexdigest()

def validate_password(pw):
    return (len(pw) >= 8 and re.search(r'[A-Z]', pw)
            and re.search(r'[a-z]', pw) and re.search(r'[^A-Za-z0-9]', pw))

def current_user():
    uid = session.get('user_id')
    if not uid:
        return None
    result = get_db().users.find_one({'id': uid}, {'_id': 0})
    return result


# ── Offer letter patterns ─────────────────────────────────────────────────────
PATTERNS = {
    'classic': {
        'label':       'Classic',
        'accent':      '#0d1b3e',
        'accent2':     '#1a56db',
        'header_bg':   '#0d1b3e',
        'row_even':    '#f0f4ff',
        'row_odd':     '#ffffff',
        'font_header': 'Helvetica-Bold',
        'layout':      'classic',
        'desc':        'Traditional corporate style with centred header and blue-accented table.',
    },
    'modern': {
        'label':       'Modern',
        'accent':      '#6d28d9',
        'accent2':     '#7c3aed',
        'header_bg':   '#6d28d9',
        'row_even':    '#f5f3ff',
        'row_odd':     '#ffffff',
        'font_header': 'Helvetica-Bold',
        'layout':      'modern',
        'desc':        'Clean modern layout with purple palette and two-column detail cards.',
    },
    'minimal': {
        'label':       'Minimal',
        'accent':      '#374151',
        'accent2':     '#6b7280',
        'header_bg':   '#f9fafb',
        'row_even':    '#f3f4f6',
        'row_odd':     '#ffffff',
        'font_header': 'Helvetica',
        'layout':      'minimal',
        'desc':        'Light and airy with grey tones — ideal for creative industries.',
    },
    'executive': {
        'label':       'Executive',
        'accent':      '#7f1d1d',
        'accent2':     '#b91c1c',
        'header_bg':   '#7f1d1d',
        'row_even':    '#fff1f2',
        'row_odd':     '#ffffff',
        'font_header': 'Helvetica-Bold',
        'layout':      'executive',
        'desc':        'Premium dark-red executive format with formal letter structure.',
    },
    'custom': {
        'label':  'Custom',
        'layout': 'custom',
        'desc':   'Write your own offer letter content in the text area below.',
    },
}


# ── PDF Generator ─────────────────────────────────────────────────────────────
def generate_offer_pdf(candidate, hr_user, pattern_key='classic', custom_text=''):
    cid      = candidate['id']
    pat      = PATTERNS.get(pattern_key, PATTERNS['classic'])
    out_path = os.path.join(LETTER_DIR, f"offer_{cid}.pdf")
    today    = datetime.now().strftime('%d %B %Y')
    company  = hr_user['company_name']
    name     = candidate.get('name', '')
    role     = candidate.get('role', '')
    joining  = candidate.get('joining_date', '')
    salary   = candidate.get('salary', '')
    emp_type = candidate.get('employment_type', 'full_time').replace('_', ' ').title()
    lh_path  = _resolve_lh_path(hr_user.get('letterhead', ''))

    top_margin = 54 * mm if lh_path else 18 * mm

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4,
                            rightMargin=22*mm, leftMargin=22*mm,
                            topMargin=top_margin, bottomMargin=22*mm)

    styles  = getSampleStyleSheet()
    accent  = colors.HexColor(pat.get('accent',  '#0d1b3e'))
    accent2 = colors.HexColor(pat.get('accent2', '#1a56db'))
    row_even = colors.HexColor(pat.get('row_even', '#f0f4ff'))
    row_odd  = colors.HexColor(pat.get('row_odd',  '#ffffff'))

    def ps(n, **kw):
        return ParagraphStyle(n, parent=styles['Normal'], **kw)

    layout = pat.get('layout', 'classic')

    # ── CUSTOM ────────────────────────────────────────────────────────────────
    if layout == 'custom':
        body_style = ps('cb', fontName='Helvetica', fontSize=10.5, leading=18,
                        textColor=colors.HexColor('#1e293b'), spaceAfter=8)
        content = custom_text or (
            f"Dear {name},\n\nThis is your offer letter for {role} at {company}.\n\n"
            f"Joining Date: {joining}\nSalary: ₹{salary}\n\nWarm regards,\nHR Department")
        story = []
        for para in content.split('\n'):
            story.append(Paragraph(para.strip() or '&nbsp;', body_style))
        _build_pdf(doc, buf, story, lh_path, out_path)
        return out_path

    # ── Shared table ──────────────────────────────────────────────────────────
    tbl_data = [
        ['Field', 'Details'],
        ['Candidate Name',     name],
        ['Designation / Role', role],
        ['Joining Date',       joining],
        ['Annual CTC',         f'₹ {salary}'],
        ['Employment Type',    emp_type],
        ['Reporting Location', 'As communicated by HR'],
    ]

    # ── CLASSIC ───────────────────────────────────────────────────────────────
    if layout == 'classic':
        tbl = _make_table(tbl_data, accent, accent2, row_even, row_odd)
        story = [
            Paragraph(company,
                      ps('co', fontName='Helvetica-Bold', fontSize=20,
                         textColor=accent, alignment=TA_CENTER, spaceAfter=2)),
            Paragraph('Human Resources Department',
                      ps('sub', fontSize=10, textColor=colors.HexColor('#64748b'),
                         alignment=TA_CENTER, spaceAfter=14)),
            HRFlowable(width='100%', thickness=2.5, color=accent2, spaceAfter=10),
            Paragraph('OFFER OF EMPLOYMENT',
                      ps('h2', fontName='Helvetica-Bold', fontSize=15,
                         textColor=accent, alignment=TA_CENTER, spaceAfter=16)),
            Paragraph(f'Date: {today}',
                      ps('dt', fontSize=10, textColor=colors.HexColor('#64748b'),
                         spaceAfter=14)),
            Paragraph(f'Dear <b>{name}</b>,',
                      ps('b', fontSize=10.5, leading=17,
                         textColor=colors.HexColor('#1e293b'), spaceAfter=8)),
            Spacer(1, 4),
            Paragraph(
                f'We are delighted to extend this offer for the position of '
                f'<b>{role}</b> at <b>{company}</b>. We believe your skills '
                f'are an excellent fit for our team.',
                ps('b2', fontSize=10.5, leading=17,
                   textColor=colors.HexColor('#1e293b'), spaceAfter=8)),
            Spacer(1, 10),
            Paragraph('Your offer details:',
                      ps('bb', fontName='Helvetica-Bold', fontSize=10.5,
                         textColor=accent, spaceAfter=6)),
            Spacer(1, 6), tbl, Spacer(1, 14),
            Paragraph('This offer is contingent upon:',
                      ps('bb2', fontName='Helvetica-Bold', fontSize=10.5,
                         textColor=accent, spaceAfter=6)),
            Paragraph('• Successful completion of background verification.',  ps('b3', fontSize=10.5, leading=17, spaceAfter=4)),
            Paragraph('• Submission of all required documents before joining.', ps('b4', fontSize=10.5, leading=17, spaceAfter=4)),
            Paragraph("• Acceptance of the company's Code of Conduct.",        ps('b5', fontSize=10.5, leading=17, spaceAfter=8)),
            Spacer(1, 12),
            Paragraph('Please accept via the <b>Accept Offer</b> button in your email.',
                      ps('b6', fontSize=10.5, leading=17, spaceAfter=24)),
            HRFlowable(width='100%', thickness=0.5,
                       color=colors.HexColor('#e2e8f0'), spaceAfter=12),
            Paragraph('Warm regards,', ps('b7', fontSize=10.5, leading=17, spaceAfter=4)),
            Paragraph(f'<b>HR Department — {company}</b>',
                      ps('bb3', fontName='Helvetica-Bold', fontSize=10.5,
                         textColor=accent, spaceAfter=4)),
            Spacer(1, 20),
            Paragraph(f'Generated by OfferFlow on {today}.',
                      ps('sg', fontSize=9, textColor=colors.HexColor('#94a3b8'),
                         alignment=TA_CENTER)),
        ]

    # ── MODERN ────────────────────────────────────────────────────────────────
    elif layout == 'modern':
        header_tbl = Table([[Paragraph(company,
                              ps('mco', fontName='Helvetica-Bold', fontSize=18,
                                 textColor=colors.white)),
                             Paragraph(f'Date: {today}',
                              ps('mdt', fontSize=9, textColor=colors.HexColor('#c4b5fd'),
                                 alignment=TA_RIGHT))]],
                           colWidths=[120*mm, 50*mm])
        header_tbl.setStyle(TableStyle([
            ('BACKGROUND', (0,0),(-1,-1), accent),
            ('ROWPADDING', (0,0),(-1,-1), 10),
            ('VALIGN',     (0,0),(-1,-1), 'MIDDLE'),
        ]))

        kv_data = [
            [_kv_cell('Role', role, accent2),
             _kv_cell('Joining Date', joining, accent2)],
            [_kv_cell('Annual CTC', f'₹ {salary}', accent2),
             _kv_cell('Employment Type', emp_type, accent2)],
            [_kv_cell('Company', company, accent2),
             _kv_cell('Reporting Location', 'As communicated by HR', accent2)],
        ]
        kv_tbl = Table(kv_data, colWidths=[83*mm, 83*mm], spaceBefore=10)
        kv_tbl.setStyle(TableStyle([
            ('ALIGN',      (0,0),(-1,-1), 'LEFT'),
            ('VALIGN',     (0,0),(-1,-1), 'TOP'),
            ('ROWPADDING', (0,0),(-1,-1), 6),
        ]))

        story = [
            header_tbl, Spacer(1, 18),
            Paragraph(f'Dear <b>{name}</b>,',
                      ps('mb', fontSize=11, leading=18, spaceAfter=8)),
            Paragraph(
                f'We are pleased to offer you the position of <b>{role}</b> at '
                f'<b>{company}</b>. Please review your offer details below.',
                ps('mb2', fontSize=10.5, leading=17, spaceAfter=14)),
            Paragraph('Offer Details',
                      ps('mh', fontName='Helvetica-Bold', fontSize=12,
                         textColor=accent, spaceAfter=8)),
            kv_tbl, Spacer(1, 18),
            Paragraph('Conditions of Offer',
                      ps('mh2', fontName='Helvetica-Bold', fontSize=12,
                         textColor=accent, spaceAfter=6)),
            Paragraph('• Background verification must be completed successfully.', ps('mc',  fontSize=10.5, leading=17, spaceAfter=4)),
            Paragraph('• All required documents must be submitted before joining.', ps('mc2', fontSize=10.5, leading=17, spaceAfter=4)),
            Paragraph("• You must accept the company's Code of Conduct.",           ps('mc3', fontSize=10.5, leading=17, spaceAfter=14)),
            HRFlowable(width='100%', thickness=1, color=accent2, spaceAfter=12),
            Paragraph('Sincerely,', ps('ms', fontSize=10.5, leading=17, spaceAfter=4)),
            Paragraph(f'<b>{company} — HR Team</b>',
                      ps('ms2', fontName='Helvetica-Bold', fontSize=10.5,
                         textColor=accent, spaceAfter=4)),
        ]

    # ── MINIMAL ───────────────────────────────────────────────────────────────
    elif layout == 'minimal':
        tbl = _make_table(tbl_data, accent, accent2, row_even, row_odd, grid=False)
        story = [
            Paragraph(company,
                      ps('mico', fontName='Helvetica', fontSize=16,
                         textColor=accent, spaceAfter=2)),
            Paragraph('Human Resources',
                      ps('misub', fontSize=9, textColor=colors.HexColor('#9ca3af'),
                         spaceAfter=20)),
            HRFlowable(width='30%', thickness=1, color=accent2, spaceAfter=20),
            Paragraph('Offer of Employment',
                      ps('mih', fontName='Helvetica', fontSize=14,
                         textColor=accent, spaceAfter=6)),
            Paragraph(f'{today}',
                      ps('midt', fontSize=9, textColor=colors.HexColor('#9ca3af'),
                         spaceAfter=20)),
            Paragraph(f'Dear {name},',
                      ps('mib', fontSize=10.5, leading=18, spaceAfter=10)),
            Paragraph(
                f'We are pleased to offer you the role of {role} at {company}.',
                ps('mib2', fontSize=10.5, leading=18, spaceAfter=16)),
            tbl, Spacer(1, 20),
            Paragraph('This offer is subject to background verification and document submission.',
                      ps('mic', fontSize=10, leading=16,
                         textColor=colors.HexColor('#6b7280'), spaceAfter=24)),
            Paragraph('Regards,',         ps('mis',  fontSize=10.5, leading=17, spaceAfter=4)),
            Paragraph(f'{company} HR',    ps('mis2', fontSize=10.5, textColor=accent, spaceAfter=4)),
        ]

    # ── EXECUTIVE ─────────────────────────────────────────────────────────────
    elif layout == 'executive':
        tbl = _make_table(tbl_data, accent, accent2, row_even, row_odd)
        story = [
            Paragraph(company.upper(),
                      ps('eco', fontName='Helvetica-Bold', fontSize=22,
                         textColor=accent, alignment=TA_CENTER, spaceAfter=2)),
            Paragraph('PRIVATE &amp; CONFIDENTIAL',
                      ps('esub', fontSize=8, textColor=colors.HexColor('#9ca3af'),
                         alignment=TA_CENTER, spaceAfter=4)),
            HRFlowable(width='100%', thickness=3, color=accent,  spaceAfter=6),
            HRFlowable(width='100%', thickness=1, color=accent2, spaceAfter=18),
            Paragraph(f'Date: {today}',
                      ps('edt', fontSize=10, spaceAfter=4)),
            Paragraph(f'Ref: OF-{cid[:8].upper()}',
                      ps('eref', fontSize=10,
                         textColor=colors.HexColor('#6b7280'), spaceAfter=18)),
            Paragraph(f'Dear <b>{name}</b>,',
                      ps('eb', fontSize=11, leading=18, spaceAfter=10)),
            Paragraph(
                f'On behalf of the Board and Management of <b>{company}</b>, it is '
                f'our honour to extend this formal Offer of Employment for the '
                f'position of <b>{role}</b>. We are confident that your experience '
                f'and expertise will be a valuable asset to our organisation.',
                ps('eb2', fontSize=10.5, leading=18, spaceAfter=14)),
            Paragraph('Terms of the Offer',
                      ps('eh', fontName='Helvetica-Bold', fontSize=11,
                         textColor=accent, spaceAfter=8)),
            tbl, Spacer(1, 16),
            Paragraph('Conditions Precedent',
                      ps('eh2', fontName='Helvetica-Bold', fontSize=11,
                         textColor=accent, spaceAfter=6)),
            Paragraph('1. Satisfactory completion of background and reference verification.', ps('ec',  fontSize=10.5, leading=17, spaceAfter=4)),
            Paragraph('2. Production of original educational and professional documents.',    ps('ec2', fontSize=10.5, leading=17, spaceAfter=4)),
            Paragraph("3. Execution of the Company's Employment Agreement and NDA.",          ps('ec3', fontSize=10.5, leading=17, spaceAfter=14)),
            Paragraph(
                'Kindly confirm your acceptance at the earliest convenience by '
                'clicking <b>Accept Offer</b> in the accompanying email. '
                'This offer shall remain valid for 48 hours.',
                ps('ec4', fontSize=10.5, leading=18, spaceAfter=24)),
            HRFlowable(width='100%', thickness=0.5,
                       color=colors.HexColor('#e2e8f0'), spaceAfter=12),
            Paragraph('Yours sincerely,', ps('es', fontSize=10.5, leading=17, spaceAfter=4)),
            Spacer(1, 24),
            Paragraph('_________________________________',
                      ps('esig',  fontSize=10, textColor=colors.HexColor('#9ca3af'), spaceAfter=2)),
            Paragraph('<b>Authorised Signatory</b>',
                      ps('esig2', fontName='Helvetica-Bold', fontSize=10.5,
                         textColor=accent, spaceAfter=2)),
            Paragraph(f'{company}',
                      ps('esig3', fontSize=10, textColor=colors.HexColor('#6b7280'))),
        ]

    else:
        story = [Paragraph('Offer letter', styles['Normal'])]

    _build_pdf(doc, buf, story, lh_path, out_path)
    return out_path


# ── PDF helpers ───────────────────────────────────────────────────────────────
def _kv_cell(label, value, accent_color):
    styles = getSampleStyleSheet()
    return Paragraph(
        f'<font color="{accent_color}" size="8"><b>{label.upper()}</b></font><br/>'
        f'<font size="10" color="#1e293b">{value}</font>',
        ParagraphStyle('kv', parent=styles['Normal'], leading=16,
                       borderPadding=(8, 8, 8, 8),
                       backColor=colors.HexColor('#f5f3ff'),
                       borderWidth=1, borderColor=colors.HexColor('#ede9fe'),
                       borderRadius=4, spaceAfter=4))


def _make_table(tbl_data, accent, accent2, row_even, row_odd, grid=True):
    tbl  = Table(tbl_data, colWidths=[68*mm, 102*mm])
    cmds = [
        ('BACKGROUND',     (0, 0), (-1,  0), accent),
        ('TEXTCOLOR',      (0, 0), (-1,  0), colors.white),
        ('FONTNAME',       (0, 0), (-1,  0), 'Helvetica-Bold'),
        ('FONTSIZE',       (0, 0), (-1,  0), 10),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [row_even, row_odd]),
        ('FONTNAME',       (0, 1), (0,  -1), 'Helvetica-Bold'),
        ('FONTNAME',       (1, 1), (1,  -1), 'Helvetica'),
        ('FONTSIZE',       (0, 1), (-1, -1), 10),
        ('TEXTCOLOR',      (0, 1), (0,  -1), accent2),
        ('ROWPADDING',     (0, 0), (-1, -1), 9),
        ('VALIGN',         (0, 0), (-1, -1), 'MIDDLE'),
    ]
    if grid:
        cmds.append(('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#e2e8f0')))
    tbl.setStyle(TableStyle(cmds))
    return tbl


def _build_pdf(doc, buf, story, lh_path, out_path):
    if lh_path and os.path.exists(lh_path):
        buf2 = io.BytesIO()
        doc2 = SimpleDocTemplate(buf2, pagesize=A4,
                                 rightMargin=22*mm, leftMargin=22*mm,
                                 topMargin=54*mm, bottomMargin=22*mm)
        doc2.build(story)
        buf2.seek(0)
        try:
            _merge_letterhead(lh_path, buf2.read(), out_path)
        except Exception as e:
            logger.warning(f'Letterhead merge failed: {e}; using plain PDF')
            doc.build(story)
            buf.seek(0)
            open(out_path, 'wb').write(buf.read())
    else:
        doc.build(story)
        buf.seek(0)
        open(out_path, 'wb').write(buf.read())


def _merge_letterhead(lh_path, content_bytes, out_path):
    try:
        from pdf2image import convert_from_path
        import tempfile
        from pypdf import PdfReader, PdfWriter

        imgs = convert_from_path(lh_path, first_page=1, last_page=1, dpi=150)
        if not imgs:
            raise ValueError('No pages in letterhead')
        tmp_img = tempfile.mktemp(suffix='.png')
        imgs[0].save(tmp_img, 'PNG')

        tmp_lh_pdf = tempfile.mktemp(suffix='.pdf')
        w_pt, h_pt = A4
        c = rlc.Canvas(tmp_lh_pdf, pagesize=A4)
        c.drawImage(tmp_img, 0, 0, width=w_pt, height=h_pt)
        c.save()
        os.unlink(tmp_img)

        lh_reader      = PdfReader(tmp_lh_pdf)
        content_reader = PdfReader(io.BytesIO(content_bytes))
        writer         = PdfWriter()

        lh_page = lh_reader.pages[0]
        for i, page in enumerate(content_reader.pages):
            if i == 0:
                lh_page.merge_page(page)
                writer.add_page(lh_page)
            else:
                writer.add_page(page)

        with open(out_path, 'wb') as f:
            writer.write(f)
        os.unlink(tmp_lh_pdf)

    except ImportError:
        open(out_path, 'wb').write(content_bytes)


# ── Preview HTML ──────────────────────────────────────────────────────────────
def letterhead_preview_html(hr_user, candidate, pattern_key='classic', custom_text=''):
    company  = hr_user['company_name']
    lh_path  = _resolve_lh_path(hr_user.get('letterhead', ''))
    today    = datetime.now().strftime('%d %B %Y')
    name     = candidate.get('name', '')
    role     = candidate.get('role', '')
    joining  = candidate.get('joining_date', '')
    salary   = candidate.get('salary', '')
    emp_type = candidate.get('employment_type', 'full_time').replace('_', ' ').title()
    email    = candidate.get('email', '')
    pat      = PATTERNS.get(pattern_key, PATTERNS['classic'])

    accent  = pat.get('accent',  '#0d1b3e')
    accent2 = pat.get('accent2', '#1a56db')

    lh_bg_style = ''
    lh_notice   = ''
    if lh_path and os.path.exists(lh_path):
        try:
            from pdf2image import convert_from_path
            import tempfile
            imgs = convert_from_path(lh_path, first_page=1, last_page=1, dpi=96)
            if imgs:
                tmp = tempfile.mktemp(suffix='.png')
                imgs[0].save(tmp, 'PNG')
                with open(tmp, 'rb') as f:
                    b64img = base64.b64encode(f.read()).decode()
                os.unlink(tmp)
                lh_bg_style = (
                    f'background-image:url("data:image/png;base64,{b64img}");'
                    f'background-size:100% 100%;background-repeat:no-repeat;'
                )
                lh_notice = (
                    '<div style="background:#fff8e1;border:1px solid #f59e0b;'
                    'border-radius:6px;padding:7px 14px;margin-bottom:12px;'
                    'font-size:11px;color:#92400e;font-family:Arial,sans-serif">'
                    '📄 <b>Letterhead Preview</b> — the offer letter PDF will '
                    'be sent with this letterhead as background.</div>'
                )
        except Exception:
            with open(lh_path, 'rb') as f:
                b64 = base64.b64encode(f.read()).decode()
            lh_notice = (
                f'<div style="margin-bottom:12px">'
                f'<div style="background:#fff8e1;border:1px solid #f59e0b;'
                f'border-radius:6px;padding:7px 14px;margin-bottom:8px;'
                f'font-size:11px;color:#92400e">📄 Letterhead attached to email.</div>'
                f'<iframe src="data:application/pdf;base64,{b64}" '
                f'width="100%" height="140" style="border:1px solid #e2e8f0;'
                f'border-radius:6px"></iframe></div>'
            )

    pat_badge = (
        f'<span style="display:inline-block;background:{accent};color:#fff;'
        f'font-size:10px;padding:2px 10px;border-radius:12px;'
        f'font-family:Arial,sans-serif;margin-bottom:14px">'
        f'{pat["label"]} Template</span>'
    )

    layout = pat.get('layout', 'classic')
    rows   = ''.join(
        f'<tr style="background:{"#f0f4ff" if i%2==0 else "#fff"}">'
        f'<td style="padding:9px 14px;font-weight:700;color:{accent2};'
        f'font-size:11.5px;width:42%;border:1px solid #e2e8f0">{k}</td>'
        f'<td style="padding:9px 14px;color:#1e293b;font-size:11.5px;'
        f'border:1px solid #e2e8f0">{v}</td></tr>'
        for i, (k, v) in enumerate([
            ('Candidate Name', name), ('Designation / Role', role),
            ('Joining Date', joining), ('Annual CTC', f'₹ {salary}'),
            ('Employment Type', emp_type), ('Reporting Location', 'As communicated by HR'),
        ])
    )

    if layout == 'custom':
        inner_html = f'''
<div style="font-family:Georgia,serif;font-size:12px;
            color:#1e293b;line-height:1.9;white-space:pre-wrap;padding:10px 0">
{custom_text or "(Your custom offer letter content will appear here.)"}
</div>'''

    elif layout == 'modern':
        kv_cards = ''.join(
            f'<div style="background:#f5f3ff;border:1px solid #ede9fe;'
            f'border-radius:6px;padding:10px 14px;margin:4px">'
            f'<div style="font-size:9px;font-weight:700;color:{accent2};'
            f'text-transform:uppercase;letter-spacing:.5px">{k}</div>'
            f'<div style="font-size:12px;color:#1e293b;margin-top:2px">{v}</div>'
            f'</div>'
            for k, v in [
                ('Role', role), ('Joining Date', joining),
                ('Annual CTC', f'₹ {salary}'), ('Employment Type', emp_type),
                ('Company', company), ('Location', 'As communicated by HR'),
            ]
        )
        inner_html = f'''
<div style="background:{accent};padding:18px 20px;border-radius:8px 8px 0 0;
            display:flex;justify-content:space-between;align-items:center">
  <span style="font-family:Arial;font-size:18px;font-weight:900;color:#fff">{company}</span>
  <span style="font-size:10px;color:rgba(255,255,255,.6)">{today}</span>
</div>
<div style="padding:18px 20px">
  <p style="font-family:Arial;font-size:12px;color:#1e293b;margin:0 0 10px">
    Dear <b>{name}</b>, we are pleased to offer you <b>{role}</b> at <b>{company}</b>.
  </p>
  <div style="display:flex;flex-wrap:wrap;gap:6px;margin-bottom:14px">
    {kv_cards}
  </div>
  <p style="font-family:Arial;font-size:11px;color:#6b7280;margin:0 0 14px">
    • Background verification required &nbsp;• Documents to be submitted before joining
  </p>
  <div style="border-top:1px solid #ede9fe;padding-top:12px;
              font-family:Arial;font-size:11.5px;color:#1e293b">
    Sincerely, <b>{company} HR Team</b>
  </div>
</div>'''

    elif layout == 'minimal':
        inner_html = f'''
<div style="padding:10px 0">
  <div style="font-family:Arial;font-size:15px;color:{accent};font-weight:400;
              margin-bottom:2px">{company}</div>
  <div style="font-family:Arial;font-size:9px;color:#9ca3af;margin-bottom:16px">
    Human Resources</div>
  <div style="height:2px;background:{accent2};width:60px;margin-bottom:16px"></div>
  <div style="font-family:Arial;font-size:13px;color:{accent};margin-bottom:4px">
    Offer of Employment</div>
  <div style="font-family:Arial;font-size:9px;color:#9ca3af;margin-bottom:18px">{today}</div>
  <p style="font-family:Georgia;font-size:12px;color:#1e293b;line-height:1.9">
    Dear {name},<br>We are pleased to offer you the role of {role} at {company}.
  </p>
  <table style="width:100%;border-collapse:collapse;margin:14px 0;font-size:11.5px">{rows}</table>
  <p style="font-family:Georgia;font-size:11px;color:#6b7280;margin:0 0 18px">
    This offer is subject to background verification and document submission.
  </p>
  <div style="font-family:Arial;font-size:11.5px;color:#1e293b">{company} HR</div>
</div>'''

    elif layout == 'executive':
        inner_html = f'''
<div style="text-align:center;font-family:Arial;margin-bottom:10px">
  <div style="font-size:20px;font-weight:900;color:{accent};
              letter-spacing:2px">{company.upper()}</div>
  <div style="font-size:8px;color:#9ca3af;letter-spacing:1px">
    PRIVATE &amp; CONFIDENTIAL</div>
</div>
<div style="height:3px;background:{accent};margin-bottom:2px"></div>
<div style="height:1px;background:{accent2};margin-bottom:16px"></div>
<div style="font-family:Arial;font-size:10px;color:#6b7280;margin-bottom:4px">
  Date: {today} &nbsp;|&nbsp; Ref: OF-PREVIEW</div>
<p style="font-family:Georgia;font-size:12px;color:#1e293b;line-height:1.9;margin:12px 0">
  Dear <b>{name}</b>,<br><br>
  On behalf of the Board and Management of <b>{company}</b>, it is our honour
  to extend this formal Offer of Employment for the position of <b>{role}</b>.
</p>
<table style="width:100%;border-collapse:collapse;margin:10px 0;font-size:11.5px">{rows}</table>
<p style="font-family:Georgia;font-size:11.5px;color:#1e293b;line-height:1.9;margin:14px 0 10px">
  <b>Conditions Precedent:</b><br>
  1. Background and reference verification.<br>
  2. Submission of original documents.<br>
  3. Execution of Employment Agreement.
</p>
<div style="border-top:1px solid #e2e8f0;padding-top:14px;font-family:Arial;
            font-size:11.5px;color:#1e293b;margin-top:14px">
  Yours sincerely,<br><br>
  <span style="font-size:13px;color:{accent}">___________________________</span><br>
  <b>Authorised Signatory</b><br>{company}
</div>'''

    else:  # classic (default)
        inner_html = f'''
<div style="text-align:center;margin-bottom:18px">
  <div style="font-family:Arial;font-size:20px;font-weight:900;color:{accent}">{company}</div>
  <div style="font-size:10px;color:#64748b;margin-top:3px">Human Resources Department</div>
  <div style="height:3px;background:linear-gradient(90deg,{accent},{accent2});
              width:70px;margin:8px auto 0;border-radius:2px"></div>
</div>
<div style="text-align:center;font-size:16px;font-weight:700;color:{accent};
            margin-bottom:14px;border-bottom:1.5px solid #e2e8f0;padding-bottom:12px">
  OFFER OF EMPLOYMENT
</div>
<div style="font-size:10px;color:#64748b;margin-bottom:10px">Date: {today}</div>
<p style="font-family:Georgia;font-size:13px;color:#1e293b;margin-bottom:8px">
  Dear <b>{name}</b>,</p>
<p style="font-family:Georgia;font-size:12px;color:#475569;line-height:1.9;margin-bottom:12px">
  We are delighted to extend this offer for the position of
  <b style="color:{accent2}">{role}</b> at <b>{company}</b>.
</p>
<table style="width:100%;border-collapse:collapse;margin-bottom:14px;font-size:11.5px">{rows}</table>
<ul style="font-size:11.5px;color:#475569;line-height:2;margin-left:18px;margin-bottom:14px">
  <li>Successful completion of background verification.</li>
  <li>Submission of all required documents before joining.</li>
  <li>Acceptance of the company's Code of Conduct.</li>
</ul>
<div style="border-top:1px solid #e2e8f0;padding-top:14px;font-family:Arial;font-size:11.5px">
  Warm regards,<br><b style="color:{accent}">HR Department — {company}</b>
</div>'''

    return f'''
<div style="position:relative;font-family:Georgia,serif;
            border:1px solid #e2e8f0;border-radius:12px;overflow:hidden;
            background:#fff;{lh_bg_style}">
  <div style="position:relative;background:rgba(255,255,255,0.88);padding:28px 32px">
    {lh_notice}
    {pat_badge}
    {inner_html}
    <div style="margin-top:16px;padding:10px;background:rgba(248,250,255,0.9);
                border-radius:6px;font-size:10px;color:#94a3b8;text-align:center">
      📧 PDF offer letter will be sent as an attachment to <b>{email}</b>
    </div>
  </div>
</div>'''


# ── Email HTML builders ───────────────────────────────────────────────────────
def offer_email_html(c, hr, accept_link, decline_link):
    company  = hr['company_name']
    name     = c.get('name', '')
    role     = c.get('role', '')
    joining  = c.get('joining_date', '')
    salary   = c.get('salary', '')
    emp_type = c.get('employment_type', 'full_time').replace('_', ' ').title()
    rows     = ''.join(
        f'<tr style="background:{"#f8faff" if i%2==0 else "#fff"}">'
        f'<td style="padding:11px 16px;font-weight:700;color:#1a56db;font-size:13px;width:38%">{k}</td>'
        f'<td style="padding:11px 16px;color:#1e293b;font-size:13px">{v}</td></tr>'
        for i, (k, v) in enumerate([
            ('Role', role), ('Joining Date', joining),
            ('Annual CTC', '₹ ' + str(salary)), ('Employment Type', emp_type)
        ]))
    return f"""<!DOCTYPE html><html><head><meta charset="UTF-8"></head>
<body style="margin:0;padding:0;background:#f0f4fa;font-family:Arial,sans-serif">
<table width="100%" cellpadding="0" cellspacing="0" style="padding:32px 0">
<tr><td align="center">
<table width="600" cellpadding="0" cellspacing="0"
  style="background:#fff;border-radius:16px;overflow:hidden;
         box-shadow:0 4px 24px rgba(0,0,0,.08)">
  <tr><td style="background:linear-gradient(135deg,#0d1b3e 0%,#1a56db 100%);
                 padding:36px 40px;text-align:center">
    <div style="font-size:26px;font-weight:900;color:#fff;letter-spacing:1px">{company}</div>
    <div style="font-size:12px;color:rgba(255,255,255,.7);margin-top:6px">Official Offer Letter</div>
  </td></tr>
  <tr><td style="padding:36px 40px">
    <h2 style="font-size:22px;color:#0d1b3e;margin:0 0 8px">🎉 Congratulations, {name}!</h2>
    <p style="color:#64748b;font-size:13.5px;line-height:1.9;margin:0 0 22px">
      We are pleased to extend an offer for the role of
      <b style="color:#1a56db">{role}</b> at <b>{company}</b>.
      Please review and <b>respond within 48 hours</b>.
    </p>
    <table width="100%" cellpadding="0" cellspacing="0"
      style="border:1px solid #e2e8f0;border-radius:10px;overflow:hidden;margin-bottom:28px">
      {rows}
    </table>
    <table cellpadding="0" cellspacing="0" style="margin:0 auto"><tr>
      <td style="padding-right:14px">
        <a href="{accept_link}"
          style="display:inline-block;background:#10b981;color:#fff;
                 padding:14px 34px;border-radius:9px;text-decoration:none;
                 font-weight:700;font-size:15px">✓ Accept Offer</a>
      </td>
      <td>
        <a href="{decline_link}"
          style="display:inline-block;background:#ef4444;color:#fff;
                 padding:14px 34px;border-radius:9px;text-decoration:none;
                 font-weight:700;font-size:15px">✕ Decline Offer</a>
      </td>
    </tr></table>
    <p style="color:#94a3b8;font-size:11px;text-align:center;margin-top:22px;line-height:1.7">
      ⏰ Offer expires after 48 hours if no response.<br>
      📎 Your offer letter PDF is attached.
    </p>
  </td></tr>
  <tr><td style="background:#f8faff;padding:14px 40px;text-align:center;
                 border-top:1px solid #e2e8f0">
    <p style="color:#94a3b8;font-size:11px;margin:0">
      Sent via OfferFlow · {company} HR Portal</p>
  </td></tr>
</table></td></tr></table></body></html>"""


def bg_verification_email_html(candidate, bg_link, company):
    name = candidate.get('name', '')
    role = candidate.get('role', '')
    return f"""<!DOCTYPE html><html><head><meta charset="UTF-8"></head>
<body style="margin:0;padding:0;background:#f0f4fa;font-family:Arial,sans-serif">
<table width="100%" cellpadding="0" cellspacing="0" style="padding:32px 0">
<tr><td align="center">
<table width="560" cellpadding="0" cellspacing="0"
  style="background:#fff;border-radius:16px;overflow:hidden;
         box-shadow:0 4px 20px rgba(0,0,0,.08)">
  <tr><td style="background:linear-gradient(135deg,#0d1b3e,#1a56db);
                 padding:30px 40px;text-align:center">
    <div style="font-size:22px;font-weight:900;color:#fff">{company}</div>
    <div style="font-size:12px;color:rgba(255,255,255,.65);margin-top:4px">
      Background Verification — Next Step</div>
  </td></tr>
  <tr><td style="padding:36px 40px">
    <div style="font-size:40px;text-align:center;margin-bottom:14px">📋</div>
    <h2 style="font-size:20px;color:#0d1b3e;margin:0 0 10px;text-align:center">
      Complete Your Background Verification</h2>
    <p style="color:#64748b;font-size:13px;line-height:1.9;margin:0 0 20px">
      Hi <b>{name}</b>,<br><br>
      Thank you for accepting the offer for <b>{role}</b> at <b>{company}</b>!
      To proceed with your onboarding, please complete the background
      verification form using the button below.
    </p>
    <div style="background:#f0f9ff;border-left:4px solid #1a56db;
                padding:14px 18px;border-radius:4px;margin-bottom:24px">
      <p style="margin:0;font-size:12.5px;color:#1e293b">
        <b>What you'll need:</b><br>
        • Personal details (Aadhaar, PAN)<br>
        • Educational qualifications<br>
        • Previous employment details (if experienced)
      </p>
    </div>
    <div style="text-align:center">
      <a href="{bg_link}"
        style="display:inline-block;background:#1a56db;color:#fff;
               padding:14px 38px;border-radius:9px;text-decoration:none;
               font-weight:700;font-size:15px">
        Start Verification →
      </a>
    </div>
    <p style="color:#94a3b8;font-size:11px;text-align:center;margin-top:20px;
              line-height:1.7">
      This link is unique to you. Please do not share it with others.
    </p>
  </td></tr>
  <tr><td style="background:#f8faff;padding:14px 40px;text-align:center;
                 border-top:1px solid #e2e8f0">
    <p style="color:#94a3b8;font-size:11px;margin:0">
      Sent via OfferFlow · {company} HR Portal</p>
  </td></tr>
</table></td></tr></table></body></html>"""


# ══════════════════════════════════════════════════════════════════════════════
#  ROUTES
# ══════════════════════════════════════════════════════════════════════════════
@app.route('/')
def index():
    return render_template('index.html')

@app.route('/login')
def login_page():
    return render_template('login.html')

@app.route('/api/login', methods=['POST'])
def api_login():
    data     = request.json or {}
    username = sanitize_string(data.get('username', ''))
    password = data.get('password', '')

    if not username or not password:
        return jsonify({'success': False, 'message': 'Username and password are required'}), 400

    user = get_db().users.find_one(
        {'username': username, 'password': hash_pw(password)}, {'_id': 0})
    if not user:
        return jsonify({'success': False, 'message': 'Invalid credentials'}), 401
    session['user_id'] = user['id']
    session.permanent  = True
    return jsonify({'success': True})

@app.route('/api/register', methods=['POST'])
def api_register():
    username     = sanitize_string(request.form.get('username', ''))
    email        = sanitize_string(request.form.get('email', ''))
    password     = request.form.get('password', '')
    confirm      = request.form.get('confirm_password', '')
    company_name = sanitize_string(request.form.get('company_name', ''))

    if not all([username, email, password, confirm, company_name]):
        return jsonify({'success': False, 'message': 'All fields are required'}), 400
    if not validate_email(email):
        return jsonify({'success': False, 'message': 'Invalid email format'}), 400

    db = get_db()
    if db.users.find_one({'username': username}):
        return jsonify({'success': False, 'message': 'Username already exists'}), 400
    if db.users.find_one({'email': email}):
        return jsonify({'success': False, 'message': 'Email already registered'}), 400

    if password != confirm:
        return jsonify({'success': False, 'message': 'Passwords do not match'}), 400
    if not validate_password(password):
        return jsonify({'success': False, 'message':
            'Password needs 8+ chars, uppercase, lowercase & special character'}), 400

    lh = request.files.get('letterhead')
    if not lh or not lh.filename:
        return jsonify({'success': False, 'message': 'Company letterhead (PDF) is required'}), 400
    if not lh.filename.lower().endswith('.pdf'):
        return jsonify({'success': False, 'message': 'Letterhead must be a PDF file'}), 400
    if not validate_file_size(lh, max_size_mb=5):
        return jsonify({'success': False, 'message': 'Letterhead file size must be less than 5MB'}), 400

    fname   = secure_filename(lh.filename)
    lh_path = os.path.join(UPLOAD_DIR, 'letterheads', fname)
    lh.save(lh_path)
    project_lh = os.path.join(BASE_DIR, 'uploads', 'letterheads', fname)
    if not os.path.exists(project_lh):
        shutil.copy2(lh_path, project_lh)

    uid = str(uuid.uuid4())
    db.users.insert_one({
        'id': uid, 'username': username, 'email': email,
        'password': hash_pw(password), 'company_name': company_name,
        'letterhead': lh_path, 'created_at': str(datetime.now()),
    })
    return jsonify({'success': True})

@app.route('/api/forgot-password', methods=['POST'])
def api_forgot_password():
    email = sanitize_string((request.json or {}).get('email', ''))
    if not email:
        return jsonify({'success': False, 'message': 'Email is required'}), 400
    if not validate_email(email):
        return jsonify({'success': False, 'message': 'Invalid email format'}), 400

    user = get_db().users.find_one({'email': email}, {'_id': 0})
    if not user:
        return jsonify({'success': False, 'message': 'Email not found'}), 404

    token = str(uuid.uuid4())
    get_db().tokens.insert_one({
        'token':   token,
        'user_id': user['id'],
        'expires': str(datetime.now() + timedelta(hours=1)),
    })
    link = f"{BASE_URL}/reset-password/{token}"
    html = f"""<!DOCTYPE html><html><body style="margin:0;padding:32px 0;
background:#f0f4fa;font-family:Arial,sans-serif">
<table width="100%" cellpadding="0" cellspacing="0"><tr><td align="center">
<table width="480" style="background:#fff;border-radius:16px;overflow:hidden">
  <tr><td style="background:linear-gradient(135deg,#0d1b3e,#1a56db);
                 padding:28px;text-align:center">
    <div style="font-size:22px;font-weight:900;color:#fff">OfferFlow</div>
    <div style="color:rgba(255,255,255,.7);font-size:12px;margin-top:4px">
      Password Reset</div>
  </td></tr>
  <tr><td style="padding:36px;text-align:center">
    <div style="font-size:44px;margin-bottom:14px">🔑</div>
    <h2 style="font-size:19px;color:#0d1b3e;margin:0 0 10px">Reset Your Password</h2>
    <p style="color:#64748b;font-size:13px;margin:0 0 26px;line-height:1.7">
      Click the button below. This link expires in <b>1 hour</b>.
    </p>
    <a href="{link}"
      style="display:inline-block;background:#1a56db;color:#fff;
             padding:13px 34px;border-radius:8px;text-decoration:none;
             font-weight:700;font-size:14px">Reset Password</a>
    <p style="color:#94a3b8;font-size:11px;margin-top:20px">
      If you did not request this, ignore this email.</p>
  </td></tr>
</table></td></tr></table></body></html>"""
    send_async(email, 'Reset Your OfferFlow Password', html,
               user_id=user['id'], email_type='password_reset')
    return jsonify({'success': True, 'message': 'Password reset link sent to your email'})

@app.route('/reset-password/<token>')
def reset_password_page(token):
    if not get_db().tokens.find_one({'token': token}):
        return redirect('/login')
    return render_template('reset_password.html', token=token)

@app.route('/api/reset-password', methods=['POST'])
def api_reset_password():
    data    = request.json or {}
    token   = data.get('token', '')
    pw      = data.get('password', '')
    confirm = data.get('confirm_password', '')

    tok = get_db().tokens.find_one({'token': token})
    if not tok:
        return jsonify({'success': False, 'message': 'Invalid or expired link'}), 400
    if pw != confirm:
        return jsonify({'success': False, 'message': 'Passwords do not match'}), 400
    if not validate_password(pw):
        return jsonify({'success': False, 'message': 'Password requirements not met'}), 400

    get_db().users.update_one({'id': tok['user_id']}, {'$set': {'password': hash_pw(pw)}})
    get_db().tokens.delete_one({'token': token})
    return jsonify({'success': True})

@app.route('/logout')
def logout():
    session.clear()
    return redirect('/')

@app.route('/dashboard')
def dashboard():
    if not current_user():
        return redirect('/login')
    return render_template('dashboard.html')

@app.route('/api/dashboard-stats')
def api_dashboard_stats():
    u = current_user()
    if not u:
        return jsonify({}), 401
    uid = u['id']
    db  = get_db()
    return jsonify({
        'total_offers':         db.candidates.count_documents({'hr_id': uid}),
        'offer_accepted':       db.candidates.count_documents({'hr_id': uid, 'offer_status': 'accepted'}),
        'offer_declined':       db.candidates.count_documents({'hr_id': uid, 'offer_status': 'declined'}),
        'action_pending':       db.candidates.count_documents({'hr_id': uid, 'offer_status': 'pending'}),
        'cancelled':            db.candidates.count_documents({'hr_id': uid, 'offer_status': 'cancelled'}),
        'verified':             db.verifications.count_documents({'hr_id': uid, 'verification_status': 'verified'}),
        'rejected':             db.verifications.count_documents({'hr_id': uid, 'verification_status': 'rejected'}),
        'verification_pending': db.verifications.count_documents({'hr_id': uid, 'verification_status': 'pending'}),
        'company_name': u['company_name'],
    })

@app.route('/api/offers')
def api_offers():
    u = current_user()
    if not u:
        return jsonify([]), 401
    status = request.args.get('status')
    query  = {'hr_id': u['id']}
    if status:
        query['offer_status'] = status
    result = [
        {'id': c['id'], 'name': c['name'], 'role': c.get('role', ''),
         'joining_date': c.get('joining_date', ''), 'email': c.get('email', ''),
         'salary': c.get('salary', ''), 'employment_type': c.get('employment_type', ''),
         'offer_status': c.get('offer_status', 'pending')}
        for c in get_db().candidates.find(query, {'_id': 0})
    ]
    return jsonify(result)

@app.route('/api/verifications')
def api_verifications():
    u = current_user()
    if not u:
        return jsonify([]), 401
    status = request.args.get('status')
    query  = {'hr_id': u['id']}
    if status:
        query['verification_status'] = status
    result = [
        {'id': v['id'], 'name': v['name'], 'email': v.get('email', ''),
         'salary': v.get('salary', ''), 'phone': v.get('phone', ''),
         'verification_status': v.get('verification_status', 'pending')}
        for v in get_db().verifications.find(query, {'_id': 0})
    ]
    return jsonify(result)

@app.route('/upload-excel')
def upload_excel_page():
    if not current_user():
        return redirect('/login')
    return render_template('upload_excel.html')

@app.route('/api/upload-excel', methods=['POST'])
def api_upload_excel():
    u = current_user()
    if not u:
        return jsonify({'success': False}), 401
    f = request.files.get('file')
    if not f or not f.filename:
        return jsonify({'success': False, 'message': 'Please upload a file'}), 400
    if not f.filename.lower().endswith('.xlsx'):
        return jsonify({'success': False, 'message': 'Please upload a .xlsx file'}), 400
    if not validate_file_size(f, max_size_mb=10):
        return jsonify({'success': False, 'message': 'File size must be less than 10MB'}), 400
    path = os.path.join(UPLOAD_DIR, 'excel', secure_filename(f.filename))
    f.save(path)
    try:
        df       = pd.read_excel(path)
        required = ['Name', 'Gmail id', 'Role', 'Joining date', 'Salary']
        missing  = [r for r in required if r not in df.columns]
        if missing:
            return jsonify({'success': False,
                'message': f'Missing columns: {", ".join(missing)}'}), 400
        return jsonify({'success': True, 'records': df.fillna('').to_dict('records'),
                        'columns': list(df.columns)})
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)}), 500

@app.route('/api/save-candidates', methods=['POST'])
def api_save_candidates():
    u = current_user()
    if not u:
        return jsonify({'success': False}), 401
    data    = request.json or {}
    new_ids = []
    db      = get_db()
    for rec in data.get('candidates', []):
        cid = str(uuid.uuid4())
        db.candidates.insert_one({
            'id': cid, 'hr_id': u['id'],
            'name':            rec.get('Name', ''),
            'email':           rec.get('Gmail id', ''),
            'role':            rec.get('Role', ''),
            'joining_date':    str(rec.get('Joining date', '')),
            'salary':          str(rec.get('Salary', '')),
            'employment_type': data.get('employment_type', 'full_time'),
            'pattern':         data.get('pattern', 'classic'),
            'custom_text':     data.get('custom_text', ''),
            'offer_status':    'pending',
            'sent_at':         str(datetime.now()),
            'extra': {k: v for k, v in rec.items()
                      if k not in ['Name', 'Gmail id', 'Role', 'Joining date', 'Salary']},
        })
        new_ids.append(cid)
    return jsonify({'success': True, 'candidate_ids': new_ids})

@app.route('/api/patterns')
def api_patterns():
    return jsonify([
        {'key': k, 'label': v['label'], 'desc': v['desc'],
         'accent': v.get('accent', '#0d1b3e'), 'is_custom': k == 'custom'}
        for k, v in PATTERNS.items()
    ])

@app.route('/api/preview-letter', methods=['POST'])
def api_preview_letter():
    u = current_user()
    if not u:
        return jsonify({'success': False, 'message': 'Not authenticated'}), 401
    data        = request.json or {}
    c           = data.get('candidate', {})
    c['employment_type'] = data.get('employment_type', 'full_time')
    c['id']     = 'preview'
    pattern_key = data.get('pattern', 'classic')
    custom_text = data.get('custom_text', '')
    try:
        html = letterhead_preview_html(u, c, pattern_key, custom_text)
        return jsonify({'success': True, 'html': html})
    except Exception as e:
        logger.error(f'Preview generation error: {e}')
        return jsonify({'success': False, 'message': str(e)}), 500

@app.route('/api/send-offer-emails', methods=['POST'])
def api_send_offer_emails():
    u = current_user()
    if not u:
        return jsonify({'success': False}), 401

    data   = request.json or {}
    cids   = data.get('candidate_ids', [])
    db     = get_db()
    sent   = 0
    failed = []

    for cid in cids:
        c = db.candidates.find_one({'id': cid}, {'_id': 0})
        if not c:
            continue
        if c.get('email_sent_at'):
            logger.info(f'Email already sent to {c["email"]}')
            continue

        accept_link  = f"{BASE_URL}/offer-response/{cid}/accept"
        decline_link = f"{BASE_URL}/offer-response/{cid}/decline"
        pattern_key  = c.get('pattern', 'classic')
        custom_text  = c.get('custom_text', '')

        try:
            pdf_path = generate_offer_pdf(c, u, pattern_key, custom_text)
        except Exception as e:
            logger.error(f'PDF generation failed for {cid}: {e}')
            pdf_path = None

        subject = f"Job Offer — {c.get('role', '')} at {u['company_name']}"
        ok = send_async(
            c['email'], subject,
            offer_email_html(c, u, accept_link, decline_link),
            attach_path=pdf_path,
            attach_name=f"Offer_Letter_{c['name'].replace(' ', '_')}.pdf",
            user_id=u['id'], candidate_id=cid, email_type='offer')

        if ok:
            db.candidates.update_one({'id': cid},
                                     {'$set': {'email_sent_at': str(datetime.now())}})
            sent += 1
        else:
            failed.append({'id': cid, 'email': c['email']})

    return jsonify({'success': True, 'sent': sent, 'failed': failed})

@app.route('/offer-response/<cid>/<action>')
def offer_response(cid, action):
    db = get_db()
    c  = db.candidates.find_one({'id': cid}, {'_id': 0})
    if not c:
        return ("<div style='font-family:sans-serif;text-align:center;margin-top:80px'>"
                "<h2>Invalid or expired link.</h2></div>"), 404

    current_status = c.get('offer_status', 'pending')
    if current_status in ['accepted', 'declined', 'cancelled']:
        return render_template(
            'offer_accepted.html' if current_status == 'accepted' else 'offer_declined.html',
            candidate=c)

    if action == 'accept':
        db.candidates.update_one({'id': cid}, {'$set': {
            'offer_status': 'accepted',
            'responded_at': str(datetime.now()),
        }})
        c = db.candidates.find_one({'id': cid}, {'_id': 0})

        existing = db.verifications.find_one({'candidate_id': cid})
        if not existing:
            vid = str(uuid.uuid4())
            db.verifications.insert_one({
                'id':                  vid,
                'candidate_id':        cid,
                'hr_id':               c['hr_id'],
                'name':                c['name'],
                'email':               c['email'],
                'salary':              c['salary'],
                'phone':               '',
                'verification_status': 'pending',
                'created_at':          str(datetime.now()),
            })
            bg_link = f"{BASE_URL}/background-verification/{vid}"
            hr      = db.users.find_one({'id': c['hr_id']}, {'_id': 0}) or {}
            company = hr.get('company_name', 'the company')
            send_async(
                c['email'],
                f'Next Step: Complete Your Background Verification — {company}',
                bg_verification_email_html(c, bg_link, company),
                user_id=c['hr_id'], candidate_id=cid, email_type='verification')

        return render_template('offer_accepted.html', candidate=c)

    elif action == 'decline':
        db.candidates.update_one({'id': cid}, {'$set': {
            'offer_status': 'declined',
            'responded_at': str(datetime.now()),
        }})
        c = db.candidates.find_one({'id': cid}, {'_id': 0})
        return render_template('offer_declined.html', candidate=c)

    return "<div style='font-family:sans-serif;text-align:center;margin-top:80px'><h2>Invalid action.</h2></div>"

@app.route('/background-verification/<vid>')
def background_verification(vid):
    v = get_db().verifications.find_one({'id': vid}, {'_id': 0})
    if not v:
        return ("<div style='font-family:sans-serif;text-align:center;margin-top:80px'>"
                "<h2>Invalid or expired link.</h2></div>"), 404
    return render_template('bg_verification.html', verification=v)

@app.route('/api/submit-verification', methods=['POST'])
def api_submit_verification():
    vid = request.form.get('verification_id')
    db  = get_db()
    v   = db.verifications.find_one({'id': vid}, {'_id': 0})
    if not v:
        return jsonify({'success': False}), 404
    step = request.form.get('step')

    if step == '1':
        fn = request.form.get('first_name', '')
        ln = request.form.get('last_name', '')
        db.verifications.update_one({'id': vid}, {'$set': {
            'first_name': fn, 'last_name': ln, 'name': f"{fn} {ln}",
            'phone':   request.form.get('phone', ''),
            'aadhaar': request.form.get('aadhaar', ''),
            'pan':     request.form.get('pan', ''),
        }})

    elif step == '2':
        db.verifications.update_one({'id': vid}, {'$set': {
            'college':        request.form.get('college', ''),
            'specialization': request.form.get('specialization', ''),
            'percentage':     request.form.get('percentage', ''),
        }})

    elif step == '3':
        ctype   = request.form.get('candidate_type', 'fresher')
        hr      = db.users.find_one({'id': v.get('hr_id', '')}, {'_id': 0}) or {}
        company = hr.get('company_name', 'the company')

        if ctype == 'experienced':
            prev_co   = request.form.get('prev_company', '')
            prev_role = request.form.get('prev_role', '')
            co_email  = request.form.get('company_email', '')
            duration  = request.form.get('duration', '')
            db.verifications.update_one({'id': vid}, {'$set': {
                'candidate_type':    ctype,
                'prev_company':      prev_co,
                'prev_role':         prev_role,
                'company_email':     co_email,
                'duration':          duration,
                'verification_status': 'pending',
                'submitted_at':      str(datetime.now()),
            }})
            v = db.verifications.find_one({'id': vid}, {'_id': 0})
            vlink = f"{BASE_URL}/company-verify/{vid}/verify"
            rlink = f"{BASE_URL}/company-verify/{vid}/reject"
            send_async(co_email,
                f"Employment Verification Request — {v['name']}",
                f"""<html><body style="font-family:Arial;margin:0;padding:32px 0;background:#f0f4fa">
<table width="100%" cellpadding="0" cellspacing="0"><tr><td align="center">
<table width="540" style="background:#fff;border-radius:16px;overflow:hidden;
                          box-shadow:0 4px 20px rgba(0,0,0,.08)">
  <tr><td style="background:linear-gradient(135deg,#0d1b3e,#1a56db);
                 padding:28px;text-align:center">
    <div style="font-size:20px;font-weight:900;color:#fff">Employment Verification Request</div>
  </td></tr>
  <tr><td style="padding:36px">
    <p style="font-size:14px;color:#1e293b;line-height:1.8">
      We are conducting a background check for <b>{v['name']}</b>, who has
      indicated employment at your organisation as <b>{prev_role}</b>
      for <b>{duration}</b>.
    </p>
    <p style="font-size:13px;color:#64748b;margin-bottom:24px">
      Please respond using the buttons below:
    </p>
    <table cellpadding="0" cellspacing="0"><tr>
      <td style="padding-right:12px">
        <a href="{vlink}"
          style="display:inline-block;background:#10b981;color:#fff;
                 padding:12px 28px;border-radius:8px;text-decoration:none;
                 font-weight:700;font-size:14px">✓ Verify Employment</a>
      </td>
      <td>
        <a href="{rlink}"
          style="display:inline-block;background:#ef4444;color:#fff;
                 padding:12px 28px;border-radius:8px;text-decoration:none;
                 font-weight:700;font-size:14px">✕ Cannot Verify</a>
      </td>
    </tr></table>
  </td></tr>
</table></td></tr></table></body></html>""",
                user_id=v.get('hr_id'), candidate_id=vid,
                email_type='employment_verification')
        else:
            db.verifications.update_one({'id': vid}, {'$set': {
                'candidate_type':      ctype,
                'verification_status': 'verified',
                'submitted_at':        str(datetime.now()),
            }})
            send_async(v['email'], '🎉 Background Verification Complete!',
                f"""<html><body style="font-family:Arial;margin:0;padding:32px 0;background:#f0f4fa">
<table width="100%" cellpadding="0" cellspacing="0"><tr><td align="center">
<table width="500" style="background:#fff;border-radius:16px;overflow:hidden">
  <tr><td style="background:linear-gradient(135deg,#10b981,#059669);
                 padding:28px;text-align:center">
    <div style="font-size:20px;font-weight:900;color:#fff">Verification Successful ✅</div>
  </td></tr>
  <tr><td style="padding:36px;text-align:center">
    <div style="font-size:44px;margin-bottom:14px">🎊</div>
    <h2 style="font-size:20px;color:#065f46">Congratulations, {v['name']}!</h2>
    <p style="color:#64748b;font-size:13px;line-height:1.8">
      Your background verification is complete. Welcome to <b>{company}</b>!
      HR will be in touch with onboarding details shortly.
    </p>
  </td></tr>
</table></td></tr></table></body></html>""",
                user_id=v.get('hr_id'), candidate_id=vid,
                email_type='verification_complete')

    v = db.verifications.find_one({'id': vid}, {'_id': 0})
    return jsonify({'success': True, 'status': v.get('verification_status')})

@app.route('/company-verify/<vid>/<action>')
def company_verify(vid, action):
    db = get_db()
    v  = db.verifications.find_one({'id': vid}, {'_id': 0})
    if not v:
        return "<div style='font-family:sans-serif;text-align:center;margin-top:80px'><h2>Invalid link.</h2></div>", 404
    hr      = db.users.find_one({'id': v.get('hr_id', '')}, {'_id': 0}) or {}
    company = hr.get('company_name', 'the company')

    if action == 'verify':
        db.verifications.update_one({'id': vid},
                                    {'$set': {'verification_status': 'verified'}})
        send_async(v['email'], '✅ Employment Verified — Welcome Aboard!',
            f"""<html><body style="font-family:Arial;margin:0;padding:32px 0;background:#f0f4fa">
<table width="100%" cellpadding="0" cellspacing="0"><tr><td align="center">
<table width="500" style="background:#fff;border-radius:16px;overflow:hidden">
  <tr><td style="background:linear-gradient(135deg,#10b981,#059669);
                 padding:28px;text-align:center">
    <div style="font-size:20px;font-weight:900;color:#fff">Background Verified ✅</div>
  </td></tr>
  <tr><td style="padding:36px;text-align:center">
    <div style="font-size:44px;margin-bottom:14px">🎉</div>
    <h2 style="font-size:18px;color:#065f46">Congratulations, {v['name']}!</h2>
    <p style="color:#64748b;font-size:13px;line-height:1.8">
      Your employment has been verified. You are now a fully verified candidate
      at <b>{company}</b>. HR will send your onboarding schedule shortly.
    </p>
  </td></tr>
</table></td></tr></table></body></html>""",
            user_id=v.get('hr_id'), candidate_id=vid,
            email_type='verification_verified')

    elif action == 'reject':
        db.verifications.update_one({'id': vid},
                                    {'$set': {'verification_status': 'rejected'}})
        send_async(v['email'], 'Update on Your Application',
            f"""<html><body style="font-family:Arial;margin:0;padding:32px 0;background:#f0f4fa">
<table width="100%" cellpadding="0" cellspacing="0"><tr><td align="center">
<table width="500" style="background:#fff;border-radius:16px;overflow:hidden">
  <tr><td style="background:#ef4444;padding:28px;text-align:center">
    <div style="font-size:20px;font-weight:900;color:#fff">Verification Update</div>
  </td></tr>
  <tr><td style="padding:36px;text-align:center">
    <h2 style="font-size:18px;color:#991b1b">Dear {v['name']},</h2>
    <p style="color:#64748b;font-size:13px;line-height:1.8">
      We were unable to verify your employment history as provided.
      We regret we cannot proceed with your application at this time.
      Thank you for your interest.
    </p>
  </td></tr>
</table></td></tr></table></body></html>""",
            user_id=v.get('hr_id'), candidate_id=vid,
            email_type='verification_rejected')

    v = db.verifications.find_one({'id': vid}, {'_id': 0})
    return render_template('company_verify_done.html', action=action, candidate=v)

@app.route('/api/download-template')
def download_template():
    u = current_user()
    if not u:
        return redirect('/login')
    cands = list(get_db().candidates.find({'hr_id': u['id']}, {'_id': 0}))
    rows  = [{'Name': c['name'], 'Gmail id': c['email'], 'Role': c['role'],
              'Joining date': c['joining_date'], 'Salary': c['salary'],
              'Employment Type': c['employment_type'], 'Status': c['offer_status']}
             for c in cands] or \
            [{'Name': '', 'Gmail id': '', 'Role': '', 'Joining date': '',
              'Salary': '', 'Employment Type': '', 'Status': ''}]
    df   = pd.DataFrame(rows)
    path = os.path.join('/tmp', 'offers_export.xlsx')
    df.to_excel(path, index=False)
    return send_file(path, as_attachment=True, download_name='offers_export.xlsx')

@app.route('/settings')
def settings_page():
    if not current_user():
        return redirect('/login')
    return render_template('settings.html')

@app.route('/api/settings')
def api_settings():
    u = current_user()
    if not u:
        return jsonify({'success': False}), 401
    return jsonify({'success': True, 'user': u})

@app.route('/api/update-profile', methods=['POST'])
def api_update_profile():
    u = current_user()
    if not u:
        return jsonify({'success': False}), 401
    data         = request.json or {}
    email        = sanitize_string(data.get('email', ''))
    company_name = sanitize_string(data.get('company_name', ''))
    if not email or not company_name:
        return jsonify({'success': False, 'message': 'All fields are required'}), 400
    if not validate_email(email):
        return jsonify({'success': False, 'message': 'Invalid email format'}), 400
    db = get_db()
    existing = db.users.find_one({'email': email, 'id': {'$ne': u['id']}})
    if existing:
        return jsonify({'success': False, 'message': 'Email already registered'}), 400
    db.users.update_one({'id': u['id']},
                        {'$set': {'email': email, 'company_name': company_name}})
    logger.info(f'User {u["id"]} updated profile')
    return jsonify({'success': True})

@app.route('/api/change-password', methods=['POST'])
def api_change_password():
    u = current_user()
    if not u:
        return jsonify({'success': False}), 401
    data             = request.json or {}
    current_password = data.get('current_password', '')
    new_password     = data.get('new_password', '')
    if not current_password or not new_password:
        return jsonify({'success': False, 'message': 'All fields are required'}), 400
    if u['password'] != hash_pw(current_password):
        return jsonify({'success': False, 'message': 'Current password is incorrect'}), 400
    if not validate_password(new_password):
        return jsonify({'success': False, 'message':
            'Password needs 8+ chars, uppercase, lowercase & special character'}), 400
    get_db().users.update_one({'id': u['id']},
                              {'$set': {'password': hash_pw(new_password)}})
    logger.info(f'User {u["id"]} changed password')
    return jsonify({'success': True})

@app.route('/api/update-letterhead', methods=['POST'])
def api_update_letterhead():
    u = current_user()
    if not u:
        return jsonify({'success': False}), 401
    lh = request.files.get('letterhead')
    if not lh or not lh.filename:
        return jsonify({'success': False, 'message': 'Letterhead file is required'}), 400
    if not lh.filename.lower().endswith('.pdf'):
        return jsonify({'success': False, 'message': 'Letterhead must be a PDF file'}), 400
    if not validate_file_size(lh, max_size_mb=5):
        return jsonify({'success': False, 'message': 'Letterhead file size must be less than 5MB'}), 400
    fname   = secure_filename(lh.filename)
    lh_path = os.path.join(UPLOAD_DIR, 'letterheads', fname)
    lh.save(lh_path)
    project_lh = os.path.join(BASE_DIR, 'uploads', 'letterheads', fname)
    if not os.path.exists(project_lh):
        shutil.copy2(lh_path, project_lh)
    get_db().users.update_one({'id': u['id']}, {'$set': {'letterhead': lh_path}})
    logger.info(f'User {u["id"]} updated letterhead')
    return jsonify({'success': True})

@app.route('/api/email-history')
def api_email_history():
    u = current_user()
    if not u:
        return jsonify({'success': False}), 401
    emails = list(get_db().email_history.find(
        {'user_id': u['id']}, {'_id': 0},
        sort=[('sent_at', -1)]))
    return jsonify({'success': True, 'emails': emails})

@app.route('/api/download-offer-letter/<cid>')
def download_offer_letter(cid):
    u = current_user()
    if not u:
        return jsonify({'success': False}), 401
    c = get_db().candidates.find_one({'id': cid}, {'_id': 0})
    if not c:
        return jsonify({'success': False, 'message': 'Candidate not found'}), 404
    if c.get('hr_id') != u['id']:
        return jsonify({'success': False, 'message': 'Unauthorized'}), 403
    pdf_path = os.path.join(LETTER_DIR, f"offer_{cid}.pdf")
    if not os.path.exists(pdf_path):
        return jsonify({'success': False,
            'message': 'Offer letter not found. Please send the offer first.'}), 404
    return send_file(pdf_path, as_attachment=True,
                     download_name=f"Offer_Letter_{c['name'].replace(' ', '_')}.pdf")

@app.route('/api/cancel-offer/<cid>', methods=['POST'])
def cancel_offer(cid):
    u = current_user()
    if not u:
        return jsonify({'success': False}), 401
    db = get_db()
    c  = db.candidates.find_one({'id': cid}, {'_id': 0})
    if not c:
        return jsonify({'success': False, 'message': 'Candidate not found'}), 404
    if c.get('hr_id') != u['id']:
        return jsonify({'success': False, 'message': 'Unauthorized'}), 403
    if c.get('offer_status') != 'pending':
        return jsonify({'success': False, 'message': 'Can only cancel pending offers'}), 400
    db.candidates.update_one({'id': cid}, {'$set': {
        'offer_status': 'cancelled',
        'cancelled_at': str(datetime.now()),
    }})
    logger.info(f'Offer cancelled for candidate {cid} by user {u["id"]}')
    return jsonify({'success': True, 'message': 'Offer cancelled successfully'})

@app.route('/api/get-candidate/<cid>')
def get_candidate(cid):
    u = current_user()
    if not u:
        return jsonify({'success': False}), 401
    c = get_db().candidates.find_one({'id': cid}, {'_id': 0})
    if not c:
        return jsonify({'success': False, 'message': 'Candidate not found'}), 404
    if c.get('hr_id') != u['id']:
        return jsonify({'success': False, 'message': 'Unauthorized'}), 403
    return jsonify({'success': True, 'candidate': c})

@app.route('/api/update-candidate/<cid>', methods=['POST'])
def update_candidate(cid):
    u = current_user()
    if not u:
        return jsonify({'success': False}), 401
    db = get_db()
    c  = db.candidates.find_one({'id': cid}, {'_id': 0})
    if not c:
        return jsonify({'success': False, 'message': 'Candidate not found'}), 404
    if c.get('hr_id') != u['id']:
        return jsonify({'success': False, 'message': 'Unauthorized'}), 403
    if c.get('email_sent_at'):
        return jsonify({'success': False,
            'message': 'Cannot edit candidate after email is sent'}), 400
    data    = request.json or {}
    updates = {}
    for field in ['name', 'email', 'role', 'joining_date', 'salary', 'employment_type']:
        if field in data:
            updates[field] = sanitize_string(data[field])
    if 'email' in updates and not validate_email(updates['email']):
        return jsonify({'success': False, 'message': 'Invalid email format'}), 400
    db.candidates.update_one({'id': cid}, {'$set': updates})
    logger.info(f'Candidate {cid} updated by user {u["id"]}')
    return jsonify({'success': True, 'message': 'Candidate updated successfully'})

# ── Debug route (remove after confirming env vars on Vercel) ──────────────────
@app.route('/api/debug-env')
def debug_env():
    return jsonify({
        'brevo_key_set': bool(BREVO_API_KEY),
        'sender_email':  SENDER_EMAIL,
        'sender_name':   SENDER_NAME,
        'base_url':      BASE_URL,
        'is_vercel':     IS_VERCEL,
        'mongo_set':     bool(MONGO_URI),
    })


# ── Email sender (synchronous — threads are killed on Vercel) ─────────────────
def send_async(to, subject, html, attach_path=None, attach_name=None,
               user_id=None, candidate_id=None, email_type='offer'):
    """Synchronous on Vercel; background thread locally for speed."""
    if IS_VERCEL:
        return send_email(to, subject, html, attach_path, attach_name,
                          user_id, candidate_id, email_type)
    t = threading.Thread(
        target=send_email,
        args=(to, subject, html, attach_path, attach_name,
              user_id, candidate_id, email_type),
        daemon=True)
    t.start()
    return True   # optimistic on local


def send_email(to, subject, html_body, attach_path=None, attach_name=None,
               user_id=None, candidate_id=None, email_type='offer'):
    if not BREVO_API_KEY:
        logger.error('BREVO_API_KEY is not set')
        return False

    email_id = str(uuid.uuid4())
    success  = False
    error_msg = None

    try:
        configuration = sib_api_v3_sdk.Configuration()
        configuration.api_key['api-key'] = BREVO_API_KEY
        api_instance  = sib_api_v3_sdk.TransactionalEmailsApi(
            sib_api_v3_sdk.ApiClient(configuration))

        attachments = []
        if attach_path and os.path.exists(attach_path):
            with open(attach_path, 'rb') as f:
                encoded_file = base64.b64encode(f.read()).decode()
            attachments.append({'content': encoded_file,
                                 'name': attach_name or 'offer_letter.pdf'})

        email_params = {
            'to':           [{'email': to}],
            'sender':       {'name': SENDER_NAME, 'email': SENDER_EMAIL},
            'subject':      subject,
            'html_content': html_body,
        }
        if attachments:
            email_params['attachment'] = attachments

        api_instance.send_transac_email(
            sib_api_v3_sdk.SendSmtpEmail(**email_params))
        logger.info(f'Email sent to {to}')
        success = True

    except ApiException as e:
        logger.error(f'Brevo API error: {e}')
        error_msg = str(e)
    except Exception as e:
        logger.error(f'General email error: {e}')
        error_msg = str(e)

    # Log to MongoDB
    try:
        get_db().email_history.insert_one({
            'id':             email_id,
            'user_id':        user_id,
            'candidate_id':   candidate_id,
            'email_type':     email_type,
            'to':             to,
            'subject':        subject,
            'sent_at':        str(datetime.now()),
            'status':         'sent' if success else 'failed',
            'error':          error_msg,
            'has_attachment': bool(attach_path),
        })
    except Exception as e:
        logger.warning(f'Failed to log email history: {e}')

    return success


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)