"""
OfferFlow — app.py  (updated)
Changes:
  1. Pattern selection: 4 named patterns (Classic, Modern, Minimal, Executive) +
     a Custom pattern that accepts free-text via textarea. Each pattern renders a
     distinctly structured PDF.
  2. Preview: offer-letter content is composited ON TOP of the letterhead PDF
     (page-1 rendered as background image) so the preview matches the real PDF.
  3. Background-verification email is sent ONLY after the candidate clicks
     "Accept Offer", which was already partially implemented but is now
     guaranteed to fire once and include the correct verification link.
"""

import os, io, json, uuid, base64, hashlib, re, threading
from datetime import datetime, timedelta

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

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'dev-secret-key-change-me')

BASE_URL      = os.environ.get('BASE_URL', 'http://localhost:5000')
BREVO_API_KEY = os.environ.get('BREVO_API_KEY', '')
SENDER_EMAIL  = os.environ.get('SENDER_EMAIL', 'noreply@offerflow.com')
SENDER_NAME   = os.environ.get('SENDER_NAME',  'OfferFlow HR')

BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(BASE_DIR, 'uploads')
DATA_DIR   = os.path.join(BASE_DIR, 'data')
LETTER_DIR = os.path.join(BASE_DIR, 'generated_letters')
for d in [UPLOAD_DIR+'/letterheads', UPLOAD_DIR+'/excel',
          UPLOAD_DIR+'/documents', DATA_DIR, LETTER_DIR]:
    os.makedirs(d, exist_ok=True)

# ── JSON helpers ──────────────────────────────────────────────────────────────
def load_json(path):
    return json.load(open(path)) if os.path.exists(path) else {}

def save_json(path, data):
    json.dump(data, open(path, 'w'), indent=2, default=str)

USERS_F  = DATA_DIR+'/users.json'
CANDS_F  = DATA_DIR+'/candidates.json'
VERIFS_F = DATA_DIR+'/verifications.json'
TOKENS_F = DATA_DIR+'/tokens.json'

def get_users():  return load_json(USERS_F)
def get_cands():  return load_json(CANDS_F)
def get_verifs(): return load_json(VERIFS_F)
def get_tokens(): return load_json(TOKENS_F)
def save_users(d):  save_json(USERS_F,  d)
def save_cands(d):  save_json(CANDS_F,  d)
def save_verifs(d): save_json(VERIFS_F, d)
def save_tokens(d): save_json(TOKENS_F, d)

def hash_pw(pw): return hashlib.sha256(pw.encode()).hexdigest()

def validate_password(pw):
    return (len(pw) >= 8 and re.search(r'[A-Z]', pw)
            and re.search(r'[a-z]', pw) and re.search(r'[^A-Za-z0-9]', pw))

def current_user():
    uid = session.get('user_id')
    return get_users().get(uid) if uid else None


# ── Email via Brevo API ───────────────────────────────────────────────────────
def send_email(to, subject, html_body, attach_path=None, attach_name=None):
    try:
        configuration = sib_api_v3_sdk.Configuration()
        configuration.api_key['api-key'] = BREVO_API_KEY
        api_instance = sib_api_v3_sdk.TransactionalEmailsApi(
            sib_api_v3_sdk.ApiClient(configuration))
        attachments = []
        if attach_path and os.path.exists(attach_path):
            with open(attach_path, 'rb') as f:
                encoded_file = base64.b64encode(f.read()).decode()
            attachments.append({'content': encoded_file,
                                 'name': attach_name or 'offer_letter.pdf'})
        send_smtp_email = sib_api_v3_sdk.SendSmtpEmail(
            to=[{'email': to}],
            sender={'name': SENDER_NAME, 'email': SENDER_EMAIL},
            subject=subject,
            html_content=html_body,
            attachment=attachments)
        api_instance.send_transac_email(send_smtp_email)
        print(f'[BREVO EMAIL SENT] → {to}')
        return True
    except ApiException as e:
        print(f'[BREVO ERROR] → {e}')
        return False
    except Exception as e:
        print(f'[GENERAL EMAIL ERROR] → {e}')
        return False

def send_async(to, subject, html, attach_path=None, attach_name=None):
    threading.Thread(target=send_email,
                     args=(to, subject, html, attach_path, attach_name),
                     daemon=True).start()


# ══════════════════════════════════════════════════════════════════════════════
#  PATTERN DEFINITIONS
#  Each pattern is a dict consumed by generate_offer_pdf() and
#  letterhead_preview_html() to change layout/colours/structure.
# ══════════════════════════════════════════════════════════════════════════════
PATTERNS = {
    'classic': {
        'label':       'Classic',
        'accent':      '#0d1b3e',
        'accent2':     '#1a56db',
        'header_bg':   '#0d1b3e',
        'row_even':    '#f0f4ff',
        'row_odd':     '#ffffff',
        'font_header': 'Helvetica-Bold',
        'layout':      'classic',      # centred header + table
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
        'layout':      'modern',       # left-aligned, two-column KV table
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
        'layout':      'minimal',      # plain, no heavy borders
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
        'layout':      'executive',    # formal letter with signature block
        'desc':        'Premium dark-red executive format with formal letter structure.',
    },
    'custom': {
        'label':  'Custom',
        'layout': 'custom',
        'desc':   'Write your own offer letter content in the text area below.',
    },
}


# ══════════════════════════════════════════════════════════════════════════════
#  PDF GENERATOR  (pattern-aware)
# ══════════════════════════════════════════════════════════════════════════════
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
    lh_path  = hr_user.get('letterhead', '')

    # Choose top-margin based on whether letterhead exists
    top_margin = 54 * mm if (lh_path and os.path.exists(lh_path)) else 18 * mm

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4,
                            rightMargin=22*mm, leftMargin=22*mm,
                            topMargin=top_margin, bottomMargin=22*mm)

    styles   = getSampleStyleSheet()
    accent   = colors.HexColor(pat.get('accent', '#0d1b3e'))
    accent2  = colors.HexColor(pat.get('accent2', '#1a56db'))
    row_even = colors.HexColor(pat.get('row_even', '#f0f4ff'))
    row_odd  = colors.HexColor(pat.get('row_odd',  '#ffffff'))

    def ps(n, **kw):
        return ParagraphStyle(n, parent=styles['Normal'], **kw)

    layout = pat.get('layout', 'classic')

    # ── CUSTOM PATTERN ────────────────────────────────────────────────────────
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

    # ── TABLE OF DETAILS (shared by non-custom patterns) ─────────────────────
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
            Paragraph('• Successful completion of background verification.', ps('b3', fontSize=10.5, leading=17, spaceAfter=4)),
            Paragraph('• Submission of all required documents before joining.', ps('b4', fontSize=10.5, leading=17, spaceAfter=4)),
            Paragraph("• Acceptance of the company's Code of Conduct.", ps('b5', fontSize=10.5, leading=17, spaceAfter=8)),
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
        # Left-aligned header bar + two-column key-value cards
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
            Paragraph('• Background verification must be completed successfully.', ps('mc', fontSize=10.5, leading=17, spaceAfter=4)),
            Paragraph('• All required documents must be submitted before joining.', ps('mc2', fontSize=10.5, leading=17, spaceAfter=4)),
            Paragraph("• You must accept the company's Code of Conduct.", ps('mc3', fontSize=10.5, leading=17, spaceAfter=14)),
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
            Paragraph(f'Regards,',
                      ps('mis', fontSize=10.5, leading=17, spaceAfter=4)),
            Paragraph(f'{company} HR',
                      ps('mis2', fontSize=10.5, textColor=accent, spaceAfter=4)),
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
            HRFlowable(width='100%', thickness=3, color=accent, spaceAfter=6),
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
            Paragraph('1. Satisfactory completion of background and reference verification.', ps('ec', fontSize=10.5, leading=17, spaceAfter=4)),
            Paragraph('2. Production of original educational and professional documents.', ps('ec2', fontSize=10.5, leading=17, spaceAfter=4)),
            Paragraph("3. Execution of the Company's Employment Agreement and NDA.", ps('ec3', fontSize=10.5, leading=17, spaceAfter=14)),
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
                      ps('esig', fontSize=10, textColor=colors.HexColor('#9ca3af'), spaceAfter=2)),
            Paragraph(f'<b>Authorised Signatory</b>',
                      ps('esig2', fontName='Helvetica-Bold', fontSize=10.5,
                         textColor=accent, spaceAfter=2)),
            Paragraph(f'{company}',
                      ps('esig3', fontSize=10, textColor=colors.HexColor('#6b7280'))),
        ]

    else:
        story = [Paragraph('Offer letter', styles['Normal'])]

    _build_pdf(doc, buf, story, lh_path, out_path)
    return out_path


# ── helpers ───────────────────────────────────────────────────────────────────
def _kv_cell(label, value, accent_color):
    """Small two-line key-value card paragraph for Modern layout."""
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
    tbl = Table(tbl_data, colWidths=[68*mm, 102*mm])
    cmds = [
        ('BACKGROUND',  (0, 0), (-1,  0), accent),
        ('TEXTCOLOR',   (0, 0), (-1,  0), colors.white),
        ('FONTNAME',    (0, 0), (-1,  0), 'Helvetica-Bold'),
        ('FONTSIZE',    (0, 0), (-1,  0), 10),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [row_even, row_odd]),
        ('FONTNAME',    (0, 1), (0, -1), 'Helvetica-Bold'),
        ('FONTNAME',    (1, 1), (1, -1), 'Helvetica'),
        ('FONTSIZE',    (0, 1), (-1, -1), 10),
        ('TEXTCOLOR',   (0, 1), (0, -1), accent2),
        ('ROWPADDING',  (0, 0), (-1, -1), 9),
        ('VALIGN',      (0, 0), (-1, -1), 'MIDDLE'),
    ]
    if grid:
        cmds.append(('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#e2e8f0')))
    tbl.setStyle(TableStyle(cmds))
    return tbl


def _build_pdf(doc, buf, story, lh_path, out_path):
    """Build PDF; if a letterhead exists, merge it as background."""
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
            print(f'Merge failed: {e}; using plain PDF')
            doc.build(story)
            buf.seek(0)
            open(out_path, 'wb').write(buf.read())
    else:
        doc.build(story)
        buf.seek(0)
        open(out_path, 'wb').write(buf.read())


def _merge_letterhead(lh_path, content_bytes, out_path):
    """
    Composite the letterhead (page 1) as a background image under the
    content PDF, then save to out_path.
    Requires: pip install pdf2image pillow pypdf
    Falls back to plain content PDF if dependencies are missing.
    """
    try:
        from pdf2image import convert_from_path
        import tempfile
        from pypdf import PdfReader, PdfWriter

        # Render letterhead page-1 to PNG
        imgs = convert_from_path(lh_path, first_page=1, last_page=1, dpi=150)
        if not imgs:
            raise ValueError('No pages in letterhead')
        tmp_img = tempfile.mktemp(suffix='.png')
        imgs[0].save(tmp_img, 'PNG')

        # Build a single-page PDF that is just the letterhead image
        tmp_lh_pdf = tempfile.mktemp(suffix='.pdf')
        w_pt, h_pt = A4
        c = rlc.Canvas(tmp_lh_pdf, pagesize=A4)
        c.drawImage(tmp_img, 0, 0, width=w_pt, height=h_pt)
        c.save()
        os.unlink(tmp_img)

        # Merge: letterhead as background, content on top
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


# ══════════════════════════════════════════════════════════════════════════════
#  PREVIEW HTML  (pattern-aware, letterhead as background)
# ══════════════════════════════════════════════════════════════════════════════
def letterhead_preview_html(hr_user, candidate, pattern_key='classic', custom_text=''):
    company  = hr_user['company_name']
    lh_path  = hr_user.get('letterhead', '')
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

    # ── Letterhead background block ───────────────────────────────────────────
    lh_bg_style = ''
    lh_notice   = ''
    if lh_path and os.path.exists(lh_path):
        # Convert first page of letterhead PDF → base64 PNG for CSS background
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
            # pdf2image unavailable — show plain iframe fallback
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

    # ── Pattern badge ─────────────────────────────────────────────────────────
    pat_badge = (
        f'<span style="display:inline-block;background:{accent};color:#fff;'
        f'font-size:10px;padding:2px 10px;border-radius:12px;'
        f'font-family:Arial,sans-serif;margin-bottom:14px">'
        f'{pat["label"]} Template</span>'
    )

    # ── Content varies by pattern ─────────────────────────────────────────────
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
  <!-- semi-transparent white overlay so content is readable over letterhead -->
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


# ══════════════════════════════════════════════════════════════════════════════
#  EMAIL HTML BUILDERS
# ══════════════════════════════════════════════════════════════════════════════
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
    """
    Email sent to the candidate after they ACCEPT the offer,
    asking them to complete background verification.
    """
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
    data = request.json or {}
    user = next((u for u in get_users().values()
                 if u['username'] == data.get('username')
                 and u['password'] == hash_pw(data.get('password', ''))), None)
    if not user:
        return jsonify({'success': False, 'message': 'Invalid credentials'}), 401
    session['user_id'] = user['id']
    return jsonify({'success': True})

@app.route('/api/register', methods=['POST'])
def api_register():
    users        = get_users()
    username     = request.form.get('username', '').strip()
    email        = request.form.get('email', '').strip()
    password     = request.form.get('password', '')
    confirm      = request.form.get('confirm_password', '')
    company_name = request.form.get('company_name', '').strip()
    if not all([username, email, password, confirm, company_name]):
        return jsonify({'success': False, 'message': 'All fields are required'}), 400
    if any(u['username'] == username for u in users.values()):
        return jsonify({'success': False, 'message': 'Username already exists'}), 400
    if any(u['email'] == email for u in users.values()):
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
    fname   = secure_filename(lh.filename)
    lh_path = os.path.join(UPLOAD_DIR, 'letterheads', fname)
    lh.save(lh_path)
    uid = str(uuid.uuid4())
    users[uid] = {'id': uid, 'username': username, 'email': email,
                  'password': hash_pw(password), 'company_name': company_name,
                  'letterhead': lh_path, 'created_at': str(datetime.now())}
    save_users(users)
    return jsonify({'success': True})

@app.route('/api/forgot-password', methods=['POST'])
def api_forgot_password():
    email = (request.json or {}).get('email', '').strip()
    user  = next((u for u in get_users().values() if u['email'] == email), None)
    if not user:
        return jsonify({'success': False, 'message': 'Email not found'}), 404
    token  = str(uuid.uuid4())
    tokens = get_tokens()
    tokens[token] = {'user_id': user['id'],
                     'expires': str(datetime.now() + timedelta(hours=1))}
    save_tokens(tokens)
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
    send_async(email, 'Reset Your OfferFlow Password', html)
    return jsonify({'success': True, 'message': 'Password reset link sent to your email'})

@app.route('/reset-password/<token>')
def reset_password_page(token):
    if token not in get_tokens():
        return redirect('/login')
    return render_template('reset_password.html', token=token)

@app.route('/api/reset-password', methods=['POST'])
def api_reset_password():
    data    = request.json or {}
    token   = data.get('token', '')
    pw      = data.get('password', '')
    confirm = data.get('confirm_password', '')
    tokens  = get_tokens()
    if token not in tokens:
        return jsonify({'success': False, 'message': 'Invalid or expired link'}), 400
    if pw != confirm:
        return jsonify({'success': False, 'message': 'Passwords do not match'}), 400
    if not validate_password(pw):
        return jsonify({'success': False, 'message': 'Password requirements not met'}), 400
    users = get_users()
    uid   = tokens[token]['user_id']
    users[uid]['password'] = hash_pw(pw)
    save_users(users)
    del tokens[token]; save_tokens(tokens)
    return jsonify({'success': True})

@app.route('/logout')
def logout():
    session.clear()
    return redirect('/')

# Dashboard
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
    uid    = u['id']
    cands  = {k: v for k, v in get_cands().items()  if v.get('hr_id') == uid}
    verifs = {k: v for k, v in get_verifs().items() if v.get('hr_id') == uid}
    return jsonify({
        'total_offers':         len(cands),
        'offer_accepted':       sum(1 for c in cands.values()  if c.get('offer_status') == 'accepted'),
        'offer_declined':       sum(1 for c in cands.values()  if c.get('offer_status') == 'declined'),
        'action_pending':       sum(1 for c in cands.values()  if c.get('offer_status') == 'pending'),
        'cancelled':            sum(1 for c in cands.values()  if c.get('offer_status') == 'cancelled'),
        'verified':             sum(1 for v in verifs.values() if v.get('verification_status') == 'verified'),
        'rejected':             sum(1 for v in verifs.values() if v.get('verification_status') == 'rejected'),
        'verification_pending': sum(1 for v in verifs.values() if v.get('verification_status') == 'pending'),
        'company_name': u['company_name'],
    })

@app.route('/api/offers')
def api_offers():
    u = current_user()
    if not u:
        return jsonify([]), 401
    status = request.args.get('status')
    result = [{'id': cid, 'name': c['name'], 'role': c.get('role', ''),
               'joining_date': c.get('joining_date', ''), 'email': c.get('email', ''),
               'salary': c.get('salary', ''), 'employment_type': c.get('employment_type', ''),
               'offer_status': c.get('offer_status', 'pending')}
              for cid, c in get_cands().items() if c.get('hr_id') == u['id']]
    if status:
        result = [r for r in result if r['offer_status'] == status]
    return jsonify(result)

@app.route('/api/verifications')
def api_verifications():
    u = current_user()
    if not u:
        return jsonify([]), 401
    status = request.args.get('status')
    result = [{'id': vid, 'name': v['name'], 'email': v.get('email', ''),
               'salary': v.get('salary', ''), 'phone': v.get('phone', ''),
               'verification_status': v.get('verification_status', 'pending')}
              for vid, v in get_verifs().items() if v.get('hr_id') == u['id']]
    if status:
        result = [r for r in result if r['verification_status'] == status]
    return jsonify(result)

# ── Upload Excel ──────────────────────────────────────────────────────────────
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
    if not f or not f.filename.endswith('.xlsx'):
        return jsonify({'success': False, 'message': 'Please upload a .xlsx file'}), 400
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
    data  = request.json or {}
    cands = get_cands()
    new_ids = []
    for rec in data.get('candidates', []):
        cid = str(uuid.uuid4())
        cands[cid] = {
            'id': cid, 'hr_id': u['id'],
            'name':          rec.get('Name', ''),
            'email':         rec.get('Gmail id', ''),
            'role':          rec.get('Role', ''),
            'joining_date':  str(rec.get('Joining date', '')),
            'salary':        str(rec.get('Salary', '')),
            'employment_type': data.get('employment_type', 'full_time'),
            'pattern':       data.get('pattern', 'classic'),
            'custom_text':   data.get('custom_text', ''),
            'offer_status':  'pending',
            'sent_at':       str(datetime.now()),
            'extra': {k: v for k, v in rec.items()
                      if k not in ['Name', 'Gmail id', 'Role', 'Joining date', 'Salary']},
        }
        new_ids.append(cid)
    save_cands(cands)
    return jsonify({'success': True, 'candidate_ids': new_ids})


# ── Pattern list (for frontend to build the selector UI) ─────────────────────
@app.route('/api/patterns')
def api_patterns():
    """Return pattern metadata so the frontend can render the selector."""
    return jsonify([
        {'key': k, 'label': v['label'], 'desc': v['desc'],
         'accent': v.get('accent', '#0d1b3e'), 'is_custom': k == 'custom'}
        for k, v in PATTERNS.items()
    ])


# ── Preview ───────────────────────────────────────────────────────────────────
@app.route('/api/preview-letter', methods=['POST'])
def api_preview_letter():
    u = current_user()
    if not u:
        return jsonify({'success': False}), 401
    data        = request.json or {}
    c           = data.get('candidate', {})
    c['employment_type'] = data.get('employment_type', 'full_time')
    c['id']     = 'preview'
    pattern_key = data.get('pattern', 'classic')
    custom_text = data.get('custom_text', '')
    return jsonify({
        'success': True,
        'html': letterhead_preview_html(u, c, pattern_key, custom_text),
    })


# ── Send offer emails ─────────────────────────────────────────────────────────
@app.route('/api/send-offer-emails', methods=['POST'])
def api_send_offer_emails():
    u = current_user()
    if not u:
        return jsonify({'success': False}), 401
    data  = request.json or {}
    cids  = data.get('candidate_ids', [])
    cands = get_cands()
    sent  = 0
    for cid in cids:
        c = cands.get(cid)
        if not c:
            continue
        if c.get('email_sent_at'):
            print(f'Email already sent to {c["email"]}')
            continue
        accept_link  = f"{BASE_URL}/offer-response/{cid}/accept"
        decline_link = f"{BASE_URL}/offer-response/{cid}/decline"
        pattern_key  = c.get('pattern', 'classic')
        custom_text  = c.get('custom_text', '')
        try:
            pdf_path = generate_offer_pdf(c, u, pattern_key, custom_text)
        except Exception as e:
            print(f'PDF failed for {cid}: {e}')
            pdf_path = None
        subject = f"Job Offer — {c.get('role', '')} at {u['company_name']}"
        send_async(
            c['email'], subject,
            offer_email_html(c, u, accept_link, decline_link),
            attach_path=pdf_path,
            attach_name=f"Offer_Letter_{c['name'].replace(' ', '_')}.pdf")
        cands[cid]['email_sent_at'] = str(datetime.now())
        sent += 1
    save_cands(cands)

    # Auto-cancel after 48 hours
    def auto_cancel(ids):
        import time
        time.sleep(172800)
        c2 = get_cands()
        changed = False
        for i in ids:
            if c2.get(i, {}).get('offer_status') == 'pending':
                c2[i]['offer_status'] = 'cancelled'
                changed = True
        if changed:
            save_cands(c2)

    threading.Thread(target=auto_cancel, args=(cids,), daemon=True).start()
    return jsonify({'success': True, 'sent': sent})


# ── Candidate offer response (Accept / Decline) ───────────────────────────────
@app.route('/offer-response/<cid>/<action>')
def offer_response(cid, action):
    cands = get_cands()
    c     = cands.get(cid)
    if not c:
        return ("<div style='font-family:sans-serif;text-align:center;margin-top:80px'>"
                "<h2>Invalid or expired link.</h2></div>"), 404

    current_status = c.get('offer_status', 'pending')
    if current_status in ['accepted', 'declined', 'cancelled']:
        return render_template(
            'offer_accepted.html' if current_status == 'accepted' else 'offer_declined.html',
            candidate=c)

    # ── ACCEPT ────────────────────────────────────────────────────────────────
    if action == 'accept':
        c['offer_status']  = 'accepted'
        c['responded_at']  = str(datetime.now())
        save_cands(cands)

        verifs   = get_verifs()
        existing = next((v for v in verifs.values() if v.get('candidate_id') == cid), None)

        if not existing:
            vid = str(uuid.uuid4())
            verifs[vid] = {
                'id':                  vid,
                'candidate_id':        cid,
                'hr_id':               c['hr_id'],
                'name':                c['name'],
                'email':               c['email'],
                'salary':              c['salary'],
                'phone':               '',
                'verification_status': 'pending',
                'created_at':          str(datetime.now()),
            }
            save_verifs(verifs)

            bg_link = f"{BASE_URL}/background-verification/{vid}"
            users   = get_users()
            hr      = users.get(c['hr_id'], {})
            company = hr.get('company_name', 'the company')

            # ── Send nicely formatted BG verification email ───────────────────
            send_async(
                c['email'],
                f'Next Step: Complete Your Background Verification — {company}',
                bg_verification_email_html(c, bg_link, company),
            )

        return render_template('offer_accepted.html', candidate=c)

    # ── DECLINE ───────────────────────────────────────────────────────────────
    elif action == 'decline':
        c['offer_status'] = 'declined'
        c['responded_at'] = str(datetime.now())
        save_cands(cands)
        return render_template('offer_declined.html', candidate=c)

    return "<div style='font-family:sans-serif;text-align:center;margin-top:80px'><h2>Invalid action.</h2></div>"


# ── Background verification form ──────────────────────────────────────────────
@app.route('/background-verification/<vid>')
def background_verification(vid):
    verifs = get_verifs()
    v      = verifs.get(vid)
    if not v:
        return ("<div style='font-family:sans-serif;text-align:center;margin-top:80px'>"
                "<h2>Invalid or expired link.</h2></div>"), 404
    return render_template('background_verification.html', verification=v)

@app.route('/api/submit-verification', methods=['POST'])
def api_submit_verification():
    vid    = request.form.get('verification_id')
    verifs = get_verifs()
    v      = verifs.get(vid)
    if not v:
        return jsonify({'success': False}), 404
    step = request.form.get('step')

    if step == '1':
        fn = request.form.get('first_name', '')
        ln = request.form.get('last_name', '')
        v.update({'first_name': fn, 'last_name': ln, 'name': f"{fn} {ln}",
                  'phone':   request.form.get('phone', ''),
                  'aadhaar': request.form.get('aadhaar', ''),
                  'pan':     request.form.get('pan', '')})
    elif step == '2':
        v.update({'college':        request.form.get('college', ''),
                  'specialization': request.form.get('specialization', ''),
                  'percentage':     request.form.get('percentage', '')})
    elif step == '3':
        ctype    = request.form.get('candidate_type', 'fresher')
        v['candidate_type'] = ctype
        users    = get_users()
        hr       = users.get(v.get('hr_id', ''), {})
        company  = hr.get('company_name', 'the company')

        if ctype == 'experienced':
            prev_co   = request.form.get('prev_company', '')
            prev_role = request.form.get('prev_role', '')
            co_email  = request.form.get('company_email', '')
            duration  = request.form.get('duration', '')
            v.update({'prev_company': prev_co, 'prev_role': prev_role,
                      'company_email': co_email, 'duration': duration})
            v['verification_status'] = 'pending'
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
</table></td></tr></table></body></html>""")
        else:
            # Fresher — auto-verified
            v['verification_status'] = 'verified'
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
</table></td></tr></table></body></html>""")
        v['submitted_at'] = str(datetime.now())

    save_verifs(verifs)
    return jsonify({'success': True, 'status': v.get('verification_status')})


# ── Company verify (employer confirms/rejects previous employment) ────────────
@app.route('/company-verify/<vid>/<action>')
def company_verify(vid, action):
    verifs = get_verifs()
    v      = verifs.get(vid)
    if not v:
        return "<div style='font-family:sans-serif;text-align:center;margin-top:80px'><h2>Invalid link.</h2></div>", 404
    users   = get_users()
    hr      = users.get(v.get('hr_id', ''), {})
    company = hr.get('company_name', 'the company')

    if action == 'verify':
        v['verification_status'] = 'verified'
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
</table></td></tr></table></body></html>""")

    elif action == 'reject':
        v['verification_status'] = 'rejected'
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
</table></td></tr></table></body></html>""")

    save_verifs(verifs)
    return render_template('company_verify_done.html', action=action, candidate=v)


# ── Export / download template ────────────────────────────────────────────────
@app.route('/api/download-template')
def download_template():
    u = current_user()
    if not u:
        return redirect('/login')
    cands = {k: v for k, v in get_cands().items() if v.get('hr_id') == u['id']}
    rows  = [{'Name': c['name'], 'Gmail id': c['email'], 'Role': c['role'],
              'Joining date': c['joining_date'], 'Salary': c['salary'],
              'Employment Type': c['employment_type'], 'Status': c['offer_status']}
             for c in cands.values()] or \
            [{'Name': '', 'Gmail id': '', 'Role': '', 'Joining date': '',
              'Salary': '', 'Employment Type': '', 'Status': ''}]
    df   = pd.DataFrame(rows)
    path = '/tmp/offers_export.xlsx'
    df.to_excel(path, index=False)
    return send_file(path, as_attachment=True, download_name='offers_export.xlsx')


if __name__ == '__main__':
    app.run(debug=True)