from flask import Flask, render_template, request, redirect, url_for, session, flash, send_file, jsonify
from models import db, Pengguna, Permohonan, Persetujuan
from config import Config
import pymysql
import os
from werkzeug.utils import secure_filename
from docx import Document
from docx.shared import Cm
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from datetime import datetime
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from io import BytesIO
from reportlab.platypus import Image as RLImage

pymysql.install_as_MySQLdb()

app = Flask(__name__)
app.config.from_object(Config)
db.init_app(app)

# ── FOLDER PATHS ──
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SURAT_FOLDER = os.path.join(BASE_DIR, 'documents', 'surat')
FOTO_FOLDER = os.path.join(BASE_DIR, 'static', 'uploads', 'foto')
LOGO_PATH = os.path.join(BASE_DIR, 'static', 'images', 'logo_pelindo.png')

os.makedirs(SURAT_FOLDER, exist_ok=True)
os.makedirs(FOTO_FOLDER, exist_ok=True)

# ── FILE VALIDATION ──
ALLOWED_EXTENSIONS = {'jpg', 'jpeg', 'png'}
MAX_FILE_SIZE = 5 * 1024 * 1024

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# ═══════════════════════════════════════════════════════════════════════════════
# SIGN UPN
# ═══════════════════════════════════════════════════════════════════════════════
@app.route('/signup', methods=['GET', 'POST'])
def signup():
    """Sign up untuk Bea Cukai dan Gate Superintendent"""
    
    if request.method == 'POST':
        try:
            nama = request.form.get('nama', '').strip()
            email = request.form.get('email', '').strip()
            nip = request.form.get('nip', '').strip()
            password = request.form.get('password', '').strip()
            confirm_password = request.form.get('confirm_password', '').strip()
            role = request.form.get('role', '').strip()
            
            # ── VALIDASI ──
            if not all([nama, email, nip, password, confirm_password, role]):
                flash('❌ Semua field harus diisi!', 'error')
                return redirect(url_for('signup'))
            
            if len(password) < 6:
                flash('❌ Password minimal 6 karakter!', 'error')
                return redirect(url_for('signup'))
            
            if password != confirm_password:
                flash('❌ Password tidak cocok!', 'error')
                return redirect(url_for('signup'))
            
            if role not in ['Bea Cukai', 'Gate Superintendent']:
                flash('❌ Role tidak valid!', 'error')
                return redirect(url_for('signup'))
            
            # ── CEK EMAIL SUDAH TERDAFTAR ──
            existing_user = Pengguna.query.filter_by(email=email).first()
            if existing_user:
                flash('❌ Email sudah terdaftar!', 'error')
                return redirect(url_for('signup'))
            
            # ── CREATE USER BARU ──
            new_user = Pengguna(
                nama=nama,
                email=email,
                nip=nip,
                password=password,
                role=role
            )
            
            db.session.add(new_user)
            db.session.commit()
            
            flash('✅ Akun berhasil dibuat! Silakan login.', 'success')
            return redirect(url_for('login'))
            
        except Exception as e:
            db.session.rollback()
            app.logger.error(f"Error signup: {str(e)}")
            flash(f'❌ Terjadi kesalahan: {str(e)}', 'error')
            return redirect(url_for('signup'))
    
    return render_template('signup.html')


# ═══════════════════════════════════════════════════════════════════════════════
# FUNCTION: GENERATE SURAT PERMOHONAN PDF - KOP OTOMATIS TIAP HALAMAN
# ═══════════════════════════════════════════════════════════════════════════════

def _terbilang_foto(n: int) -> str:
    kata = ['nol','satu','dua','tiga','empat','lima','enam','tujuh','delapan','sembilan','sepuluh']
    return kata[n] if n <= 10 else str(n)


def generate_surat_dengan_ttd_gate_pdf(data: dict, foto_files: list = None, nama_gate: str = None, nip_gate: str = None, tgl_ttd: str = None) -> str:
    """
    Generate surat permohonan PDF dari gate.
    - Kop surat otomatis tiap halaman
    - TTD Gate: nama + NIP dari session
    """
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import cm
    from reportlab.lib.enums import TA_CENTER, TA_RIGHT, TA_JUSTIFY
    from datetime import datetime
    from werkzeug.utils import secure_filename
    import os

    timestamp  = datetime.now().strftime('%Y%m%d%H%M%S')
    filename   = secure_filename(f"surat_permohonan_{data['nomor_kontainer']}_{timestamp}.pdf")
    filepath   = os.path.join(SURAT_FOLDER, filename)

    PAGE_W, PAGE_H = A4
    LEFT_M  = 2.5 * cm
    RIGHT_M = 2.5 * cm
    TOP_M   = 1.5 * cm
    BOT_M   = 2.0 * cm
    KOP_H   = 3.2 * cm
    page_w  = PAGE_W - LEFT_M - RIGHT_M

    base_dir     = os.path.dirname(os.path.abspath(__file__))
    logo_path    = os.path.join(base_dir, 'static', 'logo',   'logo_sistem_2.png')
    barcode_path = os.path.join(base_dir, 'static', 'images', 'barcode_ttd.png')

    styles = getSampleStyleSheet()
    def S(name, **kw):
        return ParagraphStyle(name, parent=styles['Normal'], **kw)

    st_center_bold = S('CB',  fontSize=12, fontName='Helvetica-Bold', alignment=TA_CENTER, spaceAfter=2)
    st_center      = S('C',   fontSize=10, fontName='Helvetica',      alignment=TA_CENTER, spaceAfter=2)
    st_body        = S('B',   fontSize=10, fontName='Helvetica',      alignment=TA_JUSTIFY, spaceAfter=4, leading=16)
    st_body_indent = S('BI',  fontSize=10, fontName='Helvetica',      alignment=TA_JUSTIFY, firstLineIndent=1*cm, spaceAfter=4, leading=16)
    st_bold        = S('Bo',  fontSize=10, fontName='Helvetica-Bold',  spaceAfter=4)
    st_small       = S('Sm',  fontSize=9,  fontName='Helvetica',      spaceAfter=2)
    st_ttd_c       = S('TC',  fontSize=10, fontName='Helvetica',      alignment=TA_CENTER, spaceAfter=2)
    st_ttd_cb      = S('TCB', fontSize=10, fontName='Helvetica-Bold', alignment=TA_CENTER, spaceAfter=2)
    st_ttd_right   = S('TR',  fontSize=10, fontName='Helvetica',      alignment=TA_RIGHT,  spaceAfter=2)
    st_ttd_small   = S('TS',  fontSize=9,  fontName='Helvetica',      alignment=TA_CENTER, spaceAfter=2)

    bulan      = datetime.now().strftime('%m')
    tahun      = datetime.now().strftime('%Y')
    nomor_urut = Permohonan.query.count() + 1
    no_surat   = f"BC/{nomor_urut:03d}/TPK/{bulan}/{tahun}"
    tgl_format = datetime.strptime(data['tanggal_pengajuan'], '%Y-%m-%d').strftime('%d %B %Y')
    kota_tgl   = f"Surabaya, {tgl_format}"

    def draw_kop(canv, doc):
        canv.saveState()
        y_top = PAGE_H - TOP_M
        if os.path.exists(logo_path):
            canv.drawImage(logo_path, LEFT_M, y_top - 2.6*cm,
                           width=2.5*cm, height=2.5*cm,
                           preserveAspectRatio=True, mask='auto')
        center_x = PAGE_W / 2
        canv.setFont('Helvetica-Bold', 12)
        canv.drawCentredString(center_x, y_top - 0.7*cm, "PT. TERMINAL PETI KEMAS NUSANTARA")
        canv.setFont('Helvetica', 10)
        canv.drawCentredString(center_x, y_top - 1.3*cm, "Terminal Kontainer dan Kepelabuhanan")
        canv.drawCentredString(center_x, y_top - 1.8*cm, "Jl. Dermaga Raya No. 88, Jakarta Utara, Indonesia")
        canv.drawCentredString(center_x, y_top - 2.3*cm, "Telp. (031) 3298631  |  mail: jayalogistik@gmail.com")
        garis_y = y_top - KOP_H + 0.2*cm
        canv.setStrokeColor(colors.HexColor('#003366'))
        canv.setLineWidth(2)
        canv.line(LEFT_M, garis_y, PAGE_W - RIGHT_M, garis_y)
        canv.restoreState()

    doc = SimpleDocTemplate(
        filepath, pagesize=A4,
        topMargin=TOP_M + KOP_H, bottomMargin=BOT_M,
        leftMargin=LEFT_M, rightMargin=RIGHT_M,
    )
    elements = []

    elements.append(Spacer(1, 0.2*cm))
    elements.append(Paragraph("<b>SURAT PERMOHONAN</b>", st_center_bold))
    elements.append(Spacer(1, 0.4*cm))

    meta_table = Table([
        ['Nomor',   f': {no_surat}'],
        ['Tanggal', f': {tgl_format}'],
        ['Perihal', ': PENGELUARAN KONTAINER EXCEPTION AREA'],
    ], colWidths=[3.5*cm, page_w - 3.5*cm])
    meta_table.setStyle(TableStyle([
        ('FONTNAME',      (0,0), (-1,-1), 'Helvetica'),
        ('FONTSIZE',      (0,0), (-1,-1), 10),
        ('FONTNAME',      (0,0), (0,-1),  'Helvetica-Bold'),
        ('VALIGN',        (0,0), (-1,-1), 'TOP'),
        ('TOPPADDING',    (0,0), (-1,-1), 3),
        ('BOTTOMPADDING', (0,0), (-1,-1), 3),
        ('LEFTPADDING',   (0,0), (-1,-1), 0),
    ]))
    elements.append(meta_table)
    elements.append(Spacer(1, 0.5*cm))

    elements.append(Paragraph("Kepada Yth.", st_body))
    elements.append(Paragraph("Kepala Bea Cukai", st_body))
    elements.append(Paragraph("di Tempat", st_body))
    elements.append(Spacer(1, 0.3*cm))

    elements.append(Paragraph("Dengan hormat,", st_body))
    elements.append(Paragraph(
        "Bersama surat ini kami mengajukan permohonan pengeluaran kontainer pada area "
        "exception untuk diproses lebih lanjut sesuai ketentuan yang berlaku. Adapun data "
        "permohonan yang diajukan adalah sebagai berikut:",
        st_body_indent
    ))
    elements.append(Spacer(1, 0.3*cm))

    data_table = Table([
        ['Nama Perusahaan',  data['nama_perusahaan']],
        ['No. Job Order',    data.get('no_job_order') or '-'],
        ['No. Booking',      data.get('no_booking') or '-'],
        ['No. Container',    data['nomor_kontainer']],
        ['No. Polisi',       data['nopol_truck']],
        ['Jenis Permohonan', data['jenis_permohonan'].replace('_',' ').title()],
    ], colWidths=[4*cm, page_w - 4*cm])
    data_table.setStyle(TableStyle([
        ('FONTNAME',      (0,0), (-1,-1), 'Helvetica'),
        ('FONTSIZE',      (0,0), (-1,-1), 10),
        ('GRID',          (0,0), (-1,-1), 0.5, colors.black),
        ('ALIGN',         (0,0), (-1,-1), 'LEFT'),
        ('VALIGN',        (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING',    (0,0), (-1,-1), 6),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
        ('LEFTPADDING',   (0,0), (-1,-1), 8),
        ('RIGHTPADDING',  (0,0), (-1,-1), 8),
        ('ROWBACKGROUNDS',(0,0), (-1,-1), [colors.white, colors.HexColor('#f5f7fa')]),
    ]))
    elements.append(data_table)
    elements.append(Spacer(1, 0.4*cm))

    elements.append(Paragraph("<b>Kendala :</b>", st_bold))
    elements.append(Paragraph(data.get('kendala', '-'), st_body))
    elements.append(Spacer(1, 0.2*cm))

    elements.append(Paragraph("<b>Catatan Khusus :</b>", st_bold))
    elements.append(Paragraph(data.get('remark') or '-', st_body))
    elements.append(Spacer(1, 0.2*cm))

    if foto_files:
        elements.append(Paragraph("<b>Dokumentasi Pendukung :</b>", st_bold))
        elements.append(Paragraph(
            f"Terlampir {len(foto_files)} ({_terbilang_foto(len(foto_files))}) foto kontainer sebagai dokumen pendukung permohonan.",
            st_body
        ))
        elements.append(Spacer(1, 0.2*cm))
        for idx, foto_filename in enumerate(foto_files, 1):
            foto_path = os.path.join(FOTO_FOLDER, foto_filename.strip())
            if os.path.exists(foto_path):
                try:
                    elements.append(Image(foto_path, width=10*cm, height=7*cm))
                    elements.append(Spacer(1, 0.3*cm))
                except Exception:
                    elements.append(Paragraph(f"[Foto {idx} tidak bisa ditampilkan]", st_small))

    if data.get('keterangan_tambahan'):
        elements.append(Paragraph("<b>Keterangan Tambahan :</b>", st_bold))
        elements.append(Paragraph(data['keterangan_tambahan'], st_body))
        elements.append(Spacer(1, 0.2*cm))

    elements.append(Spacer(1, 0.3*cm))
    elements.append(Paragraph(
        "Demikian permohonan ini disampaikan untuk dapat dipertimbangkan dan diproses "
        "sebagaimana mestinya. Atas perhatian dan kerja sama yang diberikan, kami ucapkan terima kasih.",
        st_body_indent
    ))
    elements.append(Spacer(1, 0.5*cm))

    # ══════════════════════════════════════
    # TTD — 2 KOLOM: BC (kiri) & Gate (kanan)
    # ══════════════════════════════════════
    col_w = page_w / 2

    elements.append(Paragraph(kota_tgl, st_ttd_right))
    elements.append(Spacer(1, 0.3*cm))

    gate_ttd_cell = Image(barcode_path, width=3*cm, height=3*cm) if os.path.exists(barcode_path) else Paragraph("", st_ttd_c)

    # Nama + NIP Gate di baris terakhir TTD
    nama_gate_display = nama_gate or 'Gate Superintendent'
    nip_gate_display  = nip_gate or '-'

    ttd_table = Table([
        [Paragraph("Disetujui oleh,",      st_ttd_c),  Paragraph("Diajukan oleh,",        st_ttd_c)],
        [Paragraph("Bea Cukai",            st_ttd_cb), Paragraph("Gate Superintendent",   st_ttd_cb)],
        [Paragraph("",                     st_ttd_c),  gate_ttd_cell],
        [
            Paragraph(f"", st_ttd_c),
            Paragraph(f"( {nama_gate_display} )", st_ttd_c)
        ],
        [
            Paragraph("", st_ttd_small),
            Paragraph(f"NIP: {nip_gate_display}", st_ttd_small)
        ],
    ], colWidths=[col_w, col_w])

    ttd_table.setStyle(TableStyle([
        ('ALIGN',        (0,0), (-1,-1), 'CENTER'),
        ('VALIGN',       (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING',   (0,0), (-1,-1), 4),
        ('BOTTOMPADDING',(0,0), (-1,-1), 4),
        ('LEFTPADDING',  (0,0), (-1,-1), 0),
        ('RIGHTPADDING', (0,0), (-1,-1), 0),
        ('ROWHEIGHT',    (0,2), (-1,2),  60),
    ]))
    elements.append(ttd_table)

    doc.build(elements, onFirstPage=draw_kop, onLaterPages=draw_kop)
    return filename


# ═══════════════════════════════════════════════════════════════════════════════
# FUNCTION: UPDATE SURAT - OVERLAY STEMPEL BC KE HALAMAN TERAKHIR
# ═══════════════════════════════════════════════════════════════════════════════

def update_surat_dengan_ttd_bc_pdf(file_surat, nama_bc: str, tanggal_bc: str, nip_bc: str = None):
    """
    Update surat - overlay stempel + nama + NIP BC ke halaman terakhir
    """
    from PyPDF2 import PdfReader, PdfWriter
    from reportlab.pdfgen import canvas
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import cm
    from io import BytesIO
    import os

    try:
        file_path = os.path.join(SURAT_FOLDER, file_surat)
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"File surat tidak ditemukan: {file_path}")

        base_dir     = os.path.dirname(os.path.abspath(__file__))
        stempel_path = os.path.join(base_dir, 'static', 'images', 'stempel_bc.jpg')

        pdf_reader     = PdfReader(file_path)
        pdf_writer     = PdfWriter()
        last_page_idx  = len(pdf_reader.pages) - 1
        last_page      = pdf_reader.pages[last_page_idx]

        overlay_buffer = BytesIO()
        overlay_canvas = canvas.Canvas(overlay_buffer, pagesize=A4)

        # ── STEMPEL BC ──
        if os.path.exists(stempel_path):
            try:
                overlay_canvas.drawImage(stempel_path, 140, 280, width=3*cm, height=3*cm)
            except Exception as e:
                app.logger.warning(f"Stempel image error: {str(e)}")

        # ── NAMA BC ──
        overlay_canvas.setFont("Helvetica-Bold", 9)
        overlay_canvas.drawCentredString(180, 275, f"✓ {nama_bc}")

        # ── NIP BC ──
        nip_display = nip_bc or '-'
        overlay_canvas.setFont("Helvetica", 9)
        overlay_canvas.drawCentredString(180, 248, f"NIP: {nip_display}")

        overlay_canvas.save()
        overlay_buffer.seek(0)

        overlay_pdf  = PdfReader(overlay_buffer)
        overlay_page = overlay_pdf.pages[0]
        last_page.merge_page(overlay_page)

        for page_num in range(len(pdf_reader.pages)):
            if page_num == last_page_idx:
                pdf_writer.add_page(last_page)
            else:
                pdf_writer.add_page(pdf_reader.pages[page_num])

        with open(file_path, 'wb') as output_file:
            pdf_writer.write(output_file)

        app.logger.info(f"Surat {file_surat} diupdate dengan stempel BC {nama_bc} NIP {nip_display}")
        return True

    except Exception as e:
        app.logger.error(f"Error update_surat_dengan_ttd_bc_pdf: {str(e)}")
        raise Exception(f"Gagal update surat: {str(e)}")
 
# ═══════════════════════════════════════════════════════════════════════════════
# ROUTES: GATE - MEMBUAT PERMOHONAN
# ═══════════════════════════════════════════════════════════════════════════════

def save_uploaded_photos(files, nomor_kontainer) -> list:
    """Save uploaded photos dengan naming: NOKO-[timestamp]-[nomor].jpg"""
    saved_files = []
    timestamp = datetime.now().strftime('%Y%m%d%H%M%S')
    
    # Gunakan global FOTO_FOLDER (sudah didefinisikan di atas)
    os.makedirs(FOTO_FOLDER, exist_ok=True)
    
    for idx, file in enumerate(files, 1):
        if file and file.filename:
            ext = file.filename.rsplit('.', 1)[1].lower() if '.' in file.filename else ''
            if ext in ['jpg', 'jpeg', 'png']:
                filename = secure_filename(f"{nomor_kontainer}-{timestamp}-{idx}.{ext}")
                filepath = os.path.join(FOTO_FOLDER, filename)
                file.save(filepath)
                saved_files.append(filename)
    
    return saved_files
 
@app.route('/permohonan', methods=['GET', 'POST'])
def permohonan():
    """
    Gate membuat permohonan → Generate surat dengan TTD Gate
    Status langsung 'Menunggu' untuk BC approval
    """
    if 'username' not in session or session.get('role') != 'Gate Superintendent':
        flash('Akses ditolak. Hanya Gate Superintendent yang bisa membuat permohonan.', 'error')
        return redirect(url_for('login'))
    
    if request.method == 'POST':
        data = {
            'nama_perusahaan':     request.form.get('nama_perusahaan', '').strip(),
            'no_job_order':        request.form.get('no_job_order', '').strip(),
            'nomor_kontainer':     request.form.get('nomor_kontainer', '').strip(),
            'no_booking':          request.form.get('no_booking', '').strip(),
            'nopol_truck':         request.form.get('nopol_truck', '').strip(),
            'tanggal_pengajuan':   request.form.get('tanggal_pengajuan', '').strip(),
            'jenis_permohonan':    request.form.get('jenis_permohonan', '').strip(),
            'kendala':             request.form.get('kendala', '').strip(),
            'remark':              request.form.get('remark', '').strip(),
            'keterangan_tambahan': request.form.get('keterangan_tambahan', '').strip(),
        }
 
        # Validasi field wajib
        required = [
            'nama_perusahaan', 'nomor_kontainer',
            'nopol_truck', 'tanggal_pengajuan', 'jenis_permohonan', 
            'kendala', 'remark'
        ]
        if any(not data[k] for k in required):
            flash('⚠️ Harap lengkapi semua field yang wajib diisi.', 'error')
            return render_template('permohonan.html')
 
        try:
            # ── PROSES UPLOAD FOTO ──
            foto_files = request.files.getlist('foto_kontainer')
            saved_foto_files = []
            
            if foto_files and foto_files[0].filename != '':
                saved_foto_files = save_uploaded_photos(foto_files, data['nomor_kontainer'])
                
                if not saved_foto_files:
                    flash('⚠️ Gagal menyimpan foto. Pastikan format file adalah JPG/PNG.', 'error')
                    return render_template('permohonan.html')
 
            # ── GENERATE SURAT DENGAN TTD GATE ──
            nama_gate = session.get('username')
            tgl_ttd = datetime.now().strftime('%d %B %Y')
            
            filename_surat = generate_surat_dengan_ttd_gate_pdf(
                data=data,
                foto_files=saved_foto_files,
                nama_gate=nama_gate,
                nip_gate=session.get('nip'),
                tgl_ttd=tgl_ttd
            )
 
            # ── SIMPAN KE DATABASE ──
            new_pengajuan = Permohonan(
                id_pengguna         = session.get('user_id'),
                nama_perusahaan     = data['nama_perusahaan'],
                no_job_order        = data['no_job_order'],
                nomor_kontainer     = data['nomor_kontainer'],
                no_booking          = data['no_booking'],
                nopol_truck         = data['nopol_truck'],
                tanggal_pengajuan   = data['tanggal_pengajuan'],
                jenis_permohonan    = data['jenis_permohonan'],
                kendala             = data['kendala'],
                remark              = data['remark'],
                keterangan_tambahan = data['keterangan_tambahan'],
                file_surat          = filename_surat,
                foto_dokumentasi    = ','.join(saved_foto_files) if saved_foto_files else None,
                status              = 'Menunggu',
            )
            db.session.add(new_pengajuan)
            db.session.commit()
 
            flash(
                f'✅ Pengajuan berhasil dikirim! '
                f'Surat permohonan telah digenerate dengan TTD {nama_gate}. '
                f'Permohonan sedang menunggu persetujuan dari Bea Cukai.',
                'success'
            )
            return redirect(url_for('monitoring_gate'))
 
        except Exception as e:
            db.session.rollback()
            flash(f'❌ Terjadi kesalahan: {str(e)}', 'error')
            app.logger.error(f"Error di /permohonan: {str(e)}")
            return render_template('permohonan.html')
 
    return render_template('permohonan.html')
 
 
@app.route('/dashboard_gate')
def dashboard_gate():
    """Dashboard Gate Superintendent"""
    if 'username' not in session or session.get('role') != 'Gate Superintendent':
        return redirect(url_for('login'))
    
    user_id = session.get('user_id')
    
    total_pengajuan = Permohonan.query.filter_by(id_pengguna=user_id).count()
    menunggu_pengajuan = Permohonan.query.filter_by(id_pengguna=user_id, status='Menunggu').count()
    disetujui_pengajuan = Permohonan.query.filter_by(id_pengguna=user_id, status='Disetujui').count()
    
    approved_pengajuan = Permohonan.query.filter_by(id_pengguna=user_id, status='Disetujui').order_by(
        Permohonan.created_at.desc()
    ).limit(3).all()
    
    return render_template(
        'dashboard_gate.html',
        total_pengajuan=total_pengajuan,
        menunggu_pengajuan=menunggu_pengajuan,
        disetujui_pengajuan=disetujui_pengajuan,
        approved_pengajuan=approved_pengajuan
    )
 
 
@app.route('/monitoring_gate')
def monitoring_gate():
    """Gate lihat permohonan yang mereka buat sendiri"""
    if 'username' not in session or session.get('role') != 'Gate Superintendent':
        return redirect(url_for('login'))
 
    user_id = session.get('user_id')
    pengajuan_kontainer = Permohonan.query.filter_by(id_pengguna=user_id).all()
 
    total = len(pengajuan_kontainer)
    menunggu = len([p for p in pengajuan_kontainer if p.status == 'Menunggu'])
    disetujui = len([p for p in pengajuan_kontainer if p.status == 'Disetujui'])
 
    return render_template(
        'monitoring_gate.html',
        pengajuan_kontainer=pengajuan_kontainer,
        total_pengajuan=total,
        menunggu_pengajuan=menunggu,
        disetujui_pengajuan=disetujui
    )
 
 
# ═══════════════════════════════════════════════════════════════════════════════
# ROUTES: BEAU CUKAI - MONITORING & APPROVAL
# ═══════════════════════════════════════════════════════════════════════════════
 
@app.route('/dashboard_beacukai')
def dashboard_beacukai():
    """Dashboard Bea Cukai - Overview statistik permohonan"""
    if 'username' not in session or session.get('role') != 'Bea Cukai':
        return redirect(url_for('login'))
    
    total_pengajuan = Permohonan.query.count()
    menunggu_pengajuan = Permohonan.query.filter_by(status='Menunggu').count()
    disetujui_pengajuan = Permohonan.query.filter_by(status='Disetujui').count()
    
    approved_pengajuan = Permohonan.query.filter_by(status='Disetujui').order_by(
        Permohonan.created_at.desc()
    ).limit(3).all()
    
    return render_template(
        'dashboard_beacukai.html',
        total_pengajuan=total_pengajuan,
        menunggu_pengajuan=menunggu_pengajuan,
        disetujui_pengajuan=disetujui_pengajuan,
        approved_pengajuan=approved_pengajuan
    )
 
@app.route('/monitoring_beacukai', methods=['GET', 'POST'])
def monitoring_beacukai():
    """BC lihat semua permohonan untuk review & approval"""
    if 'username' not in session or session.get('role') != 'Bea Cukai':
        return redirect(url_for('login'))
 
    status_filter = request.args.get('status', 'all')
 
    if status_filter == 'all':
        pengajuan_kontainer = Permohonan.query.order_by(Permohonan.created_at.desc()).all()
    else:
        pengajuan_kontainer = Permohonan.query.filter_by(status=status_filter).order_by(
            Permohonan.created_at.desc()
        ).all()
 
    total = Permohonan.query.count()
    menunggu = Permohonan.query.filter_by(status='Menunggu').count()
    disetujui = Permohonan.query.filter_by(status='Disetujui').count()
    ditolak = Permohonan.query.filter_by(status='Ditolak').count()
 
    return render_template(
        'monitoring_beacukai.html',
        pengajuan_kontainer=pengajuan_kontainer,
        total_pengajuan=total,
        menunggu_pengajuan=menunggu,
        disetujui_pengajuan=disetujui,
        ditolak_pengajuan=ditolak,
        status_filter=status_filter
    )
 
 
@app.route('/api/permohonan/<int:permohonan_id>/approve', methods=['POST'])
def approve_permohonan(permohonan_id):
    """BC approve permohonan → Update surat dengan TTD BC"""
    if 'username' not in session or session.get('role') != 'Bea Cukai':
        return jsonify({'success': False, 'error': 'Unauthorized'}), 403
 
    try:
        permohonan = Permohonan.query.get(permohonan_id)
        if not permohonan:
            return jsonify({'success': False, 'error': 'Permohonan tidak ditemukan'}), 404
 
        if permohonan.status == 'Disetujui':
            return jsonify({'success': False, 'error': 'Permohonan sudah disetujui sebelumnya'}), 400
 
        # ── UPDATE SURAT DENGAN TTD BC ──
        nama_bc = session.get('username')
        tanggal_bc = datetime.now().strftime('%d %B %Y')
        
        # Update file surat yang ada + tambah halaman TTD BC
        update_surat_dengan_ttd_bc_pdf(
            file_surat=permohonan.file_surat,
            nama_bc=nama_bc,
            nip_bc=session.get('nip'),
            tanggal_bc=tanggal_bc
        )
 
        # ── UPDATE STATUS ──
        permohonan.status = 'Disetujui'
        permohonan.catatan = f"Disetujui oleh {nama_bc} pada {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
        
        persetujuan = Persetujuan(
            id_permohonan=permohonan_id,
            status_persetujuan='Disetujui',
            komentar=f"Surat disetujui oleh {nama_bc}"
        )
        
        db.session.add(persetujuan)
        db.session.commit()
 
        return jsonify({
            'success': True,
            'message': '✅ Permohonan disetujui! TTD BC sudah ditambahkan ke surat.'
        })
 
    except Exception as e:
        db.session.rollback()
        app.logger.error(f"Error approve_permohonan: {str(e)}")
        return jsonify({'success': False, 'error': str(e)}), 500
 
 
@app.route('/api/permohonan/<int:permohonan_id>/reject', methods=['POST'])
def reject_permohonan(permohonan_id):
    """BC reject permohonan"""
    if 'username' not in session or session.get('role') != 'Bea Cukai':
        return jsonify({'success': False, 'error': 'Unauthorized'}), 403
 
    try:
        data = request.get_json()
        alasan = data.get('alasan', 'Tidak ada alasan yang diberikan').strip()
        
        if not alasan:
            return jsonify({'success': False, 'error': 'Alasan penolakan harus diisi'}), 400
 
        permohonan = Permohonan.query.get(permohonan_id)
        if not permohonan:
            return jsonify({'success': False, 'error': 'Permohonan tidak ditemukan'}), 404
 
        if permohonan.status == 'Ditolak':
            return jsonify({'success': False, 'error': 'Permohonan sudah ditolak sebelumnya'}), 400
 
        # ── UPDATE STATUS ──
        permohonan.status = 'Ditolak'
        permohonan.catatan = alasan
 
        persetujuan = Persetujuan(
            id_permohonan=permohonan_id,
            status_persetujuan='Ditolak',
            komentar=alasan
        )
 
        db.session.add(persetujuan)
        db.session.commit()
 
        return jsonify({
            'success': True,
            'message': '❌ Permohonan ditolak.'
        })
 
    except Exception as e:
        db.session.rollback()
        app.logger.error(f"Error reject_permohonan: {str(e)}")
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/permohonan/<int:permohonan_id>', methods=['GET'])
def get_permohonan_detail(permohonan_id):
    """Get detail permohonan as JSON - untuk modal detail di monitoring"""
    if 'username' not in session:
        return jsonify({'success': False, 'error': 'Unauthorized'}), 403
    
    try:
        permohonan = Permohonan.query.get(permohonan_id)
        if not permohonan:
            return jsonify({'success': False, 'error': 'Permohonan tidak ditemukan'}), 404
        
        return jsonify({
            'id': permohonan.id,
            'nama_perusahaan': permohonan.nama_perusahaan,
            'no_job_order': permohonan.no_job_order or '-',
            'nomor_kontainer': permohonan.nomor_kontainer,
            'no_booking': permohonan.no_booking or '-',
            'nopol_truck': permohonan.nopol_truck,
            'jenis_permohonan': permohonan.jenis_permohonan,
            'status': permohonan.status,
            'file_surat': permohonan.file_surat
        })
    except Exception as e:
        app.logger.error(f"Error get_permohonan_detail: {str(e)}")
        return jsonify({'success': False, 'error': str(e)}), 500
 
 
# ═══════════════════════════════════════════════════════════════════════════════
# ROUTES: VIEW/DOWNLOAD SURAT - HANYA GATE & BEAU CUKAI
# ═══════════════════════════════════════════════════════════════════════════════
 
@app.route('/lihat_surat/<filename>')
def lihat_surat(filename):
    """
    Display file surat PDF di browser (tidak download)
    Accessible: Gate & Bea Cukai only
    """
    # ── CHECK LOGIN & ROLE ──
    if 'username' not in session:
        return redirect(url_for('login'))
    
    role = session.get('role')
    if role not in ['Gate Superintendent', 'Bea Cukai']:
        flash('❌ Akses ditolak - hanya Gate & Bea Cukai yang bisa lihat surat', 'error')
        return redirect(url_for('login'))
    
    try:
        # ── SECURE FILENAME ──
        secure_name = secure_filename(filename)
        if not secure_name:
            flash('❌ Nama file tidak valid', 'error')
            return redirect(request.referrer or url_for('dashboard_beacukai'))
        
        # ── BUILD FILE PATH ──
        file_path = os.path.join(SURAT_FOLDER, secure_name)
        
        # ── VALIDATE PATH (PREVENT PATH TRAVERSAL) ──
        real_path = os.path.realpath(file_path)
        real_folder = os.path.realpath(SURAT_FOLDER)
        
        if not real_path.startswith(real_folder):
            flash('❌ Akses ditolak - path tidak valid', 'error')
            return redirect(request.referrer or url_for('dashboard_beacukai'))
        
        # ── CHECK FILE EXISTS ──
        if not os.path.isfile(file_path):
            app.logger.warning(f"File not found: {file_path}")
            flash('❌ File tidak ditemukan', 'error')
            return redirect(request.referrer or url_for('dashboard_beacukai'))
        
        # ── SEND FILE AS PDF (DISPLAY, TIDAK DOWNLOAD) ──
        app.logger.info(f"User {session.get('username')} ({role}) lihat: {secure_name}")
        return send_file(file_path, mimetype='application/pdf')
        
    except Exception as e:
        app.logger.error(f"Error lihat_surat: {str(e)}")
        flash(f'❌ Gagal membuka file: {str(e)}', 'error')
        return redirect(request.referrer or url_for('dashboard_beacukai'))


# ═══════════════════════════════════════════════════════════════════════════════
# AUTH & DASHBOARD ROUTES
# ═══════════════════════════════════════════════════════════════════════════════

@app.route('/', methods=['GET', 'POST'])
def login():
    if 'username' in session:
        if session['role'] == 'CS':
            return redirect(url_for('dashboard_cs'))
        elif session['role'] == 'Bea Cukai':
            return redirect(url_for('monitoring_beacukai'))
        elif session['role'] == 'Gate Superintendent':
            return redirect(url_for('dashboard_gate'))
        elif session['role'] == 'EMKL':
            return redirect(url_for('dashboard_emkl'))
        elif session['role'] == 'Exception Area':
            return redirect(url_for('dashboard_ea'))

    if request.method == 'POST':
        role = request.form['role']
        email = request.form['email']
        password = request.form['password']

        user = Pengguna.query.filter_by(email=email).first()
        if user and user.password == password:
            session['role'] = user.role
            session['username'] = user.nama
            session['user_id'] = user.id
            session['email'] = user.email
            session['nip'] = user.nip
            
            if user.role == 'CS':
                return redirect(url_for('dashboard_cs'))
            elif user.role == 'EMKL':
                return redirect(url_for('dashboard_emkl'))
            elif user.role == 'Bea Cukai':
                return redirect(url_for('monitoring_beacukai'))
            elif user.role == 'Gate Superintendent':
                return redirect(url_for('dashboard_gate'))
            elif user.role == 'Exception Area':
                return redirect(url_for('dashboard_ea'))
        else:
            flash('Email or password incorrect', 'error')
            return redirect(url_for('login'))

    return render_template('login_cs.html')

@app.route('/profile_beacukai')
def profile_beacukai():
    if 'username' not in session:
        return redirect(url_for('login'))
    if session.get('role') != 'Bea Cukai':
        return redirect(url_for('login'))
    return render_template('profile_beacukai.html')

@app.route('/profile_gate')
def profile_gate():
    if 'username' not in session:
        return redirect(url_for('login'))
    if session.get('role') != 'Gate Superintendent':
        return redirect(url_for('login'))
    return render_template('profile_gate.html')

@app.route('/delete_account', methods=['POST'])
def delete_account():
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    try:
        user = Pengguna.query.get(session['user_id'])
        if user:
            db.session.delete(user)
            db.session.commit()
            session.clear()
            flash('✅ Akun berhasil dihapus.', 'success')
    except Exception as e:
        db.session.rollback()
        flash(f'❌ Gagal menghapus akun: {str(e)}', 'error')
    
    return redirect(url_for('login'))

# Route untuk logout
@app.route('/logout')
def logout():
    session.clear()  # Menghapus session login
    return redirect(url_for('login'))  # Mengarahkan kembali ke halaman login

if __name__ == '__main__':
    app.run(debug=True)

# Taruh route ini di app.py, di antara profile_gate dan delete_account

@app.route('/edit_profile', methods=['POST'])
def edit_profile():
    if 'user_id' not in session:
        return redirect(url_for('login'))

    try:
        user = Pengguna.query.get(session['user_id'])
        if not user:
            flash('❌ User tidak ditemukan.', 'error')
            return redirect(url_for('profile_gate'))

        user.nama  = request.form.get('nama', user.nama).strip()
        user.email = request.form.get('email', user.email).strip()
        user.nip   = request.form.get('nip', user.nip).strip()

        password_baru = request.form.get('password_baru', '').strip()
        if password_baru:
            user.password = password_baru  # ganti sesuai hashing yang kamu pakai

        db.session.commit()

        # Update session supaya tampilan langsung berubah
        session['username'] = user.nama
        session['email']    = user.email
        session['nip']      = user.nip

        flash('✅ Profile berhasil diperbarui.', 'success')

    except Exception as e:
        db.session.rollback()
        flash(f'❌ Gagal memperbarui profile: {str(e)}', 'error')

    return redirect(url_for('profile_gate'))


#@app.route('/dashboard_emkl')
#def dashboard_emkl():
#    if 'username' not in session:
#        return redirect(url_for('login'))
#    return render_template('dashboard_emkl.html')

#@app.route('/dashboard_cs')
#def dashboard_cs():
#    if 'username' not in session:
#        return redirect(url_for('login'))
    
    # Get stats
#    total_pengajuan = Permohonan.query.count()
#    menunggu_pengajuan = Permohonan.query.filter_by(status='Menunggu').count()
#    disetujui_pengajuan = Permohonan.query.filter_by(status='Disetujui').count()
    
    # Get 3 approved pengajuan terbaru
#    approved_pengajuan = Permohonan.query.filter_by(status='Disetujui').order_by(Permohonan.created_at.desc()).limit(3).all()
    
#    return render_template('dashboard_cs.html', 
#                         total_pengajuan=total_pengajuan,
#                         menunggu_pengajuan=menunggu_pengajuan,
#                         disetujui_pengajuan=disetujui_pengajuan,
#                         approved_pengajuan=approved_pengajuan)

#@app.route('/dashboard_beacukai')
#def dashboard_beacukai():
#    if 'username' not in session:
#        return redirect(url_for('login'))
#    return render_template('dashboard_beacukai.html')
#@app.route('/dashboard_ea')
#def dashboard_ea():
#    if 'username' not in session:
#        return redirect(url_for('login'))
#    return render_template('dashboard_ea.html')
# ═══════════════════════════════════════════════════════════════════════════════
# CS ROUTES - Monitor semua permohonan + bisa lihat surat yang sudah disetujui
# ═══════════════════════════════════════════════════════════════════════════════

#@app.route('/monitoring_cs', methods=['GET', 'POST'])
#def monitoring_cs():
#    if 'username' not in session:
#        return redirect(url_for('login'))
#
    # CS bisa lihat semua permohonan (baik menunggu, disetujui, atau ditolak)
#    pengajuan_kontainer = Permohonan.query.all()

#    if request.method == 'POST':
#        permohonan_id = request.form.get('permohonan_id')
#        catatan = request.form.get('catatan')

#        permohonan = Permohonan.query.filter_by(id=permohonan_id).first()
#        if permohonan:
#            permohonan.catatan = catatan
#            db.session.commit()

#        return redirect(url_for('monitoring_cs'))

#    return render_template('monitoring_cs.html', pengajuan_kontainer=pengajuan_kontainer)


# ═══════════════════════════════════════════════════════════════════════════════
# EMKL ROUTES - COMMENTED (untuk nanti)
# ═══════════════════════════════════════════════════════════════════════════════

# @app.route('/monitoring_emkl', methods=['GET', 'POST'])
# def monitoring_emkl():
#     if 'username' not in session:
#         return redirect(url_for('login'))
#
#     # Filter hanya data milik user yang login
#     pengajuan_kontainer = Permohonan.query.filter_by(
#         id_pengguna=session['user_id']
#     ).all()
#
#     return render_template('monitoring_emkl.html', pengajuan_kontainer=pengajuan_kontainer)


# ═══════════════════════════════════════════════════════════════════════════════
# FILE HANDLING ROUTES
# ═══════════════════════════════════════════════════════════════════════════════

#@app.route('/lihat_surat/<filename>')
#def lihat_surat(filename):
#    if 'username' not in session:
#        return redirect(url_for('login'))
    
#    file_path = os.path.join(SURAT_FOLDER, filename)
#    if os.path.exists(file_path):
#        return send_from_directory(SURAT_FOLDER, filename)
#    else:
#        flash('File surat tidak ditemukan', 'error')
#        return redirect(url_for('monitoring_cs'))

# ═══════════════════════════════════════════════════════════════════════════════
# PROFILE & UTILITY ROUTES
# ═══════════════════════════════════════════════════════════════════════════════

#@app.route('/profile_cs')
#def profile_cs():
#    if 'username' not in session:
#        return redirect(url_for('login'))
#    return render_template('profile_cs.html')
#

# Route untuk halaman dashboard EMKL
#@app.route('/dashboard_emkl')
#def dashboard_emkl():
#    if 'username' not in session:  # Jika belum login, arahkan ke halaman login
#        return redirect(url_for('login'))
#    return render_template('dashboard_emkl.html')

#Route Monitoring Emkl
#@app.route('/monitoring_emkl', methods=['GET', 'POST'])
#def monitoring_emkl():
#    if 'username' not in session:
#        return redirect(url_for('login'))

    # Filter hanya data milik user yang login
#  pengajuan_kontainer = Permohonan.query.filter_by(
 #       id_pengguna=session['user_id']
#    ).all()

#    return render_template('monitoring_emkl.html', pengajuan_kontainer=pengajuan_kontainer)

# Route download surat (hanya bisa kalau status Disetujui)

# Route profile emkl
#@app.route('/profile_emkl')
#def profile_emkl():
#    if 'username' not in session:  # Cek apakah sudah login
#        return redirect(url_for('login'))  # Arahkan ke halaman login jika belum login

#    # Render halaman profil emkl
#    return render_template('profile_emkl.html')  # Ganti dengan nama file yang sesuai

# Route profile ea
#@app.route('/profile_ea')
#def profile_ea():
#    if 'username' not in session:  # Cek apakah sudah login
#        return redirect(url_for('login'))  # Arahkan ke halaman login jika belum login

    # Render halaman profil emkl
#    return render_template('profile_ea.html')  # Ganti dengan nama file yang sesuai

# Route untuk Pembuatan Kode
#@app.route('/pembuatan_kode', methods=['GET', 'POST'])
#def pembuatan_kode():
#    if 'username' not in session:  # Jika belum login, arahkan ke halaman login
#        return redirect(url_for('login'))
    
#    if request.method == 'POST':
#        try:
#            # Ambil data dari form
#            id_dokumen = request.form.get('id_dokumen')
#            jenis_dokumen = request.form.get('jenis_dokumen')
#            tgl_dokumen = request.form.get('tgl_dokumen')
            
            # Validasi data
#            if not all([id_dokumen, jenis_dokumen, tgl_dokumen]):
#                return render_template('pembuatan_kode.html', 
#                                      error='Semua field harus diisi!')
            
            # Generate kode unik
#            import uuid
#            kode_akses = str(uuid.uuid4())[:8].upper()
            
            # Simpan ke database (sesuaikan dengan struktur database Anda)
            # db.insert_kode(id_dokumen, kode_akses, jenis_dokumen, tgl_dokumen, session['username'])
            
#            return render_template('pembuatan_kode.html',
#                                  success=f'Kode berhasil dibuat: {kode_akses}',
#                                  kode_akses=kode_akses)
        
#        except Exception as e:
#           return render_template('pembuatan_kode.html',
#                                  error=f'Error: {str(e)}')
    
#    return render_template('pembuatan_kode.html')

# Route untuk Generate Kode (AJAX)
#@app.route('/api/generate_kode', methods=['POST'])
#def generate_kode():
#    if 'username' not in session:
#        return {'status': 'error', 'message': 'Unauthorized'}, 401
    
#    try:
#        data = request.json
#        id_dokumen = data.get('id_dokumen')
#        jenis_dokumen = data.get('jenis_dokumen')
#        
#        if not id_dokumen or not jenis_dokumen:
#            return {'status': 'error', 'message': 'Data tidak lengkap'}, 400
        
        # Generate kode
#        import uuid
#        kode_akses = str(uuid.uuid4())[:8].upper()
#        
#        return {
#            'status': 'success',
#            'kode_akses': kode_akses,
#            'waktu_dibuat': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
#        }
    
#    except Exception as e:
#        return {'status': 'error', 'message': str(e)}, 500