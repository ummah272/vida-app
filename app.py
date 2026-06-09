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
            password = request.form.get('password', '').strip()
            confirm_password = request.form.get('confirm_password', '').strip()
            role = request.form.get('role', '').strip()
            
            # ── VALIDASI ──
            if not all([nama, email, password, confirm_password, role]):
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
# FUNCTION: GENERATE SURAT DENGAN PLACEHOLDER TTD BC & GATE (1 HALAMAN)
# ═══════════════════════════════════════════════════════════════════════════════

def generate_surat_dengan_ttd_gate_pdf(data: dict, foto_files: list = None, nama_gate: str = None, tgl_ttd: str = None) -> str:
    """
    Generate surat permohonan PDF dari gate
    SUDAH INCLUDE placeholder TTD BC & TTD Gate (2 kolom, 1 halaman)
    
    Saat dibuat: Gate TTD terisi, BC placeholder kosong
    Saat BC approve: BC placeholder diisi otomatis
    """
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import cm
    from datetime import datetime
    from werkzeug.utils import secure_filename
    import os
    
    timestamp = datetime.now().strftime('%Y%m%d%H%M%S')
    filename = secure_filename(f"surat_permohonan_{data['nomor_kontainer']}_{timestamp}.pdf")
    filepath = os.path.join(SURAT_FOLDER, filename)
    
    # Create PDF
    doc = SimpleDocTemplate(filepath, pagesize=A4, topMargin=1*cm, bottomMargin=1*cm)
    elements = []
    styles = getSampleStyleSheet()
    
    # Custom styles
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Normal'],
        fontSize=14,
        textColor=colors.HexColor('#1a1a2e'),
        spaceAfter=6,
        alignment=1,  # CENTER
        fontName='Helvetica-Bold'
    )
    
    subtitle_style = ParagraphStyle(
        'CustomSubtitle',
        parent=styles['Normal'],
        fontSize=12,
        textColor=colors.HexColor('#1a1a2e'),
        spaceAfter=12,
        alignment=1,  # CENTER
        fontName='Helvetica-Bold'
    )
    
    body_style = ParagraphStyle(
        'CustomBody',
        parent=styles['Normal'],
        fontSize=10,
        textColor=colors.HexColor('#1a1a2e'),
        spaceAfter=6,
        alignment=4,  # JUSTIFY
    )
    
    info_style = ParagraphStyle(
        'Info',
        parent=styles['Normal'],
        fontSize=9,
        textColor=colors.HexColor('#1a1a2e'),
        spaceAfter=3,
    )
    
    # HEADER
    elements.append(Paragraph("PT. PELINDO III - TERMINAL PETIKEMAS", title_style))
    elements.append(Spacer(1, 0.3*cm))
    elements.append(Paragraph("_" * 80, body_style))
    elements.append(Spacer(1, 0.3*cm))
    
    elements.append(Paragraph("SURAT PERMOHONAN", title_style))
    elements.append(Spacer(1, 0.2*cm))
    elements.append(Paragraph("PENGAJUAN KONTAINER EMKL", subtitle_style))
    elements.append(Spacer(1, 0.3*cm))
    
    #nomor surat
    from datetime import datetime
    # Get bulan & tahun
    bulan = datetime.now().strftime('%m')  # 05 untuk Mei
    tahun = datetime.now().strftime('%Y')  # 2026

    # Get nomor urutan surat (total permohonan + 1)
    nomor_urut = Permohonan.query.count() + 1
    no_surat_urut = f"{nomor_urut:03d}"  # 001, 002, 003, dst

    # Format nomor surat
    no_surat = f"BC/{no_surat_urut}/TPK/{bulan}/{tahun}"

    bulan = datetime.now().strftime('%m')
    tahun = datetime.now().strftime('%Y')
    nomor_urut = Permohonan.query.count() + 1
    no_surat = f"BC/{nomor_urut:03d}/TPK/{bulan}/{tahun}"

    # Tanggal
    tgl_format = datetime.strptime(data['tanggal_pengajuan'], '%Y-%m-%d').strftime('%d %B %Y')
    
    elements.append(Paragraph(f"<b>Nomor Surat</b> : {no_surat}", info_style))
    elements.append(Paragraph(f"<b>Tanggal</b> : {tgl_format}", info_style))
    elements.append(Spacer(1, 0.3*cm))
    
    # Addressed to
    elements.append(Paragraph("Kepada Yth.", body_style))
    elements.append(Paragraph("Kepala Bea dan Cukai / Gate Pengeluaran Kontainer", body_style))
    elements.append(Paragraph("di Tempat", body_style))
    elements.append(Spacer(1, 0.2*cm))
    
    # Opening
    elements.append(Paragraph("Dengan hormat,", body_style))
    elements.append(Spacer(1, 0.1*cm))
    
    opening_text = f"""Sehubungan dengan permohonan yang telah diajukan oleh <b>{data['nama_perusahaan']}</b> dan telah ditandatangani oleh Gate Superintendent, 
    dengan ini kami sampaikan data permohonan berikut:"""
    elements.append(Paragraph(opening_text, body_style))
    elements.append(Spacer(1, 0.2*cm))
    
    # TABLE DATA
    table_data = [
        ['Nama Perusahaan', data['nama_perusahaan']],
        ['No Job Order', data['no_job_order']],
        ['Nomor Kontainer', data['nomor_kontainer']],
        ['No Booking', data['no_booking']],
        ['Nopol Truck', data['nopol_truck']],
        ['Jenis Permohonan', data['jenis_permohonan'].replace('_', ' ').title()],
        ['Tanggal Pengajuan', tgl_format],
    ]
    
    table = Table(table_data, colWidths=[3*cm, 13*cm])
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (0, -1), colors.HexColor("#4D647C")),
        ('TEXTCOLOR', (0, 0), (0, -1), colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 9),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
        ('TOPPADDING', (0, 0), (-1, -1), 8),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#f0f4f8')]),
    ]))
    elements.append(table)
    elements.append(Spacer(1, 0.3*cm))
    
    # Kendala & Remark
    elements.append(Paragraph("<b>Kendala:</b>", body_style))
    elements.append(Paragraph(data['kendala'], body_style))
    elements.append(Spacer(1, 0.2*cm))
    
    elements.append(Paragraph("<b>Remark & Catatan Khusus:</b>", body_style))
    elements.append(Paragraph(data['remark'] if data.get('remark') else "-", body_style))
    elements.append(Spacer(1, 0.3*cm))
    
    # DOKUMENTASI FOTO
    if foto_files:
        elements.append(Paragraph("<b>Dokumentasi Foto Kontainer di EA:</b>", body_style))
        elements.append(Paragraph(f"Jumlah file foto: {len(foto_files)}", info_style))
        elements.append(Spacer(1, 0.2*cm))
        
        for idx, foto_filename in enumerate(foto_files, 1):
            foto_path = os.path.join(FOTO_FOLDER, foto_filename.strip())
            if os.path.exists(foto_path):
                try:
                    img = Image(foto_path, width=10*cm, height=7*cm)
                    elements.append(Paragraph(f"<b>Foto {idx}:</b>", info_style))
                    elements.append(img)
                    elements.append(Spacer(1, 0.3*cm))
                except Exception as e:
                    elements.append(Paragraph(f"[Foto {idx} tidak bisa ditampilkan]", info_style))
    
    if data.get('keterangan_tambahan'):
        elements.append(Paragraph("<b>Keterangan Tambahan:</b>", body_style))
        elements.append(Paragraph(data['keterangan_tambahan'], body_style))
        elements.append(Spacer(1, 0.3*cm))
    
    elements.append(Spacer(1, 0.3*cm))
    
    # Closing
    elements.append(Paragraph("Demikian permohonan ini kami sampaikan. Atas perhatian dan kerja samanya, kami ucapkan terima kasih.", body_style))
    elements.append(Spacer(1, 0.5*cm))
    
    # ══════════════════════════════════════════════════════════════════
    # BAGIAN TTD - 2 KOLOM (BC & GATE) - DALAM 1 HALAMAN YANG SAMA
    # ══════════════════════════════════════════════════════════════════
    
    elements.append(Paragraph("_" * 80, info_style))
    elements.append(Spacer(1, 0.3*cm))
    
    # Signature table (2 kolom) - placeholder untuk BC & Gate
    barcode_img = RLImage(os.path.join(BASE_DIR, 'static', 'images', 'barcode_ttd.png'), 
                      width=3*cm, height=3*cm)
    stempel_img = RLImage(os.path.join(BASE_DIR, 'static', 'images', 'stempel_bc.jpg'), 
                      width=3*cm, height=3*cm)
    sig_data = [
    ['Pejabat Bea dan Cukai', 'Gate Superintendent'],
    ['', ''],
    ['', 'dengan ini menyetujui'],
    ['', barcode_img],  # ← BC kosong, Gate punya barcode
    [f'_______________', f'✓ Nama: {nama_gate if nama_gate else "_______________"}'],
    ['(____________________)', '(______________________)'],
    ]
    
    sig_table = Table(sig_data, colWidths=[7.5*cm, 7.5*cm])
    sig_table.setStyle(TableStyle([
        # Header - bold
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 10),
        
        # Semua cell align center
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        
        # Padding
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('LEFTPADDING', (0, 0), (-1, -1), 10),
        ('RIGHTPADDING', (0, 0), (-1, -1), 10),
        
        # No grid (clean look)
        ('GRID', (0, 0), (-1, -1), 0, colors.white),
        
        # Border hanya di bawah header
        ('LINEBELOW', (0, 0), (-1, 0), 1, colors.black),
    ]))
    
    elements.append(sig_table)
    elements.append(Spacer(1, 0.4*cm))
    
    # Tanggal TTD Gate
    if tgl_ttd:
        elements.append(Paragraph(f"Tanggal TTD Gate: <b>{tgl_ttd}</b>", info_style))
    
    # Build PDF
    doc.build(elements)
    
    return filename


def update_surat_dengan_ttd_bc_pdf(file_surat, nama_bc: str, tanggal_bc: str):
    """
    Update surat - OVERLAY stempel BC ke placeholder TTD BC (di halaman yang sama)
    
    Args:
        file_surat: nama file surat yang sudah ada
        nama_bc: nama Bea Cukai yang approve
        tanggal_bc: tanggal approval
    """
    from PyPDF2 import PdfReader, PdfWriter
    from reportlab.pdfgen import canvas
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import cm
    from io import BytesIO
    import os
    
    try:
        # ── PATH FILE SURAT ──
        file_path = os.path.join(SURAT_FOLDER, file_surat)
        
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"File surat tidak ditemukan: {file_path}")
        
        stempel_path = os.path.join(BASE_DIR, 'static', 'images', 'stempel_bc.jpg')
        
        # ── READ EXISTING PDF ──
        pdf_reader = PdfReader(file_path)
        pdf_writer = PdfWriter()
        
        # ── OVERLAY KE HALAMAN TERAKHIR (TTD PAGE) ──
        last_page_idx = len(pdf_reader.pages) - 1
        last_page = pdf_reader.pages[last_page_idx]
        
        # Create overlay canvas dengan stempel BC
        overlay_buffer = BytesIO()
        overlay_canvas = canvas.Canvas(overlay_buffer, pagesize=A4)
        
        # ── DRAW STEMPEL BC ──
        # Posisi: kolom kiri TTD BC (X=75, Y=300 untuk stempel image)
        if os.path.exists(stempel_path):
            try:
                # Draw image stempel (4cm x 4cm)
                overlay_canvas.drawImage(stempel_path, 150, 380, width=3*cm, height=3*cm)
            except Exception as e:
                app.logger.warning(f"Stempel image error: {str(e)}")
        
        # ── DRAW TEXT NAMA BC + TANGGAL ──
        overlay_canvas.setFont("Helvetica-Bold", 9)
        overlay_canvas.drawString(150, 365, f"✓ {nama_bc}")
        
        overlay_canvas.setFont("Helvetica", 8)
        overlay_canvas.drawString(150, 355, f"Tanggal: {tanggal_bc}")
        
        overlay_canvas.save()
        overlay_buffer.seek(0)
        
        # ── MERGE PDF + OVERLAY ──
        overlay_pdf = PdfReader(overlay_buffer)
        overlay_page = overlay_pdf.pages[0]
        last_page.merge_page(overlay_page)
        
        # ── ADD ALL PAGES TO WRITER ──
        for page_num in range(len(pdf_reader.pages)):
            if page_num == last_page_idx:
                pdf_writer.add_page(last_page)
            else:
                pdf_writer.add_page(pdf_reader.pages[page_num])
        
        # ── SAVE UPDATED PDF ──
        with open(file_path, 'wb') as output_file:
            pdf_writer.write(output_file)
        
        app.logger.info(f"Surat {file_surat} berhasil diupdate dengan stempel BC {nama_bc}")
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

# Route untuk logout
@app.route('/logout')
def logout():
    session.clear()  # Menghapus session login
    return redirect(url_for('login'))  # Mengarahkan kembali ke halaman login

if __name__ == '__main__':
    app.run(debug=True)


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