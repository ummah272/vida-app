from flask import Flask, render_template, request, redirect, url_for, session, flash, send_from_directory, jsonify
from models import db, Pengguna, Permohonan
from config import Config
import pymysql
import os
from werkzeug.utils import secure_filename
from docx import Document
from datetime import datetime
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm

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
# HELPER FUNCTIONS
# ═══════════════════════════════════════════════════════════════════════════════

def generate_surat(data: dict, foto_files: list = None) -> str:
    """Generate surat permohonan dari gate (Word format)"""
    doc = Document()
 
    judul = doc.add_heading('SURAT PERMOHONAN', level=1)
    judul.alignment = 1
 
    sub = doc.add_heading('PENGAJUAN KONTAINER EMKL', level=2)
    sub.alignment = 1
 
    doc.add_paragraph('')
 
    timestamp  = datetime.now().strftime('%Y%m%d%H%M%S')
    no_surat   = f"EMKL/{data['nomor_kontainer']}/{timestamp}"
    tgl_format = datetime.strptime(data['tanggal_pengajuan'], '%Y-%m-%d').strftime('%d %B %Y')
 
    doc.add_paragraph(f"Nomor Surat    : {no_surat}")
    doc.add_paragraph(f"Tanggal        : {tgl_format}")
    doc.add_paragraph('')
 
    doc.add_paragraph("Kepada Yth.")
    doc.add_paragraph("Kepala Bea dan Cukai / Gate Pengeluaran Kontainer")
    doc.add_paragraph("di Tempat")
    doc.add_paragraph('')
 
    doc.add_paragraph("Dengan hormat,")
    doc.add_paragraph(
        f"Yang bertanda tangan di bawah ini, kami dari {data['nama_perusahaan']} "
        f"dengan ini mengajukan permohonan "
        f"{data['jenis_permohonan'].replace('_',' ').title()} untuk:"
    )
    doc.add_paragraph('')
 
    table = doc.add_table(rows=6, cols=2)
    table.style = 'Table Grid'
    rows_data = [
        ('No Job Order',       data['no_job_order']),
        ('Nomor Kontainer',    data['nomor_kontainer']),
        ('No Booking',         data['no_booking']),
        ('Nopol Truck',        data['nopol_truck']),
        ('Jenis Permohonan',   data['jenis_permohonan'].replace('_',' ').title()),
        ('Tanggal Pengajuan',  tgl_format),
    ]
    for i, (key, val) in enumerate(rows_data):
        table.rows[i].cells[0].text = key
        table.rows[i].cells[1].text = val
 
    doc.add_paragraph('')
    doc.add_paragraph("Kendala:")
    doc.add_paragraph(data['kendala'])
 
    doc.add_paragraph('')
    doc.add_paragraph("Remark & Catatan Khusus:")
    doc.add_paragraph(data['remark'] if data.get('remark') else "-")
 
    # ── DOKUMENTASI FOTO (INSERT KE SURAT) ──
    if foto_files:
        doc.add_paragraph('')
        doc.add_paragraph("Dokumentasi Foto Kontainer di EA:")
        doc.add_paragraph(f"Jumlah file foto: {len(foto_files)}")
        doc.add_paragraph('')
        
        # Insert foto langsung ke surat
        for idx, foto_filename in enumerate(foto_files, 1):
            foto_path = os.path.join(FOTO_FOLDER, foto_filename.strip())
            if os.path.exists(foto_path):
                try:
                    doc.add_paragraph(f"Foto {idx}:")
                    doc.add_picture(foto_path, width=15*cm)  # Lebar 15cm
                    doc.add_paragraph('')
                except Exception as e:
                    doc.add_paragraph(f"[Foto {idx} tidak bisa ditampilkan: {str(e)}]")
 
    if data.get('keterangan_tambahan'):
        doc.add_paragraph('')
        doc.add_paragraph("Keterangan Tambahan:")
        doc.add_paragraph(data['keterangan_tambahan'])
 
    doc.add_paragraph('')
    doc.add_paragraph('')
    doc.add_paragraph("Demikian permohonan ini kami sampaikan. Atas perhatian dan kerja samanya, kami ucapkan terima kasih.")
    doc.add_paragraph('')
    doc.add_paragraph("Hormat kami,")
    doc.add_paragraph(data['nama_perusahaan'])
 
    filename = secure_filename(f"surat_permohonan_{data['nomor_kontainer']}_{timestamp}.docx")
    filepath = os.path.join(SURAT_FOLDER, filename)
    doc.save(filepath)
 
    return filename
 
 
def save_uploaded_photos(files, nomor_kontainer) -> list:
    """Save uploaded photos dengan naming: NOKO-[timestamp]-[nomor].jpg"""
    saved_files = []
    timestamp = datetime.now().strftime('%Y%m%d%H%M%S')
    
    for idx, file in enumerate(files, 1):
        if file and allowed_file(file.filename):
            ext = file.filename.rsplit('.', 1)[1].lower()
            filename = secure_filename(f"{nomor_kontainer}-{timestamp}-{idx}.{ext}")
            filepath = os.path.join(FOTO_FOLDER, filename)
            file.save(filepath)
            saved_files.append(filename)
    
    return saved_files
 
 
def generate_surat_persetujuan(permohonan) -> str:
    """Generate surat persetujuan dari Bea Cukai (PDF format) + AUTO APPROVED"""
    timestamp   = datetime.now().strftime('%Y%m%d%H%M%S')
    filename    = secure_filename(f"surat_persetujuan_{permohonan.nomor_kontainer}_{timestamp}.pdf")
    filepath    = os.path.join(SURAT_FOLDER, filename)
 
    doc    = SimpleDocTemplate(filepath, pagesize=A4, topMargin=2*cm, bottomMargin=1*cm)
    styles = getSampleStyleSheet()
    story  = []
 
    title_style = ParagraphStyle('title', parent=styles['Heading1'], alignment=1, fontSize=14, spaceAfter=0.3*cm)
    subtitle_style = ParagraphStyle('subtitle', parent=styles['Normal'], alignment=1, fontSize=11, spaceAfter=0.5*cm)
    normal_style = ParagraphStyle('normal', parent=styles['Normal'], fontSize=10, alignment=0)
 
    # KOP SURAT
    if os.path.exists(LOGO_PATH):
        logo = Image(LOGO_PATH, width=1.5*cm, height=1.5*cm)
        kop_table = Table([[logo, Paragraph("<b>PT. PELINDO III</b><br/>TERMINAL PETIKEMAS", styles['Normal'])]], colWidths=[2*cm, 14*cm])
        kop_table.setStyle(TableStyle([('ALIGN', (0,0), (-1,-1), 'CENTER'), ('VALIGN', (0,0), (-1,-1), 'MIDDLE')]))
        story.append(kop_table)
    else:
        story.append(Paragraph("<b>PT. PELINDO III - TERMINAL PETIKEMAS</b>", subtitle_style))
    
    story.append(Spacer(1, 0.3*cm))
    story.append(Paragraph("_" * 100, styles['Normal']))
    story.append(Spacer(1, 0.5*cm))
 
    story.append(Paragraph("SURAT PERSETUJUAN", title_style))
    story.append(Paragraph("PERMOHONAN KONTAINER EMKL", title_style))
    story.append(Spacer(1, 0.5*cm))
 
    no_surat    = f"BC/SETUJUI/{permohonan.nomor_kontainer}/{timestamp}"
    tgl_format  = datetime.now().strftime('%d %B %Y')
    tgl_pengajuan = datetime.strptime(permohonan.tanggal_pengajuan, '%Y-%m-%d').strftime('%d %B %Y')
 
    story.append(Paragraph(f"<b>Nomor Surat</b> : {no_surat}", normal_style))
    story.append(Paragraph(f"<b>Tanggal</b>     : {tgl_format}", normal_style))
    story.append(Spacer(1, 0.5*cm))
 
    story.append(Paragraph("<b>Kepada Yth.</b>", normal_style))
    story.append(Paragraph(f"{permohonan.nama_perusahaan}", normal_style))
    story.append(Paragraph("di Tempat", normal_style))
    story.append(Spacer(1, 0.5*cm))
 
    story.append(Paragraph("<b>Dengan hormat,</b>", normal_style))
    story.append(Spacer(1, 0.2*cm))
    story.append(Paragraph(
        f"Sehubungan dengan permohonan yang telah diajukan oleh <b>{permohonan.nama_perusahaan}</b>, "
        f"dengan ini kami dari pihak Bea dan Cukai menyatakan bahwa permohonan "
        f"<b>{permohonan.jenis_permohonan.replace('_',' ').title()}</b> berikut telah <b>DISETUJUI</b>:",
        normal_style
    ))
    story.append(Spacer(1, 0.4*cm))
 
    table_data = [
        ['Nama Perusahaan',  permohonan.nama_perusahaan],
        ['No Job Order',     permohonan.no_job_order],
        ['Nomor Kontainer',  permohonan.nomor_kontainer],
        ['No Booking',       permohonan.no_booking],
        ['Nopol Truck',      permohonan.nopol_truck],
        ['Jenis Permohonan', permohonan.jenis_permohonan.replace('_',' ').title()],
        ['Tanggal Pengajuan', tgl_pengajuan],
    ]
    t = Table(table_data, colWidths=[5*cm, 11*cm])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (0,-1), colors.HexColor("#8C959E")),
        ('TEXTCOLOR', (0,0), (0,-1), colors.navy),
        ('FONTNAME', (0,0), (0,-1), 'Helvetica-Bold'),
        ('FONTSIZE', (0,0), (-1,-1), 10),
        ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
        ('PADDING', (0,0), (-1,-1), 8),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.beige, colors.white]),
    ]))
    story.append(t)
    story.append(Spacer(1, 0.5*cm))
 
    # FOTO DOKUMENTASI
    if permohonan.foto_dokumentasi:
        foto_list = permohonan.foto_dokumentasi.split(',')
        story.append(Paragraph("<b>Dokumentasi Foto Kontainer di EA:</b>", normal_style))
        story.append(Spacer(1, 0.2*cm))
        
        foto_table_data = []
        for i, foto_filename in enumerate(foto_list):
            foto_path = os.path.join(FOTO_FOLDER, foto_filename.strip())
            if os.path.exists(foto_path):
                try:
                    foto_img = Image(foto_path, width=4*cm, height=3*cm)
                    foto_table_data.append([Paragraph(f"<b>Foto {i+1}</b>", styles['Normal']), foto_img])
                except:
                    pass
        
        if foto_table_data:
            foto_table = Table(foto_table_data, colWidths=[2*cm, 5*cm])
            foto_table.setStyle(TableStyle([('ALIGN', (0,0), (-1,-1), 'CENTER'), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('PADDING', (0,0), (-1,-1), 4)]))
            story.append(foto_table)
            story.append(Spacer(1, 0.5*cm))
 
    story.append(Paragraph("Demikian surat persetujuan ini kami keluarkan untuk dapat dipergunakan sebagaimana mestinya.", normal_style))
    story.append(Spacer(1, 0.8*cm))
 
    # TANDA TANGAN DENGAN AUTO APPROVED
    sign_style = ParagraphStyle('sign', parent=styles['Normal'], fontSize=9, alignment=1)
    
    signature_data = [
        [
            Paragraph("<b>Pejabat Bea dan Cukai</b>", sign_style),
            "",
            Paragraph("<b>Gate Superintendent</b>", sign_style),
        ],
        [
            Paragraph("<b style='font-size: 16px; color: green;'>✓ AUTO APPROVED ✓</b>", sign_style),
            "",
            Spacer(1, 1.5*cm),
        ],
        [
            Paragraph("(________________________)", sign_style),
            "",
            Paragraph("(________________________)", sign_style),
        ],
        [
            Paragraph("NIP: _______________", sign_style),
            "",
            Paragraph("Nama: _______________", sign_style),
        ],
    ]
    
    sig_table = Table(signature_data, colWidths=[5*cm, 3*cm, 5*cm])
    sig_table.setStyle(TableStyle([('ALIGN', (0,0), (-1,-1), 'CENTER'), ('VALIGN', (0,0), (-1,-1), 'MIDDLE')]))
    story.append(sig_table)
    story.append(Spacer(1, 1*cm))
    story.append(Paragraph(f"Tanggal: {tgl_format}", normal_style))
 
    doc.build(story)
    return filename


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


@app.route('/dashboard_emkl')
def dashboard_emkl():
    if 'username' not in session:
        return redirect(url_for('login'))
    return render_template('dashboard_emkl.html')


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


@app.route('/dashboard_gate')
def dashboard_gate():
    if 'username' not in session:
        return redirect(url_for('login'))
    
    # Get stats
    total_pengajuan = Permohonan.query.count()
    menunggu_pengajuan = Permohonan.query.filter_by(status='Menunggu').count()
    disetujui_pengajuan = Permohonan.query.filter_by(status='Disetujui').count()
    
    # Get 3 approved pengajuan terbaru
    approved_pengajuan = Permohonan.query.filter_by(status='Disetujui').order_by(Permohonan.created_at.desc()).limit(3).all()
    
    return render_template('dashboard_gate.html', 
                         total_pengajuan=total_pengajuan,
                         menunggu_pengajuan=menunggu_pengajuan,
                         disetujui_pengajuan=disetujui_pengajuan,
                         approved_pengajuan=approved_pengajuan)


#@app.route('/dashboard_ea')
#def dashboard_ea():
#    if 'username' not in session:
#        return redirect(url_for('login'))
#    return render_template('dashboard_ea.html')


# ═══════════════════════════════════════════════════════════════════════════════
# GATE ROUTES - Gate buat surat permohonan & monitoring permohonan mereka
# ═══════════════════════════════════════════════════════════════════════════════

@app.route('/permohonan', methods=['GET', 'POST'])
def permohonan():
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

        required = [
            'nama_perusahaan', 'no_job_order', 'nomor_kontainer',
            'no_booking', 'nopol_truck', 'tanggal_pengajuan', 'jenis_permohonan', 
            'kendala', 'remark'
        ]
        if any(not data[k] for k in required):
            flash('Harap lengkapi semua field yang wajib diisi.', 'error')
            return render_template('permohonan.html')

        try:
            foto_files = request.files.getlist('foto_kontainer')
            saved_foto_files = []
            
            if foto_files and foto_files[0].filename != '':
                saved_foto_files = save_uploaded_photos(foto_files, data['nomor_kontainer'])
                
                if not saved_foto_files:
                    flash('Gagal menyimpan foto. Pastikan format file adalah JPG/PNG.', 'error')
                    return render_template('permohonan.html')

            filename_surat = generate_surat(data, saved_foto_files)

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

            flash('Pengajuan berhasil dikirim! Surat permohonan telah digenerate otomatis.', 'success')
            return redirect(url_for('monitoring_gate'))

        except Exception as e:
            db.session.rollback()
            flash(f'Terjadi kesalahan: {str(e)}', 'error')
            return render_template('permohonan.html')

    return render_template('permohonan.html')


@app.route('/monitoring_gate', methods=['GET'])
def monitoring_gate():
    if 'username' not in session:
        return redirect(url_for('login'))

    # Gate hanya bisa lihat permohonan yang mereka buat sendiri
    user_id = session.get('user_id')
    pengajuan_kontainer = Permohonan.query.filter_by(id_pengguna=user_id).all()

    return render_template('monitoring_gate.html', pengajuan_kontainer=pengajuan_kontainer)


# ═══════════════════════════════════════════════════════════════════════════════
# BEA CUKAI ROUTES - Approve/Reject permohonan dari Gate
# ═══════════════════════════════════════════════════════════════════════════════

@app.route('/monitoring_beacukai', methods=['GET', 'POST'])
def monitoring_beacukai():
    if 'username' not in session:
        return redirect(url_for('login'))

    pengajuan_kontainer = Permohonan.query.filter_by(status='Menunggu').all()

    if request.method == 'POST':
        action        = request.form.get('action')
        permohonan_id = request.form.get('permohonan_id')
        catatan       = request.form.get('catatan', '').strip()

        permohonan = Permohonan.query.filter_by(id=permohonan_id).first()

        if permohonan:
            if action == 'approve':
                permohonan.status = 'Disetujui'
                permohonan.catatan = catatan
                try:
                    filename_surat = generate_surat_persetujuan(permohonan)
                    permohonan.file_surat = filename_surat
                except Exception as e:
                    flash(f'Gagal generate surat: {str(e)}', 'error')
                    db.session.rollback()
                    return redirect(url_for('monitoring_beacukai'))
                    
            elif action == 'reject':
                permohonan.status = 'Ditolak'
                permohonan.catatan = catatan
            
            db.session.commit()
            flash('Status permohonan telah diupdate', 'success')

        return redirect(url_for('monitoring_beacukai'))

    return render_template('monitoring_beacukai.html', pengajuan_kontainer=pengajuan_kontainer)


# ═══════════════════════════════════════════════════════════════════════════════
# CS ROUTES - Monitor semua permohonan + bisa lihat surat yang sudah disetujui
# ═══════════════════════════════════════════════════════════════════════════════

@app.route('/monitoring_cs', methods=['GET', 'POST'])
def monitoring_cs():
    if 'username' not in session:
        return redirect(url_for('login'))

    # CS bisa lihat semua permohonan (baik menunggu, disetujui, atau ditolak)
    pengajuan_kontainer = Permohonan.query.all()

    if request.method == 'POST':
        permohonan_id = request.form.get('permohonan_id')
        catatan = request.form.get('catatan')

        permohonan = Permohonan.query.filter_by(id=permohonan_id).first()
        if permohonan:
            permohonan.catatan = catatan
            db.session.commit()

        return redirect(url_for('monitoring_cs'))

    return render_template('monitoring_cs.html', pengajuan_kontainer=pengajuan_kontainer)


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

@app.route('/lihat_surat/<filename>')
def lihat_surat(filename):
    if 'username' not in session:
        return redirect(url_for('login'))
    
    file_path = os.path.join(SURAT_FOLDER, filename)
    if os.path.exists(file_path):
        return send_from_directory(SURAT_FOLDER, filename)
    else:
        flash('File surat tidak ditemukan', 'error')
        return redirect(url_for('monitoring_cs'))


@app.route('/download_surat/<filename>')
def download_surat(filename):
    if 'username' not in session:
        return redirect(url_for('login'))
    return send_from_directory(SURAT_FOLDER, filename, as_attachment=True)


@app.route('/lihat_foto/<filename>')
def lihat_foto(filename):
    if 'username' not in session:
        return redirect(url_for('login'))
    
    file_path = os.path.join(FOTO_FOLDER, filename)
    if os.path.exists(file_path):
        return send_from_directory(FOTO_FOLDER, filename)
    else:
        flash('File foto tidak ditemukan', 'error')
        return redirect(url_for('monitoring_cs'))


@app.route('/api/permohonan/<int:permohonan_id>')
def get_permohonan_detail(permohonan_id):
    if 'username' not in session:
        return jsonify({'error': 'Unauthorized'}), 401
    
    permohonan = Permohonan.query.filter_by(id=permohonan_id).first()
    
    if not permohonan:
        return jsonify({'error': 'Permohonan not found'}), 404
    
    return jsonify({
        'id': permohonan.id,
        'nama_perusahaan': permohonan.nama_perusahaan,
        'no_job_order': permohonan.no_job_order,
        'nomor_kontainer': permohonan.nomor_kontainer,
        'no_booking': permohonan.no_booking,
        'nopol_truck': permohonan.nopol_truck,
        'tanggal_pengajuan': permohonan.tanggal_pengajuan,
        'jenis_permohonan': permohonan.jenis_permohonan,
        'kendala': permohonan.kendala,
        'remark': permohonan.remark,
        'keterangan_tambahan': permohonan.keterangan_tambahan,
        'status': permohonan.status,
        'catatan': permohonan.catatan,
        'file_surat': permohonan.file_surat,
        'foto_dokumentasi': permohonan.foto_dokumentasi,
        'created_at': permohonan.created_at.strftime('%d-%m-%Y %H:%M') if permohonan.created_at else None,
    })


# ═══════════════════════════════════════════════════════════════════════════════
# PROFILE & UTILITY ROUTES
# ═══════════════════════════════════════════════════════════════════════════════

#@app.route('/profile_cs')
#def profile_cs():
#    if 'username' not in session:
#        return redirect(url_for('login'))
#    return render_template('profile_cs.html')


@app.route('/profile_beacukai')
def profile_beacukai():
    if 'username' not in session:
        return redirect(url_for('login'))
    return render_template('profile_beacukai.html')


#@app.route('/profile_gate')
#def profile_gate():
#    if 'username' not in session:
#        return redirect(url_for('login'))
#    return render_template('profile_gate.html')

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

# Route untuk logout
@app.route('/logout')
def logout():
    session.clear()  # Menghapus session login
    return redirect(url_for('login'))  # Mengarahkan kembali ke halaman login

if __name__ == '__main__':
    app.run(debug=True)