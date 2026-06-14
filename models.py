from datetime import datetime
from flask_sqlalchemy import SQLAlchemy

db = SQLAlchemy()

class Pengguna(db.Model):
    __tablename__ = 'pengguna'
    id       = db.Column(db.Integer, primary_key=True)
    nama     = db.Column(db.String(255), nullable=False)
    nip      = db.Column(db.String(50),  nullable=True)   # ← tambah ini
    email    = db.Column(db.String(255), unique=True, nullable=False)
    password = db.Column(db.String(255), nullable=False)
    role     = db.Column(db.String(50),  nullable=False)

class Permohonan(db.Model):
    id                  = db.Column(db.Integer, primary_key=True)
    id_pengguna         = db.Column(db.Integer, nullable=True)
    nama_perusahaan     = db.Column(db.String(200), nullable=False)
    no_job_order        = db.Column(db.String(50),  nullable=False)          # BARU
    nomor_kontainer     = db.Column(db.String(50),  nullable=False)
    no_booking          = db.Column(db.String(50),  nullable=False)          # BARU
    nopol_truck         = db.Column(db.String(30),  nullable=False)
    tanggal_pengajuan   = db.Column(db.String(20),  nullable=False)
    jenis_permohonan    = db.Column(db.String(50),  nullable=False)
    kendala             = db.Column(db.Text,        nullable=False)          # BARU (ganti alasan_permohonan)
    remark              = db.Column(db.Text,        nullable=True)           # BARU
    keterangan_tambahan = db.Column(db.Text,        nullable=True)
    file_surat          = db.Column(db.String(200), nullable=True)
    foto_dokumentasi    = db.Column(db.Text,        nullable=True)           # BARU (CSV list filename)
    status              = db.Column(db.String(30),  default='Menunggu')
    catatan             = db.Column(db.Text,        nullable=True)
    created_at          = db.Column(db.DateTime,    default=datetime.utcnow)
 
    def __repr__(self):
        return f'<Permohonan {self.nomor_kontainer}>'

class Persetujuan(db.Model):
    __tablename__ = 'persetujuan'
    id = db.Column(db.Integer, primary_key=True)
    id_permohonan = db.Column(db.Integer, db.ForeignKey('permohonan.id'), nullable=False)
    status_persetujuan = db.Column(db.String(50), nullable=False)  # Disetujui, Ditolak
    waktu_persetujuan = db.Column(db.DateTime, default=db.func.current_timestamp())
    komentar = db.Column(db.Text)

    permohonan = db.relationship('Permohonan', backref='persetujuans')