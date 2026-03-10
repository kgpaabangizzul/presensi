# 🏢 AbsensiPro - Sistem Manajemen Kehadiran Digital

Aplikasi absensi digital berbasis Flask dengan fitur lengkap: GPS tracking, foto absen, dashboard admin, grafik, dan export laporan.

## ✨ Fitur Utama

### 👤 Portal Pegawai
- **Registrasi** dengan biodata lengkap + foto profil
- **Login** setelah divalidasi admin
- **Absen Masuk/Keluar** dengan:
  - 📸 Foto selfie via kamera
  - 📍 Deteksi GPS otomatis
  - 📏 Kalkulasi jarak dari kantor
- **Riwayat Absensi** dengan filter bulan
- **Pengajuan Izin/Cuti** dengan lampiran dokumen
- **Profil** — edit data & foto

### 🛡️ Panel Admin
- **Dashboard** — statistik real-time, grafik tren 7 hari, per departemen
- **Kelola Pegawai** — validasi/tolak akun baru, lihat foto
- **Data Absensi** — filter per bulan & departemen
- **Kelola Izin** — approve/reject permohonan izin
- **Grafik & Analitik** — chart harian, pie chart, per departemen
- **Laporan & Export** — rekap bulanan + progress bar kehadiran
- **Export Excel** — file .xlsx berwarna dengan styling
- **Export PDF** — laporan A4 landscape
- **Pengaturan** — konfigurasi jam kerja, koordinat GPS kantor, radius absen

## 🚀 Cara Menjalankan

### Menggunakan Script (Rekomendasi)
```bash
chmod +x run.sh
./run.sh
```

### Manual
```bash
# Buat virtual environment
python3 -m venv venv
source venv/bin/activate       # Linux/Mac
venv\Scripts\activate          # Windows

# Install dependencies
pip install -r requirements.txt

# Jalankan
python3 app.py
```

Buka browser: **http://localhost:5000**

## 🔑 Login Default
| Role | Email | Password |
|------|-------|----------|


## 📁 Struktur Proyek
```
absensi/
├── app.py                 # Main Flask app
├── requirements.txt       # Dependencies
├── run.sh                 # Script runner
├── instance/
│   └── absensi.db         # Database SQLite (auto-created)
├── static/
│   └── uploads/photos/    # Foto pegawai & absensi
└── templates/
    ├── base.html          # Base template
    ├── layout.html        # User layout
    ├── login.html
    ├── register.html
    ├── dashboard.html
    ├── riwayat.html
    ├── izin.html
    ├── profil.html
    └── admin/
        ├── layout.html    # Admin layout
        ├── dashboard.html
        ├── pegawai.html
        ├── absensi.html
        ├── izin.html
        ├── laporan.html
        ├── grafik.html
        └── settings.html
```

## 🔧 Konfigurasi
Edit di **Admin → Pengaturan**:
- Nama perusahaan
- Jam masuk/keluar
- Koordinat GPS kantor (lat/lng)
- Radius maksimal absen (meter)

## 💡 Alur Penggunaan
1. Pegawai **daftar** dengan biodata lengkap
2. Admin **validasi** akun di panel pegawai
3. Pegawai **login** dan mulai absen harian
4. Admin monitor kehadiran via **dashboard**
5. Export **laporan Excel/PDF** tiap bulan
