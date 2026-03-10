#!/bin/bash
# Script untuk menjalankan AbsensiPro

echo "==================================="
echo "   AbsensiPro - Sistem Absensi    "
echo "==================================="

# Cek apakah venv ada
if [ ! -d "venv" ]; then
    echo "📦 Membuat virtual environment..."
    python3 -m venv venv
fi

# Aktivasi venv
source venv/bin/activate

# Install dependencies
echo "📥 Menginstall dependencies..."
pip install -r requirements.txt --quiet

# Jalankan aplikasi
echo ""
echo "🚀 Menjalankan server di http://localhost:5000"
echo "📌 Login Admin: admin@absensi.com / admin123"
echo "🛑 Tekan Ctrl+C untuk berhenti"
echo ""

python3 app.py
