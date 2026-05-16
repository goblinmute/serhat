#!/bin/bash
# Arbitraj Bot - Otomatik Kurulum Scripti
# Kullanim: curl -sSL https://raw.githubusercontent.com/goblinmute/serhat/main/install.sh | bash

set -e

echo "========================================"
echo "  Arbitraj Bot Kurulum Basliyor..."
echo "========================================"

# 1. Sistem guncelleme ve gerekli paketler
echo "[1/5] Sistem guncelleniyor..."
apt-get update -qq
apt-get install -y -qq python3-pip git screen

# 2. Eski kurulum varsa temizle
echo "[2/5] Onceki kurulum temizleniyor..."
rm -rf /root/bot

# 3. Repoyu cek
echo "[3/5] Kod indiriliyor (GitHub)..."
git clone https://github.com/goblinmute/serhat.git /root/bot

# 4. Bagimliliklar
echo "[4/5] Python kutuphaneleri kuruluyor..."
cd /root/bot
pip3 install -r requirements.txt --break-system-packages -q

# 5. Servisi baslat
echo "[5/5] Bot baslatiliyor (screen oturumu: arbitrajbot)..."
screen -dmS arbitrajbot bash -c "cd /root/bot && python3 watchdog.py 2>&1 | tee /root/bot/watchdog.log"

echo ""
echo "========================================"
echo "  KURULUM TAMAMLANDI!"
echo "========================================"
echo ""
echo "  Botu izlemek icin:"
echo "  screen -r arbitrajbot"
echo ""
echo "  Ekrandan cikmak icin (bot durmaz):"
echo "  Ctrl+A sonra D"
echo ""
echo "  Log dosyasi:"
echo "  tail -f /root/bot/watchdog.log"
echo "========================================"
