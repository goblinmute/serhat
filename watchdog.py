"""
WATCHDOG — Arbitraj Bot Bekçisi
================================
- main.py'yi sürekli çalışır durumda tutar.
- Çökerse / kapanırsa Telegram'a bildirim gönderir.
- 30 saniye bekleyip otomatik yeniden başlatır.
- Yeniden başlatmayı da Telegram'a bildirir.
"""

import subprocess
import sys
import os
import time
import requests
from datetime import datetime

# ─── Ayarlar ────────────────────────────────────────────────
TELEGRAM_BOT_TOKEN = "8932402031:AAHENNJ_-n123XoXuLlBxTHlPGfVf_8uVck"
TELEGRAM_CHAT_ID   = "8724428153"
RESTART_DELAY_SEC  = 30          # çöküş sonrası bekleme süresi
SCRIPT_DIR         = os.path.dirname(os.path.abspath(__file__))
MAIN_SCRIPT        = os.path.join(SCRIPT_DIR, "main.py")
PYTHON_EXE         = sys.executable  # watchdog ile aynı Python ortamı
# ────────────────────────────────────────────────────────────

def tg(msg: str):
    """Telegram'a mesaj gönder (hata olsa sessizce geç)."""
    try:
        url  = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
        data = {"chat_id": TELEGRAM_CHAT_ID, "text": msg, "parse_mode": "HTML"}
        requests.post(url, data=data, timeout=10)
    except Exception:
        pass

def now_str():
    return datetime.now().strftime("%d.%m.%Y %H:%M:%S")

def main():
    run_count = 0

    tg(
        "🐕 <b>WATCHDOG BAŞLATILDI</b>\n"
        "━━━━━━━━━━━━━━━━━━━━\n"
        "Bot bekçisi aktif. main.py çökerse otomatik\n"
        f"yeniden başlatılacak ({RESTART_DELAY_SEC}sn bekleyerek).\n"
        f"🕐 {now_str()}"
    )
    print(f"[{now_str()}] Watchdog başlatıldı. main.py izleniyor...")

    while True:
        run_count += 1
        start_time = datetime.now()
        print(f"\n[{now_str()}] ▶ main.py başlatılıyor... (#{run_count}. çalışma)")

        try:
            proc = subprocess.run(
                [PYTHON_EXE, MAIN_SCRIPT],
                cwd=SCRIPT_DIR
            )
            exit_code = proc.returncode
        except Exception as e:
            exit_code = -1
            print(f"[{now_str()}] subprocess hatası: {e}")

        end_time   = datetime.now()
        duration   = end_time - start_time
        dur_str    = f"{int(duration.total_seconds() // 3600)}s " \
                     f"{int((duration.total_seconds() % 3600) // 60)}dk"

        print(f"[{now_str()}] ✗ main.py durdu. Çıkış kodu: {exit_code} | Çalışma süresi: {dur_str}")

        if exit_code == 0:
            # Normal çıkış (Ctrl+C gibi kasıtlı)
            icon  = "🟡"
            title = "BOT NORMAL DURDU"
            note  = "Kullanıcı tarafından durdurulmuş olabilir."
        else:
            icon  = "🔴"
            title = "BOT ÇÖKTÜ / KESİNTİ YAŞANDI"
            note  = f"Çıkış kodu: <code>{exit_code}</code>"

        tg(
            f"{icon} <b>{title}</b>\n"
            "━━━━━━━━━━━━━━━━━━━━\n"
            f"🕐 Başlangıç : {start_time.strftime('%d.%m.%Y %H:%M:%S')}\n"
            f"🕑 Kapanış  : {end_time.strftime('%d.%m.%Y %H:%M:%S')}\n"
            f"⏱ Çalışma  : {dur_str}\n"
            f"ℹ️ {note}\n"
            f"🔄 {RESTART_DELAY_SEC} saniye sonra yeniden başlatılacak..."
        )

        print(f"[{now_str()}] {RESTART_DELAY_SEC} saniye bekleniyor...")
        time.sleep(RESTART_DELAY_SEC)

        tg(
            f"♻️ <b>BOT YENİDEN BAŞLATILIYOR</b>\n"
            "━━━━━━━━━━━━━━━━━━━━\n"
            f"🔢 #{run_count + 1}. çalışma\n"
            f"🕐 {now_str()}"
        )
        print(f"[{now_str()}] Yeniden başlatılıyor...")

if __name__ == "__main__":
    main()
