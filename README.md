# Arbitraj Bot — Ekonomi Radar Sistemi

Türkiye ekonomi haberlerini ve Polymarket kontratlarını takip eden, otomatik fırsat tespiti yapan Python botu.

## Özellikler
- RSS kaynakları üzerinden ekonomi haberi tarama (HaberTürk, NTV, TRT, CNN Türk)
- Polymarket Türkiye kontratlarını otomatik eşleştirme
- Excel'e canlı raporlama (Türkiye Gündem + Polymarket Fırsatları)
- Telegram bildirimleri (yüksek fırsat, jeopolitik risk, hata, kalp atışı)
- Gece yarısı otomatik sayfa geçişi
- Watchdog: bot çökerse otomatik yeniden başlatır

## Kurulum (VPS / Sunucu)

```bash
# 1. Repoyu çek
git clone https://github.com/goblinmute/serhat.git
cd serhat

# 2. Bağımlılıkları kur
pip install -r requirements.txt

# 3. Watchdog ile başlat (sürekli açık kalır)
python watchdog.py
```

## Dosyalar
| Dosya | Açıklama |
|---|---|
| `main.py` | Ana bot — haber tarama, Polymarket, Excel, Telegram |
| `watchdog.py` | Bekçi — bot çökerse Telegram bildir + yeniden başlat |
| `requirements.txt` | Python bağımlılıkları |

## Telegram Bildirimleri
| Mesaj | Durum |
|---|---|
| 🚀 BOT BAŞLATILDI | Başarılı başlangıç |
| 💓 SİSTEM AKTİF | Saatlik kalp atışı |
| 🟢 YÜKSEK FIRSAT | Dolar/TL + Polymarket eşleşmesi |
| 🟡 ORTA ÖNCELİKLİ | Altın/Petrol eşleşmesi |
| 🔴 KRİTİK HATA | Bot hatası |
| ⛔ BOT DURDU | Kapanış bildirimi (watchdog) |
| ♻️ YENİDEN BAŞLATILIYOR | Otomatik restart (watchdog) |
