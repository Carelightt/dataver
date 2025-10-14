import logging
import pytz
import os 
import re 

from telegram import Update
from telegram.ext import Application, CommandHandler, ContextTypes, JobQueue 
from openpyxl import load_workbook
from openpyxl.styles import PatternFill 

# --- Sabitler (Sana Özel Bilgiler) ---
# Lütfen bu bilgilerin doğru olduğundan emin ol.
TOKEN = "8484668521:AAGiVlPq_SAc5UKBXpC6F7weGFOShJDJ0yA"
YETKILI_USER_ID = 6672759317  # Senin Telegram Kullanıcı ID'n
HEDEF_GRUP_ID = -1003195011322 # Verilerin gönderileceği Telegram Grup ID'si
EXCEL_DOSYA_ADI = "veriler.xlsx"
EXCEL_BOYAMA_RENGI = "00ADD8E6" # Açık Mavi (ARGB formatında)

# Veri Etiketlerinin Sıralaması (Excel sütun sırasına göre)
VERI_ETIKETLERI = [
    "TC", 
    "Ad Soyad", 
    "Telefon Numarası", 
    "Doğum Tarihi", 
    "İl/İlçe", 
    "IBAN"
]

# Günlükleme (logging) ayarları
logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    level=logging.INFO
)
logger = logging.getLogger(__name__)


# --- Yardımcı Fonksiyon: Yetki Kontrolü ve Yetkisiz Mesajı ---

def yetkili_mi(update: Update) -> bool:
    """
    Komutu kullanan kişinin yetkili ID'ye sahip olup olmadığını kontrol eder.
    Yetkisi yoksa istenen hata mesajını gönderir.
    """
    if update.effective_user.id != YETKILI_USER_ID:
        logger.warning(
            f"Yetkisiz erişim denemesi: User ID {update.effective_user.id} - Chat ID {update.effective_chat.id}"
        )
        # İstenen yetkisiz mesajı gönderiliyor
        update.message.reply_text("Yetkiniz yoktur.")
        return False
    return True


# --- YARDIMCI EXCEL FONKSİYONU ---

def excel_durumu_hesapla():
    """Excel dosyasındaki kullanılan (boyalı) ve kalan (boyasız) satırları hesaplar."""
    if not os.path.exists(EXCEL_DOSYA_ADI):
        return None, "Hata: Excel dosyası bulunamadı."
    
    try:
        workbook = load_workbook(EXCEL_DOSYA_ADI)
        sheet = workbook.active
        
        kullanilan_sayisi = 0
        kalan_sayisi = 0
        baslangic_satiri = 2
        
        # Tüm veri satırlarını döngüye al
        for row in sheet.iter_rows(min_row=baslangic_satiri):
            # Eğer ilk hücrenin rengi boyama rengiyle eşleşiyorsa, kullanılmıştır.
            if row[0].fill.start_color.rgb == EXCEL_BOYAMA_RENGI:
                kullanilan_sayisi += 1
            else:
                kalan_sayisi += 1
                
        return kalan_sayisi, kullanilan_sayisi
        
    except Exception as e:
        logger.error(f"Excel durumu hesaplanırken hata: {e}")
        return None, f"Kritik Hata: Excel okunamadı. ({e})"


# --- KOMUT İŞLEYİCİLERİ ---

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """/start komutuna yanıt verir."""
    # start komutuna gelen yetkisiz mesaj, yetkili_mi fonksiyonu içinde zaten gönderiliyor.
    if yetkili_mi(update):
        await update.message.reply_text(
            f'Merhaba yetkili! Ben göreve hazırım. Mevcut komutlar:\n\n'
            f'  • /ver <miktar>: Veri gönderir ve işaretler.\n'
            f'  • /kalan: Verilmemiş veri sayısını söyler.\n'
            f'  • /rapor: Verilmiş veri sayısını söyler.'
        )

# YENİ KOMUT: /kalan
async def kalan_komutu_isleyici(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Kalan veri sayısını bildirir."""
    if not yetkili_mi(update):
        return

    kalan, kullanilan = excel_durumu_hesapla()
    
    if kalan is None:
        await update.message.reply_text(kullanilan) # Hata mesajı gönderiliyor
        return
        
    await update.message.reply_text(
        f" KALAN DATA SAYISI\n"
        f"Elimizdeki data sayısı: **{kalan}**"
    )

# YENİ KOMUT: /rapor
async def rapor_komutu_isleyici(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Verilmiş (boyanmış) veri sayısını bildirir."""
    if not yetkili_mi(update):
        return

    kalan, kullanilan = excel_durumu_hesapla()
    
    if kullanilan is None:
        await update.message.reply_text(kalan) # Hata mesajı gönderiliyor
        return
        
    await update.message.reply_text(
        f" **Gönderilen Veri Raporu**\n"
        f"Verilen data sayısı: **{kullanilan}**"
    )

# Güncellenmiş /ver komutu (Excel'i açma mantığı aynı kaldı)
async def ver_komutu_isleyici(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """
    /ver <miktar> komutunu işler. Yetkili tarafından gruba veri gönderir ve Excel'de işaretler.
    """
    # 1. Yetki Kontrolü
    if not yetkili_mi(update):
        return

    # 2. Miktar Kontrolü ve Ayrıştırma (Aynı)
    if not context.args or not context.args[0].isdigit():
        await update.message.reply_text("Kullanım: `/ver <miktar>`. Lütfen kaç adet veri istediğinizi sayı olarak belirtin.")
        return
    
    try:
        miktar = int(context.args[0])
        if miktar <= 0:
            await update.message.reply_text("Lütfen pozitif bir sayı girin.")
            return
    except ValueError:
        await update.message.reply_text("Miktar sayı olmalıdır.")
        return

    # 3. Excel Dosya Kontrolü (Aynı)
    if not os.path.exists(EXCEL_DOSYA_ADI):
        await update.message.reply_text(f"Hata: '{EXCEL_DOSYA_ADI}' dosyası bulunamadı. Lütfen kontrol edin.")
        return

    await update.message.reply_text(f"{miktar} adet data çekiliyor ve gruba gönderiliyor...")

    try:
        # 4. Excel İşlemleri (Aynı)
        workbook = load_workbook(EXCEL_DOSYA_ADI)
        sheet = workbook.active  
        
        veriler = []
        satir_numaralari = []
        veri_sayisi = 0
        baslangic_satiri = 2 
        
        mavi_dolgu = PatternFill(start_color=EXCEL_BOYAMA_RENGI, end_color=EXCEL_BOYAMA_RENGI, fill_type="solid")

        # Okuma ve Toplama Döngüsü (Aynı)
        for row_index, row in enumerate(sheet.iter_rows(min_row=baslangic_satiri), start=baslangic_satiri):
            
            if row[0].fill.start_color.rgb == EXCEL_BOYAMA_RENGI:
                 continue

            if veri_sayisi >= miktar:
                break
                
            # YENİ FORMATLAMA (Aynı)
            satir_verisi_duzenli = []
            hucre_degerleri = [str(cell.value).strip() if cell.value is not None else "" for cell in row]
            
            for etiket_index, etiket in enumerate(VERI_ETIKETLERI):
                if etiket_index < len(hucre_degerleri):
                    deger = hucre_degerleri[etiket_index]
                    satir_verisi_duzenli.append(f"**{etiket}**: {deger}")
            
            veriler.append("\n".join(satir_verisi_duzenli)) 
            
            satir_numaralari.append(row_index)
            veri_sayisi += 1

        if not veriler:
            await update.message.reply_text("Üzgünüm, Excel dosyasında gönderilebilecek işaretlenmemiş veri kalmadı.")
            return

        # 5. Verileri Gruba Gönder (Aynı)
        gonderilecek_mesaj = f"**{veri_sayisi}** adet yeni data:\n\n" + "\n\n---\n\n".join(veriler)
        
        await context.bot.send_message(
            chat_id=HEDEF_GRUP_ID,
            text=gonderilecek_mesaj,
            parse_mode='Markdown'
        )
        
        # 6. Excel'de Kullanılan Satırları İşaretle (Aynı)
        for satir_no in satir_numaralari:
            for col_index in range(1, sheet.max_column + 1):
                sheet.cell(row=satir_no, column=col_index).fill = mavi_dolgu

        # Değişiklikleri kaydet
        workbook.save(EXCEL_DOSYA_ADI)

        await update.message.reply_text(
            f"{veri_sayisi} adet data gruba gönderildi ve çöp kutusuna taşındı."
        )

    except Exception as e:
        logger.error(f"Excel veya Telegram işlemi sırasında kritik hata: {e}")
        await update.message.reply_text(f"❌ Bir Hata Oluştu. Detaylar loglara kaydedildi. Hata: `{e}`")


# --- Ana Fonksiyon ---

def main() -> None:
    """Botu başlatır."""
    
    try:
        application = (
            Application.builder()
            .token(TOKEN)
            .concurrent_updates(True)
            .job_queue(JobQueue()) 
            .build()
        )
        
        application.job_queue.scheduler.configure(timezone=pytz.timezone('Europe/Istanbul'))

    except Exception as e:
        logger.error(f"Bot başlatma hatası: {e}")
        logger.error("Lütfen Job Queue'yu kurduğunuzdan emin olun: 'pip install \"python-telegram-bot[job-queue]\"'")
        return 

    # Komut işleyicilerini ekle
    application.add_handler(CommandHandler("start", start))
    application.add_handler(CommandHandler("ver", ver_komutu_isleyici)) 
    application.add_handler(CommandHandler("kalan", kalan_komutu_isleyici)) # Yeni komut
    application.add_handler(CommandHandler("rapor", rapor_komutu_isleyici)) # Yeni komut

    # Bot çalışmaya başlar (sürekli güncellemeleri dinler)
    logger.info("Bot başarıyla başlatıldı ve dinlemede...")
    application.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == '__main__':
    main()