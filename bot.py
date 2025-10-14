import logging
import pytz
import os 
import re 
import json # YENİ: Satır numaralarını kaydetmek için eklendi

from telegram import Update
from telegram.ext import Application, CommandHandler, ContextTypes, JobQueue 
from openpyxl import load_workbook
# from openpyxl.styles import PatternFill # Kaldırıldı, artık boyama yapmayacağız

# --- Sabitler (Sana Özel Bilgiler) ---
# Lütfen bu bilgilerin doğru olduğundan emin ol.
TOKEN = "8484668521:AAGiVlPq_SAc5UKBXpC6F7weGFOShJDJ0yA"
YETKILI_USER_ID = 6672759317  # Senin Telegram Kullanıcı ID'n
HEDEF_GRUP_ID = -1003195011322 # Verilerin gönderileceği Telegram Grup ID'si
EXCEL_DOSYA_ADI = "veriler.xlsx"
# EXCEL_BOYAMA_RENGI = "00ADD8E6" # Boyama sabiti kaldırıldı
KULLANILANLAR_DOSYA_ADI = "kullanilanlar.txt" # YENİ: Kullanılan satırları tutacak dosya

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


# --- YARDIMCI KALICILIK FONKSİYONLARI (YENİ) ---

def kullanilan_satirlari_oku() -> set:
    """Kullanılan satır numaralarını dosyadan okur."""
    if not os.path.exists(KULLANILANLAR_DOSYA_ADI):
        return set()
    try:
        with open(KULLANILANLAR_DOSYA_ADI, 'r') as f:
            # Satır numaraları metin dosyasında JSON listesi olarak tutuluyor
            return set(json.load(f))
    except Exception as e:
        # Dosya boş veya bozuksa, boş bir küme döndür.
        logger.warning(f"Kullanılan satırlar dosyası okunamadı, sıfırdan başlıyor: {e}")
        return set()

def kullanilan_satirlari_kaydet(satirlar: set):
    """Kullanılan satır numaralarını dosyaya kaydeder."""
    try:
        with open(KULLANILANLAR_DOSYA_ADI, 'w') as f:
            # Set nesnesini JSON'a kaydetmek için listeye çeviriyoruz
            json.dump(list(satirlar), f)
        # logger.info(f"Kullanılan {len(satirlar)} satır numarası dosyaya kaydedildi.")
    except Exception as e:
        logger.error(f"Kullanılan satırlar kaydedilirken KRİTİK HATA: {e}")


# --- Yardımcı Fonksiyon: Yetki Kontrolü ve Yetkisiz Mesajı (Aynı) ---

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


# --- YARDIMCI EXCEL FONKSİYONU GÜNCELLENDİ ---

def excel_durumu_hesapla():
    """Excel dosyasındaki kullanılan (kaydedilmiş) ve kalan (kaydedilmemiş) satırları hesaplar."""
    if not os.path.exists(EXCEL_DOSYA_ADI):
        return None, "Hata: Excel dosyası bulunamadı."
    
    try:
        workbook = load_workbook(EXCEL_DOSYA_ADI)
        sheet = workbook.active
        
        # YENİ: Kullanılan satır numaralarını hafızadan oku
        kullanilan_satir_numaralari = kullanilan_satirlari_oku()
        
        kullanilan_sayisi = 0
        kalan_sayisi = 0
        baslangic_satiri = 2
        
        # Tüm veri satırlarını döngüye al
        for row_index, row in enumerate(sheet.iter_rows(min_row=baslangic_satiri), start=baslangic_satiri):
            # YENİ KONTROL: Satır, kullanılanlar listesinde mi kontrol et
            if row_index in kullanilan_satir_numaralari:
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
    if yetkili_mi(update):
        await update.message.reply_text(
            f'Merhaba yetkili! Ben göreve hazırım. Mevcut komutlar:\n\n'
            f'  • /ver <miktar>: Veri gönderir ve işaretler.\n'
            f'  • /kalan: Verilmemiş veri sayısını söyler.\n'
            f'  • /rapor: Verilmiş veri sayısını söyler.'
        )

# /kalan komutu güncellenen excel_durumu_hesapla'yı kullanacak.
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

# /rapor komutu güncellenen excel_durumu_hesapla'yı kullanacak.
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

# /ver komutu güncellendi
async def ver_komutu_isleyici(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """
    /ver <miktar> komutunu işler. Veriyi gönderir ve satır numarasını kalıcı dosyaya kaydeder.
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
        # Önce mevcut kullanılan satırları oku (Kalıcılık için)
        mevcut_kullanilanlar = kullanilan_satirlari_oku()
        
        workbook = load_workbook(EXCEL_DOSYA_ADI)
        sheet = workbook.active  
        
        veriler = []
        yeni_kullanilacak_satir_numaralari = [] # Bu oturumda kullanılanlar
        veri_sayisi = 0
        baslangic_satiri = 2 
        
        # Okuma ve Toplama Döngüsü
        for row_index, row in enumerate(sheet.iter_rows(min_row=baslangic_satiri), start=baslangic_satiri):
            
            # **MAVİ KONTROLÜ YERİNE TXT KONTROLÜ**
            if row_index in mevcut_kullanilanlar:
                 continue

            if veri_sayisi >= miktar:
                break
                
            # Veri formatlama kısmı (Aynı)
            satir_verisi_duzenli = []
            hucre_degerleri = [str(cell.value).strip() if cell.value is not None else "" for cell in row]
            
            for etiket_index, etiket in enumerate(VERI_ETIKETLERI):
                if etiket_index < len(hucre_degerleri):
                    deger = hucre_degerleri[etiket_index]
                    satir_verisi_duzenli.append(f"**{etiket}**: {deger}")
            
            veriler.append("\n".join(satir_verisi_duzenli)) 
            
            yeni_kullanilacak_satir_numaralari.append(row_index)
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
        
        # 6. Kullanılan Satırları KAYDET (Boyama Kaldırıldı)
        
        # Yeni kullanılanları mevcut listeye ekle
        mevcut_kullanilanlar.update(yeni_kullanilacak_satir_numaralari)
        
        # Kalıcı dosyaya kaydet
        kullanilan_satirlari_kaydet(mevcut_kullanilanlar)

        # Excel'i kaydetme ve boyama kodları kaldırıldı.
        # workbook.save(EXCEL_DOSYA_ADI)

        await update.message.reply_text(
            f"{veri_sayisi} adet data gruba gönderildi ve çöp kutusuna taşındı."
        )

    except Exception as e:
        logger.error(f"Kritik hata oluştu: {e}")
        await update.message.reply_text(f"❌ Bir Hata Oluştu. Detaylar loglara kaydedildi. Hata: `{e}`")


# --- Ana Fonksiyon (Aynı) ---

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
    application.add_handler(CommandHandler("kalan", kalan_komutu_isleyici))
    application.add_handler(CommandHandler("rapor", rapor_komutu_isleyici))

    # Bot çalışmaya başlar (sürekli güncellemeleri dinler)
    logger.info("Bot başarıyla başlatıldı ve dinlemede...")
    application.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == '__main__':
    main()
