import openpyxl

# Excel dosyanızın yolunu belirtin
excel_dosya_yolu0 = input("Dosya ismini yazın: ")
excel_dosya_yolu = excel_dosya_yolu0

# Excel dosyanızı yükleyin
excel_wb = openpyxl.load_workbook(excel_dosya_yolu)

# Çalışma sayfanızı belirleyin
calisma_sayfasi = excel_wb.active

# Verileri depolamak için bir sözlük oluşturun
veri_dict = {}

# Tüm satırları okuyun ve gerekli verileri sözlükte depolayın
for satir in calisma_sayfasi.iter_rows(min_row=2, values_only=True):
    musteri_adi = satir[0]
    if musteri_adi is not None:
        musteri_adi = musteri_adi.lower()
    telefon = satir[1]
    adres = satir[2]
    borc_tutari = satir[3]
    tarih = satir[4]
    
    if musteri_adi not in veri_dict:
        veri_dict[musteri_adi] = []
    veri_dict[musteri_adi].append((borc_tutari, telefon, adres, tarih))

# Kullanıcıdan müşteri adı alın ve sözlükte arayın
while True:
    musteri_adi = input("Müşteri adı (Çıkmak için 'q' tuşuna basın): ").lower()
    if musteri_adi == 'q':
        print("Program sonlandırıldı.")
        break
        
    if musteri_adi in veri_dict:
        borc_tutarlari_toplami = 0
        for borc_tutari, telefon, adres, tarih in veri_dict[musteri_adi]:
            borc_tutarlari_toplami += borc_tutari
            print(f"{musteri_adi} adlı müşterinin borcu {borc_tutari} TL.")
            print(f"Telefon No: {telefon}")
            print(f"Adres: {adres}")
            print(f"Tarih: {tarih}\n")
        print(f"{musteri_adi} adlı müşterinin toplam borcu {borc_tutarlari_toplami} TL.\n")
    else:
        print(f"{musteri_adi} adlı müşterinin borcu bulunamadı.")
