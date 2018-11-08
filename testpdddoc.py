#! python3
import os

import AcroExchPDDoc
import AcroExchRect

dosya = "ekli.pdf"
dosya_yol = os.path.abspath(dosya)
pdf_dosyası = AcroExchPDDoc.PDDoc(dosya_yol)


def bilgi(pdf_dosyası=pdf_dosyası):
    if pdf_dosyası.open():
        print("Dosya açıldı")

    pdf_dosya_ismi = pdf_dosyası.get_file_name()
    pdf_dosya_durumu = pdf_dosyası.get_flags()
    pdf_dosya_geçiçi_id = pdf_dosyası.get_instance_id()
    pdf_dosya_sayfa_durumu = pdf_dosyası.get_page_mode()
    pdf_dosya_sayfa_sayısı = pdf_dosyası.get_num_pages()
    pdf_dosya_id = pdf_dosyası.get_permanent_id()

    print("Dosya ismi: {}".format(pdf_dosya_ismi))
    print("Dosya durumu: {}".format(pdf_dosya_durumu))
    print("Dosya geçici {} ve gerçek {} kimlik numarası.".format(pdf_dosya_geçiçi_id, pdf_dosya_id))
    print("Dosya içinde {} adet sayfa bulunmaktadır.".format(pdf_dosya_sayfa_sayısı))
    print("Dosyada sayfa gösterim durumu: {}".format(pdf_dosya_sayfa_durumu))


def arayüzü_aç(pdf_dosyası=pdf_dosyası):
    sonuç = pdf_dosyası.open_avdoc()
    if sonuç is None:
        print("Arayüz açılamadı.")
    return sonuç


def arayüz():
    arayüzü_aç()
    arayüz_dosyası = arayüzü_aç()
    if arayüz_dosyası.close_doc(1):
        print("Arayüz kapandı.")


def dosya_kapat(pdf_dosyası=pdf_dosyası, yol=None):
    pdf_dosya_durumu = pdf_dosyası.get_flags()
    if pdf_dosya_durumu == 8320:
        if pdf_dosyası.close():
            print("Dosya kapandı")
    elif pdf_dosya_durumu == 8325:
        if pdf_dosyası.save(0):
            print("Dosya kayıt edildi.")
    else:
        print("Yeni bir durum oluştu.... Dosya kapatıldı.")
        pdf_dosyası.close()
    del pdf_dosyası


def dosya_durumu_ayarla(pdf_dosyası=pdf_dosyası):
    if pdf_dosyası.clear_flags():
        print("Dosya durumu temizlendi.")
    if pdf_dosyası.set_flags(8325):
        print("Dosya durumu değişti.")


def dosya_infosu_al(pdf_dosyası=pdf_dosyası):
    info = pdf_dosyası.get_info(0)  # 0 için bir değer var "kuubbkk"
    info_title = pdf_dosyası.get_info("Title")
    info_yazar = pdf_dosyası.get_info("Author")
    info_konu = pdf_dosyası.get_info("Subject")
    info_anahtar_kelime = pdf_dosyası.get_info("Keywords")
    info_oluşturan = pdf_dosyası.get_info("Creator")
    info_sağlayıcı = pdf_dosyası.get_info("Producer")
    info_oluşturma_tarihi = pdf_dosyası.get_info("CreationDate")
    güncelleme_tarihi = pdf_dosyası.get_info("ModDate")
    info_t = pdf_dosyası.get_info("Trapped")
    if info:
        print("Dosya infosu alındı.{}-{}-{}-{}-{}-{}-{}-{}-{}-{}-".format(info, info_title, info_yazar, info_konu,
                                                                               info_anahtar_kelime,
                                                                               info_oluşturan, info_sağlayıcı,
                                                                               info_oluşturma_tarihi,
                                                                               güncelleme_tarihi, info_t))


def dosya_infosu_ekle(pdf_dosyası=pdf_dosyası):
    if pdf_dosyası.set_info(1, "birinçi veri"):
        print("Dosya info eklendi.")


def dosya_sayfa_modu_değiştir(pdf_dosyası=pdf_dosyası):
    if pdf_dosyası.set_page_mode(0):
        print("Dosya sayfa gösterim modu değişti.")


def yeni_döküman_oluştur():
    temp_dosya_yol = "C:\\Users\\kubikanber\\Desktop\\PYT\\PDFisleme\\pdf_temp.pdf"
    pdf_temp = AcroExchPDDoc.PDDoc(temp_dosya_yol)
    if pdf_temp.create():
        print("yeni bir döküman oluşturuldu.")
    return pdf_temp


def sayfaları_kırp(pdf_dosyası=pdf_dosyası, başlangıç_sayfası=0, bitiş_sayfası=0, tek_veya_çift=0, kırpa_bölgesi=None):
    kırpa_bölgesi = AcroExchRect.Rect()
    kırpa_bölgesi.create_rect(1, 1, 100, 110)
    if pdf_dosyası.crop_pages(başlangıç_sayfası, bitiş_sayfası, tek_veya_çift, kırpa_bölgesi.rect):
        print("Dosya kırpıldı.")


def küçük_resimleri_yeniden_yükle(pdf_dosyası=pdf_dosyası):
    if pdf_dosyası.delete_thumbs(0, pdf_dosyası.get_num_pages() - 1):
        print("Küçük resimler silindi.")
    if pdf_dosyası.create_thumbs(0, pdf_dosyası.get_num_pages() - 1):
        print("Küçük resimler oluşturuldu.")


def sayfa_sil(pdf_dosyası=pdf_dosyası):
    if pdf_dosyası.delete_pages(1, 1):
        print("Sayfa silindi.")


def sayfa_ekle(pdf_dosyası=pdf_dosyası):
    eklenecek_dosya = "C:\\Users\\kubikanber\\Desktop\\PYT\\PDFisleme\\440-LS-008-14-58.pdf"
    if pdf_dosyası.insert_pages(0, eklenecek_dosya, 0, 1):
        print("Sayfa eklendi.")


def sayfa_yeri_değiştir(pdf_dosyası=pdf_dosyası):
    if pdf_dosyası.move_page(1, 0):
        print("Sayfa yerdeğiştirdi")


def sayfa_değiştir(pdf_dosyası=pdf_dosyası):
    eklenecek_dosya = "C:\\Users\\kubikanber\\Desktop\\PYT\\PDFisleme\\ss.pdf"
    if pdf_dosyası.replace_pages(1, eklenecek_dosya, 0, 1, 0):
        print("Dosyada sayfa değişti.")


if __name__ == '__main__':
    bilgi()
    # arayüz()  # arayüz denemesi için arayüz_içinde örneklemeler yapılacak.
    # dosya_durumu_ayarla()
    # bilgi()
    dosya_infosu_al()
    # dosya_infosu_ekle()
    # dosya_sayfa_modu_değiştir()
    # küçük_resimleri_yeniden_yükle()
    # sayfa_sil()
    # sayfa_ekle()
    # sayfa_yeri_değiştir()
    # sayfa_değiştir()  # sayfa değiştir projede boyanan sayfalar için kullanılmalı.

    # yeni_döküman = yeni_döküman_oluştur()
    # bilgi(yeni_döküman)

    # sayfaları_kırp()
    # bilgi()

    # dosya_kapat(yeni_döküman, "C:\\Users\\kubikanber\\Desktop\\PYT\\PDFisleme\\pdf_yeni.pdf")
    dosya_kapat()
print("son")
