#! python3
import os

import AcroExchPDDoc
import AcroExchRect

dosya = ".\\test\\ekli.pdf"
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


def sayfaları_al(pdf_dosyası=pdf_dosyası):
    pdf_sayfaları = []
    for sayfa in range(0, pdf_dosyası.get_num_pages()):
        pdf_sayfaları.append(pdf_dosyası.acquire_page(sayfa))
    print("{} adet sayfa alındı.".format(len(pdf_sayfaları)))
    return pdf_sayfaları


def sayfa_numarası(sayfa):
    sayfa_no = sayfa.get_number()
    print("{} nolu sayfada.".format(sayfa_no), end=" ")


def yorum_sayısı(sayfa):
    yorum_sayısı = sayfa.get_num_annots()
    print("{} adet yorum parçası bulunmaktadır.".format(yorum_sayısı), end=" ")


def sayfa_dönüklüğü(sayfa):
    sayfa_dön = sayfa.get_rotate()
    print("{} derecede gösterilmektedir.".format(sayfa_dön), end=" ")


def arka_plana_kopyala(sayfa):
    sınır = AcroExchRect.Rect()
    sınır.create_rect(0, 0, 400, 400)
    if sayfa.copy_to_clipboard(sınır.rect, 1, 1, 100):
        print("Seçilen yer kopyalandı.")


def sayfa_boyutu(sayfa):
    boyut = sayfa.get_size()
    x_noktası = boyut.x() / ((72 / 2.54) / 10)  # mm
    y_noktası = boyut.y() / ((72 / 2.54) / 10)  # mm
    print("Sayfa {:0.0f} x {:0.0f} boyutlarındadır.".format(x_noktası, y_noktası))


def sayfayı_kırp(sayfa):
    sınır = AcroExchRect.Rect()
    sınır.create_rect(0, 0, 400, 400)
    if sayfa.crop_page(sınır.rect):
        print("Sayfa kırpıldı.")


def yorum_sil(sayfa, yorum_index_numarası):
    if sayfa.remove_annot(yorum_index_numarası):
        print("yapıldı.")


def sayfayı_döndür(sayfa, döndür):
    if sayfa.set_rotate(döndür):
        print("{} derece sayfa döndürüldü.".format(döndür))


if __name__ == '__main__':
    bilgi()
    pdf_sayfaları = sayfaları_al()
    for sayfa in pdf_sayfaları:
        sayfa_numarası(sayfa)
        yorum_sayısı(sayfa)
        sayfa_dönüklüğü(sayfa)
        sayfa_boyutu(sayfa)

    # arka_plana_kopyala(pdf_sayfaları[1])
    # sayfayı_kırp(pdf_sayfaları[0])
    # yorum_sil(pdf_sayfaları[0], 0)
    # yorum_sayısı(pdf_sayfaları[0])
    # sayfayı_döndür(pdf_sayfaları[0], 0)

    dosya_kapat()
