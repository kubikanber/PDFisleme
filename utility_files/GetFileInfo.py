#! python3 this code written by kubikanber, 13.11.2018, author: kubikanber
# 1. dosya içinde ki bilgileri alma
import os

import AcroExchPDDoc


def pdf_dosya_bilgisi(pdf_dosyası):
    if pdf_dosyası.open():
        print("Dosya açıldı")

    pdf_dosya_ismi = pdf_dosyası.get_file_name()
    pdf_dosya_durumu = pdf_dosyası.get_flags()
    pdf_dosya_geçiçi_id = pdf_dosyası.get_instance_id()
    pdf_dosya_sayfa_durumu = pdf_dosyası.get_page_mode()
    pdf_dosya_sayfa_sayısı = pdf_dosyası.get_num_pages()
    pdf_dosya_id = pdf_dosyası.get_permanent_id()
    if pdf_dosya_ismi != "":
        print("Dosya ismi: {}".format(pdf_dosya_ismi))
        print("Dosya durumu: {}".format(pdf_dosya_durumu))
        # print("Dosya geçici {} ve gerçek {} kimlik numarası.".format(pdf_dosya_geçiçi_id, pdf_dosya_id))
        print("Dosya içinde {} adet sayfa bulunmaktadır.".format(pdf_dosya_sayfa_sayısı))
        print("Dosyada sayfa gösterim durumu: {}".format(pdf_dosya_sayfa_durumu))
    else:
        print("PDF dosyası bulunamadı.")
    return pdf_dosya_sayfa_sayısı


def pdf_dosya_kapat(pdf_dosyası, yol=None):
    pdf_dosya_durumu = pdf_dosyası.get_flags()
    if pdf_dosya_durumu == 8320:
        if pdf_dosyası.close():
            print("Dosya kapandı")
    elif pdf_dosya_durumu == 8325:
        if pdf_dosyası.save(0):
            print("Dosya kayıt edildi.")
    elif pdf_dosya_durumu == 15:
        if pdf_dosyası.save(1, yol):
            print("Dosya {} kayıt edildi.".format(yol))
    else:
        print("Yeni bir durum oluştu.... Dosya kapatıldı.")
        pdf_dosyası.close()
    del pdf_dosyası


def sayfaları_al(pdf_dosyası):
    pdf_sayfaları = []
    for sayfa in range(0, pdf_dosyası.get_num_pages()):
        pdf_sayfaları.append(pdf_dosyası.acquire_page(sayfa))
    print("{} adet sayfa alındı.".format(len(pdf_sayfaları)))
    return pdf_sayfaları


def pdf_sayfa_bilgisi(pdf_sayfası):
    for sayfa in pdf_sayfası:
        sayfa_no = sayfa.get_number()
        yorum_sayısı = sayfa.get_num_annots()
        sayfa_dön = sayfa.get_rotate()
        boyut = sayfa.get_size()
        x_noktası = boyut.x() / ((72 / 2.54) / 10)  # mm
        y_noktası = boyut.y() / ((72 / 2.54) / 10)  # mm
        print("{} nolu sayfada.".format(sayfa_no), end=" ")
        print("{} adet yorum parçası bulunmaktadır.".format(yorum_sayısı), end=" ")
        print("{} derecede gösterilmektedir.".format(sayfa_dön), end=" ")
        print("Sayfa {:0.0f} x {:0.0f} boyutlarındadır.".format(x_noktası, y_noktası))

        yorumlar = yorumları_al(pdf_dosyası, sayfa)

        for yorum in yorumlar:
            yorum_index = sayfa.get_annot_index(yorum.pdannot)

            yorum_tipi = yorum.get_subtype()
            yorum_içeriği = yorum.get_contents()
            yorum_rengi = yorum.get_color()
            yorum_tarihi = yorum.get_date()
            yorum_yeri = yorum.get_rect()
            yorum_başlığı = yorum.get_title()
            yorum_geçerlimi = yorum.is_valid()

            print("{:10} nolu yorum.".format(yorum_index), end=" ")
            print("Yorum tipi '{}'".format(yorum_tipi), end=" ")
            print("içeriğinde: {} ".format(yorum_içeriği), end=" ")
            print("Başlığı: {}".format(yorum_başlığı), end=" ")
            print("Rengi: {}".format(yorum_rengi), end=" ")
            print("Yorum tarihi: {}:{}:{} - {}:{}:{}:{}".format(yorum_tarihi.date(), yorum_tarihi.month(),
                                                                yorum_tarihi.year(), yorum_tarihi.hour(),
                                                                yorum_tarihi.minute(), yorum_tarihi.second(),
                                                                yorum_tarihi.millisecond()), end=" ")
            print("Yeri: {} x {} x {} x {}".format(yorum_yeri.bottom(), yorum_yeri.left(), yorum_yeri.right(),
                                                   yorum_yeri.top()), end=" ")
            print(":{}".format(yorum_geçerlimi))


def yorumları_al(pdf_dosyası, sayfa):
    yorum_sayısı = sayfa.get_num_annots()
    yorumlar = []

    dj = pdf_dosyası.get_js_object()
    dj.sync_annot_scan()
    for y in range(0, yorum_sayısı):
        yorum = sayfa.get_annot(y)
        yorumlar.append(yorum)
        index = sayfa.get_annot_index(yorum.pdannot)
        # print("{} nolu index...".format(index))

    print("{:3} adet yorum alındı.".format(len(yorumlar)))
    return yorumlar


if __name__ == '__main__':
    dosya = "..\\test\\ekli.pdf"
    dosya_yol = os.path.abspath(dosya)
    pdf_dosyası = AcroExchPDDoc.PDDoc(dosya_yol)
    pdf_dosya_bilgisi(pdf_dosyası)
    pdf_sayfaları = sayfaları_al(pdf_dosyası)
    pdf_sayfa_bilgisi(pdf_sayfaları)
    pdf_dosya_kapat(pdf_dosyası)
