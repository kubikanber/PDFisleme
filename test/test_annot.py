#! python3, this code written by kubikanber, 09.11.2018, author: kubikanber, in this test files words written Turkish
import os
from copy import copy

import AcroExchPDDoc
import AcroExchRect
from test import test_pdpage

dosya = "440-LS-008-14-58.pdf"
dosya_yol = os.path.abspath(dosya)
pdf_dosyası = AcroExchPDDoc.PDDoc(dosya_yol)


def yorumları_al(sayfa):
    yorum_sayısı = sayfa.get_num_annots()
    yorumlar = []

    dj = pdf_dosyası.get_js_object()
    dj.sync_annot_scan()
    for y in range(0, yorum_sayısı):
        yorum = sayfa.get_annot(y)
        yorumlar.append(yorum)
        index = sayfa.get_annot_index(yorum.pdannot)
        # print("{} nolu index...".format(index))

    print("{:3} adet yorum alındı".format(len(yorumlar)))
    return yorumlar


def yorum_bilgi(ty):
    yorum_index = sayfa.get_annot_index(ty.pdannot)

    yorum_tipi = ty.get_subtype()
    yorum_içeriği = ty.get_contents()
    yorum_rengi = ty.get_color()
    yorum_tarihi = ty.get_date()
    yorum_yeri = ty.get_rect()
    yorum_başlığı = ty.get_title()
    yorum_geçerlimi = ty.is_valid()

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


def yorum_bilgileri(sayfalar=None):
    if sayfalar is not None:
        sayfadaki_yorumlar = yorumları_al(sayfalar)
        for syy in sayfadaki_yorumlar:
            yorum_bilgi(syy)
    else:
        sayfadaki_yorumlar = [yorumları_al(sayfa)]

        for sy in sayfadaki_yorumlar:
            for ty in sy:
                yorum_bilgi(ty)
    return sayfadaki_yorumlar


def yorum_kopyala(yorum):
    yorum_objesi = yorum[0][0]
    k_top = yorum_objesi.get_rect().top()
    k_right = yorum_objesi.get_rect().right()
    k_left = yorum_objesi.get_rect().left()
    k_bottom = yorum_objesi.get_rect().bottom()
    yorum_ek = copy(yorum_objesi)
    yorum_ek.set_rect(k_top + 100, k_right + 100, k_left + 100, k_bottom + 100)
    kk = yorum_ek.get_rect().top()
    return yorum_ek


def yorum_oluştur(sayfa):
    sayfa_boyutu = sayfa.get_size()
    sayfa_dönüklüğü = sayfa.get_rotate()
    kkk = sayfa_boyutu.x()
    kyyyy = sayfa_boyutu.y()
    acrorect = AcroExchRect.Rect()
    acrorect.create_rect(kyyyy - 10, (kkk / 2) - 60, (kkk / 2) + 60, (kyyyy - 10) - 10)
    if sayfa.add_new_annot(0, "FreeText", acrorect.rect):
        print("Yorum eklendi.")


if __name__ == '__main__':
    test_pdpage.bilgi(pdf_dosyası)

    pdf_sayfaları = test_pdpage.sayfaları_al(pdf_dosyası)

    for sayfa in pdf_sayfaları:
        test_pdpage.sayfa_numarası(sayfa)
        test_pdpage.yorum_sayısı(sayfa)
        test_pdpage.sayfa_dönüklüğü(sayfa)
        test_pdpage.sayfa_boyutu(sayfa)
        yorum_bilgileri()

    seçilen_sayfa = pdf_sayfaları[0]
    sayfa_yorum = yorum_bilgileri(seçilen_sayfa)
    sonuç = sayfa_yorum[0].is_equal(sayfa_yorum[0].pdannot)
    # sayfa_yorum[0][0].set_open(1)
    açıkmı = sayfa_yorum[0].is_open()
    print("yorum: {}, açık mı ? {}".format(sonuç, açıkmı))
    # yeni_yorum = yorum_kopyala(sayfa_yorum)
    # pdf_sayfaları[5].add_annot(1, yeni_yorum.pdannot)
    yorum_oluştur(seçilen_sayfa)
    test_pdpage.dosya_kapat(pdf_dosyası)
    print("son")

