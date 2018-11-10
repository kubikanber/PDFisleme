#! python3, this code written by kubikanber, 09.11.2018, author: kubikanber, in this test files words written Turkish
import os

import AcroExchPDDoc
import test_pdpage

dosya = "ekli.pdf"
dosya_yol = os.path.abspath(dosya)
pdf_dosyası = AcroExchPDDoc.PDDoc(dosya_yol)


def yorumları_al(sayfa):
    yorum_sayısı = sayfa.get_num_annots()
    yorumlar = []

    dj = pdf_dosyası.get_js_object()
    dj.sync_annot_scan()
    for y in range(0, yorum_sayısı):
        yorumlar.append(sayfa.get_annot())
        index = sayfa.get_annot_index(sayfa.get_annot())
        print("{} inddex...".format(index))

    print("{} adet yorum alındı".format(len(yorumlar)))
    return yorumlar


def yorum_bilgileri(sayfalar):
    sayfadaki_yorumlar = [yorumları_al(sayfa)]

    for sy in sayfadaki_yorumlar:
        for ty in sy:
            yorum_index = sayfa.get_annot_index(ty.pdannot)
            print("{} nolu yorum.".format(yorum_index), end=" ")
            yorum_tipi = ty.get_subtype()
            yorum_içeriği = ty.get_contents()
            print("Yorum tipi '{}'".format(yorum_tipi), end=" ")
            print("içeriğinde: {} ".format(yorum_içeriği))


if __name__ == '__main__':
    test_pdpage.bilgi(pdf_dosyası)

    pdf_sayfaları = test_pdpage.sayfaları_al(pdf_dosyası)

    for sayfa in pdf_sayfaları:
        test_pdpage.sayfa_numarası(sayfa)
        test_pdpage.yorum_sayısı(sayfa)
        test_pdpage.sayfa_dönüklüğü(sayfa)
        test_pdpage.sayfa_boyutu(sayfa)
        yorum_bilgileri(sayfa)

    test_pdpage.dosya_kapat(pdf_dosyası)
    print("son")
