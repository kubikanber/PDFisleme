#! python3 this code written by kubikanber, 13.11.2018, author: kubikanber
# 1. delete all annot parts
import os

import AcroExchPDDoc
from utility_files import GetFileInfo


def tüm_yorumları_sil(pdf_dosyası):
    sayfa_sayısı = GetFileInfo.pdf_dosya_bilgisi(pdf_dosyası)
    print("{} sayfa".format(sayfa_sayısı))
    pdf_sayfaları = GetFileInfo.sayfaları_al(pdf_dosyası)
    for sayfa in pdf_sayfaları:
        sayfa_no = sayfa.get_number()
        yorum_sayısı = sayfa.get_num_annots()
        print("{} nolu sayfada.".format(sayfa_no), end=" ")
        print("{} adet yorum parçası bulunmaktadır.".format(yorum_sayısı))
        yorumlar = GetFileInfo.yorumları_al(pdf_dosyası, sayfa)
        yorum_sil(sayfa, yorumlar)


def yorum_sil(sayfa, yorumlar):
    for ys in yorumlar:
        sayfa.remove_annot(sayfa.get_annot_index(ys.pdannot))


if __name__ == '__main__':
    dosya = "tamam.pdf"
    dosya_yol = os.path.abspath(dosya)
    pdf_dosyası = AcroExchPDDoc.PDDoc(dosya_yol)
    tüm_yorumları_sil(pdf_dosyası)
    GetFileInfo.pdf_dosya_kapat(pdf_dosyası)
