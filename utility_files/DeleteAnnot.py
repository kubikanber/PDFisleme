#! python3 this code written by kubikanber, 13.11.2018, author: kubikanber
# 1. delete all annot parts
import os

import AcroExchPDDoc
from utility_files import GetFileInfo


def tüm_yorumları_sil(pdf_dosyası):
    sayfa_sayısı = GetFileInfo.pdf_dosya_bilgisi(pdf_dosyası)
    print("{} sayfa".format(sayfa_sayısı))


if __name__ == '__main__':
    dosya = "tamam.pdf"
    dosya_yol = os.path.abspath(dosya)
    pdf_dosyası = AcroExchPDDoc.PDDoc(dosya_yol)
    tüm_yorumları_sil(pdf_dosyası)
    GetFileInfo.pdf_dosya_kapat(pdf_dosyası)
