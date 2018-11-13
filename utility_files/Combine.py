#! python3 this code written by kubikanber, 13.11.2018, author: kubikanber
# 1. read files in pages folder.
# 2. combine to one file.
import os

import AcroExchPDDoc

klasör = ".\pages"
hedef_dosya = "C:\\Users\\kubikanber\\Desktop\\PYT\\PDFisleme\\utility_files\\tamam.pdf"


def klasör_okuma(klasör):
    print("Okunan klasör: {}".format(klasör))
    klasörler = [k for k in os.listdir(klasör) if os.path.isdir(os.path.join(klasör, k))]
    # print("Okundu.")
    return klasörler


def dosyaları_okuma(dosya_yol):
    print("Dosyalar okunuyor")
    dosyalar = [d for d in os.listdir(dosya_yol) if
                (d.endswith(".pdf") or d.endswith(".PDF")) and not d.startswith("~$")]
    print("{} adet dosya okundu.".format(len(dosyalar)))
    dosyalar.sort(reverse=True)
    return dosyalar


def pdf_sayfa_ekle(tek_dosya, dosyalar, sayfa_sayısı):
    eklenecek_dosya = dosyalar
    ek_sayfalar = sayfa_sayısı
    if tek_dosya.insert_pages(-1, eklenecek_dosya, 0, ek_sayfalar):
        print("Sayfa eklendi.")


def dosyaları_birleştir(dosyalar, tek_dosya):
    for dosya in dosyalar:
        dosya_yol = os.path.abspath(klasör + "\\" + dosya)
        pdf_dosyası = AcroExchPDDoc.PDDoc(dosya_yol)
        sayfa_sayısı = pdf_dosya_bilgisi(pdf_dosyası)
        pdf_dosya_kapat(pdf_dosyası)
        pdf_sayfa_ekle(tek_dosya, dosya_yol, sayfa_sayısı)


def yeni_pdf_dosyası_oluştur(dosya_ismi):
    dosya_yol = os.path.abspath("tmp")
    pdf_temp = AcroExchPDDoc.PDDoc("tmp")
    if pdf_temp.create():
        print("yeni bir döküman oluşturuldu.")
    return pdf_temp


def var_olan_pdf(dosya_ismi):
    dosya_yol = os.path.abspath(klasör + "\\" + dosya_ismi)
    pdf_dosya = AcroExchPDDoc.PDDoc(dosya_yol)
    return pdf_dosya


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
        if pdf_dosyası.save(1,yol):
            print("Dosya {} kayıt edildi.".format(yol))
    else:
        print("Yeni bir durum oluştu.... Dosya kapatıldı.")
        pdf_dosyası.close()
    del pdf_dosyası


if __name__ == '__main__':
    print("Dosya birleştirme.")
    klasör_okuma(klasör)
    dosyalar = dosyaları_okuma(klasör)
    yeni_pdf = yeni_pdf_dosyası_oluştur("temp.pdf")
    pdf_dosya_bilgisi(yeni_pdf)

    dosyaları_birleştir(dosyalar, yeni_pdf)
    pdf_dosya_bilgisi(yeni_pdf)
    pdf_dosya_kapat(yeni_pdf, hedef_dosya)
