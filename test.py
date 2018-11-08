#! python3
import os

import AcroExchPDDoc

dosya = "ekli.pdf"
pdf_dosyası = os.path.abspath(dosya)
kk = AcroExchPDDoc.PDDoc(pdf_dosyası)
if kk.open():
    print("'{}' dosyası açıldı.".format(kk.get_file_name()))  # dosya ismini almak için

    sayfa_sayısı = kk.get_num_pages()  # dosya içinde ki sayfa sayısnı öğrenmek için
    print("Dosyada {} adet sayfa bulunmaktadır.".format(sayfa_sayısı))

    print("dosya {}".format(kk.get_flags()))  # dosyanın son hali hakkında bilgi.

    sayfa_no = input("Hangi sayfayı almak istiyorsunuz? ")
    if sayfa_no == "":
        sayfa_no = 1
        # int(sayfa_no)
    elif int(sayfa_no) > sayfa_sayısı:
        print("Dosyada olmayan bir sayfa seçtiniz. Lütfen tekrar deneyin...")
        exit(1)

    # sayfa_no = int(sayfa_no)
    sayfa = kk.acquire_page(int(sayfa_no) - 1)  # belirtilen sayfayaı almak için
    print("sayfa numarası: " + str(sayfa.get_number()))  # belirtilen sayfanın gerçek sayfa numarasını almak için

    sayfadaki_yorum_sayısı = sayfa.get_num_annots()
    print("yorum sayısı: " + str(sayfadaki_yorum_sayısı))  # sayfa içindeki yorumların sayısnı öğrenmek için

    yorumlar = sayfa.get_annot(2)  # burada tüm yorumlar bir değişkene toplabilir.
    print("yorum index nosu {}".format(
        sayfa.get_annot_index(yorumlar.pdannot)))  # alınan yorumun index nosunu öğrenmek için

    print("sayfa yönü {} ".format(sayfa.get_rotate()))  # sayfanın yönünü öğrenmek için

    pdf_ölçü = sayfa.get_size()  # sayfa ölçülerini almak için
    ölçü_x = pdf_ölçü.X  # Acro point[Adobe] 1" 72 point
    ölçü_y = pdf_ölçü.Y  # Acro point[Adobe] 1" 72 point
    x = ölçü_x / ((72 / 2.54) / 10)  # mm
    y = ölçü_y / ((72 / 2.54) / 10)  # mm
    print("Sayfa {:0.1f} x {:0.1f} ölçüsünde".format(x, y))
else:
    print("Dosya açılamadı.")

print("dosya {}".format(kk.get_flags()))

if kk.close():
    print("Dosya kapatıldı.")
else:
    print("Dosya kapatılırken bir sorunla karşılaştı...")

