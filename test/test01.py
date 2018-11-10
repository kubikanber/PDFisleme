#! python3
import os

import winerror

import AcroExchAnnotJava
import AcroExchPDDoc

# from win32com.client.dynamic import ERRORS_BAD_CONTEXT
#
# # “Not implemented” Exception when using pywin32 to control Adobe Acrobat
# # https://stackoverflow.com/questions/9383307/not-implemented-exception-when-using-pywin32-to-control-adobe-acrobat
# global ERRORS_BAD_CONTEXT  ###############################
# ERRORS_BAD_CONTEXT.append(winerror.E_NOTIMPL)  ############################

mesaj00 = """1. PDF dosya/sayfa bilgileri.
        """
mesaj01 = """1. PDF dosyası ve sayfa bilgileri.
2. PDF dosyasından seçilen sayfa bilgileri.
        """


class PDFdosya:

    def __init__(self, dosya_ismi):

        self.pdf_dosya = AcroExchPDDoc.PDDoc(dosya_ismi)

    def pdf_dosya_aç(self):
        pdf_dosya_ismi = ""
        if self.pdf_dosya.open():
            pdf_dosya_ismi = self.pdf_dosya.get_file_name()
            print("{} Dosyası başarılı bir şekilde açıldı.".format(pdf_dosya_ismi))
        else:
            print("Dosya açılamadı".format())

        return pdf_dosya_ismi

    def pdf_kapat(self):
        if self.pdf_dosya.close():
            print("Dosya kapandı.")
        else:
            print("Dosya kapanırken bir hata oluştu.")

    def pdf_dosya_durumu(self):
        durum = self.pdf_dosya.get_flags()
        if durum == 8320:
            print("Dosyada herhangi bir değişiklik yok")
        elif durum == 8325:
            print("Dosyada bir değiklik olmuş.")

    def pdf_kimlik_idleri(self):
        örnek_id = self.pdf_dosya.get_instance_id()
        kalıcı_id = self.pdf_dosya.get_permanent_id()
        print("Dosyanın örnekleme kimliği {} , kalıcı kimliği {}".format(örnek_id, kalıcı_id))

    def pdf_java_varmı(self):
        pdf_java = self.pdf_dosya.get_js_object()
        if pdf_java is None:
            print("Java objesi bulunamadı.")
        else:
            print("Java objesi bulundu...")
        return pdf_java

    def pdf_sayfaları(self):
        sayfa_sayısı = self.pdf_dosya.get_num_pages()
        sayfa_gösterim_durumu = self.pdf_dosya.get_page_mode()
        print("Dosyada {} adet sayfa bulunmaktadır.".format(sayfa_sayısı))
        if sayfa_gösterim_durumu == 0:
            print("Sayfalar olduğu gibi gösteriliyor.")
        elif sayfa_gösterim_durumu == 1:
            print("Sayfalar yer imsiz ve küçük resimsiz gösteriliyor.")
        elif sayfa_gösterim_durumu == 2:
            print("Sayfala sadece küçük resimler ile birlikte gösteriliyor.")
        elif sayfa_gösterim_durumu == 3:
            print("Sayfalar sadece yer imleri ile birlikte gösteriliyor.")
        elif sayfa_gösterim_durumu == 4:
            print("Sayfalar tam ekran modunda gösteriliyor.")
        else:
            print("Sayfa gösterim durumu bilinmiyor...")
        return sayfa_sayısı, sayfa_gösterim_durumu

    def pdf_sayfa_al(self, sayfa_no):
        return self.pdf_dosya.acquire_page(sayfa_no)

    def pdf_sayfa_boyutu(self, sayfa):
        pdf_ölçü = sayfa.get_size()  # sayfa ölçülerini almak için
        ölçü_x = pdf_ölçü.x()  # Acro point[Adobe] 1" 72 point
        ölçü_y = pdf_ölçü.y()  # Acro point[Adobe] 1" 72 point
        x = ölçü_x / ((72 / 2.54) / 10)  # mm
        y = ölçü_y / ((72 / 2.54) / 10)  # mm
        print("Sayfa {:0.1f} x {:0.1f} ölçüsündedir.".format(x, y))


def işlem_girdisi(mesaj=""):
    while True:
        print(mesaj)
        işlem_no = input("işlem no:")
        if işlem_no.isnumeric():
            işlem_no = int(işlem_no)
            break
        else:
            print("Lütfen işlem numarası giriniz: ")
            print("#" * 30)
    return işlem_no


def işlem():
    işlem = işlem_girdisi(mesaj00)
    if işlem == 1:
        alt_işlem = işlem_girdisi(mesaj01)
    return str(işlem) + str(alt_işlem)


def dosya_girdisi():
    PDF_dosyası_ismi = input("PDF dosya ismini giriniz.: ")
    if PDF_dosyası_ismi is "":
        PDF_dosyası_ismi = "ekli.pdf"
    return PDF_dosyası_ismi


def işlemler(işlem_no):
    if int(işlem_no) == 11:
        dosya = dosya_girdisi()
        dosya_yolile = os.path.abspath(dosya)
        pdfdokumanı = PDFdosya(dosya_yolile)
        print("1. Dosyanın belgeden alınan ismi: {}".format(pdfdokumanı.pdf_dosya_aç()))
        print("Dosya durumu: ", end="")
        pdfdokumanı.pdf_dosya_durumu()
        pdfdokumanı.pdf_kimlik_idleri()
        pdfdokumanı.pdf_java_varmı()
        pdfdokumanı.pdf_sayfaları()
        pdfdokumanı.pdf_kapat()

    if int(işlem_no) == 12:
        i = 0
        yorum_sayısı = 0
        dosya = dosya_girdisi()
        dosya_yolile = os.path.abspath(dosya)

        pdfdokumanı = PDFdosya(dosya_yolile)
        pdfdokumanı.pdf_dosya_aç()
        sayfa_sayısı, sayfa_durumu = pdfdokumanı.pdf_sayfaları()
        pdf_jave = pdfdokumanı.pdf_dosya.get_js_object()
        pdf_jave.sync_annot_scan()
        j_annot = pdf_jave.get_annots()
        j_rot = pdf_jave.get_page_rotation()
        for ji in j_annot:
            ja = AcroExchAnnotJava.AnnotJava(ji)
            print("J_ tipi: {} ".format(ja.get_type()), end=" ")
            print("sayfa: {:0.0f}".format(ja.get_set_page() + 1), end=" ")
            print("J_ içindeki: {}".format(ja.get_set_contents()))
            # ja.annotjava.contents = "değişim01"
        for i in range(0, sayfa_sayısı):
            sayfa = pdfdokumanı.pdf_sayfa_al(i)
            print("{} sayfada ".format(sayfa.get_number()), end=" ")
            yorum_sayısı = sayfa.get_num_annots()
            print(" {} adet yorum bulunmaktadır.".format(yorum_sayısı), end=" ")
            print(" Sayfa {} derece dönüktür.".format(sayfa.get_rotate()), end=" ")
            pdfdokumanı.pdf_sayfa_boyutu(sayfa)

            for y in range(0, yorum_sayısı):
                # input("bas")
                 # pdf_jave.sync_annot_scan()
                yorum = sayfa.get_annot(y)
                yorum_index = sayfa.get_annot_index(yorum.pdannot)
                # while True:
                #
                #     yorum_index = sayfa.get_annot_index(yorum.pdannot)
                #     if yorum_index == y:
                #         break
                print("{} nolu yorum.".format(yorum_index), end=" ")
                print("Yorum tipi '{}'. ".format(yorum.get_subtype()), end=" ")
                print("İçerik: {}".format(yorum.get_contents()))

        while True:
            sayfa_no = input("Hangi sayfayı almak istiyorsunuz? ")
            if sayfa_no == "":
                sayfa_no = 1
                break
            elif int(sayfa_no) > sayfa_sayısı:
                print("Dosyada olmayan bir sayfa seçtiniz. Lütfen tekrar deneyin...")
            elif int(sayfa_no) <= sayfa_sayısı:
                break

        print(sayfa_no)
        pdfdokumanı.pdf_kapat()
        del pdfdokumanı


if __name__ == "__main__":
    print("Dosya direct çalıştı.")
    print("#" * 30)
    print("# PDF işleme")
    print("#" * 30)
    işlem_no = işlem()
    işlemler(işlem_no)

    # dosya_ismi = input("PDF dosya ismini giriniz: ")
    # pdf_dosya = os.path.abspath(dosya_ismi)
    # pdf01 = PDFdosya(pdf_dosya)
    # print("Açılan {} içinde".format(pdf01.pdf_dosya_ismi()))
    #
    # pdf01.pdf_kapat()
    print("#son#")
