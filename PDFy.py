import os
from win32com.client.dynamic import Dispatch

src = os.path.abspath("ekli.pdf")


def pdf_aç_kapa(dosya):
    app = Dispatch("AcroExch.App")  # adobe pdf programı üzerinde çalışma
    app.Show()  # Show() ve Hide() programı göster veya gösterme
    avdoc = Dispatch("AcroExch.AVDoc")
    pddoc = Dispatch("AcroExch.PDDoc")
    aform = Dispatch("AFormAut.App")

    if avdoc.Open(dosya, ""):
        js = 'f = this.addField("newField", "text", 0, [250,650,20,600]);' + ' f.value = "any Text"'
        # + ' f.flatten'

        annot = 'annot = this.addAnnot(page:0, type:"Text",author: "kubikk", point:[300,400], strokeColor: color.yellow, contents: "vay anasınaa", noteIcon: "Help")'

        aform.Fields.ExecuteThisJavaScript(js)  # java script kodu çalıştır.

    # avdoc = app.GetActiveDoc
    # pageview = kk.GetAVPageView
    # pds = app.Open(dosya, dosya)
    # kk = pds.PDAnnot
    # print(pageview)
    # annot = Dispatch("AcroExch.PDAnnot")
    dokuman_sayi = app.GetNumAvDocs()  # acık olan döküman sayısını veriyor.
    dokuman_dil = app.GetLanguage()
    pddoc.Open(dosya)
    sayfa_sayısı = pddoc.GetNumPages()
    kj = []
    sayfa = pddoc.AcquirePage(0)
    kty = sayfa.GetAnnot(2)  # LPDISPATCH GetAnnot(long nIndex);
    hrty = kty.GetContents()
    jurt = sayfa.GetAnnotIndex(kty)  # long GetAnnotIndex(LPDISPATCH iPDAnnot);

    # avdoc.Close(True)
    # app.CloseAllDocs()  # açık olan tüm pdfleri kapatır.
    # app.Exit()  # programı kapatır. önce tüm pdfler kapatılması gerekiyor.

    print("k: " + str(dokuman_sayi) + dokuman_dil + str(sayfa_sayısı))


def pdf_te(dosya):
    pdf_aç_kapa(src)


pdf_aç_kapa(src)
