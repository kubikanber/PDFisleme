import os
import time
import py4j
import winerror
from py4j.java_gateway import JavaGateway

from win32com.client.dynamic import Dispatch
from win32com.client.dynamic import ERRORS_BAD_CONTEXT

dosya_yol = os.path.abspath("ekli.pdf")

# app = Dispatch("AcroExch.App")
# avdoc = Dispatch("AcroExch.AVDoc")

# AcroExch.PDDoc
# The underlying PDF representation of a document. This is a creatable interface. There is a correspondence between
# a PDDoc object and an ASFile object (an opaque representation of an open file made available through an interface
#  encapsulating Acrobat’s access to file services), and the PDDoc object is the hidden object behind every AVDoc
# object. An ASFile object may have zero or more underlying files, so a PDF file does not always correspond to a single
#  disk file. For example, an ASFile object may provide access to PDF data in a database.
#
# Through PDDoc objects, your application can perform most of the Document menu items from Acrobat (delete pages,
# replace pages, and so on), create and delete thumbnails, and set and retrieve document information fields.
#
# Methods
# The PDDoc object has the following methods.
# The PDDoc object has the following methods.
#
# Method
#
# Description
#
# AcquirePage
#
# Acquires the specified page.
#
# ClearFlags
#
# Clears a document’s flags.
#
# Close
#
# Closes a file.
#
# Create
#
# Creates a new AcroExch.PDDoc.
#
# CreateTextSelect
#
# Creates a text selection from the specified rectangle on the specified page.
#
# CreateThumbs
#
# Creates thumbnail images for the specified page range in a document.
#
# CropPages
#
# Crops the pages in a specified range in a document.
#
# DeletePages
#
# Deletes pages from a file.
#
# DeleteThumbs
#
# Deletes thumbnail images from the specified pages in a document.
#
# GetFileName
#
# Gets the name of the file associated with this AcroExch.PDDoc.
#
# GetFlags
#
# Gets a document’s flags.
#
# GetInfo
#
# Gets the value of a specified key in the document’s Info dictionary.
#
# GetInstanceID
#
# Gets the instance ID (the second element) from the ID array in the document’s trailer.
#
# GetJSObject
#
# Gets a dual interface to the JavaScript object associated with the PDDoc.
#
# GetNumPages
#
# Gets the number of pages in a file.
#
# GetPageMode
#
# Gets a value indicating whether the Acrobat application is currently displaying only pages, pages and thumbnails, or pages and bookmarks.
#
# GetPermanentID
#
# Gets the permanent ID (the first element) from the ID array in the document’s trailer.
#
# InsertPages
#
# Inserts the specified pages from the source document after the indicated page within the current document.
#
# MovePage
#
# Moves a page to another location within the same document.
#
# Open
#
# Opens a file.
#
# OpenAVDoc
#
# Opens a window and displays the document in it.
#
# ReplacePages
#
# Replaces the indicated pages in the current document with those specified from the source document.
#
# Save
#
# Saves a document.
#
# SetFlags
#
# Sets a document’s flags indicating whether the document has been modified, whether the document is a temporary document and should be deleted when closed, and the version of PDF used in the file.
#
# SetInfo
#
# Sets the value of a key in a document’s Info dictionary.
#
# SetPageMode
#
# Sets the page mode in which a document is to be opened: display only pages, pages and thumbnails, or pages and bookmarks.

pddoc = Dispatch("AcroExch.PDDoc")
# rect = Dispatch("AcroExch.Rect")

# pdf_dosya = avdoc.Open(dosya_yol, "")

# dosya işlemleri için dosyanın açılması gerekiyor.
pdf_dosya_acıldımı = pddoc.Open(dosya_yol)  # VARIANT_BOOL Open(BSTR szFullPath);

# Açılan dosyanın ismini almak için:
dosya_ismi = pddoc.GetFileName()  # BSTR GetFileName();

bayrak = pddoc.GetFlags()
bilgi = pddoc.GetInfo(1)  # BSTR GetInfo(BSTR szInfoKey);
dosya_id = pddoc.GetInstanceID()  # BSTR GetInstanceID();
dosya_id_p = pddoc.GetPermanentID()  # BSTR GetPermanentID();
sayfa_sayısı = pddoc.GetNumPages()  # long GetNumPages();

sayfa_modu = pddoc.GetPageMode()  # long GetPageMode();
# Returns
# The current page mode. Will be one of the following values:
#
# PDDontCare: 0 — leave the view mode as it is
#
# PDUseNone: 1 — display without bookmarks or thumbnails
#
# PDUseThumbs: 2 — display using thumbnails
#
# PDUseBookmarks: 3 — display using bookmarks
#
# PDFullScreen: 4 — display in full screen mode
# gateway = JavaGateway()
# gateway.jvm.java.util.Random()
pdf_sayfalar = []

# “Not implemented” Exception when using pywin32 to control Adobe Acrobat
# https://stackoverflow.com/questions/9383307/not-implemented-exception-when-using-pywin32-to-control-adobe-acrobat
global ERRORS_BAD_CONTEXT  ###############################
ERRORS_BAD_CONTEXT.append(winerror.E_NOTIMPL)  ############################

for sayfa in range(0, sayfa_sayısı):
    print("sayfa: {} ".format(sayfa))
    pdf_sayfalar.append(pddoc.AcquirePage(sayfa))  # LPDISPATCH AcquirePage(long nPage); ilk sayfa: 0
    yeni_pdf = pddoc.Create()
    js = "a = this.syncAnnotScan();"
    # Açılan dosyadan sayfaları almak için burada sayfa sayısını edindikten sonra sayfalara ulaşabiliriz:
    # AcroExch.PDPage
    # A single page in the PDF representation of a document. This is a non-creatable interface. Just as PDF files are
    # partially composed of their pages, PDDoc objects are composed of PDPage objects. A page contains a series of objects
    # representing the objects drawn on the page (PDGraphic objects), a list of resources used in drawing the page,
    # annotations (PDAnnot objects), an optional thumbnail image of the page, and the threads used in any articles that
    #  occur on the page. The first page in a PDDoc object is page 0.
    #
    # Methods
    # The PDPage object has the following methods.
    döküman = pdf_sayfalar[sayfa].GetDoc()
    pdf_sss = döküman.AcquirePage(sayfa)
    pdf_java = döküman.GetJSObject()
    # yeni_pdf.InsertPages(0, pdf_sss, sayfa, 1, 0)
    yorum_sayısı = pdf_sss.GetNumAnnots()

    deneme_j = pdf_java.getPageRotation()
    deneme_jj = pdf_java.filesize
    dememe_jjj = pdf_java.syncAnnotScan()

    # PDPage.GetAnnot(2).GetContents

    # sayfadaki yorum objelerini almak için;
    yorumlar = []
    for i in range(0, yorum_sayısı):
        pdf_java.syncAnnotScan()
        yorumlar.append(pdf_sayfalar[sayfa].GetAnnot(i))

    # alınan yorumlardan gerekli bilgileri almak için
    # AcroExch.PDAnnot
    # An annotation on a page in a PDF file. This is a non-creatable interface. Acrobat applications have two built-in
    # annotation types: PDTextAnnot and PDLinkAnnot. The object provides access to the physical attributes of the
    # annotation. Plug-ins may add movie and Widget (form field) annotations, and developers can define new annotation
    # subtypes by creating new annotation handlers.
    #
    # Methods
    # The PDAnnot object has the following methods.
    # yorum içinde alabildiğimiz metotdlar
    içindeki = []
    obj_kordinat = []
    k = 0
    renk = []
    tarih = []
    sub_tipi = []
    baslık = []
    index = []
    for v in yorumlar:
        kordinat = []
        içindeki.append(v.GetContents())
        kordinat.append(v.GetRect().Bottom)
        kordinat.append(v.GetRect().Left)
        kordinat.append(v.GetRect().Right)
        kordinat.append(v.GetRect().Top)
        renk.append(v.GetColor())
        tarih.append(v.GetDate().Date)
        index.append(pdf_sayfalar[sayfa].GetAnnotIndex(v))

        # AcroExch.Time
        # Defines a specified time, accurate to the millisecond.
        #
        # Properties
        # The Time object has the following properties.
        #
        # Property
        #
        # Description
        #
        # Date
        #
        # Gets or sets the date from an AcroTime.
        #
        # Hour
        #
        # Gets or sets the hour from an AcroTime.
        #
        # Millisecond
        #
        # Gets or sets the milliseconds from an AcroTime.
        #
        # Minute
        #
        # Gets or sets the minutes from an AcroTime.
        #
        # Month
        #
        # Gets or sets the month from an AcroTime.
        #
        # Second
        #
        # Gets or sets the seconds from an AcroTime.
        #
        # Year
        #
        # Gets or sets the year from an AcroTime.
        sub_tipi.append(v.GetSubtype())
        baslık.append(v.GetTitle())

        obj_kordinat.append(kordinat)
        print(içindeki)
        print(kordinat)
        print(index)

# Açılan dosyayı kapatmak için:
dosya_kapandımı = pddoc.Close()  # VARIANT_BOOL Close();
del pddoc

print("son")
