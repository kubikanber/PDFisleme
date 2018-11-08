#! python3
# https://help.adobe.com/en_US/acrobat/acrobat_dc_sdk/2015/HTMLHelp/index.html#t=Acro12_MasterBook%2FIAC_API_OLE_Objects%2FAcroExch_PDDoc.htm
# AcroExch.PDPage
# A single page in the PDF representation of a document. This is a non-creatable interface.
# Just as PDF files are partially composed of their pages, PDDoc objects are composed of PDPage objects.
# A page contains a series of objects representing the objects drawn on the page (PDGraphic objects),
# a list of resources used in drawing the page, annotations (PDAnnot objects), an optional thumbnail image of the page,
# and the threads used in any articles that occur on the page. The first page in a PDDoc object is page 0.
#
# Methods
# The PDPage object has the following methods.
#
# Method# #     Description
#
# AddAnnot          # Adds a specified annotation at a specified location in the page’s annotation array
#                       #önçe get çözümlenecek
#
# AddNewAnnot       # Creates a new text annotation and adds it to the page. ##################### önçe get çözümlenecek
#
# CopyToClipboard   # Copies a PDF image to the clipboard without requiring an hWnd or hDC from the client.
#
# CreatePageHilite  # Creates a text selection from a list of character offsets and character counts on a single page.
#
# CreateWordHilite  # Creates a text selection from a list of word offsets and word counts on a single page.
#
# CropPage          # Crops the page.
#
# Draw              # Deprecated. Draws page contents into a specified window.
#
# DrawEx            # Draws page contents into a specified window.
#
# GetAnnot          # Gets the specified annotation from the page’s array of annotations. #############################
#
# GetAnnotIndex     # Gets the index (within the page’s annotation array) of the specified annotation. ################
#
# GetDoc            # Gets the AcroExch.PDDoc associated with the page. ################################################
#
# GetNumAnnots      # Gets the number of annotations on the page. ######################################################
#
# GetNumber         # Gets the page number of the current page. The first page in a document is page zero. #############
#
# GetRotate         # Gets the rotation value, in degrees, for the current page. #######################################
#
# GetSize           # Gets a page’s width and height in points. ########################################################
#
# RemoveAnnot       # Removes the specified annotation from the page’s annotation array.
#
# SetRotate         # Sets the rotation, in degrees, for the current page. #############################################

# Need AcquirePage object. bu sınıfı çalıştırmak için AcroExch.PDDoc dan AcquirePage objesi alınması gerekmektedir.
import AcroExchPDAnnot
import AcroExchPoint


class PDPage:

    def __init__(self, acquire_page):
        self.pdpage = acquire_page

    # GetNumber
    # Gets the page number of the current page. The first page in a document is page zero.
    #
    # Syntax
    # long GetNumber();
    #
    # Returns
    # The page number of the current page. The first page in a PDDoc object is page 0.

    def get_number(self):
        return self.pdpage.GetNumber()

    # GetNumAnnots
    # Gets the number of annotations on the page.
    #
    # Annotations that have associated pop-up windows, such as a strikeout, count as two annotations.
    # Also note that widget annotations (Acrobat form fields) are included.
    #
    # Syntax
    # long GetNumAnnots();
    #
    # Returns
    # The number of annotations on the page.
    def get_num_annots(self):
        return self.pdpage.GetNumAnnots()

    # GetAnnot
    # Gets the specified annotation from the page’s array of annotations.
    #
    # Syntax
    # LPDISPATCH GetAnnot(long nIndex);
    #
    # Parameters
    # nIndex
    #
    # Index (in the page’s annotation array) of the annotation to be retrieved.
    # The first annotation in the array has an index of zero.
    #
    # Returns
    # The LPDISPATCH for the AcroExch.PDAnnot object.
    # Buradan sınıf bağlantısı yapılacak........
    def get_annot(self, yorum_numarası=0):
        pdannot = AcroExchPDAnnot.PDAnnot(self.pdpage.GetAnnot(yorum_numarası))
        return pdannot

    # GetAnnotIndex
    # Gets the index (within the page’s annotation array) of the specified annotation.
    #
    # Syntax
    # long GetAnnotIndex(LPDISPATCH iPDAnnot);
    #
    # Parameters
    # iPDAnnot
    #
    # LPDISPATCH for the AcroExch.PDAnnot whose index is obtained.
    # iPDAnnot contains the instance variable m_lpDispatch, which contains the LPDISPATCH.
    #
    # Returns
    # The annotation’s index.
    def get_annot_index(self, yorum):
        return self.pdpage.GetAnnotIndex(yorum)

    # GetDoc
    # Gets the AcroExch.PDDoc associated with the page.
    #
    # Syntax
    # LPDISPATCH GetDoc();
    #
    # Returns
    # The LPDISPATCH for the page’s AcroExch.PDDoc.
    # sayfadan tüm dökümanı alıyor. tüm sayfaları geri almak için bir yöntem
    def get_doc(self):
        return self.pdpage.GetDoc()

    # GetRotate
    # Gets the rotation value, in degrees, for the current page.
    #
    # Syntax
    # short GetRotate();
    #
    # Returns
    # Rotation value.
    def get_rotate(self):
        return self.pdpage.GetRotate()

    # SetRotate
    # Sets the rotation, in degrees, for the current page.
    #
    # Syntax
    # VARIANT_BOOL SetRotate(short nRotate);
    #
    # Parameters
    # nRotate
    #
    # Rotation value of 0, 90, 180, or 270.
    #
    # Returns
    # 0 if the Acrobat application does not support editing, -1 otherwise.
    def set_rotate(self, derece):
        return self.pdpage.SetRotate(derece)

    # GetSize
    # Gets a page’s width and height in points.
    #
    # Syntax
    # LPDISPATCH GetSize();
    #
    # Returns
    # The LPDISPATCH for an AcroExch.Point containing the width and height, measured in points.
    # Point x contains the width, point y the height.
    def get_size(self):
        rec_size = AcroExchPoint.Point(self.pdpage.GetSize())
        return rec_size

    # AddAnnot
    # Adds a specified annotation at a specified location in the page’s annotation array.
    #
    # Syntax
    # VARIANT_BOOL AddAnnot(long nIndexAddAfter, LPDISPATCH iPDAnnot);
    #
    # Parameters
    # nIndexAddAfter
    #
    # Location in the page’s annotation array to add the annotation.
    # The first annotation on a page has an index of zero.
    #
    # iPDAnnot
    #
    # The LPDISPATCH for the AcroExch.PDAnnot to add.
    # iPDAnnot contains the instance variable m_lpDispatch, which contains the LPDISPATCH.
    #
    # Returns
    # 0 if the Acrobat application does not support editing, -1 otherwise
    # burada AcroExch.PDAnnot objesi kullanılacak.
    # Annot objesi önceden oluşturulup sonra sayfa eklenmesi
    def add_annot(self, yorum_nodan_sonra, iPDAnnot):
        return self.pdpage.AddAnnot(yorum_nodan_sonra, iPDAnnot)

    # AddNewAnnot
    # Creates a new text annotation and adds it to the page.
    #
    # The newly-created text annotation is not complete until PDAnnot.SetContents has been called to
    # fill in the /Contents key.
    #
    # Syntax
    # LPDISPATCH AddNewAnnot(long nIndexAddAfter, BSTR szSubType,
    #                  LPDISPATCH iAcroRect);
    #
    # Parameters
    # nIndexAddAfter
    #
    # Location in the page’s annotation array after which to add the annotation.
    # The first annotation on a page has an index of zero.
    #
    # szSubType
    #
    # Subtype of the annotation to be created. Must be text.
    #
    # iAcroRect
    #
    # The LPDISPATCH for the AcroExch.Rect bounding the annotation’s location on the page.
    # iAcroRect contains the instance variable m_lpDispatch, which contains the LPDISPATCH.
    #
    # Returns
    # The LPDISPATCH for an AcroExch.PDAnnot object, or NULL if the annotation could not be added.
    # burada son yorumdan sonra yorum tipi bildirilerek oluşturulan acrorect objesi ile ekleme yapılıyor.
    def add_new_annot(self, yorum_nodan_sonra, yorum_tipi, iAcroRect):
        return self.pdpage.AddNewAnnot(yorum_nodan_sonra, yorum_tipi, iAcroRect)

    # CopyToClipboard
    # Copies a PDF image to the clipboard without requiring an hWnd or hDC from the client.
    # This method is only available on 32-bit systems.
    #
    # Syntax
    # VARIANT_BOOL CopyToClipboard(LPDISPATCH boundRect,
    #
    #                  short nXOrigin,short nYOrigin,
    #
    #                  short nZoom);
    #
    # Parameters
    # boundRect
    #
    # The LPDISPATCH for the AcroExch.Rect bounding rectangle in device space coordinates.
    # boundRect contains the instance variable m_lpDispatch, which contains the LPDISPATCH.
    #
    # nXOrigin
    #
    # The x–coordinate of the portion of the page to be copied.
    #
    # nYOrigin
    #
    # The y–coordinate of the portion of the page to be copied.
    #
    # nZoom
    #
    # Zoom factor at which the page is copied, specified as a percent.
    # For example, 100 corresponds to a magnification of 1.0.
    #
    # Returns
    # -1 if the page is successfully copied, 0 otherwise.
    def copy_to_clipboard(self, acrorect, nXOrigin, nYOrigin, nZoom=1):
        return self.pdpage.CopyToClipboard(acrorect, nXOrigin, nYOrigin, nZoom)
