#! python3
# https://help.adobe.com/en_US/acrobat/acrobat_dc_sdk/2015/HTMLHelp/index.html#t=Acro12_MasterBook%2FIAC_API_OLE_Objects%2FAcroExch_PDDoc.htm
# AcroExch.PDDoc
# The underlying PDF representation of a document. This is a creatable interface. There is a correspondence between
# a PDDoc object and an ASFile object (an opaque representation of an open file made available through an interface
# encapsulating Acrobat’s access to file services), and the PDDoc object is the hidden object behind every AVDoc object.
#  An ASFile object may have zero or more underlying files, so a PDF file does not always correspond to a single disk
# file. For example, an ASFile object may provide access to PDF data in a database.
#
# Through PDDoc objects, your application can perform most of the Document menu items from Acrobat (delete pages,
# replace pages, and so on), create and delete thumbnails, and set and retrieve document information fields.
#
# Methods
# The PDDoc object has the following methods.
#
# Method############# Description#######################################################################################

# AcquirePage       # Acquires the specified page.
#
# ClearFlags        # Clears a document’s flags.
#
# Close             # Closes a file.
#
# Create            # Creates a new AcroExch.PDDoc.
#
# CreateTextSelect  # Creates a text selection from the specified rectangle on the specified page. ?????????????????????
#
# CreateThumbs      # Creates thumbnail images for the specified page range in a document. #####????????????????????????
#
# CropPages         # Crops the pages in a specified range in a document.
#
# DeletePages       # Deletes pages from a file.
#
# DeleteThumbs      # Deletes thumbnail images from the specified pages in a document.
#
# GetFileName       # Gets the name of the file associated with this AcroExch.PDDoc.
#
# GetFlags          # Gets a document’s flags.
#
# GetInfo           # Gets the value of a specified key in the document’s Info dictionary.
#
# GetInstanceID     # Gets the instance ID (the second element) from the ID array in the document’s trailer.
#
# GetJSObject       # Gets a dual interface to the JavaScript object associated with the PDDoc. Burada
#                       Doc objelerine JavaScript komutları ile ulaşmak için
#                       https://stackoverflow.com/questions/9383307/not-implemented-exception-when-using-pywin32-to-control-adobe-acrobat
#
# GetNumPages       # Gets the number of pages in a file.
#
# GetPageMode       # Gets a value indicating whether the Acrobat application is currently displaying only pages,
#                       pages and thumbnails, or pages and bookmarks.
#
# GetPermanentID    # Gets the permanent ID (the first element) from the ID array in the document’s trailer.
#
# InsertPages       # Inserts the specified pages from the source document after the indicated page within the current
#                       document.
#
# MovePage          # Moves a page to another location within the same document.
#
# Open              # Opens a file.
#
# OpenAVDoc         # Opens a window and displays the document in it.
#
# ReplacePages      # Replaces the indicated pages in the current document with those specified from the
#                       source document.
#
# Save              # Saves a document.
#
# SetFlags          # Sets a document’s flags indicating whether the document has been modified, whether the document
#                       is a temporary document and should be deleted when closed, and the version of PDF used
#                       in the file.
#
# SetInfo           # Sets the value of a key in a document’s Info dictionary.
#
# SetPageMode       # Sets the page mode in which a document is to be opened: display only pages, pages and thumbnails,
#                       or pages and bookmarks.
########################################################################################################################


from win32com.client.dynamic import Dispatch

import AcroExchAVDoc
import AcroExchDocJava
import AcroExchPDPage
import AcroExchRect


class PDDoc:

    def __init__(self, file):
        self.pddoc = Dispatch("AcroExch.PDDoc")
        self.file = file
        # self.dosya = os.path.abspath(dosya) # dosya ismi yolu ile birlikte tamamem olması gerekmektedir.

    # Open
    # Opens a file. A new instance of AcroExch.PDDoc must be created for each open PDF file.
    #
    # Syntax
    # VARIANT_BOOL Open(BSTR szFullPath);
    #
    # Parameters
    # szFullPath
    #
    # Full path of the file to be opened.
    #
    # Returns
    # -1 if the document was opened successfully, 0 otherwise.
    def open(self):
        return self.pddoc.Open(self.file)

    # Close
    # Closes a file.
    #
    # Note:If PDDoc and AVDoc are constructed with the same file, PDDoc.Close destroys both objects
    # (which closes the document in the viewer).
    #
    # Syntax
    # VARIANT_BOOL Close();
    #
    # Returns
    # -1 if the document was closed successfully, 0 otherwise.

    def close(self):
        return self.pddoc.Close()

    # OpenAVDoc
    # Opens a window and displays the document in it.
    #
    # Syntax
    # LPDISPATCH OpenAVDoc(BSTR szTitle);
    #
    # Parameters
    # szTitle
    #
    # The title to be used for the window. A default title is used if szTitle is NULL or an empty string.
    #
    # Returns
    # The LPDISPATCH for the AcroExch.AVDoc that was opened, or NULL if the open fails.
    # OpenAVDoc ile açılan dosya yine AVDoc.Close() yöntemi ile kapatılmalıdır.
    def open_avdoc(self, title=""):
        avdoc = AcroExchAVDoc.AVDoc(self.file, pddoc=self.pddoc.OpenAVDoc(title))
        return avdoc

    # GetFileName
    # Gets the name of the file associated with this AcroExch.PDDoc.
    #
    # Syntax
    # BSTR GetFileName();
    #
    # Returns
    # The file name, which can currently contain up to 256 characters.
    def get_file_name(self):
        return self.pddoc.GetFileName()

    # Save
    # Saves a document.
    #
    # Syntax
    # VARIANT_BOOL Save(short nType, BSTR szFullPath);
    #
    # Parameters
    # nType
    #
    # Specifies the way in which the file should be saved.
    #
    # nType is a logical OR of one or more of the following flags:
    #
    # 0 = PDSaveIncremental — Write changes only, not the complete file. This will always result in a larger file,
    # even if objects have been deleted.
    #
    # 1 = PDSaveFull — Write the entire file to the filename specified by szFullPath.
    #
    # 2 = PDSaveCopy — Write a copy of the file into the file specified by szFullPath, but keep using the old file.
    # This flag can only be specified if PDSaveFull is also used.
    #
    # 3 = PDSaveCollectGarbage — Remove unreferenced objects; this often reduces the file size, and its usage is
    # encouraged. This flag can only be specified if PDSaveFull is also used.
    #
    # 4 = PDSaveLinearized — Save the file optimized for the web, providing hint tables. This allows the PDF file to be
    # byte-served. This flag can only be specified if PDSaveFull is also used.
    #
    # Note:If you save a file optimized for the web using the PDSaveLinearized flag, you must follow this sequence:
    #
    # 1.  Open the PDF file with PDDoc.Open.
    #
    # 2.  Call PDDoc.Save using the PDSaveLinearized flag.
    #
    # 3.  Call PDDoc.Close.
    #
    # This allows batch optimization of files.
    #
    # szFullPath
    #
    # The new path to the file, if any.
    #
    # Returns
    # -1 if the document was successfully saved. Returns 0 if it was not or if the Acrobat application does not support
    # editing.
    # kayıt tipi çalışiyor mu ??????
    def save(self, type, full_path=""):
        return self.pddoc.Save(type, full_path)

    # GetFlags
    # Gets a document’s flags. The flags indicate whether the document has been modified, whether the document is a
    # temporary document and should be deleted when closed, and the version of PDF used in the file.
    #
    # Syntax
    # long GetFlags();
    #
    # Returns
    # The document’s flags, containing an OR of the following:
    #
    # Flag                      ## Description
    #
    # PDDocNeedsSave            ## Document has been modified and needs to be saved.
    #
    # PDDocRequiresFullSave     ## Document cannot be saved incrementally; it must be written using PDSaveFull.
    #
    # PDDocIsModified           ## Document has been modified slightly (such as bookmarks or text annotations have been
    #                               opened or closed),but not in a way that warrants saving.
    #
    # PDDocDeleteOnClose        ## Document is based on a temporary file that must be deleted when the document is
    #                                closed or saved.
    #
    # PDDocWasRepaired          ## Document was repaired when it was opened.
    #
    # PDDocNewMajorVersion      ## Document’s major version is newer than current.
    #
    # PDDocNewMinorVersion      ## Document’s minor version is newer than current.
    #
    # PDDocOldVersion           ## Document’s version is older than current.
    #
    # PDDocSuppressErrors       ## Don’t display errors.
    # Dökümanda bir değişiklik olduysa 8325
    # döküman kayıt edildikten sonra 8320
    # yeni bir döküman oluşturulduğunda 15 içinde sayfa yok.
    def get_flags(self):
        return self.pddoc.GetFlags()

    # ClearFlags
    # Clears a document’s flags. The flags indicate whether the document has been modified, whether the document is a
    # temporary document and should be deleted when closed, and the version of PDF used in the file. This method can be
    #  used only to clear, not to set, the flag bits.
    #
    # Syntax
    # VARIANT_BOOL ClearFlags(long nFlags);
    #
    # Parameters
    # nFlags
    #
    # Flags to be cleared. See PDDoc.GetFlags for a description of the flags. The flags PDDocWasRepaired,
    # PDDocNewMajorVersion, PDDocNewMinorVersion, and PDDocOldVersion are read-only and cannot be cleared.
    #
    # Returns
    # Always returns -1.
    def clear_flags(self, flags=0):
        return self.pddoc.ClearFlags(flags)

    # SetFlags
    # Sets a document’s flags indicating whether the document has been modified, whether the document is a temporary
    # document and should be deleted when closed, and the version of PDF used in the file. This method can be used
    # only to set, not to clear, the flag bits.
    #
    # Syntax
    # VARIANT_BOOL SetFlags(long nFlags);
    #
    # Parameters
    # nFlags
    #
    # Flags to be set. See PDDoc.GetFlags for a description of the flags. The flags PDDocWasRepaired,
    # PDDocNewMajorVersion, PDDocNewMinorVersion, and PDDocOldVersion are read-only and cannot be set.
    # 8320 dosya kapalı iken.
    # 8325 dosyada bir değişiklik olduysa
    #
    # Returns
    # Always returns -1.
    def set_flags(self, flags=0):
        return self.pddoc.SetFlags(flags)

    # GetInfo
    # Gets the value of a specified key in the document’s Info dictionary. A maximum of 512 bytes are returned.
    #
    # Syntax
    # BSTR GetInfo(BSTR szInfoKey);
    #
    # Parameters
    # szInfoKey
    #
    # The key whose value is obtained.
    #
    # Returns
    # The string if the value was read successfully. Returns an empty string if the key does not exist or its value
    # cannot be read. (Title, Author, Subject, Keywords, Creator, Producer, CreationDate, ModDate, Trapped)
    # http://pdf-file.nnn2.com/?p=103
    def get_info(self, info):  # info tobe numeric or text
        return self.pddoc.GetInfo(info)

    # SetInfo
    # Sets the value of a key in a document’s Info dictionary.
    #
    # Syntax
    # VARIANT_BOOL SetInfo(BSTR szInfoKey, BSTR szBuffer);
    #
    # Parameters
    # szInfoKey
    #
    # The key whose value is set.
    #
    # szBuffer
    #
    # The value to be assigned to the key.
    #
    # Returns
    # -1 if the value was added successfully, 0 if it was not or if the Acrobat application does not support editing.
    # belge özelliklerinde Özel sekmesinde "Özel özellikler" ekleme
    def set_info(self, info_key, value):
        return self.pddoc.SetInfo(info_key, value)

    # GetInstanceID
    # Gets the instance ID (the second element) from the ID array in the document’s trailer.
    #
    # Syntax
    # BSTR GetInstanceID();
    #
    # Returns
    # A string whose maximum length is 32 characters, containing the document’s instance ID.
    def get_instance_id(self):
        return self.pddoc.GetInstanceID()

    # GetPermanentID
    # Gets the permanent ID (the first element) from the ID array in the document’s trailer.
    #
    # Syntax
    # BSTR GetPermanentID();
    #
    # Returns
    # A string whose maximum length is 32 characters, containing the document’s permanent ID.
    def get_permanent_id(self):
        return self.pddoc.GetPermanentID()

    # GetJSObject
    # Gets a dual interface to the JavaScript object associated with the PDDoc. This allows automation clients full
    # access to both built-in and user-defined JavaScript methods available in the document. For more information on
    #  working with JavaScript, see Developing Applications Using Interapplication Communication.
    #
    # Syntax
    # LDispatch* GetJSObject();
    #
    # Returns
    # The interface to the JavaScript object if the call succeeded, NULL otherwise.
    # Burada adobe DOC objesi JavaScript komutlarını çalıştırmak için bir obje oluşturulur.
    # https://stackoverflow.com/questions/9383307/not-implemented-exception-when-using-pywin32-to-control-adobe-acrobat
    def get_js_object(self):
        docjava = AcroExchDocJava.DocJava(self.pddoc.GetJSObject())
        return docjava

    # GetNumPages
    # Gets the number of pages in a file.
    #
    # Syntax
    # long GetNumPages();
    #
    # Returns
    # The number of pages, or -1 if the number of pages cannot be determined.
    # sayfa sayısını döndürür
    def get_num_pages(self):
        return self.pddoc.GetNumPages()

    # AcquirePage
    # Acquires the specified page.
    #
    # Syntax
    # LPDISPATCH AcquirePage(long nPage);
    #
    # Parameters
    # nPage
    #
    # The number of the page to acquire. The first page is page 0.
    #
    # Returns
    # The LPDISPATCH for the AcroExch.PDPage object for the acquired page.
    # Returns NULL if the page could not be acquired.
    # herzaman ilk sayfa 0
    # Burada sayfa objesi alındıktan sonra AcroExch.PDPage geçerlidir
    def acquire_page(self, page_number=0):
        pdpage = AcroExchPDPage.PDPage(self.pddoc.AcquirePage(page_number))
        return pdpage

    # GetPageMode
    # Gets a value indicating whether the Acrobat application is currently displaying only pages, pages and thumbnails,
    #  or pages and bookmarks.
    #
    # Syntax
    # long GetPageMode();
    #
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
    # First argument (long nPageMode): Value of
    # PDF initial page display mode
    #     0: Page only
    #     1: Page only *
    #     2: Thumbnail image panel and page *
    #     3: Bookmark panel and Page *
    #     4: Open in full screen mode. <Caution>
    #     It is reflected as 5: 1.
    #     6: Layers panel and page *
    #     7: Attachment panel and page *
    #     8: Ignored after this value
    def get_page_mode(self):
        return self.pddoc.GetPageMode()

    # SetPageMode
    # Sets the page mode in which a document is to be opened: display only pages, pages and thumbnails, or pages and
    # bookmarks.
    #
    # Syntax
    # VARIANT_BOOL SetPageMode(long nPageMode);
    #
    # Parameters
    # nPageMode
    #
    # The page mode to be set. Possible values:
    #
    # PDDontCare: 0 — leave the view mode as it is
    #
    # PDUseNone: 1 — display without bookmarks or thumbnails
    #
    # PDUseThumbs: 2 — display using thumbnails
    #
    # PDUseBookmarks: 3 — display using bookmarks
    def set_page_mode(self, page_mode_value):
        return self.pddoc.SetPageMode(page_mode_value)

    # Create
    # Creates a new AcroExch.PDDoc.
    #
    # Syntax
    # VARIANT_BOOL Create();
    #
    # Returns
    # -1 if the document is created successfully, 0 if it is not or if the Acrobat application does not support editing.
    # yenibir döküman oluşturmak için
    def create(self):
        return self.pddoc.Create()

    # CreateThumbs
    # Creates thumbnail images for the specified page range in a document.
    #
    # Syntax
    # VARIANT_BOOL CreateThumbs(long nFirstPage, long nLastPage);
    #
    # Parameters
    # nFirstPage
    #
    # First page for which thumbnail images are created. The first page in a PDDoc object is page 0.
    #
    # nLastPage
    #
    # Last page for which thumbnail images are created.
    #
    # Returns
    # -1 if thumbnail images were created successfully, 0 if they were not or if the Acrobat application
    # does not support editing.
    def create_thumbs(self, start_page, end_page):
        return self.pddoc.CreateThumbs(start_page, end_page)

    # DeleteThumbs
    # Deletes thumbnail images from the specified pages in a document.
    #
    # Syntax
    # VARIANT_BOOL DeleteThumbs(long nStartPage, long nEndPage);
    #
    # Parameters
    # nStartPage
    #
    # First page whose thumbnail image is deleted. The first page in a PDDoc object is page 0.
    #
    # nEndPage
    #
    # Last page whose thumbnail image is deleted.
    #
    # Returns
    # -1 if the thumbnails were deleted, 0 if they were not deleted or if the Acrobat application does not
    # support editing.
    def delete_thumbs(self, start_page, end_page):
        return self.pddoc.DeleteThumbs(start_page, end_page)

    # CropPages
    # Crops the pages in a specified range in a document. This method ignores the request if either the width
    # or height of the crop box is less than 72 points (one inch).
    #
    # Syntax
    # VARIANT_BOOL CropPages(long nStartPage, long nEndPage,
    #                  short nEvenOrOddPagesOnly,
    #                  LPDISPATCH iAcroRect);
    #
    # Parameters
    # nStartPage
    #
    # First page that is cropped. The first page in a PDDoc object is page 0.
    #
    # nEndPage
    #
    # Last page that is cropped.
    #
    # nEvenOrOddPagesOnly
    #
    # Value indicating which pages in the range are cropped. Must be one of the following:
    #
    # 0 — crop all pages in the range
    #
    # 1 — crop only odd pages in the range
    #
    # 2 — crop only even pages in the range
    #
    # iAcroRect
    #
    # An LPDISPATCH for a CAcroRect specifying the cropping rectangle, which is specified in user space.
    #
    # Returns
    # -1 if the pages were cropped successfully, 0 otherwise.
    # Burada iAcroRect objesi AcroExchRect sınıfı ile tanımlanması gerekiyor.
    def crop_pages(self, start_page=0, end_page=0, even_or_odd=0, iAcroRect=AcroExchRect.Rect().create_rect()):
        return self.pddoc.CropPages(start_page, end_page, even_or_odd, iAcroRect)

    # DeletePages
    # Deletes pages from a file.
    #
    # Syntax
    # VARIANT_BOOL DeletePages(long nStartPage, long nEndPage);
    #
    # Parameters
    # nStartPage
    #
    # The first page to be deleted. The first page in a PDDoc object is page 0.
    #
    # nEndPage
    #
    # The last page to be deleted.
    #
    # Returns
    # -1 if the pages were successfully deleted. Returns 0 if they were not or if the Acrobat application
    #  does not support editing.
    def delete_pages(self, start_page, end_page):
        return self.pddoc.DeletePages(start_page, end_page)

    # InsertPages
    # Inserts the specified pages from the source document after the indicated page within the current document.
    #
    # Syntax
    # VARIANT_BOOL InsertPages(long nInsertPageAfter,
    #
    #                  LPDISPATCH iPDDocSource,long nStartPage,
    #
    #                  long nNumPages, long bBookmarks);
    #
    # Parameters
    # nInsertPageAfter
    #
    # The page in the current document after which pages from the source document are inserted. The first page in a
    # PDDoc object is page 0.
    #
    # iPDDocSource
    #
    # The LPDISPATCH for the AcroExch.PDDoc containing the pages to insert. iPDDocSource contains the instance
    # variable m_lpDispatch, which contains the LPDISPATCH.
    #
    # nStartPage
    #
    # The first page in iPDDocSource to be inserted into the current document.
    #
    # nNumPages
    #
    # The number of pages to be inserted.
    #
    # bBookmarks
    #
    # If a positive number, bookmarks are copied from the source document. If 0, they are not.
    #
    # Returns
    # -1 if the pages were successfully inserted. Returns 0 if they were not or if the Acrobat application does not
    # support editing.
    def insert_pages(self, page_after=0, added_PDDoc=None, start_page=0, pages=0, bookmarks=0):
        ek = PDDoc(added_PDDoc)
        ek.open()
        return self.pddoc.InsertPages(page_after, ek.pddoc, start_page, pages, bookmarks)

    # MovePage
    # Moves a page to another location within the same document.
    #
    # Syntax
    # VARIANT_BOOL MovePage(long nMoveAfterThisPage,
    #
    #                  long nPageToMove);
    #
    # Parameters
    # nMoveAfterThisPage
    #
    # The page being moved is placed after this page number. The first page in a PDDoc object is page 0.
    #
    # nPageToMove
    #
    # Page number of the page to be moved.
    #
    # Returns
    # 0 if the Acrobat application does not support editing, -1 otherwise.
    def move_page(self, page_after, page):
        return self.pddoc.MovePage(page_after, page)

    # ReplacePages
    # Replaces the indicated pages in the current document with those specified from the source document.
    # No links or bookmarks are copied from iPDDocSource, but text annotations may optionally be copied.
    #
    # Syntax
    # VARIANT_BOOL ReplacePages(long nStartPage,
    #
    #                  LPDISPATCH iPDDocSource,
    #                  long nStartSourcePage, long nNumPages,
    #                  long bMergeTextAnnotations);
    #
    # Parameters
    # nStartPage
    #
    # The first page within the source file to be replaced. The first page in a PDDoc object is page 0.
    #
    # iPDDocSource
    #
    # The LPDISPATCH for the AcroExch.PDDoc containing the new copies of pages that are replaced.
    # iPDDocSource contains the instance variable m_lpDispatch, which contains the LPDISPATCH.
    #
    # nStartSourcePage
    #
    # The first page in iPDDocSource to use as a replacement page.
    #
    # nNumPages
    #
    # The number of pages to be replaced.
    #
    # bMergeTextAnnotations
    #
    # If a positive number, text annotations from iPDDocSource are copied. If 0, they are not.
    #
    # Returns
    # -1 if the pages were successfully replaced. Returns 0 if they were not or if the Acrobat application does not
    # support editing.
    def replace_pages(self, start_page, iPDDocSource, source_start_page, source_pages,
                      merge_text_annotations=1):
        re = PDDoc(iPDDocSource)
        re.open()
        return self.pddoc.ReplacePages(start_page, re.pddoc, source_start_page,
                                       source_pages, merge_text_annotations)
