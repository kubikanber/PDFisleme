#! python3, this code written by kubikanber, 09.11.2018, author: kubikanber
# Doc Java
# This object provides the interface between a PDF document open in the viewer and the JavaScript interpreter.
# It provides methods and properties for accessing the PDF document.
#
# You can access the Doc object from JavaScript in a variety of ways:
#
# The this object usually points to the Doc object of the underlying document.
#
# Some properties and methods, such as extractPages, app.activeDocs, and app. openDoc, return the Doc object.
#
# The Doc object can often be accessed through event objects, which are created for each event by which a JavaScript
# is executed:
#
# For mouse, focus, blur, calculate, validate, and format events, event.target returns the Field object that initiated
# the event. You can then access the Doc object through the Doc method of the Field object.
#
# For all other events, event.target points to the Doc.
#
# Doc properties #######################################################################################################
#
# alternatePresentations    # external  # mouseX    # securityHandler   # author    # filesize  # mouseY
#
# selectedAnnots    # baseURL   # hidden    # noautocomplete    # sounds    # bookmarkRoot  # hostContainer # numFields
#
# spellDictionaryOrder  # calculate # icons # numPages  # spellLanguageOrder    # collection    # info  # numTemplates
#
# subject   # creationDate  # innerAppWindowRect    # path  # templates # creator   # innerDocWindowRect
#
# outerAppWindowRect    # title # dataObjects   # isModal   # outerDocWindowRect    # URL   # delay # keywords
#
# pageNum   # viewState # dirty # layout    # pageWindowRect    # xfa   # disclosed # media # permStatusReady
#
# XFAForeground # docID # metadata  # producer  # zoom  # documentFileName  # modDate   # requiresFullSave  # zoomType
#
# dynamicXFAForm
#
########################################################################################################################
#
# Doc methods ##########################################################################################################
# addAnnot  # getAnnot3D    # mailDoc   # addField  # getAnnots # mailForm  # addIcon   # getAnnots3D   # movePage
#
# addLink   # getColorConvertAction     # newPage   # addRecipientListCryptFilter       # getDataObject # openDataObject
#
# addRequirement            # getDataObjectContents # preflight # addScript # getField  # print         # addThumbnails
#
# getIcon   # removeDataObject          # addWatermarkFromFile  # getLegalWarnings      # removeField
#
# addWatermarkFromText      # getLinks  # removeIcon            # addWeblinks           # getNthFieldName
#
# removeLinks   # applyRedactions   # getNthTemplate    # removePreflightAuditTrail # bringToFront  # getOCGs
#
# removeRequirement # calculateNow  # getOCGOrder       # removeScript  # certifyInvisibleSign  # getPageBox
#
# removeTemplate    # closeDoc      # getPageLabel      # removeThumbnails  # colorConvertPage  # getPageNthWord
#
# removeWeblinks    # createDataObject  # getPageNthWordQuads   # replacePages  # createTemplate    # getPageNumWords
#
# resetForm # deletePages   # getPageRotation   # saveAs    # deleteSound   # getPageTransition # scroll
#
# embedDocAsDataObject  # getPreflightAuditTrail    # selectPageNthWord # embedOutputIntent # getPrintParams
#
# setAction # encryptForRecipients  # getSound  # setDataObjectContents # encryptUsingPolicy    # getTemplate
#
# setOCGOrder   # exportAsFDF   # getURL    # setPageAction # exportAsFDFStr    # getUserUnitSize   # setPageBoxes
#
# exportAsText  # gotoNamedDest # setPageLabels # exportAsXFDF  # importAnFDF   # setPageRotations  # exportAsXFDFStr
#
# importAnXFDF  # setPageTabOrder   # exportDataObject  # importDataObject  # setPageTransitions    # exportXFAData
#
# importIcon    # spawnPageFromTemplate # extractPages  # importSound   # submitForm    # flattenPages  # importTextData
#
# syncAnnotScan # getAnnot  # importXFAData # timestampSign # insertPages   # validatePreflightAuditTrail
########################################################################################################################

import winerror

from win32com.client.dynamic import ERRORS_BAD_CONTEXT

# “Not implemented” Exception when using pywin32 to control Adobe Acrobat
# https://stackoverflow.com/questions/9383307/not-implemented-exception-when-using-pywin32-to-control-adobe-acrobat
import AcroExchAnnotJava

global ERRORS_BAD_CONTEXT  ###############################
ERRORS_BAD_CONTEXT.append(winerror.E_NOTIMPL)  ############################


class DocJava:

    def __init__(self, get_js_object):
        self.docjava = get_js_object

    # Gets the rotation of the specified page. See also setPageRotations.
    #
    # Parameters
    # nPage
    #
    # (optional) The 0-based index of the page. The default is 0, the first page in the document.
    #
    # Returns
    # The rotation value of 0, 90, 180, or 270.
    def get_page_rotation(self):
        return self.docjava.getPageRotation()

    # Guarantees that all annotations will be scanned by the time this method returns.
    #
    # To show or process annotations for the entire document, all annotations must have been detected.
    # Normally, a background task runs that examines every page and looks for annotations during idle time,
    # as this scan is a time-consuming task. Much of the annotation behavior works gracefully even when the full list
    # of annotations is not yet acquired by background scanning.
    #
    # In general, you should call this method if you want the entire list of annotations.
    def sync_annot_scan(self):
        self.docjava.syncAnnotScan()

    # Gets an array of Annotation objects satisfying specified criteria. See also getAnnot and syncAnnotScan.
    #
    # Parameters
    # nPage :
    #
    #  (optional) A 0-based page number. If specified, gets only annotations on the given page. If not specified,
    # gets annotations that meet the search criteria from all pages.
    #
    # nSortBy
    #
    # (optional) A sort method applied to the array. Values are:
    #
    #   ANSB_None — (default) Do not sort; equivalent to not specifiying this parameter.
    #
    #   ANSB_Page — Use the page number as the primary sort criteria.
    #
    #   ANSB_Author — Use the author as the primary sort criteria.
    #
    #   ANSB_ModDate — Use the modification date as the primary sort criteria.
    #
    # A NSB_Type — Use the annotation type as the primary sort criteria.
    #
    # bReverse
    #
    # (optional) If true, causes the array to be reverse sorted with respect to nSortBy.
    #
    # nFilterBy
    #
    # (optional) Gets only annotations satisfying certain criteria. Values are:
    #
    #   ANFB_ShouldNone — (default) Get all annotations. Equivalent of not specifying this parameter.
    #
    #   ANFB_ShouldPrint — Only include annotations that can be printed.
    #
    #   ANFB_ShouldView — Only include annotations that can be viewed.
    #
    #   ANFB_ShouldEdit — Only include annotations that can be edited.
    #
    #   ANFB_ShouldAppearInPanel — Only annotations that appear in the annotations pane.
    #
    #   ANFB_ShouldSummarize — Only include annotations that can be included in a summary.
    #
    #   ANFB_ShouldExport — Only include annotations that can be included in an export.
    #
    # Returns
    # An array of Annotation objects, or null if none are found.
    def get_annots(self, nPage=None, nSortBy="ANSB_None", bReverse="False", nFilterBy="ANFB_ShouldNone"):
        if nPage is None:
            annots = self.docjava.getAnnots()
        else:
            annots = self.docjava.getAnnots(nPage, '', nSortBy, '', bReverse, nFilterBy)
        return annots
