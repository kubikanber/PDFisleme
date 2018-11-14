#! python3, this code written by kubikanber, 09.11.2018, author: kubikanber
# AcroExch.PDAnnot
# An annotation on a page in a PDF file. This is a non-creatable interface.
# Acrobat applications have two built-in annotation types: PDTextAnnot and PDLinkAnnot.
# The object provides access to the physical attributes of the annotation.
# Plug-ins may add movie and Widget (form field) annotations, and developers can define new annotation subtypes by
# creating new annotation handlers.
#
# Methods
# The PDAnnot object has the following methods.
#
# Method ######## Description ##########################################################################################
#
# GetColor      # Gets an annotation’s color.
#
# GetContents   # Gets a text annotation’s contents.
#
# GetDate       # Gets an annotation’s date.
#
# GetRect       # Gets an annotation’s bounding rectangle.
#
# GetSubtype    # Gets an annotation’s subtype.
#
# GetTitle      # Gets a text annotation’s title.
#
# IsEqual       # Determines whether an annotation is the same as the specified annotation.
#
# IsOpen        # Tests whether a text annotation is open.
#
# IsValid       # Tests whether an annotation is still valid.
#
# Perform       # Performs a link annotation’s action.
#
# SetColor      # Sets an annotation’s color.
#
# SetContents   # Sets a text annotation’s contents.
#
# SetDate       # Sets an annotation’s date.
#
# SetOpen       # Opens or closes a text annotation.
#
# SetRect       # Sets an annotation’s bounding rectangle.
#
# SetTitle      # Sets a text annotation’s title.
########################################################################################################################
from win32com.client.dynamic import Dispatch

import AcroExchRect
import AcroExchTime


class PDAnnot:

    def __init__(self, get_annot):
        self.pdannot = get_annot

    # GetColor
    # Gets an annotation’s color.
    #
    # Syntax
    # long GetColor();
    #
    # Returns
    # The annotation’s color, a long value of the form 0x00BBGGRR where the first byte from the right (RR)
    # is a relative value for red, the second byte (GG) is a relative value for green, and the third byte (BB)
    # is a relative value for blue. The high-order byte must be 0.

    # Some samples of RGB are shown below.
    #
    # Please use it with method to set color number or get it. http://pdf-file.nnn2.com/?p=145
    def get_color(self):
        return self.pdannot.GetColor()

    # GetContents
    # Gets a text annotation’s contents.
    #
    # Syntax
    # BSTR GetContents();
    #
    # Returns
    # The annotation’s contents.
    def get_contents(self):
        return self.pdannot.GetContents()

    # GetDate
    # Gets an annotation’s date.
    #
    # Syntax
    # LPDISPATCH GetDate();
    #
    # Returns
    # The LPDISPATCH for an AcroExch.Time object containing the date.
    # Sonuç Date objesi dönüyor. bunun için AcroExch.Time object ile bağlanılacak.
    def get_date(self):
        annot_date = AcroExchTime.AcroTime(self.pdannot.GetDate())
        return annot_date

    # GetRect
    # Gets an annotation’s bounding rectangle.
    #
    # Syntax
    # LPDISPATCH GetRect();
    #
    # Returns
    # The LPDISPATCH for an AcroExch.Rect containing the annotation’s bounding rectangle
    # burada rect objesi dönüyor. AcroExch.Rect ojesinden değerler alınacak.
    def get_rect(self):
        acrorect = AcroExchRect.Rect(self.pdannot.GetRect())
        return acrorect

    # GetSubtype
    # Gets an annotation’s subtype.
    #
    # Syntax
    # BSTR GetSubtype();
    #
    # Returns
    # The annotation’s subtype. The built-in subtypes are Text and Link.
    # There are four types: Text, Popup, Link, Highlight.
    def get_subtype(self):
        return self.pdannot.GetSubtype()

    # GetTitle
    # Gets a text annotation’s title.
    #
    # Syntax
    # BSTR GetTitle();
    #
    # Returns
    # The annotation’s title.
    def get_title(self):
        return self.pdannot.GetTitle()

    # IsEqual
    # Determines whether an annotation is the same as the specified annotation.
    #
    # Syntax
    # VARIANT_BOOL IsEqual(LPDISPATCH PDAnnot);
    #
    # Parameters
    # PDAnnot
    #
    # The LPDISPATCH for the AcroExch.PDAnnot to be tested. PDAnnot contains the instance variable m_lpDispatch,
    #  which contains the LPDISPATCH.
    #
    # Returns
    # -1 if the annotations are the same, 0 otherwise.
    def is_equal(self, pdannot):
        return self.pdannot.IsEqual(pdannot)

    # IsOpen
    # Tests whether a text annotation is open.
    #
    # Syntax
    # VARIANT_BOOL IsOpen();
    #
    # Returns
    # -1 if open, 0 otherwise.
    def is_open(self):  # not understood
        return self.pdannot.IsOpen()

    # IsValid
    # Tests whether an annotation is still valid.
    # This method is intended only to test whether the annotation has been deleted,
    # not whether it is a completely valid annotation object.
    #
    # Syntax
    # VARIANT_BOOL IsValid();
    #
    # Returns
    # -1 if the annotation is valid, 0 otherwise.
    def is_valid(self):
        return self.pdannot.IsValid()

    # Perform
    # Performs a link annotation’s action.
    #
    # Syntax
    # VARIANT_BOOL Perform(LPDISPATCH iAcroAVDoc);
    #
    # Parameters
    # iAcroAVDoc
    #
    # The LPDISPATCH for the AcroExch.AVDoc in which the annotation is located.
    # iAcroAVDoc contains the instance variable m_lpDispatch, which contains the LPDISPATCH.
    #
    # Returns
    # -1 if the action was executed successfully, 0 otherwise.
    def perform(self, acroavdoc):  # todo: future
        return self.pdannot.Perform(acroavdoc)

    # SetColor
    # Sets an annotation’s color.
    #
    # Syntax
    # VARIANT_BOOL SetColor(long nRGBColor);
    #
    # Parameters
    # nRGBColor
    #
    # The color to use for the annotation.
    #
    # Returns
    # -1 if the annotation’s color was set, 0 if the Acrobat application does not support editing.
    #
    # nRGBColor is a long value with the form 0x00BBGGRR where the first byte from the right (RR)
    # is a relative value for red, the second byte (GG) is a relative value for green, and the third byte (BB)
    # is a relative value for blue. The high-order byte must be 0.
    def set_color(self, rgb_color):
        return self.pdannot.SetColor(rgb_color)

    # SetContents
    # Sets a text annotation’s contents.
    #
    # Syntax
    # VARIANT_BOOL SetContents(BSTR szContents);
    #
    # Parameters
    # szContents
    #
    # The contents to use for the annotation.
    #
    # Returns
    # 0 if the Acrobat application does not support editing, -1 otherwise.
    def set_contents(self, contents):
        return self.pdannot.SetContents(contents)

    # SetDate
    # Sets an annotation’s date.
    #
    # Syntax
    # VARIANT_BOOL SetDate(LPDISPATCH iAcroTime);
    #
    # Parameters
    # iAcroTime
    #
    # The LPDISPATCH for the date and time to use for the annotation. iAcroTime’s instance variable m_lpDispatch
    # contains this LPDISPATCH.
    #
    # Returns
    # -1 if the date was set, 0 if the Acrobat application does not support editing.
    def set_date(self, acrotime):
        return self.pdannot.SetDate(acrotime)

    # SetOpen
    # Opens or closes a text annotation.
    #
    # Syntax
    # VARIANT_BOOL SetOpen(long bIsOpen);
    #
    # Parameters
    # bIsOpen
    #
    # If a positive number, the annotation is open. If 0, the annotation is closed.
    # 0 yada 1 olaçak
    #
    # Returns
    # Always returns -1.
    def set_open(self, is_open):
        return self.pdannot.SetOpen(is_open)

    # SetRect
    # Sets an annotation’s bounding rectangle.
    #
    # Syntax
    # VARIANT_BOOL SetRect(LPDISPATCH iAcroRect);
    #
    # Parameters
    # iAcroRect
    #
    # The LPDISPATCH for the bounding rectangle (AcroExch.Rect) to set. iAcroRect contains
    # the instance variable m_lpDispatch, which contains the LPDISPATCH.
    #
    # Returns
    # -1 if a rectangle was supplied, 0 otherwise.
    def set_rect(self, bottom_point, left_point, right_point, top_point):
        acrorect = AcroExchRect.Rect()
        acrorect.create_rect(bottom_point, left_point, right_point, top_point)
        return self.pdannot.SetRect(acrorect.rect)

    # SetTitle
    # Sets a text annotation’s title.
    #
    # Syntax
    # VARIANT_BOOL SetTitle(BSTR szTitle);
    #
    # Parameters
    # szTitle
    #
    # The title to use.
    #
    # Returns
    # -1 if the title was set, 0 if the Acrobat application does not support editing.
    def set_title(self, title):
        return self.pdannot.SetTitle(title)
