#! python3 this code written by kubikanber, 09.11.2018, author: kubikanber
# AcroExch.AVDoc
# A view of a PDF document in a window. This is a creatable interface. There is one AVDoc object per displayed document.
#  Unlike a PDDoc object, an AVDoc object has a window associated with it.
#
# Methods
# The AVDoc object has the following methods.
#
# Method                # Description
#
# BringToFront          # Brings the window to the front.
#
# ClearSelection        # Clears the current selection.
#
# Close                 # Closes a document.
#
# FindText              # Finds the specified text, scrolls so that it is visible, and highlights it.
#
# GetAVPageView         # Gets the AcroExch.AVPageView associated with an AcroExch.AVDoc.
#
# GetFrame              # Gets the rectangle specifying the window’s size and location.
#
# GetPDDoc              # Gets the AcroExch.PDDoc associated with an AcroExch.AVDoc.
#
# GetTitle              # Gets the window’s title.
#
# GetViewMode           # Gets the current document view mode (pages only, pages and thumbnails,or pages and bookmarks).
#
# IsValid               # Determines whether the AcroExch.AVDoc is still valid.
#
# Maximize              # Maximizes the window if bMaxSize is a positive number.
#
# Open                  # Opens a file.
#
# OpenInWindow          # Opens a PDF file and displays it in a user-specified window.
#
# OpenInWindowEx        # Opens a PDF file and displays it in a user-specified window.
#
# PrintPages            # Prints a specified range of pages displaying a print dialog box.
#
# PrintPagesEx          # Prints a specified range of pages, displaying a print dialog box.
#
# PrintPagesSilent      # Prints a specified range of pages without displaying any dialog box.
#
# PrintPagesSilentEx    # Prints a specified range of pages without displaying any dialog box.
#
# SetFrame              # Sets the window’s size and location.
#
# SetTextSelection      # Sets the document’s selection to the specified text selection.
#
# SetTitle              # Sets the window’s title.
#
# SetViewMode           # Sets the mode in which the document will be viewed (pages only, pages and thumbnails,
#                           or pages and bookmarks)
#
# ShowTextSelect        # Changes the view so that the current text selection is visible.
from win32com.client.dynamic import Dispatch


class AVDoc:

    def __init__(self, file, pddoc=None):
        if pddoc is None:
            self.avdoc = Dispatch("AcroExch.AVDoc")
        else:
            self.avdoc = pddoc
        self.file = file

    # Open
    # Opens a file. A new instance of AcroExch.AVDoc must be created for each displayed PDF file.
    #
    # Note:An application must explicitly close any AVDoc that it opens by calling AVDoc.Close (the destructor for the
    # AcroExch.AVDoc class does not call AVDoc.Close).
    #
    # Syntax
    # VARIANT_BOOL Open(BSTR szFullPath, BSTR szTempTitle);
    #
    # Parameters
    # szFullPath
    #
    # The full path of the file to open.
    #
    # szTempTitle
    #
    # An optional title for the window in which the file is opened. If szTempTitle is NULL or the empty string,
    # it is ignored. Otherwise, szTempTitle is used as the window title.
    #
    # Returns
    # -1 if the file was opened successfully, 0 otherwise.
    def open(self, temp_title=""):
        return self.avdoc.Open(self.file, temp_title)

    # Close
    # Closes a document. You can close all open AVDoc objects by calling App.CloseAllDocs.
    #
    # To reuse an AVDoc object, close it with AVDoc.Close, then use the AVDoc object’s LPDISPATCH for
    # AVDoc.OpenInWindow.
    #
    # Syntax
    # VARIANT_BOOL Close(long bNoSave);
    #
    # Parameters
    # bNoSave
    #
    # If a positive number, the document is closed without saving it. If 0 and the document has been modified,
    # the user is asked whether or not the file should be saved.
    #
    # Returns
    # Always returns -1, even if no document is open.
    def close_doc(self, no_save=0):
        return self.avdoc.Close(no_save)
