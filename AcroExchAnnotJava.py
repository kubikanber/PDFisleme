#! python3, this code written by kubikanber, 09.11.2018, author: kubikanber
# Annot Java
# Annotation
# This object represents an Acrobat annotation. Annotations can be created using the Acrobat annotation tool or
# by using the Doc object method addAnnot.
#
# Before an annotation can be accessed, it must be bound to a JavaScript variable through a Doc object method such as
# getAnnot:
#
#    var a = this.getAnnot(0, "Important");
#
# The script can then manipulate the annotation named “Important” on page 1 (0-based page numbering system)
# by means of the variable a. For example, the following code first stores the type of annotation in the variable
# thetype, then changes the author to “John Q. Public”.
#
#    var thetype = a.type;          // read property
#
#    a.author = "John Q. Public";   // write property
#
# Another way of accessing the Annotation object is through the Doc object getAnnots method.
#
# Note:In Adobe Reader 5.1 or later, you can get the value of any annotation property except contents.
# The ability to set these properties depends on Comments document rights, as indicated by the C icon.
#
# Note:The user interface in Acrobat refers to annotations as comments.

# Annotation types
# Annotations are of different types, as reflected in the type property. Each type is listed in the table below,
# along with all documented properties returned by the getProps method.
#
# Annotation type
#
# Properties
#
# Caret:
#
# author, borderEffectIntensity, borderEffectStyle, caretSymbol, contents, creationDate, delay, hidden, inReplyTo,
# intent, lock, modDate, name, noView, opacity, page, popupOpen, popupRect, print, readOnly, rect, refType,
# richContents, rotate, seqNum, strokeColor, style, subject, toggleNoView, type, width
#
# Circle:
#
# author, borderEffectIntensity, borderEffectStyle, contents, creationDate, dash, delay, fillColor, hidden, inReplyTo,
#  intent, lock, modDate, name, noView, opacity, page, popupOpen, popupRect, print, readOnly, rect, refType,
# richContents, rotate, seqNum, strokeColor, style, subject, toggleNoView, type, width
#
# FileAttachment:
#
# attachIcon, attachment, author, borderEffectIntensity, borderEffectStyle, cAttachmentPath, contents, creationDate,
# delay, hidden, inReplyTo, intent, lock, modDate, name, noView, opacity, page, point, print, readOnly, rect, refType,
# richContents, rotate, seqNum, strokeColor, style, subject, toggleNoView, type, width
#
# FreeText:
#
# alignment, author, borderEffectIntensity, borderEffectStyle, callout, contents, creationDate, dash, delay, fillColor,
#  hidden, inReplyTo, intent, lineEnding, lock, modDate, name, noView, opacity, page, print, readOnly, rect, refType,
# richContents, richDefaults, rotate, seqNum, strokeColor, style, subject, textFont, textSize, toggleNoView, type, width
#
# Highlight:
#
# author, borderEffectIntensity, borderEffectStyle, contents, creationDate, delay, hidden, inReplyTo, intent, lock,
# modDate, name, noView, opacity, page, popupOpen, popupRect, print, quads, readOnly, rect, refType, richContents,
# rotate, seqNum, strokeColor, style, subject, toggleNoView, type, width
#
# Ink:
#
# author, borderEffectIntensity, borderEffectStyle, contents, creationDate, dash, delay, gestures, hidden, inReplyTo,
# intent, lock, modDate, name, noView, opacity, page, popupOpen, popupRect, print, readOnly, rect, refType,
# richContents, rotate, seqNum, strokeColor, style, subject, toggleNoView, type, width
#
# Line:
#
# arrowBegin, arrowEnd, author, borderEffectIntensity, borderEffectStyle, contents, creationDate, dash, delay,
# doCaption, fillColor, hidden, inReplyTo, intent, leaderExtend, leaderLength, lock, modDate, name, noView, opacity,
# page, points, popupOpen, popupRect, print, readOnly, rect, refType, richContents, rotate, seqNum, strokeColor, style,
#  subject, toggleNoView, type, width
#
# Polygon:
#
# author, borderEffectIntensity, borderEffectStyle, contents, creationDate, dash, delay, fillColor, hidden, inReplyTo,
# intent, lock, modDate, name, noView, opacity, page, popupOpen, popupRect, print, readOnly, rect, refType,
# richContents, rotate, seqNum, strokeColor, style, subject, toggleNoView, type, vertices, width
#
# PolyLine:
#
# arrowBegin, arrowEnd, author, borderEffectIntensity, borderEffectStyle, contents, creationDate, dash, delay,
# fillColor, hidden, inReplyTo, intent, lock, modDate, name, noView, opacity, page, popupOpen, popupRect, print,
# readOnly, rect, refType, richContents, rotate, seqNum, strokeColor, style, subject, toggleNoView, type, vertices,
# width
#
# Redact:
#
# In addition to all the usual properites of a markup type annotation (see for example, the property list of the
# Highlight type), the properties alignment, overlayText and repeat are particular to the Redact annotation.
#
# Sound:
#
# author, borderEffectIntensity, borderEffectStyle, contents, creationDate, delay, hidden, inReplyTo, intent, lock,
# modDate, name, noView, opacity, page, point, print, readOnly, rect, refType, richContents, rotate, seqNum, soundIcon,
#  strokeColor, style, subject, toggleNoView, type, width
#
# Square:
#
# author, borderEffectIntensity, borderEffectStyle, contents, creationDate, dash, delay, fillColor, hidden, inReplyTo,
# intent, lock, modDate, name, noView, opacity, page, popupOpen, popupRect, print, readOnly, rect, refType,
# richContents, rotate, seqNum, strokeColor, style, subject, toggleNoView, type, width
#
# Squiggly:
#
# author, borderEffectIntensity, borderEffectStyle, contents, creationDate, delay, hidden, inReplyTo, intent, lock,
# modDate, name, noView, opacity, page, popupOpen, popupRect, print, quads, readOnly, rect, refType, richContents,
#  rotate, seqNum, strokeColor, style, subject, toggleNoView, type, width
#
# Stamp:
#
# AP, author, borderEffectIntensity, borderEffectStyle, contents, creationDate, delay, hidden, inReplyTo, intent, lock,
# modDate, name, noView, opacity, page, popupOpen, popupRect, print, readOnly, rect, refType, rotate, seqNum,
# strokeColor, style, subject, toggleNoView, type
#
# StrikeOut:
#
# author, borderEffectIntensity, borderEffectStyle, contents, creationDate, delay, hidden, inReplyTo, intent, lock,
# modDate, name, noView, opacity, page, popupOpen, popupRect, print, quads, readOnly, rect, refType, richContents,
# rotate, seqNum, strokeColor, style, subject, toggleNoView, type, width
#
# Text:
#
# author, borderEffectIntensity, borderEffectStyle, contents, creationDate, delay, hidden, inReplyTo, intent, lock,
# modDate, name, noView, noteIcon, opacity, page, point, popupOpen, popupRect, print, readOnly, rect, refType,
#  richContents, rotate, seqNum, state, stateModel, strokeColor, style, subject, toggleNoView, type, width
#
# Underline:
#
# author, borderEffectIntensity, borderEffectStyle, contents, creationDate, delay, hidden, inReplyTo, intent, lock,
# modDate, name, noView, opacity, page, popupOpen, popupRect, print, quads, readOnly, rect, refType, richContents,
# rotate, seqNum, strokeColor, style, subject, toggleNoView, type, width

# Annotation properties
# The PDF Reference documents all Annotation properties and specifies how they are stored.
#
# Some property values are stored in the PDF document as names and others are stored as strings
# (see the PDF Reference for details). A property stored as a name can have only 127 characters.
#
# Examples of properties that have a 127-character limit include AP, beginArrow,
# endArrow, attachIcon, noteIcon, and soundIcon.
#
# The Annotation properties are: #######################################################################################
#
# alignment     # doCaption     # overlayText   # seqNum        # AP        # fillColor # page          # soundIcon
#
# attachIcon    # inReplyTo     # arrowBegin    # gestures      # point     # state     # arrowEnd      # hidden
#
# points        # stateModel    # popupOpen     # strokeColor   # author    # intent    # popupRect     # style
#
# borderEffectIntensity         # leaderExtend  # print         # subject   # borderEffectStyle         # leaderLength
#
# quads         # textFont      # callout       # lineEnding    # rect      # textSize  # caretSymbol   # lock
#
# readOnly      # toggleNoView  # contents      # modDate       # refType   # type      # creationDate  # name
#
# repeat        # vertices      # dash          # noteIcon      # richContents          # width         # delay
#
# noView        # richDefaults  # doc           # opacity       # rotate
########################################################################################################################


class AnnotJava:

    def __init__(self, get_annot):
        self.annotjava = get_annot

    # The type of annotation. The type of an annotation can only be set within the object-literal argument of the Doc
    # object addAnnot method. The valid values are:
    #
    # Text       #FreeText   #Line          #Square     #Circle     #Polygon    #PolyLine           #Highlight
    #
    # Underline  #Squiggly   #StrikeOut     #Stamp      #Caret      #Ink        #FileAttachment     #Sound
    #
    #
    # Type
    # String
    #
    # Access
    #  R
    #
    # Annotations
    #  All
    def get_type(self):
        return self.annotjava.type

    # The page on which the annotation resides.
    #
    # Type
    # Integer
    #
    # Access
    # R/W
    #
    # Annotations
    # All
    def get_set_page(self, value=None):
        if value is not None:
            self.annotjava.page = value
        return self.annotjava.page

    # Accesses the contents of any annotation that has a pop-up window. For sound and file attachment annotations,
    # specifies the text to be displayed as a description.
    #
    # Type
    # String
    #
    # Access
    # R/W
    #
    # Annotations
    # All
    def get_set_contents(self, value):
        if value is not None:
            self.annotjava.contents = value
        return self.annotjava.contents
