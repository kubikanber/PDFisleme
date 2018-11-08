#! python 3
import os
import sys
from binascii import b2a_hex

import pdfminer
from pdfminer.pdfparser import PDFParser, PDFDocument, PDFNoOutlines
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import PDFPageAggregator
from pdfminer.layout import LAParams, LTTextBox, LTTextLine, LTFigure, LTImage


def with_pdf(pdf_doc, pdf_pwd, fn, *args):
    """Open the pdf document, and apply the function, returning the results"""
    result = None
    try:
        # open the pdf file
        fp = open(pdf_doc, 'rb')

        # create a parser object associated with the file object
        parser = PDFParser(fp)

        # create a PDFDocument object that stores the document structure
        doc = PDFDocument()

        # connect the parser and document objects
        parser.set_document(doc)
        doc.set_parser(parser)

        # supply the password for initialization
        doc.initialize(pdf_pwd)

        if doc.is_extractable:
            # apply the function and return the result
            result = fn(doc, *args)
        else:
            print("Dosya uygun değil.")
        # close the pdf file
        fp.close()
    except IOError:
        # the file doesn't exist or similar problem
        print("Dosya açılamadı.")
    return result


def _parse_toc(doc):
    """With an open PDFDocument object, get the table of contents (toc) data
    [this is a higher-order function to be passed to with_pdf()]"""
    toc = []
    try:
        outlines = doc.get_outlines()
        for (level, title, dest, a, se) in outlines:
            toc.append((level, title))
    except PDFNoOutlines:
        print("boş geçti")
        pass
    return toc


def _parse_pages(doc):
    """With an open PDFDocument object, get the pages and parse each one
    [this is a higher-order function to be passed to with_pdf()]"""
    rsrcmgr = PDFResourceManager()
    laparams = LAParams()
    device = PDFPageAggregator(rsrcmgr, laparams=laparams)
    interpreter = PDFPageInterpreter(rsrcmgr, device)
    for page in doc.get_pages():
        interpreter.process_page(page)
        # receive the LTPage object for this page
        layout = device.get_result()
        # layout is an LTPage object which may contain child objects like LTTextBox, LTFigure, LTImage, etc.
        print("sayfalar işlendi")


def _parse_pages_t(doc, images_folder):
    """With an open PDFDocument object, get the pages, parse each one, and return the entire text
    [this is a higher-order function to be passed to with_pdf()]"""
    rsrcmgr = PDFResourceManager()
    laparams = LAParams()
    device = PDFPageAggregator(rsrcmgr, laparams=laparams)
    interpreter = PDFPageInterpreter(rsrcmgr, device)

    text_content = []  # a list of strings, each representing text collected from each page of the doc
    for i, page in enumerate(doc.get_pages()):
        interpreter.process_page(page)
        # receive the LTPage object for this page
        layout = device.get_result()
        # layout is an LTPage object which may contain child objects like LTTextBox, LTFigure, LTImage, etc.
        text_content.append(parse_lt_objs(layout._objs, (i + 1), images_folder))

    return text_content


def parse_lt_objs(lt_objs, page_number, images_folder, text=[]):
    """Iterate through the list of LT* objects and capture the text or image data contained in each"""
    text_content = []

    for lt_obj in lt_objs:
        if isinstance(lt_obj, LTTextBox) or isinstance(lt_obj, LTTextLine):
            # text
            text_content.append(lt_obj.get_text())
        elif isinstance(lt_obj, LTImage):
            # an image, so save it to the designated folder, and note it's place in the text
            saved_file = save_image(lt_obj, page_number, images_folder)
            if saved_file:
                # use html style <img /> tag to mark the position of the image within the text
                text_content.append('<img src="' + os.path.join(images_folder, saved_file) + '" />')
            else:
                print(sys.stderr, "Error saving image on page", page_number, lt_obj.__repr__)
        elif isinstance(lt_obj, LTFigure):
            # LTFigure objects are containers for other LT* objects, so recurse through the children
            text_content.append(parse_lt_objs(lt_obj._objs, page_number, images_folder, text_content))

    return "\n".join(text_content)


def save_image(lt_image, page_number, images_folder):
    """Try to save the image data from this LTImage object, and return the file name, if successful"""
    result = None
    if lt_image.stream:
        file_stream = lt_image.stream.rawdata
        file_ext = determine_image_type(file_stream[0:4])
        if file_ext:
            file_name = "".join([str(page_number), "_", lt_image.name, file_ext])
            if write_file(images_folder, file_name, lt_image.stream.rawdata, flags="wb"):
                result = file_name
    return result


def determine_image_type(stream_first_4_bytes):
    """Find out the image file type based on the magic number comparison of the first 4 (or 2) bytes"""
    file_type = None
    bytes_as_hexx = b2a_hex(stream_first_4_bytes)
    bytes_as_hex = str(bytes_as_hexx,"utf-8")
    if bytes_as_hex.startswith("ffd8"):
        file_type = ".jpeg"
    elif bytes_as_hex.startswith("89504e47"):
        file_type = ".png"
    elif bytes_as_hex.startswith("47494638"):
        file_type = ".gif"
    elif bytes_as_hex.startswith("424d'"):
        file_type = ".bmp"
    return file_type


def write_file(folder, filename, filedata, flags='w'):
    """Write the file data to the folder and filename combination
    (flags: 'w' for write text, 'wb' for write binary, use 'a' instead of 'w' for append)"""
    result = False
    if os.path.isdir(folder):
        try:
            file_obj = open(os.path.join(folder, filename), flags)
            file_obj.write(filedata)
            file_obj.close()
            result = True
        except IOError:
            pass

    return result


dosya = "simple1.pdf"


def get_toc(pdf_doc, pdf_pwd=''):
    """Return the table of contents (toc), if any, for this pdf file"""
    return with_pdf(pdf_doc, pdf_pwd, _parse_toc)


def get_pages(pdf_doc, pdf_pwd=""):
    """Process each of the pages in this pdf file"""
    with_pdf(pdf_doc, pdf_pwd, _parse_pages)


def get_pages_t(pdf_doc, pdf_pwd="", images_folder=".\\tmp"):
    """Process each of the pages in this pdf file and print the entire text to stdout"""
    print("\n\n".join(with_pdf(pdf_doc, pdf_pwd, _parse_pages_t, *tuple([images_folder]))))
    print("tamam")


get_toc(dosya)
get_pages(dosya)
get_pages_t(dosya)
