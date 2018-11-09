#! python 3
import sys

from pdfminer.pdfparser import PDFParser, PDFDocument
from pdfminer.pdftypes import PDFValueError

pages = []
pdf_pwd = ""


def extract(objid, obj):
    global pages
    if isinstance(obj, dict):
        # Type is PDFObRef type
        if obj.get("Type") and obj["Type"].name == "Page":
            pages.append(objid)
        elif obj.get("C"):
            pr = obj["P"]
        elif obj.get("Type") and obj["Type"].name == "Annot":
            pages.append(objid)

        try:
            pi = pages.index(pr.objid) + 1
        except:
            pi = -1
    print(objid, pi, obj["Subj"], obj["T"], obj["Contents"])


fp = open("simple1.pdf", "rb")
parser = PDFParser(fp)
doc = PDFDocument()
parser.set_document(doc)
doc.set_parser(parser)
doc.initialize(pdf_pwd)

visited = set()
for xref in doc.xrefs:
    for objid in xref.get_objids():
        if objid in visited: continue
        visited.add(objid)
        try:
            obj = doc.getobj(objid)
            if obj is None: continue
            extract(objid, obj)
            print("oldu.")
        except:
            print(sys.stderr, "not found: %r")
