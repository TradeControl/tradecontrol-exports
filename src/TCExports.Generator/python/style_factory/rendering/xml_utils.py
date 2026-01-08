from lxml import etree as ET

OFFICE_NS = "urn:oasis:names:tc:opendocument:xmlns:office:1.0"
STYLE_NS  = "urn:oasis:names:tc:opendocument:xmlns:style:1.0"
NUMBER_NS = "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"
TABLE_NS  = "urn:oasis:names:tc:opendocument:xmlns:table:1.0"
TEXT_NS   = "urn:oasis:names:tc:opendocument:xmlns:text:1.0"
FO_NS     = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"

def q(ns: str, local: str) -> ET.QName:
    return ET.QName(ns, local)