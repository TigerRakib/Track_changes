import os
from zipfile import ZipFile
import xml.etree.ElementTree as ET
from docx import Document
from fuzzywuzzy import process
from shutil import copyfile

#  File Paths 
english_docx = "[Track Changes] KFS - AXA Global Strategic Bonds_E.docx"  # English with tracked changes
chinese_docx = "KFS - AXA Global Strategic Bonds_C 8.53.14 AM.docx"  # Chinese original
output_docx = "updated_chinese_with_changes.docx"  # Output path

#  Namespace 
ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

#  Extract document.xml content 
def get_document_xml(docx_path):
    with ZipFile(docx_path, 'r') as docx:
        xml = docx.read('word/document.xml')
    return ET.fromstring(xml)


english_xml = get_document_xml(english_docx)
chinese_xml = get_document_xml(chinese_docx)


