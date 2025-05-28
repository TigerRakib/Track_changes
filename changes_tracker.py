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
# Extract track changes
def get_tracked_changes(xml_root):
    changes = []
    for change_type in ['ins', 'del']:
        for tag in xml_root.findall(f'.//w:{change_type}', ns):
            texts = tag.findall('.//w:t', ns)
            full_text = ''.join([t.text for t in texts if t.text])
            if full_text.strip():
                changes.append((change_type, full_text.strip()))
    return changes
# Extract all paragraphs 
def get_paragraphs(xml_root):
    paras = []
    for para in xml_root.findall('.//w:p', ns):
        texts = para.findall('.//w:t', ns)
        full_text = ''.join([t.text for t in texts if t.text])
        if full_text.strip():
            paras.append(full_text.strip())
    return paras

english_xml = get_document_xml(english_docx)
chinese_xml = get_document_xml(chinese_docx)
tracked_changes = get_tracked_changes(english_xml)
chinese_paras = get_paragraphs(chinese_xml)
print(chinese_paras)


