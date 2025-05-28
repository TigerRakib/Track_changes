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
# Apply change to paragraph using fuzzy match 
def match_change(change_text, paragraphs, threshold=60):
    match = process.extractOne(change_text, paragraphs)
    if match and match[1] >= threshold:
        return match[0]
    return None
# Apply insert/delete tag to text in XML 
def apply_changes_to_chinese(chinese_doc, changes, matched_paras):
    doc = Document(chinese_doc)
    for para in doc.paragraphs:
        original_text = para.text
        for change_type, change_text in changes:
            matched_text = matched_paras.get(change_text)
            if matched_text and matched_text in original_text:
                if change_type == 'ins':
                    para.text = original_text.replace(
                        matched_text,
                        f"{matched_text}（新增內容: {change_text}）"
                    )
                elif change_type == 'del':
                    para.text = original_text.replace(
                        matched_text,
                        f"{matched_text}（刪除內容: {change_text}）"
                    )
    return doc

english_xml = get_document_xml(english_docx)
chinese_xml = get_document_xml(chinese_docx)
tracked_changes = get_tracked_changes(english_xml)
chinese_paras = get_paragraphs(chinese_xml)
matched = {}
for _, text in tracked_changes:
    match = match_change(text, chinese_paras)
    if match:
        matched[text] = match

updated_chinese_doc = apply_changes_to_chinese(chinese_docx, tracked_changes, matched)

updated_chinese_doc.save(output_docx)
