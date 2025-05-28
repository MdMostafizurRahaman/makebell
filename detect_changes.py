import zipfile
import shutil
import os
from lxml import etree as ET
import tempfile

# Namespaces for WordprocessingML
NS = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
}

def extract_document_xml(docx_path):
    with zipfile.ZipFile(docx_path) as z:
        with z.open('word/document.xml') as f:
            xml = f.read()
    return xml

def write_document_xml(docx_path, new_xml_bytes, output_path):
    # Copy original docx to output
    shutil.copy(docx_path, output_path)
    with zipfile.ZipFile(output_path, 'a') as z:
        # Overwrite document.xml with new content
        z.writestr('word/document.xml', new_xml_bytes)

def parse_changes_from_en_docx(docx_path):
    """
    Parse English DOCX with tracked changes.
    Return list of changes as dicts:
    {'paragraph': int, 'run': int, 'type': str, 'text': str}
    Types: add, delete, replace, format, format_bold, format_list
    """
    changes = []
    with zipfile.ZipFile(docx_path) as docx_zip:
        with docx_zip.open("word/document.xml") as document_xml:
            tree = ET.parse(document_xml)
            root = tree.getroot()
            paragraphs = root.findall('.//w:body/w:p', NS)

            for p_idx, para in enumerate(paragraphs):
                is_list = para.find('.//w:numPr', NS) is not None
                runs = para.findall('.//w:r', NS)
                prev_del_text = None
                for r_idx, run in enumerate(runs):
                    text_el = run.find('.//w:t', NS)
                    if text_el is None or not text_el.text:
                        continue
                    text = text_el.text.strip()

                    rpr = run.find('.//w:rPr', NS)
                    is_bold = rpr is not None and rpr.find('.//w:b', NS) is not None
                    is_formatted = rpr is not None

                    # Check if run or parent has <w:ins> or <w:del>
                    # We check ancestors by climbing up parents
                    parent = run.getparent()
                    change_type = None
                    while parent is not None and parent.tag != '{%s}p' % NS['w']:
                        tag_local = ET.QName(parent).localname
                        if tag_local == 'ins':
                            if prev_del_text:
                                change_type = 'replace'
                            else:
                                change_type = 'add'
                            break
                        elif tag_local == 'del':
                            prev_del_text = text
                            change_type = 'delete'
                            break
                        parent = parent.getparent()

                    if change_type == 'replace':
                        changes.append({'paragraph': p_idx, 'run': r_idx, 'type': 'replace', 'text': prev_del_text + " -> " + text})
                        prev_del_text = None
                    elif change_type in ('add', 'delete'):
                        changes.append({'paragraph': p_idx, 'run': r_idx, 'type': change_type, 'text': text})
                    else:
                        # No add/del parent, but maybe formatting
                        if is_bold:
                            changes.append({'paragraph': p_idx, 'run': r_idx, 'type': 'format_bold', 'text': text})
                        elif is_formatted:
                            changes.append({'paragraph': p_idx, 'run': r_idx, 'type': 'format', 'text': text})

                if is_list:
                    # For lists mark format_list
                    para_text = ''.join(t.text for t in para.findall('.//w:t', NS) if t.text)
                    if para_text:
                        changes.append({'paragraph': p_idx, 'run': None, 'type': 'format_list', 'text': para_text})
    return changes

def insert_track_changes_to_cn_docx(en_changes, cn_docx_path, output_path):
    with zipfile.ZipFile(cn_docx_path) as docx_zip:
        with docx_zip.open("word/document.xml") as document_xml:
            tree = ET.parse(document_xml)
            root = tree.getroot()
            paragraphs = root.findall('.//w:body/w:p', NS)

            for change in en_changes:
                p_idx = change['paragraph']
                r_idx = change['run']
                ctype = change['type']
                if p_idx >= len(paragraphs):
                    continue
                para = paragraphs[p_idx]
                runs = para.findall('.//w:r', NS)

                if ctype == 'add':
                    if r_idx is not None and r_idx < len(runs):
                        run = runs[r_idx]
                        # Wrap run inside <w:ins>
                        ins = ET.Element('{%s}ins' % NS['w'])
                        run_copy = ET.fromstring(ET.tostring(run))
                        ins.append(run_copy)
                        para.replace(run, ins)
                    else:
                        ins = ET.Element('{%s}ins' % NS['w'])
                        r = ET.Element('{%s}r' % NS['w'])
                        t = ET.Element('{%s}t' % NS['w'])
                        t.text = change['text']
                        r.append(t)
                        ins.append(r)
                        para.append(ins)

                elif ctype == 'delete':
                    if r_idx is not None and r_idx < len(runs):
                        run = runs[r_idx]
                        dele = ET.Element('{%s}del' % NS['w'])
                        run_copy = ET.fromstring(ET.tostring(run))
                        dele.append(run_copy)
                        para.replace(run, dele)

                elif ctype == 'replace':
                    if r_idx is not None and r_idx < len(runs):
                        run = runs[r_idx]
                        # Wrap original run as deleted
                        dele = ET.Element('{%s}del' % NS['w'])
                        run_copy = ET.fromstring(ET.tostring(run))
                        dele.append(run_copy)
                        para.replace(run, dele)
                        # Insert new run with inserted text after dele element
                        ins = ET.Element('{%s}ins' % NS['w'])
                        r = ET.Element('{%s}r' % NS['w'])
                        t = ET.Element('{%s}t' % NS['w'])
                        # New text is after '->' in change text
                        new_text = change['text'].split('->')[-1].strip()
                        t.text = new_text
                        r.append(t)
                        ins.append(r)
                        # Find dele index
                        dele_idx = para.index(dele)
                        para.insert(dele_idx + 1, ins)
                    else:
                        # If no run found, just add inserted text
                        ins = ET.Element('{%s}ins' % NS['w'])
                        r = ET.Element('{%s}r' % NS['w'])
                        t = ET.Element('{%s}t' % NS['w'])
                        new_text = change['text'].split('->')[-1].strip()
                        t.text = new_text
                        r.append(t)
                        ins.append(r)
                        para.append(ins)

                elif ctype == 'format_bold':
                    if r_idx is not None and r_idx < len(runs):
                        run = runs[r_idx]
                        rPr = run.find('{%s}rPr' % NS['w'])
                        if rPr is None:
                            rPr = ET.SubElement(run, '{%s}rPr' % NS['w'])
                        b = rPr.find('{%s}b' % NS['w'])
                        if b is None:
                            ET.SubElement(rPr, '{%s}b' % NS['w'])

                elif ctype == 'format':
                    if r_idx is not None and r_idx < len(runs):
                        run = runs[r_idx]
                        rPr = run.find('{%s}rPr' % NS['w'])
                        if rPr is None:
                            rPr = ET.SubElement(run, '{%s}rPr' % NS['w'])
                        highlight = rPr.find('{%s}highlight' % NS['w'])
                        if highlight is None:
                            highlight = ET.SubElement(rPr, '{%s}highlight' % NS['w'])
                            highlight.set('{%s}val' % NS['w'], 'yellow')

                elif ctype == 'format_list':
                    numPr = para.find('{%s}numPr' % NS['w'])
                    if numPr is None:
                        numPr = ET.SubElement(para, '{%s}numPr' % NS['w'])
                    ilvl = numPr.find('{%s}ilvl' % NS['w'])
                    if ilvl is None:
                        ilvl = ET.SubElement(numPr, '{%s}ilvl' % NS['w'])
                    ilvl.set('{%s}val' % NS['w'], '0')
                    numId = numPr.find('{%s}numId' % NS['w'])
                    if numId is None:
                        numId = ET.SubElement(numPr, '{%s}numId' % NS['w'])
                    numId.set('{%s}val' % NS['w'], '1')

    new_xml = ET.tostring(root, encoding='UTF-8', xml_declaration=True, pretty_print=True)
    write_document_xml(cn_docx_path, new_xml, output_path)


def main(en_docx_path, cn_docx_path, output_path):
    print(f"Parsing English DOCX for changes...")
    en_changes = parse_changes_from_en_docx(en_docx_path)
    print(f"Detected {len(en_changes)} changes in English DOCX")

    print(f"Inserting changes into Chinese DOCX...")
    insert_track_changes_to_cn_docx(en_changes, cn_docx_path, output_path)
    print(f"Output saved to {output_path}")

if __name__ == "__main__":
    import sys
    if len(sys.argv) != 4:
        print("Usage: python script.py english_with_track_changes.docx chinese_without_changes.docx output_chinese_with_track_changes.docx")
    else:
        en_docx_path = sys.argv[1]
        cn_docx_path = sys.argv[2]
        output_path = sys.argv[3]
        main(en_docx_path, cn_docx_path, output_path)
