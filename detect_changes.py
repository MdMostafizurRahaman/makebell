import zipfile
import xml.etree.ElementTree as ET
from docx import Document
from googletrans import Translator
import difflib

def extract_tracked_changes(docx_path):
    changes = []
    with zipfile.ZipFile(docx_path) as docx_zip:
        with docx_zip.open("word/document.xml") as document_xml:
            tree = ET.parse(document_xml)
            root = tree.getroot()
            namespaces = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

            for para_index, para in enumerate(root.findall('.//w:p', namespaces)):
                inserted = para.findall('.//w:ins//w:t', namespaces)
                deleted = para.findall('.//w:del//w:t', namespaces)
                if inserted:
                    text = ''.join([t.text for t in inserted if t.text])
                    changes.append({
                        'paragraph': para_index,
                        'type': 'insert',
                        'text': text
                    })
                if deleted:
                    text = ''.join([t.text for t in deleted if t.text])
                    changes.append({
                        'paragraph': para_index,
                        'type': 'delete',
                        'text': text
                    })
    return changes

def translate_to_chinese(texts):
    translator = Translator()
    translations = []
    for text in texts:
        try:
            translated = translator.translate(text, src='en', dest='zh-cn').text
            translations.append(translated)
        except Exception as e:
            translations.append(f"Error translating: {text}")
    return translations

def load_chinese_paragraphs(docx_path):
    doc = Document(docx_path)
    return [p.text.strip() for p in doc.paragraphs if p.text.strip()]

def match_translations_to_chinese(translations, chinese_paragraphs):
    matches = []
    for translation in translations:
        best_match = None
        best_ratio = 0
        best_index = -1
        for i, para in enumerate(chinese_paragraphs):
            ratio = difflib.SequenceMatcher(None, translation, para).ratio()
            if ratio > best_ratio:
                best_ratio = ratio
                best_match = para
                best_index = i
        matches.append({
            'translated_text': translation,
            'matched_paragraph': best_match,
            'paragraph_index': best_index,
            'similarity': round(best_ratio, 2)
        })
    return matches

# === Main Program ===
if __name__ == "__main__":
    english_docx_path = "edited_en.docx"
    chinese_docx_path = "original_cn.docx"

    tracked_changes = extract_tracked_changes(english_docx_path)
    changed_texts = [change['text'] for change in tracked_changes]

    translated_texts = translate_to_chinese(changed_texts)
    chinese_paragraphs = load_chinese_paragraphs(chinese_docx_path)
    matches = match_translations_to_chinese(translated_texts, chinese_paragraphs)

    print("\n=== Change Report ===")
    for change, match in zip(tracked_changes, matches):
        print(f"\nChange Type: {change['type']}")
        print(f"English Text: {change['text']}")
        print(f"Translated Chinese: {match['translated_text']}")
        print(f"Matched Chinese Paragraph (#{match['paragraph_index']}): {match['matched_paragraph']}")
        print(f"Similarity: {match['similarity']}")
