import difflib
import time
import win32com.client as win32
from googletrans import Translator

translator = Translator()

def extract_changes_from_word(path):
    word = win32.gencache.EnsureDispatch('Word.Application')
    word.Visible = False
    doc = word.Documents.Open(path)
    doc.TrackRevisions = True

    changes = []

    for rev in doc.Revisions:
        try:
            if rev.Type in [1, 2]:  # 1 = insert, 2 = delete
                changes.append({
                    'type': 'insert' if rev.Type == 1 else 'delete',
                    'text': rev.Range.Text.strip(),
                    'context': rev.Range.Paragraphs(1).Range.Text.strip()
                })
        except Exception as e:
            print(f"Error reading revision: {e}")
    
    doc.Close(False)
    word.Quit()
    return changes

def translate_text(text):
    try:
        result = translator.translate(text, src='en', dest='zh-cn')
        return result.text
    except Exception as e:
        print(f"Translation failed: {text} -> {e}")
        return None


def find_best_match(target, paragraph_list):
    max_score = 0
    best = None
    for p in paragraph_list:
        score = difflib.SequenceMatcher(None, target, p).ratio()
        if score > max_score:
            max_score = score
            best = p
    return best

def apply_changes_to_chinese(chinese_doc_path, changes):
    word = win32.gencache.EnsureDispatch('Word.Application')
    word.Visible = False
    doc = word.Documents.Open(chinese_doc_path)
    doc.TrackRevisions = True

    paras = [p.Range.Text.strip() for p in doc.Paragraphs]

    for change in changes:
        zh_text = translate_text(change['text'])
        zh_context = translate_text(change['context'])

        if not zh_text or not zh_context:
            continue

        best_para = find_best_match(zh_context, paras)
        if not best_para:
            continue

        for p in doc.Paragraphs:
            if best_para.strip() == p.Range.Text.strip():
                rng = p.Range
                if change['type'] == 'delete':
                    start = rng.Text.find(zh_text)
                    if start >= 0:
                        delete_range = rng.Duplicate
                        delete_range.SetRange(rng.Start + start, rng.Start + start + len(zh_text))
                        delete_range.Delete()
                elif change['type'] == 'insert':
                    rng.InsertAfter(f"{zh_text}")
                break

    output_path = chinese_doc_path.replace('.docx', '_with_tracked_changes.docx')
    doc.SaveAs(output_path)
    doc.Close()
    word.Quit()
    return output_path

def main():
    import os
    base_dir = os.path.dirname(os.path.abspath(__file__))
    english_doc = os.path.join(base_dir, "edited_en.docx")
    chinese_doc = os.path.join(base_dir, "original_cn.docx")

    print("📥 Extracting tracked changes from English...")
    changes = extract_changes_from_word(english_doc)

    print("🌐 Translating and mapping changes to Chinese document...")
    result_path = apply_changes_to_chinese(chinese_doc, changes)

    print(f"\n✅ Done! Output saved as: {result_path}")

if __name__ == "__main__":
    main()
