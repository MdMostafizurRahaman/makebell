import difflib
import win32com.client as win32
from googletrans import Translator

translator = Translator()

def extract_changes_from_word(path):
    word = win32.gencache.EnsureDispatch('Word.Application')
    word.Visible = False
    doc = word.Documents.Open(path)
    changes = []

    raw_revs = list(doc.Revisions)
    skip_next = False

    for i, rev in enumerate(raw_revs):
        if skip_next:
            skip_next = False
            continue

        try:
            rev_type = rev.Type
            rev_text = rev.Range.Text.strip()
            context = rev.Range.Paragraphs(1).Range.Text.strip()
            formatted = rev.Range.Font.Bold

            # Check for a replace pattern (delete + insert)
            if rev_type == 2 and i + 1 < len(raw_revs):
                next_rev = raw_revs[i + 1]
                if next_rev.Type == 1:
                    next_text = next_rev.Range.Text.strip()
                    next_context = next_rev.Range.Paragraphs(1).Range.Text.strip()
                    if context == next_context:
                        changes.append({
                            'type': 'replace',
                            'text_deleted': rev_text,
                            'text_inserted': next_text,
                            'context': context,
                            'bold': bool(next_rev.Range.Font.Bold)
                        })
                        skip_next = True
                        continue  # Skip normal handling

            # Regular insert/delete/format/etc.
            if rev_type == 1:
                changes.append({'type': 'insert', 'text': rev_text, 'context': context, 'bold': bool(formatted)})
            elif rev_type == 2:
                changes.append({'type': 'delete', 'text': rev_text, 'context': context, 'bold': bool(formatted)})
            elif rev_type in (3, 4, 5, 6):
                changes.append({'type': 'format', 'text': rev_text, 'context': context, 'bold': bool(formatted)})
        except Exception as e:
            print(f"⚠️ Error on revision {i}: {e}")

    doc.Close(False)
    word.Quit()
    return changes

def translate_text(text):
    try:
        return translator.translate(text, src='en', dest='zh-cn').text
    except Exception as e:
        print(f"❌ Translation failed for '{text}': {e}")
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

change_count = {
    'insert': 0,
    'delete': 0,
    'replace': 0,
    'format': 0,
    'bold': 0,
    'skipped': 0
}

def apply_changes_to_chinese(chinese_doc_path, changes):
    word = win32.gencache.EnsureDispatch('Word.Application')
    word.Visible = False
    doc = word.Documents.Open(chinese_doc_path)
    doc.TrackRevisions = True

    paras = [p.Range.Text.strip() for p in doc.Paragraphs]

    for change in changes:
        if 'type' not in change:
            change_count['skipped'] += 1
            continue
        if change['type'] == 'replace':
            zh_old = translate_text(change.get('text_deleted'))
            zh_new = translate_text(change.get('text_inserted'))
            zh_context = translate_text(change.get('context'))
        else:
            zh_text = translate_text(change.get('text'))
            zh_context = translate_text(change.get('context'))

        if not zh_text or not zh_context:
            change_count['skipped'] += 1
            continue

        best_para = find_best_match(zh_context, paras)
        if not best_para:
            change_count['skipped'] += 1
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
                        change_count['delete'] += 1

                elif change['type'] == 'insert':
                    rng.InsertAfter(f"{zh_text}")
                    change_count['insert'] += 1

                elif change['type'] == 'replace':
                    start = rng.Text.find(zh_old)
                    if start >= 0:
                        replace_range = rng.Duplicate
                        replace_range.SetRange(rng.Start + start, rng.Start + start + len(zh_old))
                        replace_range.Text = zh_new
                        change_count['replace'] += 1
                    else:
                        change_count['skipped'] += 1

                elif change['type'] == 'format':
                    formatted_text = f"[FORMATTED:{zh_text}]"
                    if change.get('bold'):
                        formatted_text = f"[BOLD:{zh_text}]"
                        change_count['bold'] += 1
                    else:
                        change_count['format'] += 1
                    rng.InsertAfter(formatted_text)

                break


    print("\n📊 Change Summary:")
    for k, v in change_count.items():
        print(f"  - {k}: {v}")

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
