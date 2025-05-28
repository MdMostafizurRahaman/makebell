# English-to-Chinese Track Changes Automation

## Overview

This script automates the process of transferring tracked changes from an English Word document (with Track Changes enabled) to its corresponding Chinese translation (original version). The output is a revised Chinese document with the same tracked changes (insertions, deletions, formatting) programmatically applied and visible in Microsoft Word.

## Features

- **Fully automated:** No manual editing required. The script extracts, matches, and applies changes end-to-end.
- **Track Changes preserved:** The output Chinese document displays all modifications using Word’s Track Changes, mirroring the English revision history.
- **Reproducible:** The process is deterministic and can be validated on any document pair with the same structure.

## How It Works

1. **Extracts tracked changes** from the English `.docx` using Microsoft Word automation.
2. **Translates** the changed segments and their context to Chinese (using Google Translate for alignment; translation quality is not evaluated).
3. **Matches** the translated context to the corresponding paragraph in the Chinese document.
4. **Applies the change** (insertion, deletion, or formatting) to the Chinese document using Word automation, with Track Changes enabled.
5. **Outputs** a new Chinese `.docx` with all changes visible in Track Changes.

## Requirements

- **Python 3.12** (not 3.13+)
- **Microsoft Word** (required for COM automation)
- **Dependencies:**  
  Install with:
  ```sh
  pip install pywin32 googletrans
  ```

## Usage

1. Place your files in the project folder:
    - `edited_en.docx` — English document with tracked changes
    - `original_cn.docx` — Chinese document (original version)

2. Run the script:
    ```sh
    python main.py
    ```

3. The output will be:
    - `original_cn_with_tracked_changes.docx` — Chinese document with tracked changes applied

## Notes

- **Close all Microsoft Word windows** before running the script.
- Only Python 3.12 or lower is supported due to library compatibility.
- The script is designed for reproducibility and can be validated on any document pair with the same structure.
- Translation is used only for alignment; translation quality is not part of the assessment.

## Evaluation Criteria

- **Accuracy:** Correct identification and application of tracked changes.
- **Consistency:** All modifications are tracked and visible in the output.
- **Automation:** No manual steps required; the process is fully script-driven.
- **Clarity:** The code and process are clear and easy to follow.

---

If you have any questions or need more time, please reach out.  
We look forward to your review and feedback!