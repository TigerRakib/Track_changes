
# Word Track Changes Translator Tool

This tool automates the transfer of tracked changes from an English Word document to its corresponding Chinese translation, preserving edits with visible annotations.

##  Use Case

Given:
- An English `.docx` file with **tracked changes**
- A Chinese `.docx` file that is a direct translation of the **original** English version

This tool will:
1. Extract tracked changes (insertions & deletions) from the English file.
2. Fuzzy match the changed segments to the Chinese content.
3. Insert Chinese annotations indicating insertions/deletions.
4. Save a new `.docx` Chinese file with those changes annotated.

##  File Structure

```
project/
│
├──[Track Changes] KFS - AXA Global Strategic Bonds_E.docxenglish.docx   # English Word doc with tracked changes
├──KFS - AXA Global Strategic Bonds_C 8.53.14 AM.docx.docx               # Chinese version (original translation)
├──track_changer.py                                                      # Python script (main logic)
├──updated_chinese_with_changes.docx                                     # Output file (auto-generated)
└──README.md                                                             # This file
```

## Requirements

Install the required Python packages:

```bash
pip install python-docx lxml fuzzywuzzy python-Levenshtein
```

## Usage

1. Place your Word documents in the project directory and rename them:
   - `english.docx`
   - `chinese.docx`

2. Run the script:

```bash
python script.py
```

3. Output:
   - `updated_chinese_with_changes.docx` will contain visible annotations like:

     - `（新增內容: XX）` → Inserted content
     - `（刪除內容: XX）` → Deleted content

##  Example Annotation

If a phrase was inserted in English and matched to its Chinese equivalent, the Chinese document will include:

```
原有文本（新增內容: 插入內容）
```

If something was deleted:

```
原有文本（刪除內容: 刪除的內容）
```

## Note on Track Changes in Word

This version uses visible annotations in parentheses to simulate tracked changes.


## Limitations

- Fuzzy matching relies on partial text alignment; translation consistency improves accuracy.
- Manual validation recommended for critical documents.
- This script does not preserve Word-specific styling or formatting.

