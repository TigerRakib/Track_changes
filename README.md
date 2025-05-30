# 📄 DOCX Translator & Modifier

A powerful and flexible Python tool to automate translations and text modifications within `.docx` files. This script allows you to translate document content, remove specially formatted text, and insert custom-formatted text while preserving the original layout and style.

---

## ✨ Features

- **🌐 Full Document Translation**  
  Translate all editable text in a DOCX document into a specified target language (e.g., Simplified Chinese).

- **❌ Selective Text Removal**  
  Automatically detect and remove text segments formatted in **blue color** and with a **strikethrough effect**.

- **➕ Custom Text Insertion**  
  Add user-defined paragraphs into the document. These will be translated and styled in **blue and underlined**.

- **🧬 Format Preservation**  
  By directly modifying `document.xml` using `lxml`, the script maintains existing styles (fonts, sizes, bolding, etc.) as much as possible.

---

## 🛠️ Prerequisites

Make sure you have Python 3 installed, then install the required dependencies:

```bash
pip install lxml deep-translator python-docx
```

---

## 🚀 How It Works

1. **Unzip DOCX File**  
   `.docx` files are treated as ZIP archives. The script unzips them to a temporary folder.

2. **Parse Main Content**  
   It parses `word/document.xml` to locate and manipulate all text content.

3. **Translate Existing Text**  
   All `<w:t>` elements are extracted and sent to Google Translate via `deep_translator`.

4. **Remove Specific Runs**  
   Deletes `<w:r>` (run) elements that are both blue-colored and strikethrough.

5. **Translate & Insert Custom Text**  
   Your specified text is translated and added to the end of the document, formatted in blue and underlined.

6. **Save & Repack**  
   Updates the document XML, re-zips the folder into a `.docx` file, and deletes temporary files.

---

## 💻 Usage

1. **Save the Script**  
   Save the script as `track_changer.py`.

2. **Place Your Input File**  
   Place your `.docx` file (e.g., `demo.docx`) in the same directory.

3. **Run the Script**

```bash
track_changer.py
```

---

## 🔧 Configuration Example

Edit the configuration at the end of the script to customize file paths, translation text, and target language:

```python
# === Example usage ===
input_file = "demo.docx"                 # Path to your input updated english DOCX file
output_file = "output_full_chinese.docx" # Output file path
text_to_add = ""
target_language = "zh-CN"                # Target language code (e.g., 'en', 'zh-CN', 'es')

modify_docx_preserve_format(input_file, output_file, text_to_add, target_language)
```

---

## ⚠️ Troubleshooting

- **Script Freezes During Translation**  
  - Check your internet connection.
  - Add a delay (e.g., `time.sleep(0.5)`) between translations to avoid Google rate limits.
  - Update `deep_translator`:  
    ```bash
    pip install --upgrade deep-translator
    ```

- **Output DOCX Is Corrupted**  
  - Ensure XML is well-formed and correctly modified.
  - Make sure files are properly re-zipped into `.docx` format.

- **No Translation Occurs**  
  - Verify `target_language` is valid.
  - Test `deep_translator` with a small string separately to confirm it's working.

---

## 🤝 Contributing

Contributions are welcome!  
Fork the repo, report issues, or submit pull requests for improvements and new features.

---

## 📄 License

This project is open-source and available under the [MIT License](LICENSE).

---

## 📬 Contact

Have questions or suggestions? Feel free to open an issue or start a discussion.

May Allah bless you. 
