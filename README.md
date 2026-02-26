# ODF Batch Styler

A CLI tool for automated, template-based styling of OpenDocument (`.odt`) files.

## 🚀 Overview

The **ODF Batch Styler** allows you to decouple document design from document processing. By using a "Golden Template" (a reference `.odt` file), you can define complex styles visually in a word processor and apply them programmatically to hundreds of files using regex patterns defined in a JSON configuration.

### Key Features

* **Template-Based**: Import Paragraph and Character styles directly from a reference document.
* **Intelligent Application**: Automatically detects if a style is a "Text" style (applying to specific words via `set_span`) or a "Paragraph" style (applying to the entire block).
* **Batch Processing**: Supports glob patterns (e.g., `docs/*.odt`) for bulk editing.
* **Dry Run Mode**: Preview targeted files without modifying them.
* **Data-Driven**: Logic is controlled entirely via `rules.json`.

## 🛠 Installation

1. **Clone the repository**:
```bash
git clone <your-repo-url>
cd odf-batch-styler
```

2. **Install dependencies**:
```bash
pip install -r requirements.txt
```

## 📖 User Guide

### 1. Prepare your Styles

Open your word processor (e.g., LibreOffice) and create a document named `template.odt`. Define the styles you want (e.g., a character style named `Critical` with red bold text).

### 2. Configure the Rules (`rules.json`)

Create a `rules.json` to map your patterns to your styles.

```json
{
  "modifications": [
    {
      "type": "import_style",
      "style_name": "Standard",
      "family": "paragraph",
      "source_file": "template_reference.odt"
    },
    {
      "type": "import_style",
      "style_name": "HighlightStyle",
      "family": "text",
      "source_file": "template_reference.odt"
    },
    {
      "type": "regex_span_styler",
      "rules": [
        {
          "pattern": "\\+\\+ IMPORTANT: (.*)",
          "style_name": "AlertStyle"
        }
      ]
    }
  ]
}
```

*Note: Remember to double-escape regex characters in JSON (e.g., `\\d` for digits).*

### 3. Run the Batch

```bash
python odt_editor.py "invoices/*.odt" --config rules.json --suffix "_PROCESSED"
```

## 💻 Developer Guide

### Architecture

The project follows the **Strategy Pattern**. The `BatchProcessor` manages the file system and document lifecycle, while `DocumentModifier` subclasses handle specific XML manipulations.

### Core Components

* **`StyleImporter`**: Injects XML style definitions from the reference document into the target document's manifest.
* **`RegexStyler`**: Utilizes `body.get_paragraphs(content=regex)` for high-performance searching. It dynamically chooses between `set_span` (for text-family styles) and paragraph-level style assignment.