# ODF Batch Styler

A CLI tool for automated, template-based styling of OpenDocument (`.odt`) files.

## Overview

ODF Batch Styler decouples document design from document processing. You define styles visually in a reference `.odt` file (the "golden template"), then apply them to many documents using regex rules from a JSON configuration.

### Key Features

* **Template-based**: imports paragraph and character styles from a reference document.
* **Automatic style kind**: a *text* (character) style is applied only to the matching text; a *paragraph* style is applied to the whole paragraph.
* **Capture groups**: if a pattern has a group, only the group is styled (see [Rules](#regex-rules)).
* **Batch processing**: accepts glob patterns (e.g. `docs/*.odt`).
* **Safe output**: originals are never modified; results are written to `<name>_EDITED.odt` (configurable).
* **Dry run**: shows what each rule would do without writing any file.
* **Per-file summary**: one results table per processed file.

## Installation

The project uses [uv](https://docs.astral.sh/uv/) and depends on [odfdo](https://github.com/jdum/odfdo).

```bash
git clone <your-repo-url>
cd odf-batch-styler
uv sync
```

## Usage

```bash
uv run styler.py "docs/*.odt"                         # uses rules.json next to styler.py
uv run styler.py a.odt "docs/*.odt"                   # several files and/or globs
uv run styler.py "docs/*.odt" --config my_rules.json
uv run styler.py "docs/*.odt" --dry-run               # preview, writes nothing
uv run styler.py "docs/*.odt" --suffix "_PROCESSED"   # custom output suffix
```

| Option | Default | Description |
|---|---|---|
| `file` ... | (required) | One or more `.odt` paths or glob patterns. Quote globs so the shell doesn't expand them. Non-`.odt` matches are skipped. |
| `--config` | `rules.json` next to `styler.py` | Path to the JSON rules file. Relative `source_file` entries are resolved against this file's folder, so the tool works from any directory. |
| `--dry-run` | off | Run all rules in memory and report results, but save nothing. |
| `--suffix` | `_EDITED` | Appended to the output file name. |

Behaviour worth knowing:

* Output is written next to the input: `report.odt` becomes `report_EDITED.odt`.
* Files whose name already ends with the suffix are skipped, so re-running on a folder doesn't process previous output.
* Each file is processed independently; the summary counts are not carried over between files.

## Using from Windows (WSL)

The tool runs inside WSL, but you can launch it from Windows on files that live on `C:\`.

* [styler.bat](styler.bat) forwards its arguments to [styler-wsl.sh](styler-wsl.sh), which converts Windows paths (`C:\...`, `\\server\...`) to WSL paths and runs the script.
* **Drag & drop**: drop one or more `.odt` files onto `styler.bat` (or onto a shortcut to it). The window stays open until you press a key.
* **Command line** (cmd/PowerShell), with any extra arguments:
  ```bat
  styler.bat --dry-run "C:\Users\me\Documents\report.odt"
  styler.bat --suffix _PROCESSED "C:\docs\a.odt" "C:\docs\b.odt"
  ```
* Output goes next to each input file, as usual.

Notes: `styler.bat` assumes the repo is at `/home/trevi/odf-batch-styler` in your default WSL distro (edit the path inside it otherwise). Set `STYLER_NO_PAUSE=1` to skip the final pause when scripting. Dropping very many files at once can hit the Windows command-line length limit.

## Configuration

### 1. Prepare the reference document

In LibreOffice (or any ODF editor), create a document containing the styles you want, for example a character style `Critical` (red, bold) and a paragraph style `Heading_Custom`. Reference and input `.odt` files are git-ignored, so keep your own copies locally.

> Style names use their ODF internal form: spaces become `_20_` (e.g. "Testo libretto" is `Testo_20_libretto`).

### 2. Write `rules.json`

Rules are grouped in `modifications`, which run **in order**. Typically: first import the styles, then apply them.

```json
{
  "modifications": [
    {
      "type": "import_style",
      "rules": [
        { "style_name": "Critical",       "family": "text",      "source_file": "styleref.odt" },
        { "style_name": "Heading_Custom", "family": "paragraph", "source_file": "styleref.odt" }
      ]
    },
    {
      "type": "regex_styler",
      "rules": [
        { "pattern": "\\+IMPORTANT: (.*)", "style_name": "Critical" },
        { "pattern": "^Chapter \\d+",      "style_name": "Heading_Custom" }
      ]
    }
  ]
}
```

Remember to double-escape backslashes in JSON (`\\d` for `\d`). [template.json](template.json) is a minimal example; copy it to `rules.json` (git-ignored, so your personal rules stay local) and adapt it.

#### Style import (`import_style`)

| Field | Description |
|---|---|
| `style_name` | Name of the style in the reference document. |
| `family` | `text` or `paragraph`. |
| `source_file` | Path to the reference `.odt`. |

The result is reported per style as `Success`, `Not found in source` or `Error: ...`.

#### Regex rules (`regex_styler`)

This rule type applies both text and paragraph styles. For every paragraph matching `pattern`, the style kind decides what happens:

| Style kind | Effect |
|---|---|
| Text style | Applied to the matched text only. |
| Paragraph style | Applied to the whole paragraph; existing character spans in it are cleared first. |
| Not found in the document | Rule is skipped and reported as `Skipped: style not found in document`. |

A style must exist in the target document before a rule can use it, so import it first (or make sure it's already there). If a name exists as both a text and a paragraph style, the text style wins.

**Capture groups (text styles).** If the pattern contains a group, only the text captured by group 1 is styled, not the whole match. Pattern `\\+IMPORTANT: (.*)` styles what follows the label, not the label itself. Only group 1 of the first match in each paragraph is used, and all occurrences of that text in the paragraph are styled.

### Reading the output

```
Processing: AO33.odt
  Applied 'Versetti' -> 5
  ...
=================================================================
MODIFICATION                                            | RESULT
-----------------------------------------------------------------
Applied 'Versetti'                                      | 5
=================================================================
```

The number is how many spans (text styles) or paragraphs (paragraph styles) were changed. Rules that share a style name are added together in the table.

## Developer Guide

### Architecture

A `BatchProcessor` manages files and the document lifecycle; `DocumentModifier` subclasses perform the individual modifications (Strategy pattern). Each modifier's `apply()` returns a list of `(label, result)` tuples used for reporting.

* **`StyleImporter`**: copies a style definition from a reference document into the target.
* **`RegexStyler`**: `apply()` determines the style kind (text, paragraph, unknown) and delegates to `_apply_text_style()`, `_apply_paragraph_style()` or `_handle_unknown_style()`.
* **`main()`**: parses the CLI arguments and builds modifiers from the JSON config.

### Adding a new modifier

1. Subclass `DocumentModifier` and implement `apply()`.
2. Register a new `type` in `main()` that maps config rules to your class.

## License

See [LICENSE](LICENSE).
