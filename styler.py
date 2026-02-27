import glob
import json
import argparse
import re
from pathlib import Path
from odfdo import Document, Style, Span, remove_tree
from abc import ABC, abstractmethod

# =============================================================================
# BASE ARCHITECTURE
# =============================================================================

class DocumentModifier(ABC):
    """Abstract Base Class for ODF modification strategies."""
    def __init__(self, doc: Document):
        self.doc = doc

    @abstractmethod
    def apply(self) -> list[tuple[str, str | int]]:
        """Executes the modification and returns a list of action logs."""
        pass

# =============================================================================
# CONCRETE MODIFIERS
# =============================================================================

class StyleImporter(DocumentModifier):
    """
    Imports a style definition from a source ODT into the current document.
    """
    def __init__(self, doc: Document, style_name: str, family: str, source_file: str):
        super().__init__(doc)
        self.style_name = style_name
        self.family = family  # 'paragraph' or 'text'
        self.source_file = source_file

    def apply(self) -> list:
        """Extracts the style from reference and injects it into target."""
        try:
            ref_doc = Document(self.source_file)
            source_style = ref_doc.get_style(self.family, self.style_name)
            
            if type(source_style) == Style:
                self.doc.insert_style(source_style)
                return [(f"Import {self.family} style: {self.style_name}", "Success")]
            
            return [(f"Import {self.family} style: {self.style_name}", "Not found in source")]
        except Exception as e:
            return [(f"Import {self.family} style: {self.style_name}", f"Error: {str(e)}")]

class RegexStyler(DocumentModifier):
    """
    Applies an EXISTING style to text matching a regex pattern.
    Does not create or define styles; only applies them by name.
    """
    def __init__(self, doc: Document, style_name:str, regex: str):
        super().__init__(doc)
        self.style_name = style_name
        self.regex = regex

    def apply(self) -> list:
        """Applies named styles to matching regex groups."""
        body = self.doc.body
        match_count = 0
        
        # Apply style to all matching paragraphs
        for para in body.get_paragraphs(content=self.regex):
            # Check whether a text style
            if self.doc.get_style(family='text', name_or_element=self.style_name):
                # Then use odfdo's set_span to appliy style_name to matching pattern
                try:
                    group1 = re.search(self.regex, para.text_recursive).group(1)
                    self.regex = group1
                except IndexError:
                    pass
                spans = para.set_span(self.style_name, regex=self.regex)
                match_count += len(spans)
            else:
                # Else, format the whole paragraph 
                # Clear paragraph styles
                remove_tree(para, Span)
                para.style = self.style_name
                match_count += 1
                
        return [(f"Applied '{self.style_name}'", match_count)]

# =============================================================================
# PROCESSOR ENGINE
# =============================================================================

class BatchProcessor:
    """Manages the batch modification of multiple ODF files."""
    def __init__(self, file_pattern: str, dry_run: bool = False):
        self.files = glob.glob(file_pattern)
        self.dry_run = dry_run
        self.modifier_configs = []
        self.totals = {}

    def add_modifier_config(self, modifier_class: type, **kwargs):
        self.modifier_configs.append((modifier_class, kwargs))

    def run(self, output_suffix: str = "_EDITED"):
        if not self.files:
            print("No files matched the pattern.")
            return

        for file_path in self.files:
            path = Path(file_path)
            output_path = path.parent / f"{path.stem}{output_suffix}{path.suffix}"
            
            if self.dry_run:
                print(f"[DRY RUN] Would process: {path.name}")

            doc = Document(file_path)
            print(f"\nProcessing: {path.name}")
            
            for mod_class, kwargs in self.modifier_configs:
                modifier = mod_class(doc, **kwargs)
                logs = modifier.apply()
                for label, info in logs:
                    # print(f"  {label} -> {info}")
                    if isinstance(info, int):
                        self.totals[label] = self.totals.get(label, 0) + info
                    else:
                        self.totals[label] = info
            
            if not self.dry_run:
                doc.save(str(output_path))

            self._print_summary()

    def _print_summary(self):
        print("\n" + "="*65)
        print(f"{'MODIFICATION':<55} | {'RESULT'}")
        print("-" * 65)
        for label, info in self.totals.items():
            print(f"{label:<55} | {info}")
        print("="*65)

# =============================================================================
# MAIN ENTRY POINT
# =============================================================================

def main():
    parser = argparse.ArgumentParser(description="Template-Based ODF Styler")
    parser.add_argument("pattern", help="File glob pattern (e.g. '*.odt')")
    parser.add_argument("--config", default="rules.json", help="Path to JSON config")
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--suffix", default="_EDITED")
    
    args = parser.parse_args()

    try:
        with open(args.config, 'r') as f:
            config_data = json.load(f)
    except Exception as e:
        print(f"Error: Could not load {args.config} ({e})")
        return

    processor = BatchProcessor(args.pattern, dry_run=args.dry_run)

    for mod in config_data.get('modifications', []):
        if mod['type'] == 'import_style':
            for rule in mod['rules']:
                processor.add_modifier_config(
                    StyleImporter, 
                    style_name=rule['style_name'], 
                    family=rule['family'],
                    source_file=rule['source_file']
                )
        elif mod['type'] == 'regex_span_styler':
            for rule in mod['rules']:
                processor.add_modifier_config(
                    RegexStyler, 
                    style_name=rule['style_name'],
                    regex=rule['pattern']
                )

    processor.run(output_suffix=args.suffix)

if __name__ == "__main__":
    main()