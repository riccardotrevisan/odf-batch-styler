import glob
import json
import argparse
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
    def apply(self) -> list:
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
                return [(f"Import {self.family}: {self.style_name}", "Success")]
            
            return [(f"Import {self.family}: {self.style_name}", "Not found in source")]
        except Exception as e:
            return [(f"Import {self.family}: {self.style_name}", f"Error: {str(e)}")]

class RegexStyler(DocumentModifier):
    """
    Applies an EXISTING style to text matching a regex pattern.
    Does not create or define styles; only applies them by name.
    """
    def __init__(self, doc: Document, rules: list):
        super().__init__(doc)
        self.rules = rules

    def apply(self) -> list:
        """Applies named styles to matching regex groups."""
        results = []
        body = self.doc.body
        
        for rule in self.rules:
            match_count = 0
            style_name = rule['style_name']
            regex = rule['pattern']
            
            # Apply style to all matching paragraphs
            for para in body.get_paragraphs(content=regex):
                # Clear paragraph styles
                remove_tree(para, Span)
                # Check whether a text style
                if self.doc.get_style(family='text', name_or_element=style_name):
                    # Then use odfdo's set_span to appliy style_name to matching pattern
                    match_count += len(para.set_span(style_name, regex=regex))
                else:
                    # Else, format the whole paragraph 
                    para.style = style_name
                    match_count += 1
                
            results.append((f"Applied '{style_name}'", match_count))
        return results

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
                continue

            doc = Document(file_path)
            print(f"\nProcessing: {path.name}")
            
            for mod_class, kwargs in self.modifier_configs:
                modifier = mod_class(doc, **kwargs)
                logs = modifier.apply()
                for label, info in logs:
                    print(f"  {label} -> {info}")
                    if isinstance(info, int):
                        self.totals[label] = self.totals.get(label, 0) + info

            doc.save(str(output_path))
        
        if not self.dry_run:
            self._print_summary()

    def _print_summary(self):
        print("\n" + "="*56)
        print(f"{'MODIFICATION':<40} | {'TOTAL MATCHES'}")
        print("-" * 56)
        for label, count in self.totals.items():
            print(f"{label:<40} | {count}")
        print("="*56)

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

    for mod_cfg in config_data.get('modifications', []):
        m_type = mod_cfg.get('type')
        if m_type == 'import_style':
            processor.add_modifier_config(
                StyleImporter, 
                style_name=mod_cfg['style_name'], 
                family=mod_cfg.get('family', 'paragraph'),
                source_file=mod_cfg['source_file']
            )
        elif m_type == 'regex_span_styler':
            processor.add_modifier_config(RegexStyler, rules=mod_cfg['rules'])

    processor.run(output_suffix=args.suffix)

if __name__ == "__main__":
    main()