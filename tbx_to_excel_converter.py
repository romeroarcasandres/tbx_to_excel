#!/usr/bin/env python3
"""
TBX to Excel Converter
Converts TermBase eXchange (.tbx) files to Excel (.xlsx) format with columnar structure
"""

import xml.etree.ElementTree as ET
import pandas as pd
import argparse
import sys
from pathlib import Path
from typing import List, Dict, Any, Optional
from collections import defaultdict


class TBXConverter:
    def __init__(self, tbx_file_path: str):
        """
        Initialize the TBX converter
        
        Args:
            tbx_file_path (str): Path to the TBX file
        """
        self.tbx_file_path = Path(tbx_file_path)
        self.namespace = {}
        self.terms_data = []
        self.entries_data = {}  # Store entries grouped by entry_id
        self.available_fields = set()
        self.selected_fields = []
        self.field_mappings = {}  # Maps original field names to user-chosen names
        
    def _scan_available_fields(self) -> None:
        """Scan the TBX file to identify all available data fields"""
        print("Scanning TBX file to identify available data fields...")
        
        try:
            tree = ET.parse(self.tbx_file_path)
            root = tree.getroot()
            
            # Register namespaces
            self._register_namespaces(root)
            
            # Find all termEntry elements
            term_entries = root.findall('.//termEntry')
            if not term_entries and 'default' in self.namespace:
                term_entries = root.findall(f'.//{self.namespace["default"]}termEntry')
            
            # If still not found, try without namespace prefix
            if not term_entries:
                for elem in root.iter():
                    if 'termentry' in elem.tag.lower():
                        term_entries.append(elem)
            
            # Scan through all entries to find available fields
            for entry in term_entries[:3]:  # Sample first 3 entries for performance
                # Find language groups (langSet or langGrp)
                lang_groups = []
                lang_groups.extend(entry.findall('.//langSet'))
                lang_groups.extend(entry.findall('.//langGrp'))
                
                for lang_grp in lang_groups:
                    # Find term groups (tig or termGrp)
                    term_groups = []
                    term_groups.extend(lang_grp.findall('.//tig'))
                    term_groups.extend(lang_grp.findall('.//termGrp'))
                    
                    for term_grp in term_groups:
                        # Scan all child elements
                        for elem in term_grp.iter():
                            tag_name = elem.tag.lower()
                            
                            # Always include the basic term
                            if tag_name == 'term':
                                self.available_fields.add('term')
                            
                            # Include termNote elements with their types
                            elif 'note' in tag_name:
                                note_type = elem.get('type', 'note')
                                self.available_fields.add(f"termNote_{note_type}")
                            
                            # Include descrip elements with their types
                            elif 'descrip' in tag_name:
                                descrip_type = elem.get('type', 'description')
                                self.available_fields.add(f"descrip_{descrip_type}")
                            
                            # Include other relevant elements
                            elif tag_name in ['definition', 'context', 'example', 'note']:
                                self.available_fields.add(tag_name)
                
                # Also check entry-level descriptions
                for elem in entry.iter():
                    if 'descrip' in elem.tag.lower():
                        descrip_type = elem.get('type', 'description')
                        self.available_fields.add(f"entry_descrip_{descrip_type}")
            
            # Add standard fields that are always available
            self.available_fields.update(['entry_id', 'language', 'term'])
            
        except Exception as e:
            print(f"Error scanning file: {e}")
            # Add some default fields as fallback
            self.available_fields.update([
                'entry_id', 'language', 'term', 'termNote_status', 
                'termNote_forbidden', 'termNote_preferred'
            ])
    
    def _interactive_field_selection(self) -> None:
        """Interactive field selection by user"""
        if not self.available_fields:
            self._scan_available_fields()
        
        sorted_fields = sorted(list(self.available_fields))
        
        print(f"\nFound {len(sorted_fields)} available data fields in your TBX file:")
        print("=" * 60)
        
        for i, field in enumerate(sorted_fields, 1):
            print(f"{i:2d}. {field}")
        
        print("\nWhich fields would you like to include in the Excel output?")
        print("Enter the numbers separated by commas (e.g., 1,3,5,7) or type 'all' for all fields:")
        
        while True:
            user_input = input("> ").strip()
            
            if user_input.lower() == 'all':
                self.selected_fields = sorted_fields[:]
                break
            
            try:
                # Parse comma-separated numbers
                selected_numbers = [int(x.strip()) for x in user_input.split(',')]
                
                # Validate numbers
                invalid_numbers = [n for n in selected_numbers if n < 1 or n > len(sorted_fields)]
                if invalid_numbers:
                    print(f"Invalid numbers: {invalid_numbers}. Please enter numbers between 1 and {len(sorted_fields)}")
                    continue
                
                # Convert numbers to field names
                self.selected_fields = [sorted_fields[n-1] for n in selected_numbers]
                break
                
            except ValueError:
                print("Invalid input. Please enter numbers separated by commas or type 'all'")
        
        print(f"\nSelected {len(self.selected_fields)} fields:")
        for field in self.selected_fields:
            print(f"  - {field}")
    
    def _interactive_field_renaming(self) -> None:
        """Interactive field renaming by user"""
        if not self.selected_fields:
            return
        
        print(f"\nWould you like to keep the original field names or rename them? (keep/rename)")
        
        while True:
            choice = input("> ").strip().lower()
            if choice in ['keep', 'rename']:
                break
            print("Please enter 'keep' or 'rename'")
        
        if choice == 'keep':
            # Use original names
            self.field_mappings = {field: field for field in self.selected_fields}
            print("Using original field names.")
            return
        
        print(f"\nWould you like to rename each field individually? (y/n)")
        
        while True:
            individual_choice = input("> ").strip().lower()
            if individual_choice in ['y', 'n', 'yes', 'no']:
                break
            print("Please enter 'y' or 'n'")
        
        if individual_choice in ['n', 'no']:
            print("Using original field names.")
            self.field_mappings = {field: field for field in self.selected_fields}
            return
        
        # Individual renaming
        print(f"\nRenaming fields individually:")
        print("Press Enter to keep the original name for any field.")
        
        for field in self.selected_fields:
            print(f"\nCurrent name: '{field}'")
            new_name = input(f"New name (or press Enter to keep '{field}'): ").strip()
            
            if new_name:
                self.field_mappings[field] = new_name
                print(f"  Renamed '{field}' → '{new_name}'")
            else:
                self.field_mappings[field] = field
                print(f"  Keeping '{field}'")
        
        print(f"\nFinal column mappings:")
        for original, new in self.field_mappings.items():
            if original != new:
                print(f"  '{original}' → '{new}'")
            else:
                print(f"  '{original}' (unchanged)")
    
    def configure_extraction(self) -> None:
        """Run the interactive configuration process"""
        print("=" * 60)
        print("TBX TO EXCEL CONVERTER - CONFIGURATION")
        print("=" * 60)
        
        self._interactive_field_selection()
        self._interactive_field_renaming()
        
        print(f"\nConfiguration complete!")
        print(f"Ready to extract {len(self.selected_fields)} fields from your TBX file.")
        print("=" * 60)
        
    def _register_namespaces(self, root: ET.Element) -> None:
        """Register XML namespaces found in the document"""
        # Common TBX namespaces
        common_namespaces = {
            '': 'http://www.lisa.org/TBX-Specification.33.0.html',
            'xml': 'http://www.w3.org/XML/1998/namespace',
            'xlink': 'http://www.w3.org/1999/xlink'
        }
        
        # Extract namespaces from root element
        for prefix, uri in root.attrib.items():
            if prefix.startswith('xmlns'):
                ns_prefix = prefix[6:] if prefix.startswith('xmlns:') else ''
                common_namespaces[ns_prefix] = uri
                
        # Register namespaces
        for prefix, uri in common_namespaces.items():
            if prefix:
                ET.register_namespace(prefix, uri)
                self.namespace[prefix] = f"{{{uri}}}"
            else:
                self.namespace['default'] = f"{{{uri}}}"
    
    def _extract_term_info(self, lang_grp: ET.Element) -> Dict[str, Any]:
        """
        Extract term information from a langGrp/langSet element
        
        Args:
            lang_grp: langGrp/langSet XML element
            
        Returns:
            Dictionary containing term information
        """
        term_info = {}
        
        # Get language code - try multiple attributes
        lang_code = (lang_grp.get('xml:lang') or 
                    lang_grp.get('lang') or 
                    lang_grp.get('{http://www.w3.org/XML/1998/namespace}lang', ''))
        term_info['language'] = lang_code
        
        print(f"  Processing language group: {lang_code}")
        
        # Find termGrp or tig elements - try multiple approaches
        term_groups = []
        
        # Direct search for tig (which your file uses)
        term_groups.extend(lang_grp.findall('.//tig'))
        term_groups.extend(lang_grp.findall('tig'))
        
        # Also try termGrp for other TBX variants
        term_groups.extend(lang_grp.findall('.//termGrp'))
        term_groups.extend(lang_grp.findall('termGrp'))
        
        # Try with various namespaces
        for ns_prefix, ns_uri in self.namespace.items():
            if ns_prefix != 'default':
                continue
            term_groups.extend(lang_grp.findall(f'.//{ns_uri}tig'))
            term_groups.extend(lang_grp.findall(f'.//{ns_uri}termGrp'))
        
        # Remove duplicates
        term_groups = list({id(tg): tg for tg in term_groups}.values())
        
        print(f"    Found {len(term_groups)} term groups")
        
        terms = []
        seen_terms = set()  # To avoid duplicate terms
        
        for term_grp in term_groups:
            # Initialize term data with all selected fields
            term_data = {field: '' for field in self.selected_fields if field != 'entry_id'}
            term_data['language'] = lang_code
            
            # Extract term text - try multiple approaches
            term_elem = None
            
            # Direct search
            term_elem = term_grp.find('.//term')
            if term_elem is None:
                term_elem = term_grp.find('term')
            
            # Try with namespace
            if term_elem is None and 'default' in self.namespace:
                term_elem = term_grp.find(f'.//{self.namespace["default"]}term')
            
            # If still not found, look for any element with 'term' in the name
            if term_elem is None:
                for child in term_grp.iter():
                    if 'term' in child.tag.lower():
                        term_elem = child
                        break
            
            if term_elem is not None:
                term_text = (term_elem.text or '').strip()
                
                # Skip empty terms or duplicates
                if not term_text or term_text in seen_terms:
                    continue
                
                seen_terms.add(term_text)
                print(f"      Found term: '{term_text}'")
                
                if 'term' in self.selected_fields:
                    term_data['term'] = term_text
                
                # Extract data for all selected fields
                for elem in term_grp.iter():
                    tag_name = elem.tag.lower()
                    elem_text = (elem.text or '').strip()
                    
                    if not elem_text:  # Skip empty elements
                        continue
                    
                    # Handle termNote elements
                    if 'note' in tag_name:
                        note_type = elem.get('type', 'note')
                        field_name = f"termNote_{note_type}"
                        if field_name in self.selected_fields:
                            term_data[field_name] = elem_text
                    
                    # Handle descrip elements  
                    elif 'descrip' in tag_name:
                        descrip_type = elem.get('type', 'description')
                        field_name = f"descrip_{descrip_type}"
                        if field_name in self.selected_fields:
                            term_data[field_name] = elem_text
                    
                    # Handle other elements
                    elif tag_name in self.selected_fields:
                        term_data[tag_name] = elem_text
                        
                terms.append(term_data)
            else:
                print(f"      Warning: No term element found in term group")
        
        term_info['terms'] = terms
        return term_info
    
    def parse_tbx(self) -> None:
        """Parse the TBX file and extract terminology data"""
        try:
            tree = ET.parse(self.tbx_file_path)
            root = tree.getroot()
            
            print(f"Root element: {root.tag}")
            
            # Register namespaces
            self._register_namespaces(root)
            print(f"Registered namespaces: {list(self.namespace.keys())}")
            
            # Find all termEntry elements
            term_entries = root.findall('.//termEntry')
            if not term_entries and 'default' in self.namespace:
                term_entries = root.findall(f'.//{self.namespace["default"]}termEntry')
            
            # If still not found, try without namespace prefix
            if not term_entries:
                for elem in root.iter():
                    if 'termentry' in elem.tag.lower():
                        term_entries.append(elem)
            
            print(f"Found {len(term_entries)} term entries")
            
            # Process each entry and store by entry_id
            for i, entry in enumerate(term_entries, start=1):
                entry_id = entry.get('id', f'entry_{i}')
                print(f"Processing entry {i+1}: {entry_id}")
                
                # Initialize entry data
                entry_data = {
                    'entry_id': entry_id,
                    'languages': {}  # Will store language -> [terms] mapping
                }
                
                # Extract entry-level description fields
                for elem in entry.iter():
                    if 'descrip' in elem.tag.lower():
                        descrip_type = elem.get('type', 'description')
                        field_name = f"entry_descrip_{descrip_type}"
                        if field_name in self.selected_fields and elem.text and elem.text.strip():
                            entry_data[field_name] = elem.text.strip()
                    elif 'subject' in elem.tag.lower():
                        if 'entry_subject' in self.selected_fields and elem.text and elem.text.strip():
                            entry_data['entry_subject'] = elem.text.strip()
                
                # Find language groups
                lang_groups = []
                
                # Try langSet first (which is what your file uses)
                lang_groups.extend(entry.findall('.//langSet'))
                lang_groups.extend(entry.findall('langSet'))
                
                # Try langGrp
                if not lang_groups:
                    lang_groups.extend(entry.findall('.//langGrp'))
                    lang_groups.extend(entry.findall('langGrp'))
                
                # Try with namespace
                if not lang_groups and 'default' in self.namespace:
                    lang_groups.extend(entry.findall(f'.//{self.namespace["default"]}langSet'))
                    lang_groups.extend(entry.findall(f'.//{self.namespace["default"]}langGrp'))
                
                # If no standard elements found, look for elements with 'lang' in the name
                if not lang_groups:
                    for elem in entry.iter():
                        if 'lang' in elem.tag.lower() and ('grp' in elem.tag.lower() or 'set' in elem.tag.lower()):
                            lang_groups.append(elem)
                
                print(f"  Found {len(lang_groups)} language groups")
                
                # Extract terms for each language
                for lang_grp in lang_groups:
                    lang_info = self._extract_term_info(lang_grp)
                    lang_code = lang_info['language']
                    if lang_code and lang_info['terms']:
                        entry_data['languages'][lang_code] = lang_info['terms']
                
                print(f"  Languages with terms: {list(entry_data['languages'].keys())}")
                
                # Store the entry data
                self.entries_data[entry_id] = entry_data
                
                if not entry_data['languages']:
                    print(f"  Warning: No terms found for entry {entry_id}")
            
            print(f"\nTotal entries processed: {len(self.entries_data)}")
            
            # Convert entries to flat rows for Excel
            self._flatten_entries_to_rows()
            
            if not self.terms_data:
                print("\nDEBUG: Let's examine the structure of your TBX file...")
                print("Sample elements found:")
                for i, elem in enumerate(root.iter()):
                    if i < 20:  # Show first 20 elements
                        attrs = {k: v for k, v in elem.attrib.items()}
                        text = elem.text.strip() if elem.text and elem.text.strip() else ""
                        text = text[:50] + "..." if len(text) > 50 else text
                        print(f"  {elem.tag}: {attrs} -> '{text}'")
                    else:
                        break
                            
        except ET.ParseError as e:
            print(f"Error parsing XML: {e}")
            sys.exit(1)
        except FileNotFoundError:
            print(f"File not found: {self.tbx_file_path}")
            sys.exit(1)
        except Exception as e:
            print(f"Unexpected error: {e}")
            sys.exit(1)
    
    def _flatten_entries_to_rows(self) -> None:
        """Convert the structured entries data into flat rows for Excel"""
        print("\nFlattening entries into Excel rows...")
        
        self.terms_data = []
        
        for entry_id, entry_data in self.entries_data.items():
            print(f"Processing entry: {entry_id}")
            
            # Start with entry-level data
            row = {}
            
            # Add entry ID if selected
            if 'entry_id' in self.selected_fields:
                row['entry_id'] = entry_id
            
            # Add entry-level description fields
            for field in self.selected_fields:
                if field.startswith('entry_') and field != 'entry_id':
                    if field in entry_data:
                        row[field] = entry_data[field]
                    else:
                        row[field] = ''
            
            # Process languages and their terms
            languages = entry_data['languages']
            
            if not languages:
                # Entry with no terms - still add the row with entry-level data
                self.terms_data.append(row)
                continue
            
            # For each language, find the maximum number of terms
            max_terms_per_lang = {}
            for lang_code, terms in languages.items():
                max_terms_per_lang[lang_code] = len(terms)
            
            print(f"  Languages: {list(languages.keys())}")
            print(f"  Terms per language: {max_terms_per_lang}")
            
            # Add columns for each language and term combination
            for lang_code, terms in languages.items():
                for term_idx, term_data in enumerate(terms):
                    # For each selected field (except entry-level ones)
                    for field in self.selected_fields:
                        if field.startswith('entry_') or field == 'entry_id':
                            continue  # Skip entry-level fields, already handled
                        
                        # Create column name: language_field or language_field_1, language_field_2, etc.
                        if term_idx == 0:
                            col_name = f"{lang_code}_{field}"
                        else:
                            col_name = f"{lang_code}_{field}_{term_idx + 1}"
                        
                        # Get the value for this field
                        value = term_data.get(field, '')
                        row[col_name] = value
            
            # Add the completed row
            self.terms_data.append(row)
            print(f"  Created row with {len(row)} columns")
        
        print(f"\nTotal rows created: {len(self.terms_data)}")
        
        # Show column structure
        if self.terms_data:
            sample_row = self.terms_data[0]
            print(f"Sample columns ({len(sample_row)}):")
            for i, col_name in enumerate(sorted(sample_row.keys())):
                print(f"  {i+1:2d}. {col_name}")
                if i >= 10:  # Limit output
                    print(f"  ... and {len(sample_row) - 11} more columns")
                    break
    
    def to_excel(self, output_path: Optional[str] = None) -> str:
        """
        Convert parsed data to Excel format
        
        Args:
            output_path: Path for output Excel file (optional)
            
        Returns:
            Path to the created Excel file
        """
        if not self.terms_data:
            print("No terminology data found. Please run parse_tbx() first.")
            return ""
        
        print(f"Converting {len(self.terms_data)} rows to Excel...")
        
        # Create DataFrame
        df = pd.DataFrame(self.terms_data)
        print(f"Original columns: {list(df.columns)}")
        
        # Apply field name mappings (rename columns) - FIXED VERSION
        if self.field_mappings:
            print(f"Applying field mappings: {self.field_mappings}")
            
            # Create a proper rename dictionary
            rename_dict = {}
            
            for current_col in df.columns:
                new_name = None
                
                # Check for language-prefixed columns (e.g., "de_term", "en_termNote_status", "de_term_2")
                if '_' in current_col:
                    # Handle columns with term numbers (e.g., "de_term_2")
                    parts = current_col.split('_')
                    
                    if len(parts) >= 2:
                        lang_prefix = parts[0]
                        
                        # Check if last part is a number (for multiple terms)
                        term_number = ""
                        if len(parts) >= 3 and parts[-1].isdigit():
                            term_number = f"_{parts[-1]}"
                            base_field = '_'.join(parts[1:-1])  # Everything between lang and number
                        else:
                            base_field = '_'.join(parts[1:])  # Everything after language
                        
                        # Check if the base field has a mapping
                        if base_field in self.field_mappings:
                            new_name = f"{lang_prefix}_{self.field_mappings[base_field]}{term_number}"
                        # Check if the full column name has a direct mapping
                        elif current_col in self.field_mappings:
                            new_name = self.field_mappings[current_col]
                
                # Check for direct mapping (non-language-prefixed columns)
                if new_name is None and current_col in self.field_mappings:
                    new_name = self.field_mappings[current_col]
                
                # Use the new name if found, otherwise keep original
                if new_name:
                    rename_dict[current_col] = new_name
                    print(f"  Renaming: '{current_col}' → '{new_name}'")
                else:
                    rename_dict[current_col] = current_col
                    print(f"  Keeping: '{current_col}'")
            
            # Apply the renaming
            df = df.rename(columns=rename_dict)
            print(f"Final columns after renaming: {list(df.columns)}")
        
        # Generate output filename if not provided
        if not output_path:
            output_path = self.tbx_file_path.with_suffix('.xlsx')
        else:
            output_path = Path(output_path)
        
        # Ensure parent directory exists
        output_path.parent.mkdir(parents=True, exist_ok=True)
        
        # Write to Excel
        try:
            print(f"Writing Excel file to: {output_path}")
            
            # Use xlsxwriter engine for better compatibility
            with pd.ExcelWriter(str(output_path), engine='xlsxwriter') as writer:
                df.to_excel(writer, sheet_name='Terminology', index=False)
                
                # Get the workbook and worksheet objects
                workbook = writer.book
                worksheet = writer.sheets['Terminology']
                
                # Auto-adjust column widths
                for i, column in enumerate(df.columns):
                    # Calculate the maximum length in the column
                    max_length = len(str(column))  # Header length
                    for value in df[column].astype(str):
                        max_length = max(max_length, len(value))
                    
                    # Set width with some padding, but cap at reasonable maximum
                    adjusted_width = min(max_length + 2, 50)
                    worksheet.set_column(i, i, adjusted_width)
            
            # Verify file was created
            if output_path.exists():
                file_size = output_path.stat().st_size
                print(f"✓ Successfully converted TBX to Excel: {output_path}")
                print(f"✓ File size: {file_size:,} bytes")
                print(f"✓ Total entries: {len(df)}")
                print(f"✓ Columns: {len(df.columns)}")
            else:
                print("✗ Error: Excel file was not created")
                return ""
            
            return str(output_path)
            
        except Exception as e:
            print(f"✗ Error writing Excel file: {e}")
            print(f"Error type: {type(e).__name__}")
            
            # Try alternative approach with openpyxl
            try:
                print("Trying alternative Excel writer (openpyxl)...")
                with pd.ExcelWriter(str(output_path), engine='openpyxl') as writer:
                    df.to_excel(writer, sheet_name='Terminology', index=False)
                
                if output_path.exists():
                    print(f"✓ Successfully created Excel file with openpyxl: {output_path}")
                    return str(output_path)
                else:
                    print("✗ Failed to create file with openpyxl as well")
                    
            except Exception as e2:
                print(f"✗ Alternative Excel writer also failed: {e2}")
                
            return ""
    
    def get_summary(self) -> Dict[str, Any]:
        """Get summary statistics of the converted data"""
        if not self.terms_data:
            return {}
        
        df = pd.DataFrame(self.terms_data)
        
        summary = {
            'total_entries': len(df),
            'columns': list(df.columns),
            'languages_detected': [],
        }
        
        # Detect languages from column names
        for col in df.columns:
            if '_' in col and not col.startswith('entry_'):
                lang = col.split('_')[0]
                if lang not in summary['languages_detected']:
                    summary['languages_detected'].append(lang)
        
        return summary


def main():
    """Main function to run the TBX to Excel converter"""
    parser = argparse.ArgumentParser(description='Convert TBX files to Excel format')
    parser.add_argument('input_file', help='Path to the TBX file')
    parser.add_argument('-o', '--output', help='Output Excel file path (optional)')
    parser.add_argument('-s', '--summary', action='store_true', 
                       help='Show summary statistics')
    parser.add_argument('--auto', action='store_true',
                       help='Skip interactive configuration and use all fields')
    
    args = parser.parse_args()
    
    # Validate input file exists
    if not Path(args.input_file).exists():
        print(f"Error: Input file '{args.input_file}' not found.")
        sys.exit(1)
    
    # Create converter instance
    converter = TBXConverter(args.input_file)
    
    print(f"Processing TBX file: {args.input_file}")
    
    if args.auto:
        print("Auto mode: using all available fields with original names")
        converter._scan_available_fields()
        converter.selected_fields = sorted(list(converter.available_fields))
        converter.field_mappings = {field: field for field in converter.selected_fields}
    else:
        # Interactive configuration
        converter.configure_extraction()
    
    print(f"\nStarting conversion...")
    
    # Parse TBX file
    converter.parse_tbx()
    
    if not converter.terms_data:
        print("Error: No data was extracted from the TBX file.")
        sys.exit(1)
    
    # Convert to Excel
    output_file = converter.to_excel(args.output)
    
    if not output_file:
        print("Error: Failed to create Excel file.")
        sys.exit(1)
    
    # Show summary if requested
    if args.summary or not args.auto:
        summary = converter.get_summary()
        print(f"\n" + "="*60)
        print(f"CONVERSION SUMMARY")
        print(f"="*60)
        print(f"Total entries: {summary.get('total_entries', 0)}")
        print(f"Languages detected: {', '.join(summary.get('languages_detected', []))}")
        print(f"Final columns: {len(summary.get('columns', []))}")
        if summary.get('columns'):
            print("Column names:")
            for col in summary.get('columns', []):
                print(f"  - {col}")
    
    return output_file


if __name__ == "__main__":
    main()