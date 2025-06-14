#!/usr/bin/env python3

"""
XLSXtract - Extract text from Excel files for password generation
Copyright (C) 2024 Garland Glessner <gglessner@gmail.com>

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program.  If not, see <https://www.gnu.org/licenses/>.

Author: Garland Glessner
Email: gglessner@gmail.com
"""

import argparse
import os
import sys
import warnings
import shutil
from pathlib import Path
from openpyxl import load_workbook
from typing import Set, Generator, List
import time

# Suppress openpyxl data validation warning
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

def get_terminal_width() -> int:
    """Get the width of the terminal window."""
    try:
        return shutil.get_terminal_size().columns
    except:
        return 80  # Default fallback width

def truncate_for_display(text: str, max_length: int) -> str:
    """Truncate text to fit terminal width, accounting for the 'Extracting: ' prefix."""
    prefix = "Extracting: "
    max_display_length = get_terminal_width() - len(prefix) - 3  # -3 for safety margin
    if len(text) > max_display_length:
        return text[:max_display_length] + "..."
    return text

def find_xlsx_files(directory: str, filename_pattern: str = None) -> Generator[Path, None, None]:
    """Recursively find all .xlsx files in the given directory.
    If filename_pattern is provided, only return files matching that name."""
    directory_path = Path(directory)
    if filename_pattern:
        # If filename_pattern doesn't end with .xlsx, add it
        if not filename_pattern.lower().endswith('.xlsx'):
            filename_pattern += '.xlsx'
        # Return only files matching the exact name
        return (f for f in directory_path.rglob("*.xlsx") if f.name.lower() == filename_pattern.lower())
    return directory_path.rglob("*.xlsx")

def extract_text_from_xlsx(xlsx_path: Path, split_words: bool, max_length: int, show_progress: bool = False) -> Set[str]:
    """Extract all text from cells in an Excel file."""
    text_values = set()
    word_count = 0
    
    try:
        # Load the workbook
        wb = load_workbook(xlsx_path, read_only=True, data_only=True)
        
        # Iterate through all sheets
        for sheet in wb:
            for row in sheet.rows:
                for cell in row:
                    if cell.value and isinstance(cell.value, str):
                        # Clean the text
                        cleaned_text = cell.value.strip()
                        if cleaned_text:
                            if split_words:
                                # Split on whitespace and add each word
                                words = cleaned_text.split()
                                # Only add words that are within the length limit
                                for word in words:
                                    if len(word) <= max_length:
                                        text_values.add(word)
                                        word_count += 1
                                        if show_progress:
                                            print(f"\rExtracting: {truncate_for_display(word, get_terminal_width())}", end='', flush=True)
                            else:
                                # Only add the full text if it's within the length limit
                                if len(cleaned_text) <= max_length:
                                    text_values.add(cleaned_text)
                                    word_count += 1
                                    if show_progress:
                                        print(f"\rExtracting: {truncate_for_display(cleaned_text, get_terminal_width())}", end='', flush=True)
    
    except Exception as e:
        print(f"\nError processing {xlsx_path}: {str(e)}")
    
    return text_values, word_count

def process_xlsx_file(xlsx_path: Path, output_file, split_words: bool, show_progress: bool, max_length: int) -> int:
    """Process a single XLSX file and write its text content to the output file."""
    print(f"\nProcessing: {xlsx_path}")
    
    # Extract text from the Excel file
    text_values, word_count = extract_text_from_xlsx(xlsx_path, split_words, max_length, show_progress)
    
    if show_progress:
        print()  # New line after done with word display
    else:
        print(f"Found {word_count} words")
    
    # Write values to output file
    for text in sorted(text_values):
        output_file.write(f"{text}\n")
    
    return len(text_values)

def main():
    parser = argparse.ArgumentParser(description='Extract text from XLSX files for password generation.')
    parser.add_argument('-d', '--directory', required=True, help='Directory to scan for XLSX files')
    parser.add_argument('-o', '--output', default='passwords.txt', help='Output file for extracted text (default: passwords.txt)')
    parser.add_argument('-w', '--split-words', action='store_true', help='Split cell contents on spaces into individual words')
    parser.add_argument('-p', '--progress', action='store_true', help='Show real-time progress of each word being extracted (slower)')
    parser.add_argument('-l', '--max-length', type=int, default=32, help='Maximum length of words to extract (default: 32)')
    parser.add_argument('-f', '--filename', help='Only process files with this exact name (e.g., "Config.xlsx")')
    
    args = parser.parse_args()
    
    # Ensure the directory exists
    if not os.path.isdir(args.directory):
        print(f"Error: Directory '{args.directory}' does not exist")
        return
    
    # Statistics tracking
    total_files = 0
    total_words = 0
    skipped_words = 0
    all_words = set()
    
    # First pass: collect all words
    print("First pass: Collecting all words...")
    print(f"Maximum word length: {args.max_length} characters")
    if args.filename:
        print(f"Only processing files named: {args.filename}")
    if args.progress:
        print("Showing real-time word extraction (this will be slower)")
    xlsx_files = list(find_xlsx_files(args.directory, args.filename))
    
    if not xlsx_files:
        if args.filename:
            print(f"No files named '{args.filename}' found in {args.directory}")
        else:
            print(f"No XLSX files found in {args.directory}")
        return
    
    print(f"Found {len(xlsx_files)} XLSX files")
    
    # Process each file and collect words
    for xlsx_path in xlsx_files:
        words = process_xlsx_file(xlsx_path, None, args.split_words, args.progress, args.max_length)
        all_words.update(words)
        total_files += 1
        total_words += len(words)
    
    # Write final sorted and unique list
    print("\nWriting final sorted and unique word list...")
    with open(args.output, 'w', encoding='utf-8') as output_file:
        for word in sorted(all_words):
            output_file.write(f"{word}\n")
    
    # Print statistics
    print("\nProcessing complete!")
    print(f"Statistics:")
    print(f"- Files processed: {total_files}")
    print(f"- Total words found: {total_words}")
    print(f"- Unique words written: {len(all_words)}")
    print(f"- Maximum word length: {args.max_length} characters")
    if args.filename:
        print(f"- Filename filter: {args.filename}")
    print(f"- Results written to: {args.output}")

if __name__ == "__main__":
    main() 