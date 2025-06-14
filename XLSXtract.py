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
from pathlib import Path
from openpyxl import load_workbook
from typing import Set, Generator, List
import time

def find_xlsx_files(directory: str) -> Generator[Path, None, None]:
    """Recursively find all .xlsx files in the given directory."""
    directory_path = Path(directory)
    return directory_path.rglob("*.xlsx")

def extract_text_from_xlsx(xlsx_path: Path, split_words: bool) -> Set[str]:
    """Extract all text from cells in an Excel file."""
    text_values = set()
    
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
                                text_values.update(words)
                            else:
                                text_values.add(cleaned_text)
    
    except Exception as e:
        print(f"\nError processing {xlsx_path}: {str(e)}")
    
    return text_values

def print_progress(file_path: Path, words: List[str], delay: float = 0.1):
    """Print progress of word extraction with animation."""
    print(f"\nProcessing: {file_path}")
    for word in words:
        # Clear the line and print new word
        print(f"\033[K", end='')  # ANSI escape sequence to clear the line
        print(f"\rExtracting: {word}", end='', flush=True)
        time.sleep(delay)
    print()  # New line after done

def process_xlsx_file(xlsx_path: Path, output_file, split_words: bool, show_progress: bool) -> int:
    """Process a single XLSX file and write its text content to the output file."""
    # Extract text from the Excel file
    text_values = extract_text_from_xlsx(xlsx_path, split_words)
    
    if show_progress:
        # Show progress animation with the words
        print_progress(xlsx_path, sorted(text_values))
    else:
        print(f"Processing: {xlsx_path}")
    
    # Write values to output file
    for text in sorted(text_values):
        output_file.write(f"{text}\n")
    
    return len(text_values)

def main():
    parser = argparse.ArgumentParser(description='Extract text from XLSX files for password generation.')
    parser.add_argument('-d', '--directory', required=True, help='Directory to scan for XLSX files')
    parser.add_argument('-o', '--output', default='passwords.txt', help='Output file for extracted text (default: passwords.txt)')
    parser.add_argument('-w', '--split-words', action='store_true', help='Split cell contents on spaces into individual words')
    parser.add_argument('-p', '--progress', action='store_true', help='Show real-time progress of word extraction')
    
    args = parser.parse_args()
    
    # Ensure the directory exists
    if not os.path.isdir(args.directory):
        print(f"Error: Directory '{args.directory}' does not exist")
        return
    
    # Statistics tracking
    total_files = 0
    total_words = 0
    all_words = set()
    
    # First pass: collect all words
    print("First pass: Collecting all words...")
    xlsx_files = list(find_xlsx_files(args.directory))
    
    if not xlsx_files:
        print(f"No XLSX files found in {args.directory}")
        return
    
    print(f"Found {len(xlsx_files)} XLSX files")
    
    # Process each file and collect words
    for xlsx_path in xlsx_files:
        words = extract_text_from_xlsx(xlsx_path, args.split_words)
        all_words.update(words)
        total_files += 1
        total_words += len(words)
        
        if args.progress:
            print_progress(xlsx_path, sorted(words))
        else:
            print(f"Processed: {xlsx_path} - Found {len(words)} words")
    
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
    print(f"- Results written to: {args.output}")

if __name__ == "__main__":
    main() 