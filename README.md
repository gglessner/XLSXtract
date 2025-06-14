# XLSXtract

A Python tool that extracts text from Excel (.xlsx) files for use as potential passwords.

## Author

Garland Glessner (gglessner@gmail.com)

## License

This project is licensed under the GNU General Public License v3.0 - see the [LICENSE](LICENSE) file for details.

## Description

XLSXtract recursively scans a directory for .xlsx files, extracts all text content from the cells, and writes unique text values to an output file. Each text value is written on a new line, making it suitable for use as a password list.

## Installation

1. Ensure you have Python 3.6 or higher installed
2. Install the required dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

Basic usage:
```bash
python XLSXtract.py -d /path/to/directory
```

This will scan the specified directory and its subdirectories for .xlsx files and write the extracted text to `passwords.txt` in the current directory.

### Options

- `-d, --directory`: Directory to scan for XLSX files (required)
- `-o, --output`: Output file name (default: passwords.txt)
- `-w, --split-words`: Split cell contents on spaces into individual words
- `-p, --progress`: Show real-time progress of word extraction

Examples:

1. Basic usage with custom output file:
```bash
python XLSXtract.py -d /path/to/directory -o custom_passwords.txt
```

2. Split words and show progress:
```bash
python XLSXtract.py -d /path/to/directory -w -p
```

3. All options:
```bash
python XLSXtract.py -d /path/to/directory -o passwords.txt -w -p
```

## Features

- Recursively scans directories for .xlsx files
- Extracts text from all cells in all sheets
- Option to split cell contents on spaces into individual words
- Real-time progress display of word extraction
- Removes duplicates using a set
- Final output is sorted alphabetically
- Handles errors gracefully
- Uses read-only mode for better memory efficiency
- UTF-8 encoding support
- Detailed statistics on processing results

## Notes

- The tool only extracts text values from cells
- Empty cells and non-text values are ignored
- Text values are stripped of leading/trailing whitespace
- When using `-w/--split-words`, each word from a cell becomes a separate entry
- Each unique word appears only once in the final output file
- Statistics are shown at the end of processing:
  - Number of files processed
  - Total words found
  - Number of unique words written
  - Output file location 