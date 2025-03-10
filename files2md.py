#!/usr/bin/env python3

"""
Enhanced File to Markdown Converter

This script combines multiple files into a single Markdown document, treating each file
as an "escaped attachment". It supports PDF, DOCX, plain text, and XLSX files, with
recursive directory searching and filename pattern matching.

Dependencies:
- pdfminer.six: For extracting text from PDFs
- pdf2image: For converting PDFs to images (used in OCR)
- pytesseract: For OCR on PDF images
- pypandoc: For converting DOCX to Markdown
- pandas: For reading XLSX files as dataframes
- openpyxl (engine): For parsing XLSX in pandas
- (Optional) tabulate: If your pandas version doesn't include built-in to_markdown.

Usage:
python3 files2md [-r] [--name-regex REGEX] [--name-code] [--name-docs] file1.pdf dir1 file2.docx > output.md
"""

import argparse
import re
import sys
import os

from pdfminer.high_level import extract_text
from pdf2image import convert_from_path
import pytesseract
import pypandoc
import pandas as pd  # <-- New import for XLSX handling

# Common file patterns
CODE_FILE_PATTERN = r'\.(py|js|java|c|cpp|h|hpp|cs|rb|go|rs|php|html|css|sql|sh|bash|ps1|rkt|hs|scala|ml|elm|clj|ex|exs|erl|fs|fsx|lisp|scm|sml|swift|kt|kts|groovy|pl|pm|t|lua|jl|dart|d|nim|cr|r|R|asm|s|zig|v|ada|f90|f95|f03|f08|pas|cob|cobol|vb|vba|vbs|tcl|hx|m|mm|ts|coffee|ts|ls|cljc|cljs|raku|bf|md|txt)$|^(Makefile|Dockerfile|Rakefile|Gemfile|Vagrantfile|CMakeLists\.txt)$'
DOCS_FILE_PATTERN = r'\.(md|txt|rst|tex|rtf|odt|doc|docx|pdf|epub|csv|tsv|json|xml|yaml|yml|ini|cfg|conf|log)$'

def escape_backticks(text):
    """
    Escape backticks within text by increasing the surrounding backtick count.
    
    Args:
        text (str): The input text to escape.
    
    Returns:
        str: The input text surrounded by an appropriate number of backticks.
    """
    backtick_pattern = r'(`+)'
    max_len = max((len(m.group(1)) for m in re.finditer(backtick_pattern, text)), default=0)
    surrounding_backticks = '`' * max(3, max_len + 1)
    return f"{surrounding_backticks}\n{text}\n{surrounding_backticks}"

def convert_docx_to_markdown(docx_path):
    """
    Convert a DOCX file to Markdown using pypandoc with GitHub Flavored Markdown (GFM) format.
    
    Args:
        docx_path (str): Path to the DOCX file.
    
    Returns:
        str: Converted Markdown text, or an error message if conversion fails.
    """
    try:
        markdown_text = pypandoc.convert_file(
            docx_path,
            'gfm',  # GitHub Flavored Markdown
            format='docx',
            extra_args=[
                '--wrap=none',  # Don't wrap lines
                '--standalone',  # Include header/footer if present
                '--markdown-headings=atx'  # Use # style headers
            ]
        )
        return markdown_text
    except (pypandoc.PandocError, OSError) as e:
        return f"Failed to convert DOCX to Markdown: {str(e)}"

def convert_xlsx_to_markdown(xlsx_path):
    """
    Convert an XLSX file into GitHub-flavored Markdown tables, one for each sheet.

    Args:
        xlsx_path (str): Path to the XLSX file.

    Returns:
        str: A string containing one or more GitHub-flavored Markdown tables
             (plus headings for each sheet) for the entire workbook.
    """
    try:
        # Read all sheets into a dictionary: {sheet_name: DataFrame}
        xls = pd.read_excel(xlsx_path, sheet_name=None)
    except Exception as e:
        return f"Failed to read XLSX file: {str(e)}"

    # Build a markdown representation of all sheets
    markdown_parts = []
    for sheet_name, df in xls.items():
        # Convert DataFrame to a GitHub-flavored Markdown table
        md_table = df.to_markdown(index=False, tablefmt="github")
        markdown_parts.append(f"### Sheet: {sheet_name}\n\n{md_table}")

    return "\n\n".join(markdown_parts)

def extract_text_from_pdf(pdf_path):
    """
    Extract text from a PDF file, using OCR if text extraction fails.
    
    Args:
        pdf_path (str): Path to the PDF file.
    
    Returns:
        str: Extracted text from the PDF.
    """
    text = extract_text(pdf_path)
    if text.strip() == '':
        # If no text was extracted, try OCR
        images = convert_from_path(pdf_path)
        text = ''.join([pytesseract.image_to_string(image) for image in images])
    return text

def process_files(filenames, name_regex=None):
    """
    Process multiple files and return their contents as a Markdown formatted string.
    
    Args:
        filenames (list): A list of filenames to process.
        name_regex (str): Regular expression to filter filenames.
    
    Returns:
        str: A Markdown formatted string containing the contents of all processed files.
    """
    markdown_output = []
    for filename in filenames:
        if name_regex and not re.search(name_regex, filename):
            continue
        
        try:
            # Distinguish file type by extension
            if filename.lower().endswith('.pdf'):
                contents = extract_text_from_pdf(filename)
                escaped = escape_backticks(contents)
                markdown_output.append(f"<!-- file-attachment: {filename} -->")
                markdown_output.append(f"## Attached file: {filename}")
                markdown_output.append(escaped)

            elif filename.lower().endswith('.docx'):
                contents = convert_docx_to_markdown(filename)
                escaped = escape_backticks(contents)
                markdown_output.append(f"<!-- file-attachment: {filename} -->")
                markdown_output.append(f"## Attached file: {filename}")
                markdown_output.append(escaped)

            elif filename.lower().endswith('.xlsx'):
                contents = convert_xlsx_to_markdown(filename)
                # For tables to render properly, do not wrap them in code fences:
                markdown_output.append(f"<!-- file-attachment: {filename} -->")
                markdown_output.append(f"## Attached file: {filename}")
                markdown_output.append(contents)

            else:
                # Plain text or unknown extension
                with open(filename, 'r', encoding='utf-8') as file:
                    contents = file.read()
                escaped = escape_backticks(contents)
                markdown_output.append(f"<!-- file-attachment: {filename} -->")
                markdown_output.append(f"## Attached file: {filename}")
                markdown_output.append(escaped)

        except IOError as e:
            print(f"Error: Could not read file {filename}. Skipping.", file=sys.stderr)
            print(e, file=sys.stderr)
    
    return '\n\n'.join(markdown_output)

def find_files(paths, recursive=False, name_regex=None):
    """
    Find files in the given paths, optionally recursing into subdirectories.
    
    Args:
        paths (list): List of file and directory paths to search.
        recursive (bool): Whether to search subdirectories recursively.
        name_regex (str): Regular expression to filter filenames.
    
    Returns:
        list: List of file paths matching the criteria.
    """
    found_files = []
    for path in paths:
        if os.path.isfile(path):
            found_files.append(path)
        elif os.path.isdir(path):
            if recursive:
                for root, _, files in os.walk(path):
                    for file in files:
                        file_path = os.path.join(root, file)
                        if name_regex is None or re.search(name_regex, file_path):
                            found_files.append(file_path)
            else:
                for item in os.listdir(path):
                    item_path = os.path.join(path, item)
                    if os.path.isfile(item_path) and (name_regex is None or re.search(name_regex, item_path)):
                        found_files.append(item_path)
    return found_files

def main():
    """
    Main function to parse command-line arguments and process files.
    """
    # Verify that pandoc is available (for DOCX -> MD conversion)
    try:
        pypandoc.get_pandoc_version()
    except OSError:
        print("Error: pandoc is not installed. Please install pandoc first.", file=sys.stderr)
        sys.exit(1)

    parser = argparse.ArgumentParser(
        description="Combine files into a single markdown document, with recursive directory searching and filename pattern matching."
    )
    parser.add_argument("paths", metavar="PATH", type=str, nargs='+',
                        help="List of files or directories to process.")
    parser.add_argument("-r", "--recursive", action="store_true",
                        help="Recursively search directories for files.")
    parser.add_argument("--name-regex", type=str,
                        help="Regular expression to match filenames.")
    parser.add_argument("--name-code", action="store_true",
                        help="Use a predefined regex to match common code file extensions.")
    parser.add_argument("--name-docs", action="store_true",
                        help="Use a predefined regex to match common document file extensions.")
    args = parser.parse_args()
    
    # Validate that only one naming argument is used
    if sum([bool(args.name_regex), args.name_code, args.name_docs]) > 1:
        print("Error: Can only use one of --name-regex, --name-code, or --name-docs at a time.", file=sys.stderr)
        sys.exit(1)
    
    # Select the appropriate filename pattern
    name_regex = None
    if args.name_regex:
        name_regex = args.name_regex
    elif args.name_code:
        name_regex = CODE_FILE_PATTERN
    elif args.name_docs:
        name_regex = DOCS_FILE_PATTERN
    
    # Gather and process files
    files = find_files(args.paths, recursive=args.recursive, name_regex=name_regex)
    markdown_content = process_files(files, name_regex=name_regex)
    
    # Print final combined markdown
    print(markdown_content)

if __name__ == "__main__":
    main()
