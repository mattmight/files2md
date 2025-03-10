# files2md

Convert multiple files to a single Markdown document with each file properly formatted and escaped.

## Features

- Combines multiple files into a single Markdown document
- Supports various file formats:
  - PDF (with fallback to OCR for image-based PDFs)
  - DOCX (converted to GitHub Flavored Markdown)
  - XLSX (each sheet converted to Markdown tables)
  - Plain text files (properly escaped)
- Recursive directory searching
- Filename pattern matching via regex
- Built-in patterns for common code and documentation files

## Installation

```bash
# Clone this repository
git clone https://github.com/yourusername/files2md.git
cd files2md

# Install dependencies
pip install -r requirements.txt

# Make sure pandoc is installed
# On macOS:
brew install pandoc
# On Ubuntu/Debian:
# apt-get install pandoc
# On Windows with Chocolatey:
# choco install pandoc

# Install the script
make install
```

## Usage

```bash
# Basic usage
files2md file1.pdf file2.docx > combined.md

# Recursively search directories
files2md -r project_dir > project_docs.md

# Filter only code files 
files2md -r --name-code project_dir > all_code.md

# Filter only documentation files
files2md -r --name-docs docs_dir > all_docs.md

# Use custom regex pattern
files2md -r --name-regex "\.py$" project_dir > python_files.md
```

## Examples

### Combine documentation for a project

```bash
files2md -r --name-docs project_dir > project_documentation.md
```

### Create a code review document

```bash
files2md -r --name-code src_dir > code_review.md
```

### Compile multiple PDFs into a single document

```bash
files2md *.pdf > compiled_pdfs.md
```

## Dependencies

- pdfminer.six - For extracting text from PDFs
- pdf2image - For converting PDFs to images (used in OCR)
- pytesseract - For OCR on PDF images
- pypandoc - For converting DOCX to Markdown
- pandas - For reading XLSX files as dataframes
- openpyxl - For parsing XLSX in pandas
- tabulate (Optional) - If your pandas version doesn't include built-in to_markdown

## License

MIT