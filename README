# DOCX PDF Extractor

A Go utility that extracts embedded PDF files from Microsoft Word (.docx) documents.

## Features

- Extracts PDF files embedded as OLE objects in DOCX files
- Automatically detects PDF filename from EMF metadata
- Saves extracted PDFs with sanitized filenames
- Supports verbose logging mode

## Usage

```bash
go run main.go [-v|--verbose] <docx_file>
```

### Examples

```bash
# Basic extraction
go run main.go document.docx

# With verbose output
go run main.go -v document.docx
```

## How It Works

1. Opens the DOCX file as a ZIP archive
2. Parses `word/document.xml` to find embedded objects
3. Identifies PDF objects by Adobe ProgID (`Acrobat.Document.DC`)
4. Extracts PDF filename from associated EMF files
5. Retrieves PDF data from OLE2-formatted `.bin` files
6. Saves PDFs with original filenames (sanitized for filesystem)

## Dependencies

- `github.com/IntelligenceX/fileconversion/ole2` - OLE2 file format handling

## Requirements

- Go 1.x
- Input file must have `.docx` extension
- Embedded PDFs must be Adobe Acrobat objects

## Output

Extracted PDF files are saved in the current directory with their original names. Non-alphanumeric characters in filenames are replaced with underscores for safety.

## TODOS:
1. Write out to a directory 
2. Implement logging 
3. Better CLI args tooling 
4. Better error handling 