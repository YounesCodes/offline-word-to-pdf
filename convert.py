"""
Word to PDF Converter - Minimalist CLI Tool
Converts .docx files to PDF while preserving layout and quality.
"""

import argparse
import sys
from pathlib import Path

try:
    from docx2pdf import convert
except ImportError:
    print("Error: docx2pdf is not installed.")
    print("Install it with: pip install docx2pdf")
    sys.exit(1)


def main():
    parser = argparse.ArgumentParser(
        description="Convert Word documents (.docx) to PDF",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python convert.py document.docx                    # Convert single file
  python convert.py document.docx -o output.pdf     # Specify output name
  python convert.py ./documents/                     # Convert all .docx in folder
  python convert.py ./docs/ -o ./pdfs/              # Convert folder to folder
        """
    )
    
    parser.add_argument(
        "input",
        help="Input .docx file or directory containing .docx files"
    )
    
    parser.add_argument(
        "-o", "--output",
        help="Output PDF file or directory (optional)",
        default=None
    )
    
    parser.add_argument(
        "-q", "--quiet",
        action="store_true",
        help="Suppress output messages"
    )
    
    args = parser.parse_args()
    
    input_path = Path(args.input)
    output_path = Path(args.output) if args.output else None
    
    # Validate input
    if not input_path.exists():
        print(f"Error: '{args.input}' does not exist.")
        sys.exit(1)
    
    # Handle single file
    if input_path.is_file():
        if input_path.suffix.lower() != ".docx":
            print(f"Error: '{args.input}' is not a .docx file.")
            sys.exit(1)
        
        if not args.quiet:
            print(f"Converting: {input_path.name}")
        
        try:
            convert(str(input_path), str(output_path) if output_path else None)
            
            if not args.quiet:
                if output_path:
                    out_file = output_path.resolve()
                else:
                    out_file = input_path.with_suffix(".pdf").resolve()
                print(f"Saved: {out_file}")
        except Exception as e:
            print(f"Error converting file: {e}")
            sys.exit(1)
    
    # Handle directory
    elif input_path.is_dir():
        docx_files = list(input_path.glob("*.docx"))
        
        if not docx_files:
            print(f"No .docx files found in '{args.input}'")
            sys.exit(1)
        
        if output_path and not output_path.exists():
            output_path.mkdir(parents=True)
        
        if not args.quiet:
            print(f"Found {len(docx_files)} file(s) to convert...")
        
        try:
            convert(str(input_path), str(output_path) if output_path else None)
            
            if not args.quiet:
                out_dir = (output_path or input_path).resolve()
                print(f"Saved {len(docx_files)} PDF(s) to: {out_dir}")
        except Exception as e:
            print(f"Error converting files: {e}")
            sys.exit(1)
    
    else:
        print(f"Error: '{args.input}' is not a valid file or directory.")
        sys.exit(1)


if __name__ == "__main__":
    main()
