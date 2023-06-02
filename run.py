import argparse
from decimal import ROUND_HALF_UP, Decimal
import os
import docx
from docx.document import Document
from docx.shared import Mm
from docx2pdf import convert
from papersize import parse_papersize


def get_paper_size(page_format):
    return parse_papersize(page_format, 'mm')


def modify_document(input_file: str, page_format: str, convert_to_pdf: bool):
    # Extract the directory path from the input file
    input_dir = os.path.dirname(input_file)

    # Load the document
    doc: Document = docx.Document(input_file)

    # Remove margins
    sections = doc.sections
    for section in sections:
        section.left_margin = Mm(0)
        section.right_margin = Mm(0)
        section.top_margin = Mm(0)
        section.bottom_margin = Mm(0)

    # Get the paper size from the input file
    page_width, page_height = get_paper_size(page_format)

    # Convert to the specified page format
    for section in sections:
        section.page_width = Mm(page_width.quantize(
            Decimal('1.0'), rounding=ROUND_HALF_UP))
        section.page_height = Mm(page_height.quantize(
            Decimal('1.0'), rounding=ROUND_HALF_UP))

    # Save the modified document
    output_file = os.path.join(input_dir, 'output.docx') if input_file.endswith(
        '.docx') else os.path.join(input_dir, 'output.doc')
    doc.save(output_file)
    print(f"Modified document saved as '{output_file}'.")

    # Convert to PDF if requested
    if convert_to_pdf:
        pdf_output = os.path.splitext(output_file)[0] + '.pdf'
        convert(output_file, pdf_output)
        print(f"Converted to PDF: '{pdf_output}'.")


if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        description='Modify Word document margins and page size.')
    parser.add_argument(
        'input_file', help='Path to the input Word document file')
    parser.add_argument('--format', default='A5',
                        help='Page format (default: A5)')
    parser.add_argument('--pdf', action='store_true',
                        help='Convert the modified document to PDF')
    args = parser.parse_args()

    modify_document(args.input_file, args.format, args.pdf)
