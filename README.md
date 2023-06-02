# docx-scaler

docx-scaler is a Python command-line tool for modifying Microsoft Word documents (.docx) by removing margins and scaling the page size. It allows you to customize the page format and optionally convert the modified document to PDF.

## Features

- Remove margins from Word documents.
- Scale the page size to different formats (A4, A5, etc.).
- Convert the modified document to PDF (optional).

## Installation

1. Clone the repository:

   ```bash
   git clone https://github.com/killbus/docx-scaler.git
   ```

2. Install the dependencies:

    ```bash
    pip install -r requirements.txt
    ```

## Usage

```bash
python run.py input.docx --format A5 --pdf
```

- `input.docx`: Path to the input Word document file.
- `--format`: Specify the page format (default: A5).
- `--pdf`: Convert the modified document to PDF (optional).

The modified document will be saved as output.docx in the same directory as the input file. If the `--pdf` flag is provided, the modified document will also be converted to PDF and saved as output.pdf.

## License

This project is licensed under the MIT License.

Feel free to contribute to the project by submitting issues or pull requests.

## Acknowledgements

docx-scaler uses the following libraries:

- [python-docx](https://python-docx.readthedocs.io/) for working with Word documents in Python.
- [papersize](https://pypi.org/project/papersize/) for retrieving paper size information.
- [docx2pdf](https://pypi.org/project/docx2pdf/) for converting Word documents to PDF.
