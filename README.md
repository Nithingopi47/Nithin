# Word to PDF Converter with Drag-and-Drop Interface

This project provides a simple and user-friendly application for converting Microsoft Word documents (.doc or .docx) to PDF format using a graphical user interface with drag-and-drop functionality.

The Word to PDF Converter is designed to streamline the process of converting Word documents to PDF format. It leverages the power of the `docx2pdf` library for the conversion process and presents a clean, intuitive interface built with Tkinter. Users can simply drag and drop their Word documents into the application window, and the converter will automatically process the file and generate a PDF in the same directory.

Key features of this application include:
- Drag-and-drop interface for easy file selection
- Support for both .doc and .docx file formats
- Automatic PDF generation in the same directory as the input file
- Clear status updates and error messages
- Robust error handling for various scenarios

## Repository Structure

- `drag_drop_converter.py`: Main script that implements the GUI and drag-and-drop functionality
- `word_to_pdf_converter.py`: Core module containing the conversion logic
- `test_word_to_pdf_converter.py`: Test suite for the conversion module
- `README.md`: This file, containing project documentation

## Usage Instructions

### Installation

1. Ensure you have Python 3.6 or higher installed on your system.
2. Install the required dependencies:

```bash
pip install docx2pdf tkinterdnd2
```

### Getting Started

To run the application:

1. Navigate to the project directory in your terminal.
2. Execute the following command:

```bash
python drag_drop_converter.py
```

3. The application window will appear, showing a drop zone for Word documents.
4. Drag and drop a .doc or .docx file into the drop zone.
5. The application will convert the file and display the path of the generated PDF.

### Common Use Cases

1. Converting a single Word document:
   - Drag and drop the Word document into the application window.
   - The PDF will be generated in the same directory as the original file.

2. Batch converting multiple documents:
   - Drag and drop each Word document into the application window one at a time.
   - Each PDF will be generated in its respective original directory.

### Testing & Quality

To run the test suite:

```bash
pytest test_word_to_pdf_converter.py
```

This will execute all the tests in the `test_word_to_pdf_converter.py` file, ensuring the core conversion functionality works as expected.

### Troubleshooting

1. Issue: File not found error
   - Error message: "The file {input_path} does not exist"
   - Solution: Ensure the file path is correct and the file exists in the specified location.

2. Issue: Invalid file format
   - Error message: "The file must be a Word document (.doc or .docx)"
   - Solution: Check that you're trying to convert a valid Word document. Only .doc and .docx files are supported.

3. Issue: Conversion error
   - Error message: "Error converting document: {specific error message}"
   - Solution: 
     1. Ensure the Word document is not corrupted or password-protected.
     2. Check that you have write permissions in the directory where the PDF will be saved.
     3. Verify that the `docx2pdf` library is correctly installed.

To enable verbose logging for debugging:

1. Open `word_to_pdf_converter.py` in a text editor.
2. Add the following lines at the beginning of the file:

```python
import logging
logging.basicConfig(level=logging.DEBUG)
```

3. Run the application again. Detailed logs will be printed to the console.

## Data Flow

The Word to PDF Converter application follows a straightforward data flow:

1. User drags and drops a Word document onto the application window.
2. The application checks if the file is a valid Word document (.doc or .docx).
3. If valid, the `convert_word_to_pdf` function is called with the input file path.
4. The function checks for file existence and format validity.
5. The `docx2pdf.convert` function is used to convert the Word document to PDF.
6. The resulting PDF is saved in the same directory as the input file.
7. The application displays the path of the generated PDF or an error message.

```
[User Input] -> [GUI] -> [File Validation] -> [Conversion Function] -> [docx2pdf Library] -> [PDF Output]
     ^                                              |
     |                                              v
     +------------------------------------------[Status Display]
```

Note: The conversion process is handled by the `docx2pdf` library, which may have its own internal data flow and processing steps not visible to this application.