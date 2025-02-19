import pytest
from word_to_pdf_converter import convert_word_to_pdf
from unittest.mock import patch
import os

class TestWordToPdfConverter:
    def test_convert_word_to_pdf_1(self):
        """
        Test that convert_word_to_pdf raises FileNotFoundError when the input file does not exist,
        and ValueError when the input file is not a Word document.
        """
        # Test for non-existent file
        with pytest.raises(FileNotFoundError):
            convert_word_to_pdf("non_existent_file.docx")

        # Test for file with incorrect extension
        invalid_file = "invalid_file.txt"
        with open(invalid_file, 'w') as f:
            f.write("This is not a Word document")

        try:
            with pytest.raises(ValueError):
                convert_word_to_pdf(invalid_file)
        finally:
            if os.path.exists(invalid_file):
                os.remove(invalid_file)

    def test_convert_word_to_pdf_2(self):
        """
        Test that convert_word_to_pdf raises a ValueError when the input file exists but is not a Word document.
        """
        test_file = 'test_file.txt'
        with open(test_file, 'w') as f:
            f.write('This is a test file')

        try:
            with pytest.raises(ValueError) as exc_info:
                convert_word_to_pdf(test_file)
            assert str(exc_info.value) == "The file must be a Word document (.doc or .docx)"
        finally:
            if os.path.exists(test_file):
                os.remove(test_file)

    def test_convert_word_to_pdf_3(self):
        """
        Test that convert_word_to_pdf raises a FileNotFoundError when the input file does not exist.
        """
        with pytest.raises(FileNotFoundError):
            convert_word_to_pdf('non_existent_file.docx')

    def test_convert_word_to_pdf_example_negative_test(self):
        """
        Negative test cases for `convert_word_to_pdf`
        """
        # 1. Test with empty input
        with pytest.raises(FileNotFoundError):
            convert_word_to_pdf("")

        # 2. Test with non-existent file
        with pytest.raises(FileNotFoundError):
            convert_word_to_pdf("non_existent_file.docx")

        # 3. Test with incorrect file format
        test_file = "test_file.txt"
        try:
            with open(test_file, 'w') as f:
                f.write("test content")
            with pytest.raises(ValueError):
                convert_word_to_pdf(test_file)
        finally:
            if os.path.exists(test_file):
                os.remove(test_file)

        # 4. Test exception during conversion
        with patch('word_to_pdf_converter.convert', side_effect=Exception("Conversion error")):
            with pytest.raises(Exception) as context:
                convert_word_to_pdf("test_file.docx")
            assert "Conversion error" in str(context.value)

        # 5. Test with file path containing special characters
        with pytest.raises(FileNotFoundError):
            convert_word_to_pdf("test@file#.docx")
