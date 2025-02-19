from docx2pdf import convert
import os

def convert_word_to_pdf(input_path):
    """
    Convert a Word document to PDF
    
    Args:
        input_path (str): Path to the Word document
        
    Returns:
        str: Path to the generated PDF file
    """
    # Check if file exists
    if not os.path.exists(input_path):
        raise FileNotFoundError(f"The file {input_path} does not exist")
    
    # Check if it's a Word document
    if not input_path.lower().endswith(('.doc', '.docx')):
        raise ValueError("The file must be a Word document (.doc or .docx)")
    
    # Generate output PDF path
    output_path = os.path.splitext(input_path)[0] + '.pdf'
    
    try:
        # Convert the document
        convert(input_path, output_path)
        return output_path
    except Exception as e:
        raise Exception(f"Error converting document: {str(e)}")

if __name__ == "__main__":
    # Example usage
    try:
        word_file = input("Enter the path to your Word document: ")
        pdf_file = convert_word_to_pdf(word_file)
        print(f"Successfully converted! PDF saved as: {pdf_file}")
    except Exception as e:
        print(f"Error: {str(e)}")