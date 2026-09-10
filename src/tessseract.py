from PIL import Image
import pytesseract

# Set the path to the Tesseract executable
# This is required if tesseract is not in your system's PATH.
# Example for Windows:
# pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
pytesseract.pytesseract.tesseract_cmd = (
    r"C:\Users\hilton.netta\AppData\Local\Programs\Tesseract-OCR\tesseract.exe"
)


def extract_text_from_image(image_path):
    """
    Extracts text and numbers from a given image file.

    Args:
        image_path (str): The path to the image file (e.g., 'invoice.jpg').

    Returns:
        str: The extracted text, or None if an error occurs.
    """
    try:
        # Open the image file using Pillow
        with Image.open(image_path) as img:
            # Use pytesseract to perform OCR on the image
            extracted_text = pytesseract.image_to_string(img)
            return extracted_text
    except FileNotFoundError:
        print(f"Error: The file at {image_path} was not found.")
        return None
    except Exception as e:
        print(f"An error occurred: {e}")
        return None


if __name__ == "__main__":
    # Replace 'your_image.jpg' with the path to your JPEG image
    image_file = (
        r"C:\Users\hilton.netta\OneDrive - Prescient\Documents\ocr_test_image.jpg"
    )

    extracted_data = extract_text_from_image(image_file)

    if extracted_data:
        print("--- Extracted Text and Numbers ---")
        print(extracted_data)

        # Optional: You can process the extracted data further
        # For example, to find only numbers:
        import re

        numbers = re.findall(r"\d+", extracted_data)
        print("\n--- Extracted Numbers Only ---")
        print(numbers)
