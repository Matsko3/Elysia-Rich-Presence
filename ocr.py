import pytesseract
from PIL import Image

def extract_text_from_image(image):
    extracted_text = pytesseract.image_to_string(image, config='--psm 6')
    return extracted_text

def compare_text(image_text, extracted_text):
    return image_text.strip() == extracted_text.strip()
