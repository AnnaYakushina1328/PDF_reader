import cv2
import pytesseract
import pandas as pd
import os
from pdf2image import convert_from_path


# путь к Tesseract OCR
pytesseract.pytesseract.tesseract_cmd = 'C:/Program Files/Tesseract-OCR/tesseract.exe'
pdf_path='C:\\Users\\yacuc\\Downloads\\Исходные данные\\Печатный\\2.pdf'


def preprocess_image(image_path):
    image = cv2.imread(image_path, cv2.IMREAD_GRAYSCALE)
    blurred = cv2.medianBlur(image, 3)
    _, processed_image = cv2.threshold(blurred, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    return processed_image


def extract_text_from_image(image):
    # распознаем текст с помощью pytesseract
    text = pytesseract.image_to_string(image)
    return text


def extract_data_from_text(text):
    # извлекаем строки, содержащие цифры, для демонстрации
    lines = text.splitlines()
    data = [line for line in lines if any(char.isdigit() for char in line)]
    return data


def save_data_to_file(data, output_file):
    # преобразуем данные в табличный формат
    df = pd.DataFrame(data, columns=["Extracted Data"])
    # сохраняем в Excel файл
    df.to_excel(output_file, index=False)


def process_pdf(pdf_path, output_folder):

    # конвертируем страницы PDF в изображения
    #images = convert_from_path(pdf_path)

    try:
        images = convert_from_path(pdf_path)
        if not images:
            print("Не удалось конвертировать PDF в изображения.")
            return
    except Exception as e:
        print(f"Ошибка при конвертации PDF: {e}")
        return
    all_data = []

    # Обрабатываем каждую страницу
    for i, page_image in enumerate(images):
        image_path = os.path.join(output_folder, f"page_{i + 1}.jpg")
        page_image.save(image_path, "JPEG")
        print(f"Изображение сохранено: {image_path}")

        # предобработка изображения
        processed_image = preprocess_image(image_path)

        # распознавание текста
        text = extract_text_from_image(processed_image)
        print(f"Распознанный текст на странице {i + 1}:\n{text}")

        # извлечение данных
        data = extract_data_from_text(text)
        all_data.extend(data)

        # сохраняем все данные в файл
        if all_data:
            save_data_to_file(all_data, os.path.join(output_folder, "extracted_data.xlsx"))
            print("Данные сохранены в файл extracted_data.xlsx")
        else:
            print("Не удалось извлечь данные.")

if __name__ == "__main__":
    pdf_path = r'C:\Users\yacuc\Downloads\Исходные данные\Печатный\2.pdf'  # PDF-файл
    output_folder = 'output'  # папка для сохранения результата
    os.makedirs(output_folder, exist_ok=True)  # создает папку, если ее нет

    process_pdf(pdf_path, output_folder)
