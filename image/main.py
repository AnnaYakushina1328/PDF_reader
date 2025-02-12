import cv2
import pytesseract
import pandas as pd
import os
from pdf2image import convert_from_path
from PIL import Image

# устанавливаем новый лимит пикселей
Image.MAX_IMAGE_PIXELS = None


# путь к Tesseract OCR
pytesseract.pytesseract.tesseract_cmd = 'C:/Program Files/Tesseract-OCR/tesseract.exe'


def preprocess_image(image_path):
    image = cv2.imread(image_path, cv2.IMREAD_GRAYSCALE)

    # снижение шума с помощью GaussianBlur
    image = cv2.GaussianBlur(image, (5, 5), 0)

    # гистограммное выравнивание для улучшения контраста
    image = cv2.equalizeHist(image)

    # Бинаризация с фиксированным порогом
    _, processed_image = cv2.threshold(image, 127, 255, cv2.THRESH_BINARY)

    return processed_image


def rotate_image(image_path):
    try:
        # Чтение изображения через OpenCV
        image = cv2.imread(image_path)

        # Поворот на 90 градусов против часовой стрелки
        rotated_image = cv2.rotate(image, cv2.ROTATE_90_COUNTERCLOCKWISE)

        # Сохранение файла
        cv2.imwrite(image_path, rotated_image)
    except Exception as e:
        print(f"Ошибка при повороте изображения {image_path}: {e}")


def extract_text_from_image(image):
    # настройка Tesseract с разными параметрами PSM для лучшей сегментации
    config = '--psm 6 -c preserve_inter-word_spaces=1'  # PSM в зависимости от структуры
    text = pytesseract.image_to_string(image, config=config, lang='rus')
    return text


def extract_data_from_text(text):
    lines = text.splitlines()
    data = [line for line in lines if any(char.isdigit() for char in line)]
    return data


def save_data_to_file(data, output_file):
    df = pd.DataFrame(data, columns=["Extracted Data"])
    df.to_excel(output_file, index=False)


def process_pdf(pdf_path, output_folder):
    try:
        images = convert_from_path(pdf_path, dpi=300)  # увеличиваем DPI для лучшего качества
        if not images:
            print("Не удалось конвертировать PDF в изображения.")
            return
    except Exception as e:
        print(f"Ошибка при конвертации PDF: {e}")
        return
    all_data = []

    for i, page_image in enumerate(images):
        image_path = os.path.join(output_folder, f"page_{i + 1}.jpg")
        page_image.save(image_path, "JPEG", quality=95)  # увеличиваем качество сохраненного изображения
        rotate_image(image_path)
        print(f"Изображение сохранено: {image_path}")

        processed_image = preprocess_image(image_path)

        # распознавание текста
        text = extract_text_from_image(processed_image)
        print(f"Распознанный текст на странице {i + 1}: \n{text}")

        # извлечение данных
        data = extract_data_from_text(text)
        all_data.extend(data)

    # сохранение данных
    if all_data:
        save_data_to_file(all_data, os.path.join(output_folder, "extracted_data.xlsx"))
        print("Данные сохранены в файл extracted_data.xlsx")
    else:
        print("Не удалось извлечь данные.")


if __name__ == "__main__":
    pdf_path = r'C:\Users\yacuc\Downloads\Исходные данные\Печатный\2.pdf'  # PDF-файл
    output_folder = 'output'  # папка для сохранения результата
    os.makedirs(output_folder, exist_ok=True)

    process_pdf(pdf_path, output_folder)
