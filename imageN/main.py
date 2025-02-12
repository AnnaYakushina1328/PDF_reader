import fitz  # PyMuPDF
import re
import tensorflow as tf
from tensorflow.keras.models import Model
from tensorflow.keras.layers import Input, Conv2D, MaxPooling2D, Flatten, Dense, LSTM, Reshape


def extract_text_from_pdf(pdf_path):
    doc = fitz.open(pdf_path)
    text = ""
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        text += page.get_text()
    return text

pdf_path = 'path_to_pdf_file.pdf'
extracted_text = extract_text_from_pdf(pdf_path)
print(extracted_text)


def preprocess_text(text):
    # удаление лишних пробелов и символов
    text = re.sub(r'\s+', ' ', text)
    return text

processed_text = preprocess_text(extracted_text)
print(processed_text)


def create_complex_model(input_shape, vocab_size, embedding_dim):
    # входной слой
    inputs = Input(shape=input_shape)

    # сверточные слои
    x = Conv2D(32, (3, 3), activation='relu', padding='same')(inputs)
    x = MaxPooling2D((2, 2))(x)
    x = Conv2D(64, (3, 3), activation='relu', padding='same')(x)
    x = MaxPooling2D((2, 2))(x)

    # преобразование данных для LSTM
    x = Reshape((input_shape[0] // 4, -1))(x)

    # рекуррентные слои
    x = LSTM(128, return_sequences=True)(x)
    x = LSTM(64)(x)

    # полносвязные слои
    x = Dense(128, activation='relu')(x)
    outputs = Dense(vocab_size, activation='softmax')(x)

    # создание модели
    model = Model(inputs, outputs)
    model.compile(optimizer='adam', loss='sparse_categorical_crossentropy', metrics=['accuracy'])

    return model

# параметры модели
input_shape = (100, 10, 1)  # пример размера входных данных
vocab_size = 1000           # размер словаря
embedding_dim = 64

model = create_complex_model(input_shape, vocab_size, embedding_dim)
model.summary()

# данные для обучения
# X_train, y_train = ...

# обучение модели
# model.fit(X_train, y_train, epochs=10, batch_size=32)
