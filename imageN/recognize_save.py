import pandas as pd


def recognize_and_save_to_excel(model, data, output_path):
    # распознавание данных
    recognized_data = model.predict(data)

    # преобразование распознанных данных в DataFrame
    df = pd.DataFrame(recognized_data)

    # сохранение в Excel
    df.to_excel(output_path, index=False)

# пример использования
# recognize_and_save_to_excel(model, new_data, 'recognized_output.xlsx')
