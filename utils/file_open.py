import os


def xlsx_filename(folder_path):
    for file_name in os.listdir(folder_path):
        if file_name.endswith('.xlsx'):  # если файл - xlsx
            file_path = os.path.join(folder_path, file_name)  # получаем полный путь к файлу
            return file_path  # возвращаем путь к файлу
    else:
        return None  # если xlsx файлы не найдены