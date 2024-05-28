import subprocess
import os
import shutil

def convert_pdf_2_html(input_file):
# Путь к исполняемому файлу
    path_to_exe = "C:\\Users\\potte\\OneDrive\\Рабочий стол\\pdf2htmlEX-win32-0.14.6-upx-with-poppler-data\\pdf2htmlEX.exe"
    directory_exe = os.path.dirname(path_to_exe)

    directory_path = os.path.dirname(input_file)
    file_name = os.path.splitext(os.path.basename(input_file))[0]
   
    # Параметры командной строки
    save_path_html = "html\\file.html"
    params = [input_file, save_path_html]
    # params = [ "D:\\Work\\Учебник для СПЕЦ ТЕХНИКИ\\38.01.05 ЭБ 2024\\Материалы для зачета\\Теоретические вопросы для зачета.pdf"]

    # Запуск процесса
    subprocess.call([path_to_exe] + params)
    source = save_path_html
    dest = os.path.join(directory_path, file_name + ".html")
   
    
    print(source)
    print(dest)
    
    shutil.move(source,dest)
    
def find_pdfs(directory):
    pdf_files = []

    # Проходим по всем файлам и папкам в указанной директории
    for root, dirs, files in os.walk(directory):
        for file in files:
            if file.lower().endswith('.pdf'):
                pdf = os.path.join(root, file)
                # print(pdf)
                pdf_files.append(pdf)

    return pdf_files



paths = [ "D:\\Work\\Учебник для СПЕЦ ТЕХНИКИ\\Новые учебники\\10.05.05 КЭ 2024",
            "D:\\Work\\Учебник для СПЕЦ ТЕХНИКИ\\Новые учебники\\10.05.05 ТЗИ 2024",
            "D:\\Work\\Учебник для СПЕЦ ТЕХНИКИ\\Новые учебники\\40.03.02 УР(бк) 2024"]
path = paths[1]

pdf_files = find_pdfs(path)
trobule_pdfs = []


# convert_pdf_2_html(pdf_files[0])
for pdf_file in pdf_files:

    file_name = os.path.basename(pdf_file)  # Получить только имя файла
    file_name_without_extension = os.path.splitext(file_name)[0]
    directory = os.path.dirname(pdf_file)  # Получить путь к директории, в которой находится файл
    full_path_without_extension = os.path.join(directory, file_name_without_extension)  # Собрать полный путь без расширения

    html = (full_path_without_extension + ".html")
    if (os.path.exists(html)):
        print(" already exist", html)
        continue
    
    for trobule_pdf in trobule_pdfs:
        if (trobule_pdf == pdf_file):
            print("TROUBLE PDF: ", pdf_file)
            continue
    print("START CONVERT: " + pdf_file)
    convert_pdf_2_html(pdf_file)
    print("SUCCESS CONVERT: " + pdf_file)


print("END")