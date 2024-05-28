<h1>Репозиторий содержит 2 скрипта</h1>

1) <h2>2)main.py</h2> 
Конвертирует файлы 
```
.doc, .docx, .ppt, .pptx -> .pdf
```
и
```
.mov -> mp4
```
Для запуска скрипта нужна программа ffmpeg, скачать тут
```
https://ffmpeg.org/
```
далее указать путь до exe
```
def convert_mov_to_mp4(mov_path, mp4_path):
    try:
        path_to_exe = "ffmpeg-7.0-full_build/bin/ffmpeg.exe"
        params = ["-i", mov_path, "-qscale", "0", "-c" ,"copy", mp4_path , "-y"]      

        # Запуск процесса
        print([path_to_exe] + params)
        subprocess.call([path_to_exe] + params)
        return True
    except Exception as e:
        print(f"Error converting {mov_path} to PDF: {e}")
        return False

```
Некоторые файлы бесконечно конвертит. В логах смотрим какой файл удаляем его и конвертим в ручную

<h2>2)pdfToHtml.py</h2> 
Конверитуре файлы 
```
.pdf -> .html
```
Нужно скачать программу pdf2htmlEX и указать путь до exe
```
https://pdf2htmlex.github.io/pdf2htmlEX/
```

И указать путь до exe
```
def convert_pdf_2_html(input_file):
# Путь к исполняемому файлу
    path_to_exe = "C:\\Users\\potte\\OneDrive\\Рабочий стол\\pdf2htmlEX-win32-0.14.6-upx-with-poppler-data\\pdf2htmlEX.exe"
    ...
```







