import os
import shutil
from docx2pdf import convert
import ppt2pdf
import doc2docx
from pptx import Presentation
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from io import BytesIO
import comtypes.client
import subprocess

# def convert_ppt_to_pdf(ppt_path, pdf_path):
#      try:
#         ppt2pdf.convert(ppt_path, pdf_path)
#         return True
#      except Exception as e:
#         print(f"Error converting {ppt_path} to PDF: {e}")
#         return False
    
def convert_ppt_to_pdf(inputFileName, outputFileName, formatType=32):
    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    powerpoint.Visible = 1

    if outputFileName[-3:] != 'pdf':
        outputFileName += ".pdf"

    deck = powerpoint.Presentations.Open(inputFileName)
    deck.SaveAs(outputFileName, formatType)  # formatType = 32 для ppt to pdf
    deck.Close()
    powerpoint.Quit()
    print("ppt", outputFileName)
    
    
def convert_doc_to_pdf(doc_path, pdf_path):
    try:
    #    doc2docx.convert(doc_path, pdf_path)
       file_name = os.path.splitext(pdf_path)[0]
       docx_filename = file_name + ".docx"
       pdf_filename = file_name + ".pdf"
       
       doc2docx.convert(doc_path, docx_filename)
       
       convert(docx_filename, pdf_filename)
       os.remove(docx_filename)
       print("file_name", file_name)
       print("doc", doc_path)
       print("pdf", pdf_path)
    except Exception as e:
        print(f"Error converting {doc_path} to PDF: {e}")
        return False

def convert_docx_to_pdf(docx_path, pdf_path):
    try:
        convert(docx_path, pdf_path)
        return True
    except Exception as e:
        print(f"Error converting {docx_path} to PDF: {e}")
        return False
    
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

def main(source_folder, destination_folder):
    if not os.path.exists(destination_folder):
        os.makedirs(destination_folder)

    for file in os.listdir(source_folder):
            if file.endswith(".pdf"):
                continue
                shutil.copy(os.path.join(source_folder, file), destination_folder)
            elif file.endswith(".docx"):
                docx_path = os.path.join(source_folder, file)
                pdf_path = os.path.join(destination_folder, os.path.splitext(file)[0] + ".pdf")
                convert_docx_to_pdf(docx_path, pdf_path)
            elif file.endswith(".doc"):
                doc_path = os.path.join(source_folder, file)
                pdf_path = os.path.join(destination_folder, os.path.splitext(file)[0] + ".pdf")
                convert_doc_to_pdf(doc_path, pdf_path)
            if file.endswith(".pptx"):
                doc_path = os.path.join(source_folder, file)
                pdf_path = os.path.join(destination_folder, os.path.splitext(file)[0] + ".pdf")
                convert_ppt_to_pdf(doc_path, pdf_path)
                
            if file.endswith(".ppt"):
                doc_path = os.path.join(source_folder, file)
                pdf_path = os.path.join(destination_folder, os.path.splitext(file)[0] + ".pdf")
                convert_ppt_to_pdf(doc_path, pdf_path)
                
            if file.endswith(".mov"):
                doc_path = os.path.join(source_folder, file)
                pdf_path = os.path.join(destination_folder, os.path.splitext(file)[0] + ".mp4")
                convert_mov_to_mp4(doc_path, pdf_path)
                

def traverse_directory(path):
    # Просматриваем все файлы и папки в указанной директории
    for item in os.listdir(path):
        item_path = os.path.join(path, item)
        # Если текущий элемент - директория, рекурсивно просматриваем её содержимое
        if os.path.isdir(item_path):
            print("Папка:", item_path)
            main(item_path, item_path)
            traverse_directory(item_path)
            
if __name__ == "__main__":
    
    source_folder = "data"
    destination_folder = 'result'
    # 
    
    # main(source_folder, destination_folder)
    
    paths = [ "D:\\Work\\Учебник для СПЕЦ ТЕХНИКИ\\Новые учебники\\10.05.05 КЭ 2024",
             "D:\\Work\\Учебник для СПЕЦ ТЕХНИКИ\\Новые учебники\\10.05.05 ТЗИ 2024",
             "D:\\Work\\Учебник для СПЕЦ ТЕХНИКИ\\Новые учебники\\40.03.02 УР(бк) 2024"]
    
    path = paths[2]
    traverse_directory(path)
            
    if not os.path.exists(destination_folder):
        os.makedirs(destination_folder)


print(path)
print("END")

