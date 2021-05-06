from pdf2image import convert_from_path
import glob,os
import os, subprocess
import shutil
from PIL import Image
import pytesseract as pt
import os
import _pyinstaller_hooks_contrib

#TRANSFORMO TUDO QUE ESTA EM PDF PARA JPG NA PASTA DE IMGNF

def pdf_convert():
    pdf_dir = r"C:\\Users\\pcordeiro\\Desktop\\Projeto abordin\\NFS\\"
    os.chdir(pdf_dir)
    for pdf_file in glob.glob(os.path.join(pdf_dir, "*.pdf")):
        pages = convert_from_path(pdf_file, 500)
        count= 0
        for page in pages:
            count += 1
            page.save(pdf_file[:-2] + "-" + str(count)+ ".jpg", 'JPEG')

    caminho_original = 'C:\\Users\\pcordeiro\\Desktop\\Projeto abordin\\NFS\\'
    caminho_novo = 'C:\\Users\\pcordeiro\\Desktop\\Projeto abordin\\IMGNF'
    try:
        os.mkdir(caminho_novo)
    except FileExistsError as e:
        pass

    for root, dirs, files in os.walk (caminho_original):
        for file in files:
            old_file_path = os.path.join(root,file)
            new_file_path = os.path.join(caminho_novo, file)

            if 'p-1.jpg' in file:
                shutil.move(old_file_path, new_file_path)
    return
pdf_convert()

def main():
    # path for the folder for getting the raw images
    path = "C:\\Users\\pcordeiro\\Desktop\\Projeto abordin\\IMGNF"

    # path for the folder for getting the output
    tempPath = "C:\\Users\\pcordeiro\\Desktop\\Projeto abordin\\NOTEPAD\\"


    # iterating the images inside the folder
    for imageName in os.listdir(path):
        inputPath = os.path.join(path, imageName)
        img = Image.open(inputPath)

        # applying ocr using pytesseract for python
        text = pt.image_to_string(img, lang="eng")



        fullTempPath = os.path.join(tempPath, 'time_' + imageName + ".txt")
        print(text)

        # saving the  text for every image in a separate .txt file
        file1 = open(fullTempPath, "w")
        file1.write(text)
        file1.close()


if __name__ == '__main__':
    main()
