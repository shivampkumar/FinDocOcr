# -*- coding: utf-8 -*-

from pytesseract import pytesseract
from PIL import Image
import re, cv2, os, glob, sys
import numpy as np
from autocorrect import spell
from docx import Document
from docx.shared import Inches
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import tempfile
import numpy as np
import sys, glob, os
from wand.image import Image as Img
sys.version


def add_to_doc(doc_path, list_of_pages):
    doc_path = 'finale/' + doc_path
    #print(doc_path)
    document = Document()
    #changing the page margins

    style = document.styles['Normal']
    font = style.font
    font.name = 'Menlo'
    font.size = Pt(6)
    paragraph_format = style.paragraph_format
    paragraph_format.line_spacing = 0.9
    sections = document.sections
    
    for section in sections:
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        section.left_margin = Inches(0.5)
        section.right_margin =Inches(0.5)
        
    for page in list_of_pages:
        para = document.add_paragraph(page)
        #para.aligment = WD_ALIGN_PARAGRAPH.LEFT
        document.add_page_break()
        
    document.save(doc_path)
        

def string_processing(files):
    list_of_pages = list()
    for name in files:
        str = ""
        #Get the doc name from the image name
        doc_path = name[-15:-6] + '.docx'
        #print(name)
        img = Image.open(name)
        data = pytesseract.image_to_data(img, config = '--psm 4 --oem 1')
        #print(data)
        #with open("lol.txt",'w') as f: f.write(str(data))
        '''
        Columns in order
        level, page_num, block_num, par_num, line_num, word_num, left, top, width, height, conf, text
        '''
        
        new_data = data
        new_data = new_data.replace('\t\n', '\tNoneValue\t')
        new_data = new_data.replace('\n', '\t')
        
        split_data = re.split('\t', new_data)
        if(len(split_data) % 12 == 11):
            split_data.append('   ')
        
        i = 12
        line_no = -1
        spend = 0
        prevlinelimit = 0
        c = 0
        cg = 0
        splen=18
        cha = 0
        prevmone=0
        countconsecmone = 0
        while (True):
            if (i >= len(split_data)):
                break
            temp = split_data[i + 11]

            if (temp == "="):
                i = i + 12
                #print("YSAAAAY")
                continue
            if ("_" in temp):
                i = i + 12
                #print("YSAAAAY")
                continue

            if  (prevmone == 1):
                line_no = split_data[i+7]
                if (countconsecmone < 4):
                    str += '\n'
                spend = 0
                cha = 0
       
            if (abs(int(line_no) - int(split_data[i + 7])) > 40):
                str += '\n'
                line_no = split_data[i + 7]
                spend = 0
                cha = 0
                countconsecmone = 0

            if (split_data[i + 10] == '-1'):
                if (countconsecmone < 4):
                    str = str + '\n'
                prevmone=1
                countconsecmone = countconsecmone + 1
            elif(temp.isspace() or temp==""):
                if (countconsecmone < 4):
                    str = str + '\n'
                prevmone=1
                countconsecmone = countconsecmone + 1
            else:
                prevmone=0
                if re.match("[0-9][0-9][0-9][0-9]20[0-9][0-9]",temp) is not None or re.match("[0-9][0-9][0-9]20[0-9][0-9]",temp) is not None:
                    if len(temp) == 7:
                        temp=temp[0:1]+'/'+temp[1:3]+"/"+temp[3:]
                    elif len(temp)==8:
                        temp=temp[0:2]+'/'+temp[2:4]+"/"+temp[4:]
                if re.match("[$][0-9]+.*",temp) is not None:
                    temp= temp[0:1]+ " "+ temp[1:]
                if temp=="S" or temp== "S$" or temp=="s":
                    temp="$"
                if((temp[-1]==',' or temp[-1]=='.') and len(temp)>1):
                    temp=temp[:-1]   
                if re.match("([.][.])+",temp) is not None or re.match("[_]+",temp) is not None or re.match("[-]+",temp) is not None or re.match("[—]+",temp) is not None:
                    temp=""
                if(re.match("[0-9]+", temp)):
                    temp = temp.replace("i", "1")
                    temp = temp.replace("s", "5")
                    temp = temp.replace("O", "0")
                    temp = temp.replace("o", "0")
                if re.match("[A-Za-z]+", temp) is not None and re.match("[A-Z]+", temp) is None:
                    temp = spell(temp)
                
                #str += temp + ' ' ye line aapne kyu add kiya mujhe nahi samjha
                temp=temp+' '
                temp = temp.replace('"','')
                temp = temp.replace("'",'')
                temp = temp.replace('“','')
                temp = temp.replace('_','')
                nu = int(split_data[i + 10])
                chbeg = int(split_data[i+6])//splen
                sp = int((chbeg-cha))
                while (sp > 0):
                    str += ' '
                    cha=cha+1
                    sp = sp - 1
                str = str + temp
                cha+=len(temp)
                spend = nu + int(split_data[i + 8])
            i = i + 12
        #str=str.rtrim()
        list_of_pages.append(str)
    #print("adding to doc")
    #print(list_of_pages)
    add_to_doc(doc_path, list_of_pages)
            

def load_data():
    dirName = 'pic/'
    listOfFile = os.listdir(dirName)
    #print(os.listdir(dirName))
    #print((listOfFile))
    for files in listOfFile:
        fullPath = os.path.join(dirName, files) + '/'
        fullPath += '*.jpg'
        files = sorted(glob.glob(fullPath))
        print(files)
        string_processing(files)
        
        
#os.mkdir("pice")
def generate_imfrompdf():   
    path = 'Dataset/*.pdf'
    files = glob.glob(path)
    print(files)
    for name in files:
        #print(name)
        i=0
        with Img(filename=name, resolution=300) as img:
            name=name[8:]
            name=name[:-4]
            #print(name)
            os.mkdir('pic/'+name)
            img.compression_quality = 100
            pic = 'pic/'+name+'/'+name+'.jpg'
            img.save(filename=pic)
            
def preprocess_pdfimg():
    path = 'pic/'
    data = sorted(glob.glob(path+'/*'+'/*.jpg'))
    for file in data:
        #print(file)
        img=cv2.imread(file)
        cv2.imwrite(file,preprocess(img))
        img= process_image_for_ocr(file)
        cv2.imwrite(file,img)



IMAGE_SIZE = 1800
BINARY_THREHOLD = 180


def process_image_for_ocr(file_path):
    # TODO : Implement using opencv
    temp_filename = set_image_dpi(file_path)
    im_new = remove_noise_and_smooth(temp_filename)
    return im_new


def set_image_dpi(file_path):
    im = Image.open(file_path)
    length_x, width_y = im.size
    factor = max(1, int(IMAGE_SIZE / length_x))
    size = factor * length_x, factor * width_y
    # size = (1800, 1800)
    im_resized = im.resize(size, Image.ANTIALIAS)
    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.jpg')
    temp_filename = temp_file.name
    im_resized.save(temp_filename, dpi=(300, 300))
    return temp_filename


def image_smoothening(img):
    ret1, th1 = cv2.threshold(img, BINARY_THREHOLD, 255, cv2.THRESH_BINARY)
    ret2, th2 = cv2.threshold(th1, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    blur = cv2.GaussianBlur(th2, (1, 1), 0)
    ret3, th3 = cv2.threshold(blur, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    return th3

def remove_noise_and_smooth(file_name):
    img = cv2.imread(file_name, 0)
    filtered = cv2.adaptiveThreshold(img.astype(np.uint8), 255, cv2.ADAPTIVE_THRESH_MEAN_C, cv2.THRESH_BINARY, 41, 3)
    kernel = np.ones((1, 1), np.uint8)
    opening = cv2.morphologyEx(filtered, cv2.MORPH_OPEN, kernel)
    closing = cv2.morphologyEx(opening, cv2.MORPH_CLOSE, kernel)
    img = image_smoothening(img)
    or_image = cv2.bitwise_or(img, closing)
    return or_image

def preprocess(img):
    img=cv2.cvtColor(img,cv2.COLOR_BGR2GRAY)
    img = cv2.bitwise_not(img)
    th2 = cv2.adaptiveThreshold(img,255, cv2.ADAPTIVE_THRESH_MEAN_C,cv2.THRESH_BINARY,15,-2)
    cv2.imwrite("th2.jpg", th2)
    horizontal = th2
    vertical = th2
    rows,cols = horizontal.shape
    horizontalsize = cols // 20
    horizontalStructure = cv2.getStructuringElement(cv2.MORPH_RECT, (horizontalsize,1))
    horizontal = cv2.erode(horizontal, horizontalStructure, (-1, -1))
    horizontal = cv2.dilate(horizontal, horizontalStructure, (-1, -1))
    #inverse the image, so that lines are black for masking
    #perform bitwise_and to mask the lines with provided mask
    horizontal_inv = cv2.bitwise_not(horizontal)
    masked_img = cv2.bitwise_and(img, img, mask=horizontal_inv)
    #reverse the image back to normal
    masked_img_inv = cv2.bitwise_not(masked_img)
    img= masked_img_inv
    return img

def main():
    #pytesseract.tesseract_cmd = r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"
    for folder in os.listdir("./pic"):
        print(folder)
        data=glob.glob('./pic/'+folder+'/*.jpg')
        print(data)
        for file in data:
            #print(file)
            os.remove(file)
        os.rmdir('./pic/'+folder)
    filelist=glob.glob(os.path.join("./finale", "*.docx"))
    for f in filelist:
        os.remove(f)
    generate_imfrompdf()
    preprocess_pdfimg()
    load_data()
    print("\n\n\t\t\t***Succesfully converted all PDFs.***\n\t\t\t***Check the finale folder for outputs.***")

if __name__ == '__main__':
    main()