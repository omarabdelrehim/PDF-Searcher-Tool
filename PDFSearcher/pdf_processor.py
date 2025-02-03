import os
from PyPDF2 import PdfReader

from file_manager import move_file ,move_file2
import pdfplumber
from PyQt5.QtWidgets import QMessageBox
import pytesseract
from pdf2image import convert_from_path
from PIL import Image
import pytesseract
import fitz  # PyMuPDF
from PIL import Image
import io

import re



def process_pdfs_in_folder(folder_path,otherword,Page_No,page_end):
    print("Process Start  OCR ,Please Wait")
    """
    Processes all PDF files in the given folder, checking for specific content related to 'Back Margin'.

    Args:
        folder_path (str): Path to the folder containing PDF files.

    Returns:
        list: A list of dictionaries containing the file name and back margin status.
        
    """
    
 
    results = []
    result2={}
    for filename in os.listdir(folder_path):
            if filename.endswith(".pdf"):
                file_path = os.path.join(folder_path, filename)
                try:
                    with open(file_path, "rb") as f:
                        reader = PdfReader(f)
                    #has_back_margin = check_back_margin(reader)
                    #has_marketing_support= check_Marketing(reader)
                    
                                    #"Back Margin": "Yes" if has_back_margin else "No",
                                    #"Marketing Support" :"Yes" if has_marketing_support else "No"}
                        
                        for i in otherword:
                                has_other= check_other(reader,i,Page_No,page_end,file_path)
                                result2.update({ i: "Yes" if has_other[0] else "No"})
                                
                                
                                 
                        Dic={"File Name": filename,"Check" :has_other[1]}
                        Dic.update(result2)
                        results.append(Dic) 
                        
        #move_file(results,folder_path, os.path.join(folder_path, 'Contract with Back Margin'))
       # move_file2(results,folder_path, os.path.join(folder_path,  'Marketing Support'))
                except Exception as e:
                    print(f"Error reading {file_path}: {e}")
                    for i in otherword:
                        result2.update({ i: "Check Failed"})
                    Dic={"File Name": filename,"Check" :"Failed"}
                    Dic.update(result2)
                    results.append(Dic)  
                        
                    
            
    print(results)
    print("Process Complete")
    return results

def process_pdfs_in_folder_without_OCR(folder_path,otherword,Page_No,page_end):
    print("Process Start  Without OCR ,Please Wait")
    """
    Processes all PDF files in the given folder, checking for specific content related to 'Back Margin'.

    Args:
        folder_path (str): Path to the folder containing PDF files.

    Returns:
        list: A list of dictionaries containing the file name and back margin status.
        
    """
    

    results = []
    result2={}
    for filename in os.listdir(folder_path):
            if filename.endswith(".pdf"):
                file_path = os.path.join(folder_path, filename)
                try:
                    with open(file_path, "rb") as f:
                        reader = PdfReader(f)
                    #has_back_margin = check_back_margin(reader)
                    #has_marketing_support= check_Marketing(reader)
                    
                                    #"Back Margin": "Yes" if has_back_margin else "No",
                                    #"Marketing Support" :"Yes" if has_marketing_support else "No"}
                        x= max(0, min(Page_No, len(reader.pages)-1))
                        v = max(0, min(page_end, len(reader.pages)-1))
                        
                        
                        for i in otherword:
                                has_other= check_other_without_ocr(reader,i,x,v,file_path)
                                result2.update({ i: "Yes" if has_other[0] else "No"})
                                
                                
                                 
                        Dic={"File Name": filename,"Check" :has_other[1]}
                        Dic.update(result2)
                        results.append(Dic) 
                        
        #move_file(results,folder_path, os.path.join(folder_path, 'Contract with Back Margin'))
       # move_file2(results,folder_path, os.path.join(folder_path,  'Marketing Support'))
                except Exception as e:
                    print(f"Error reading {file_path}: {e}")
                    for i in otherword:
                        result2.update({ i: "Check Failed"})
                    Dic={"File Name": filename,"Check" :"Failed"}
                    Dic.update(result2)
                    results.append(Dic)  
                        
                    
    print("Process Complete")
    
    return results


def check_other_without_ocr(reader,otherword,Page_No,page_end,file_path):
    
    if Page_No >len(reader.pages):
        x=len(reader.pages)
    else: 
        x=Page_No
        
    if page_end >len(reader.pages):
        v=len(reader.pages)
    else: 
        v=page_end
    text=""
    Result=[]
    otherword_lower=""
    for page in range(x,v):
            page = reader.pages[page]
            
            otherword_lower=str(otherword).lower()
            text += page.extract_text()
    
    
    if  re.search(rf'\b{re.escape(otherword_lower)}\b', text.lower(), re.IGNORECASE):
                Result=[True,"Sucess"]
    
    else:
        Result= [False,"Sucess"]
  
    
    return Result

def check_Marketing(reader):
    """
    Check if the given PDF file contains the phrase related to back margin.

    Args:
        file_path (str): Path to the PDF file.

    Returns:
        bool: True if the PDF contains the phrase, False otherwise.
    """
    try:
        for page in range(14, len(reader.pages)):
                page = reader.pages[page]
                text = page.extract_text()
                if "marketing support" in text.lower():
                    return True
                
    except Exception as e:
        print(f"Error reading {reader}: {e}")
    
    return False

def check_other(reader,otherword,Page_No,page_end,file_path):
    print("Process in progress  OCR ,Please Wait")
    """
    Check if the given PDF file contains the phrase related to back margin.

    Args:
        file_path (str): Path to the PDF file.

    Returns:
        bool: True if the PDF contains the phrase, False otherwise.
    """
    tesseract_path = os.path.join(os.getcwd(), "tesseract", "tesseract.exe")
    pytesseract.pytesseract.tesseract_cmd = tesseract_path
    tessdata_dir = os.path.join(os.getcwd(), "tesseract", "tessdata")
    os.environ["TESSDATA_PREFIX"] = tessdata_dir
    
    if Page_No >len(reader.pages):
        x=len(reader.pages)
    else: 
        x=Page_No
        
    if page_end >len(reader.pages):
        v=len(reader.pages)
    else: 
        v=page_end
    text=""
    Result=[]
    otherword_lower=""
    for page in range(x,v):
            page = reader.pages[page]
            
            otherword_lower=str(otherword).lower()
            text += page.extract_text()
    
    
    if  re.search(rf'\b{re.escape(otherword_lower)}\b', text.lower(), re.IGNORECASE):
                Result=[True,"Sucess"]
    else:
        text= pdf_to_images(file_path,x,v) 
        if  re.search(rf'\b{re.escape(otherword_lower)}\b', text.lower(), re.IGNORECASE):
            Result=[True,"Sucess"]
        else:
            Result= [False,"Sucess"]
    return Result

                
                
    
        

def pdf_to_images(pdf_path,x,v):
    # Open the PDF file
    print(" OCR Process in progress  OCR ,Please Wait")
    
    doc = fitz.open(pdf_path)
    images = []
    start=0
    end=0
    if x > doc.page_count:
        x=doc.page_count
    else: 
        start=x
        
    if v > doc.page_count:
        end=doc.page_count
    else: 
        end=v
    
    print(start)
    print(end)
    text=""


    # Iterate over each page in the PDF
    for page_num in range(start,end):
       
        page = doc.load_page(page_num)  # Load the page
        pix = page.get_pixmap(dpi=200)  # Render page to a Pixmap (image)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)  # Convert to PIL Image
        img = img.convert("L")
        
        images.append(img)
    for image in images:
        text += pytesseract.image_to_string(image)
    
    return text
    


import pdfplumber
import re
import os

def check_other2(file_path, otherwords):
    try:
        if not os.path.exists(file_path):
            print(f"File does not exist: {file_path}")
            return False
        
        with pdfplumber.open(file_path) as pdf:
            print(f"Opened PDF with {len(pdf.pages)} pages.")
            for page_num in range(14, len(pdf.pages)):
                page = pdf.pages[page_num]
                tables = page.extract_tables()

                # Loop through the tables on the page
                for table in tables:
                    for row in table:
                        for cell in row:
                            # Make sure the cell is a string before using re.search
                            if cell:
                                cell = str(cell)  # Convert the cell to a string if it's not None
                                for word in otherwords:
                                    if re.search(rf'\b{re.escape(word)}\b', cell, re.IGNORECASE):
                                        print(f"Found '{word}' in table on page {page_num+1}")
                                        return True
                            else:
                                print(f"Skipping empty or None cell at row {row}, page {page_num+1}")
    except Exception as e:
        print(f"Error reading {file_path}: {e}")
    
    return False

