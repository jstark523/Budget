# inconsistent results unless I utilized two dfs and workbooks

from PIL import Image
from pytesseract import pytesseract
import os
import pandas as pd
import numpy as np

df = pd.DataFrame(
{
})

df2 = pd.DataFrame(
    {     
    }
)

#Define path to tessaract.exe
path_to_tesseract = r'/usr/local/Cellar/tesseract/5.2.0/bin/tesseract'

#Define path to image
jasonImgPath = 'jasonSpending/'
melImgPath = 'melSpending/'

col = 0

col2 = 0

#Point tessaract_cmd to tessaract.exe
pytesseract.tesseract_cmd = path_to_tesseract

for root, dirs, file_names in os.walk(jasonImgPath):
    #Iterate over each file_name in the folder
    for file_name in file_names:
        #Open image with PIL
        jImg = Image.open(jasonImgPath + file_name)

        #Extract text from image
        jText = pytesseract.image_to_string(jImg)

        #Comma Seperate at new lines
        # jText = jText.split('\n')
        jText = [x.strip() for x in jText.split("\n")]


        #Create Dictionary, Series, and column name
        jDic = {'Jason': jText}
        jSer = pd.Series(data=jDic)
        cName = "Jason" + str(col)
        
        #Add Column and increment
        df.insert(loc=col,column= cName, value = jSer)
        col+=1

        

for root, dirs, file_names in os.walk(melImgPath):
    #Iterate over each file_name in the folder
    for file_name in file_names:
        #Open image with PIL
        mImg = Image.open(melImgPath + file_name)

        #Extract text from image
        mText = pytesseract.image_to_string(mImg)

        #Comma Seperate at new lines
        # mText = mText.split('\n')

        mText = [x.strip() for x in mText.split("\n")]


        #Create Dictionary, Series, and column name
        mDic = {'Mel': mText}
        mSer = pd.Series(data=mDic)
        cName = "Mel" + str(col2)

        #Add Column and increment
        df2.insert(loc=col2,column= cName, value = mSer)
        col2+=1

#create workbook
df.to_excel("jason.xlsx", sheet_name='Jason') 
df2.to_excel("mel.xlsx", sheet_name='Mel')
