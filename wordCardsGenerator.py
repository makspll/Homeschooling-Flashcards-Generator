import xlsxwriter
import os
import shutil
from PIL import Image
from icrawler.builtin import GoogleImageCrawler

def CalculateImageScaleRatioToFitSize(width, height, desiredSquareSize):
    heightRatio = desiredSquareSize/height
    widthRatio = desiredSquareSize/width

    biggestResize = min(heightRatio,widthRatio)
    print(biggestResize)
    return biggestResize

def InsertImageAtCellIfFound(row,col,word):
    foundPicture = False
    for extension in supportedExtensions:
        imagePath = GetImagePathFromWord(word,extension)
        if os.path.exists(imagePath):
            image = Image.open(imagePath)
            width,height = image.size

            scale = CalculateImageScaleRatioToFitSize(width,height, imageMaxSquareSize)
            print("Loaded:" + imagePath)
            worksheet.insert_image(row,0,imagePath,{'x_scale':scale,'y_scale':scale})
            foundPicture = True
            break

def GetImagePathFromWord(word, extension):
    imagePath = os.path.join(imagesDirName, word.lower() + extension)
    return imagePath

    return foundPicture

workbookName = "WordCards.xlsx"
worksheetName = "WordCards"
wordsFile = "Words.txt"
imagesDirName = "WordPictures"

if not os.path.exists(imagesDirName):
    os.makedirs(imagesDirName)

startingWordColumnIdx = 1
verticalGridPadding = 1

imageMaxSquareSize = 400
excellColumnWidth = (imageMaxSquareSize / 75) * 10 * 1.5

letterGridWidth = 2

supportedExtensions = [
    ".png",
    ".jpg",
    ".bmp",
]

# find words
wordsRaw = []

with open(wordsFile,'r') as words:
    line = words.readline()
    while line != "" and line != '\n':
        
        # remove whitespace at start and end
        formatedLine = line.strip()

        # trim by comma, if it's there
        if(len(formatedLine) > 0 and formatedLine[-1] == ','):
            formatedLine = formatedLine[:-1]

        wordsRaw.append(formatedLine)
        line = words.readline()

with xlsxwriter.Workbook(workbookName) as workbook:

    worksheet = workbook.add_worksheet(worksheetName)
    # set height    
    worksheet.set_default_row(imageMaxSquareSize / 2)


    all_cell_format = workbook.add_format({'bold':True})

    worksheet.set_column(startingWordColumnIdx,999,letterGridWidth,all_cell_format)
    letter_cell_format = workbook.add_format({'bold':True})
    letter_cell_format.set_border(1)
    letter_cell_format.set_align("center")
    letter_cell_format.set_align("vcenter")

    #iterate and write row by row
    wordIdx = 0
    row = 0
    col = 0

    for wordIdx in range(len(wordsRaw)):
        word = wordsRaw[wordIdx]
        while col < len(word):
            if(word[col] == ','):
                break
            worksheet.write(row,col + startingWordColumnIdx, word.upper()[col],letter_cell_format)
            col+= 1

        foundPicture = InsertImageAtCellIfFound(row,col,word)
        
        col = 0
        row += verticalGridPadding
    worksheet.set_column(0,0,excellColumnWidth)

