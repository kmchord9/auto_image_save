import win32clipboard
import win32con
from PIL import ImageGrab, Image
import datetime
import time
import pywintypes
from pptx import Presentation
from pptx.util import Inches, Pt 
from pptx.util import Cm
import os

SAVE_PATH = ".\\images\\"

def pptxAddImage(imgPath, text=None):
    now = datetime.datetime.now()
    pptxFileName = now.strftime('%Y%m%d')
    pptxSavePath = f"{SAVE_PATH}{pptxFileName}.pptx"

    if os.path.exists(pptxSavePath):
        prs = Presentation(pptxSavePath)
    else:
        prs = Presentation()

    sld0 = prs.slides.add_slide(prs.slide_layouts[6])

    left = top = width = height = Inches(0.5)
    txBox = sld0.shapes.add_textbox(left, top, width, height) 

    #testBox
    pa = txBox.text_frame.paragraphs[0]

    if text==None:
        pa.text = "This is text inside a textbox"
    else:
        pa.text=text
    pa.font.size = Pt(28)

    pic0 = sld0.shapes.add_picture(imgPath,Cm(1), Cm(3))

    prs.save(pptxSavePath) 

    return


def saveResizedImg(img):
    now = datetime.datetime.now()
    fname = now.strftime('%Y%m%d%H%M%S')
    resizedImg = imgResize(img)
    imgPath = f'{SAVE_PATH}{fname}.png'
    resizedImg.save(imgPath)

    return imgPath

def imgResize(img):
    MAX_WIDTH = 680
    MAX_HEIGHT = 450
     
    imgWidth, imgHeight = img.size

    if imgWidth>MAX_WIDTH or imgHeight>MAX_HEIGHT:
        xRatio = MAX_WIDTH/imgWidth
        yRatio = MAX_HEIGHT/imgHeight

        if xRatio < yRatio:
            size = (MAX_WIDTH, round(imgHeight*xRatio))
        else:
            size = (round(imgWidth*yRatio), MAX_HEIGHT)
        resizedImg = img.resize(size)

        return resizedImg
    else:
        return img


def main():
    win32clipboard.OpenClipboard()
    if win32clipboard.IsClipboardFormatAvailable(win32con.CF_DIB):
        clip0 = win32clipboard.GetClipboardData(win32con.CF_DIB)
    else:
        clip0=""
    win32clipboard.CloseClipboard()
    try:
        while True:
            win32clipboard.OpenClipboard()
            if win32clipboard.IsClipboardFormatAvailable(win32con.CF_DIB):
                clip1 = win32clipboard.GetClipboardData(win32con.CF_DIB)
                if clip0!=clip1:
                    img = ImageGrab.grabclipboard()
                    imgPath = saveResizedImg(img)
                    pptxAddImage(imgPath)
                    clip0=clip1
                    continue
            win32clipboard.CloseClipboard()
            time.sleep(0.5)
    except pywintypes.error as e:
        print(e)
        time.sleep(1)
        saveResizedImg(img)
        main()     

    except KeyboardInterrupt as e:
        print(e)

if __name__ == "__main__":
    main()