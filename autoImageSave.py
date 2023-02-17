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
import sys
import re

SAVE_PATH = ".\\images\\"

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


def main(title=None):
    win32clipboard.OpenClipboard()
    if win32clipboard.IsClipboardFormatAvailable(win32con.CF_DIB):
        clip0 = win32clipboard.GetClipboardData(win32con.CF_DIB)
    win32clipboard.CloseClipboard()
    try:
        while True:
            win32clipboard.OpenClipboard()
            if win32clipboard.IsClipboardFormatAvailable(win32con.CF_DIB):
                clip1 = win32clipboard.GetClipboardData(win32con.CF_DIB)
                win32clipboard.CloseClipboard()
                time.sleep(0.5)
                if clip0!=clip1:
                    img = ImageGrab.grabclipboard()
                    imgPath = saveResizedImg(img)
                    clip0=clip1
                    continue

    except pywintypes.error as e:
        print(e)
        time.sleep(1)
        saveResizedImg(img)
        main()     

    except KeyboardInterrupt as e:
        print(e)

if __name__ == "__main__":
        main()