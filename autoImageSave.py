import win32clipboard
import win32con
from PIL import ImageGrab
import datetime
import time
import pywintypes

SAVE_PATH = ".\\images\\"

def saveImg(img):
    now = datetime.datetime.now()
    fname = now.strftime('%Y%m%d%H%M%S')
    imgPath = f'{SAVE_PATH}{fname}.png'
    img.save(imgPath)

    return imgPath

def main():
    try:
        win32clipboard.OpenClipboard()
        if win32clipboard.IsClipboardFormatAvailable(win32con.CF_DIB):
            clip0 = win32clipboard.GetClipboardData(win32con.CF_DIB)
        else:
            clip0=""  
    finally:
        win32clipboard.CloseClipboard()

    while True:
        try:
            win32clipboard.OpenClipboard()
            if win32clipboard.IsClipboardFormatAvailable(win32con.CF_DIB):
                clip1 = win32clipboard.GetClipboardData(win32con.CF_DIB)
                if clip0!=clip1:
                    img = ImageGrab.grabclipboard()
                    imgPath = saveImg(img)
                    print(f"saved:{imgPath}")
                    clip0=clip1
                    continue

        except pywintypes.error as e:
            print(e)
            time.sleep(1)
            continue
        
        except KeyboardInterrupt as e:
            print(e)

        else:
            win32clipboard.CloseClipboard()

        time.sleep(0.5)

if __name__ == "__main__":
        main()