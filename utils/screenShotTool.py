import win32con
import win32gui
from datetime import datetime
import keyboard
import time
import sys
import mss
import numpy as np
import cv2 as cv
import os

# Global Variable
GLB_FLG_KILL = False
GLB_FLG_SCREEN = False

# SETTINGS
SECONDSTOMICRO = 1000000
FPSMICRO = 1000000
SCALEFORPREVIEW = 0.5
TARGETNAME = "Playnite"
PATHROOT = "utils/img/"
FOLDEROUTPUT = "Batch-1"


# Object for DPS counting
class FPS_counter:
    timeloop: datetime = None
    FPSact: int = 0

    def __init__(self, initialTime: datetime):
        self.timeloop = initialTime

    def update(self):
        timeGap = datetime.now() - self.timeloop
        self.timeloop = datetime.now()
        self.FPSact = int(FPSMICRO/(timeGap.seconds *
                                    SECONDSTOMICRO + timeGap.microseconds))
        print(f"FPS : {self.FPSact}")

# Object for target and shot 1 window


class ScreenShoter:

    monitorResolution = (0, 0)
    targetName: str = ""
    targetedWindow: int = None

    fl_kill: bool = False

    tg_left: int = 0
    tg_top: int = 0
    tg_width: int = 0
    tg_height: int = 0

    SS_OFFSET_LEFT = 8
    SS_OFFSET_TOP = 31

    def __init__(self, targetName: str):
        self.targetName = targetName
        self.updateTarget()

    def updateTarget(self):
        self.tg_left, self.tg_top, self.tg_width, self.tg_height = self._setTragetWindow()

    def _setTragetWindow(self):
        try:
            self.targetedWindow = win32gui.FindWindow(None, self.targetName)

        except Exception as e:
            msg = e.args[2]
            sys.exit("Crash : " + str(msg))

        return win32gui.GetWindowRect(self.targetedWindow)

    def shot(self):
        shot = mss.mss()
        screenSource = shot.grab(
            {'left': self.tg_left + self.SS_OFFSET_LEFT, 'top': self.tg_top + self.SS_OFFSET_TOP, 'width': self.tg_width, 'height': self.tg_height})
        image = np.array(screenSource)
        image = cv.cvtColor(image, cv.IMREAD_COLOR)
        return image

    def getResolution(self):
        return (self.tg_width, self.tg_height)


# Script for screen shot tool ==================================

def save():
    global GLB_FLG_SCREEN
    GLB_FLG_SCREEN = True


def close():
    global GLB_FLG_KILL
    GLB_FLG_KILL = True


def checkFolder(pathFolder: str):
    if not os.path.exists(pathFolder):
        os.makedirs(pathFolder)
        print(f"folder {pathFolder} created")


def main():
    global GLB_FLG_KILL
    global GLB_FLG_SCREEN
    # Define hotKey to interupt while
    keyboard.add_hotkey('ctrl + r', save)
    keyboard.add_hotkey('ctrl + q', close)

    fullpath = PATHROOT + FOLDEROUTPUT + "/"
    checkFolder(fullpath)

    # Definition du la fenetre source
    screenShoter = ScreenShoter(TARGETNAME)
    FPScount = FPS_counter(datetime.now())

    # definition de la fenetre de preview
    windowPreview = "Monitoring ScreenShoter"

    # Boucle d'aquisition
    while(True):
        # Update target position / take picture
        screenShoter.updateTarget()
        actFrame = screenShoter.shot()

        # Preview scale by 2
        cv.imshow(windowPreview, cv.resize(actFrame, None, fx=SCALEFORPREVIEW,
                                           fy=SCALEFORPREVIEW, interpolation=cv.INTER_LINEAR))

        # Update FPS and print it
        FPScount.update()
        cv.waitKey(1)

        if GLB_FLG_SCREEN:
            cv.imwrite(
                f"{fullpath} {FOLDEROUTPUT}- {format(datetime.now().date())} - {format(datetime.now().time(),'%H%M%S')}.jpg", actFrame)
            print(
                f"[INFO] Screen takken {FOLDEROUTPUT} - {format(datetime.now().date())} - {format(datetime.now().time(),'%H:%M:%S')}.jpg")
            GLB_FLG_SCREEN = False

        if GLB_FLG_KILL:
            cv.destroyAllWindows()
            break


if __name__ == "__main__":
    main()
