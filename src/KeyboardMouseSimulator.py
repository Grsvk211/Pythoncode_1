import pyautogui
from pynput.mouse import Button, Controller as mouseController
from pynput.keyboard import Key, Controller as keyboardController

import InputConfigParser as ICF
import logging
pyautogui.PAUSE = 2.5
mouse = mouseController()
keyboard = keyboardController()


def showWindow(appTitle):
    title = ".*" + appTitle + "*."
    from pywinauto.application import Application
    app = Application(backend="uia")
    logging.info(ICF.getExcelPath())
    app = app.connect(path=ICF.getExcelPath())
    top_window = app.window(title_re=title, visible_only=False)
    # logging.info("top_window.restore().get_focus()", top_window.restore().has_focus())
    # [Fix me] Trying to set focus for an application which is already in focus
    # top_window.restore().set_focus()
    top_window.maximize()


def getMonitorResolution():
    screenWidth, screenHeight = pyautogui.size()
    screenSize = [screenWidth, screenHeight]
    return screenSize


def getCurrentXYMousePos():
    currentMouseX, currentMouseY = pyautogui.position()
    mouseXY = [currentMouseX, currentMouseY]
    return mouseXY


def moveMouse(x, y):
    # pyautogui.moveTo(x, y)
    mouse.position = (x, y)


def mouseClick():
    # pyautogui.click()
    # Press and release
    mouse.press(Button.left)
    mouse.release(Button.left)


def triggerHotkey(hotKey1, hotKey2):
    pyautogui.hotkey('hotKey1', 'hotKey2')


def triggerKey(key):
    pyautogui.press(key)


def getCoordinatesByImage(img):
    coords = pyautogui.locateOnScreen(img)
    return coords


def getCoordinatesOfChrome():
    pass


def openChrome():
    x, y = getCoordinatesOfChrome()
    moveMouse(x, y)
    mouseClick()


def maximiseWindow(wId):
    pass


def pasteString(str):
    keyboard.type(str)


def pressEnter():
    # Press and release enter
    keyboard.press(Key.enter)
    keyboard.release(Key.enter)


def pressTab():
    # Press and release enter
    keyboard.press(Key.tab)
    keyboard.release(Key.tab)


def rightClick():
    # Press and release enter
    keyboard.press(Key.shift)
    keyboard.press(Key.f10)
    keyboard.release(Key.shift)
    keyboard.release(Key.f10)


def downArrow():
    # Press and release enter
    keyboard.press(Key.down)
    keyboard.release(Key.down)


def upArrow():
    # Press and release enter
    keyboard.press(Key.up)
    keyboard.release(Key.up)


def rightArrow():
    # Press and release enter
    keyboard.press(Key.right)
    keyboard.release(Key.right)


def leftArrow():
    # Press and release enter
    keyboard.press(Key.left)
    keyboard.release(Key.left)


def addNewRow():
    # Press [alt + I + R] and release
    keyboard.press(Key.alt)
    keyboard.press('I')
    keyboard.release('I')
    keyboard.release(Key.alt)
    keyboard.press('R')
    keyboard.release('R')