# IDE / SHELL needs to be run as Admin to work!
from datetime import datetime
import threading

import PySimpleGUI as pg
import pyautogui
import time
import win32gui
import sys

interval = 0
gui_done = False
WAIT = 2
i = 0
stop = False

def build_gui():
    window = pg.Window('App', layout=[[pg.Text('Insert New Application Name: ', size=(25, 1)),
                                       pg.In(size=(25, 1), default_text='Neuer Tab - Google Chrome',
                                             key='application_name_new'), pg.Image('Example_Name.PNG'),
                                       pg.Text('(Example)')],
                                      [pg.Text('Insert Old Application Name: ', size=(25, 1)),
                                       pg.In(size=(25, 1), default_text='App', key='application_name_old')],
                                      [pg.Text('Insert Duration (min): ', size=(25, 1)),
                                       pg.In(size=(25, 1), default_text='5', key='duration')],
                                      [pg.Button(button_text='Start', key='btn_start')]], margins=(200, 200)).read()
    return window

def build_confirm():
    window_close = pg.Window('Confirm', layout=[[pg.Button(button_text='Stop', key='btn_stop')]], margins=(30,30)).read()
    return window_close


def run(hwnd, hwnd_old, interval, event):
    print(interval)
    while not stop:
        event.wait(2)
        win32gui.SetForegroundWindow(hwnd)
        event.wait(2)
        pyautogui.keyDown('w')
        event.wait(5)
        pyautogui.keyUp('w')
        event.wait(2)
        win32gui.SetForegroundWindow(hwnd_old)
        event.wait(interval)

def btn_start(e):
    print(threading.enumerate())
    hwnd = win32gui.FindWindowEx(0, 0, 0, value['application_name_new'])
    hwnd_old = win32gui.FindWindowEx(0, 0, 0, value['application_name_old'])
    interval = float(value['duration']) * 60
    run(hwnd, hwnd_old, interval, e)


if __name__ == '__main__':
    e = threading.Event()
    window = build_gui()
    while True:
        event, value = window
        if event == 'btn_start':
            print(threading.enumerate())
            threading.Thread(target=btn_start, args=(e,)).start()
            break
    window_end = build_confirm()
    while True:
        event, value = window_end
        if event == 'btn_stop':
            e.set()
            stop = True
            sys.exit()
