# IDE / SHELL needs to be run as Admin to work!
from datetime import datetime
import threading

import PySimpleGUI as pg
import pyautogui
import time
import win32gui
import sys

from PySimpleGUI import WIN_CLOSED

instructions = '1. Enter the name of the window you wish to switch to from the taskbar or select it from the list on the right: the neccessary input for the' \
               ' given example would be \'Google Chrome - Neuer Tab\'.\n' \
               '2. Enter the name of the window to switch back to (optional).\n' \
               '3. Enter the interval in minutes, in which the application should execute the specified keystroke.\n' \
               '4. Enter the desired key to be pressed in the switched to application.\n'
interval = 0
gui_done = False
WAIT = 2
i = 0
stop = False
list = []

def winEnumHandler(hwnd, x):
        if win32gui.IsWindowVisible(hwnd):
            s = win32gui.GetWindowText(hwnd)
            if len(s) >= 1:
                list.append(win32gui.GetWindowText(hwnd))

print(win32gui.EnumWindows(winEnumHandler, None))

def build_gui(typo):
    if typo:
        window = pg.Window('App', layout=[[pg.Text('Instructions', size=(120,1))], [pg.Text('Typo in Application Name!', key='typo', text_color='red')], [pg.Text(instructions, size=(120,6))],[pg.Text('Insert New Application Name: ', size=(25, 1)),
                                           pg.In(size=(25, 1), default_text='Neuer Tab - Google Chrome',
                                                 key='application_name_new'), pg.Image('Example_Name.PNG'),
                                           pg.Text('(Example)'), pg.Listbox(list, size=(30, 6))],
                                          [pg.Text('Insert Old Application Name: ', size=(25, 1)),
                                           pg.In(size=(25, 1), default_text='App', key='application_name_old')],
                                          [pg.Text('Insert Duration (min): ', size=(25, 1)),
                                           pg.In(size=(25, 1), default_text='5', key='duration')],
                                          [pg.Button(button_text='Start', key='btn_start', size=(10, 0), button_color='green'),
                                            [pg.Text('', size=(10, 0))],
                                           [pg.Button(button_text='Stop', key='btn_stop', size=(10,0), button_color='red')]]], margins=(0, 0))
    else:
        window = pg.Window('App',
                           layout=[[pg.Text('Instructions', size=(120, 1))], [pg.Text(instructions, size=(120, 6))],
                                   [pg.Text('Insert New Application Name: ', size=(25, 1)),
                                    pg.In(size=(25, 1), default_text='Neuer Tab - Google Chrome',
                                          key='application_name_new'), pg.Image('Example_Name.PNG'),
                                    pg.Text('(Example)'), pg.Listbox(list, size=(30, 6))],
                                   [pg.Text('Insert Old Application Name: ', size=(25, 1)),
                                    pg.In(size=(25, 1), default_text='App', key='application_name_old')],
                                   [pg.Text('Insert Duration (min): ', size=(25, 1)),
                                    pg.In(size=(25, 1), default_text='5', key='duration')],
                                   [pg.Button(button_text='Start', key='btn_start', size=(10, 0), button_color='green'),
                                    pg.Text('', size=(10, 0)),
                                    pg.Button(button_text='Stop', key='btn_stop', size=(10, 0), button_color='red')]],
                           margins=(0, 0))

    return window

def catch_spelling_error(text_in):
    for s in list:
        if s == text_in:
            return True
    return False


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
    window = build_gui(False)
    while True:
        event, value = window.read()
        print(event)
        if event == 'btn_start':
            print(threading.enumerate())
            print(value['application_name_new'])
            if catch_spelling_error(value['application_name_new']):
                threading.Thread(target=btn_start, args=(e,)).start()
                window.refresh()
        if event == 'btn_stop' or event == WIN_CLOSED:
            e.set()
            stop = True
            sys.exit()