# IDE / SHELL needs to be run as Admin to work!
import threading

import PySimpleGUI as pg
import pyautogui
import time
import win32gui

interval = 0
gui_done = False
WAIT = 2
i = 0

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

def run(hwnd, hwnd_old, interval, repetitions):
    i = 0
    print(interval)
    while i < repetitions:
        i += 1
        win32gui.SetForegroundWindow(hwnd)
        threading.Timer(WAIT, pyautogui.keyDown('w'))
        time.sleep(WAIT)
        threading.Timer(WAIT, pyautogui.keyUp('w'))
        threading.Timer(WAIT, win32gui.SetForegroundWindow(hwnd_old))
        threading.Timer(interval * 1000, print(i))
        print("done")

def btn_start():
    hwnd = win32gui.FindWindowEx(0, 0, 0, value['application_name_new'])
    hwnd_old = win32gui.FindWindowEx(0, 0, 0, value['application_name_old'])
    interval = float(value['duration']) * 60
    repetitions = 5
    worker = threading.Thread(target=run, args=(hwnd, hwnd_old, interval, repetitions))
    print('reached')
    worker.start()


window = build_gui()
while True:
    event, value = window
    if event == 'btn_start':
        btn_start()
        break
print("done")
# hwnd = win32gui.FindWindowEx(0,0,0, 'FINAL FANTASY XIV')
# hwnd_old = win32gui.FindWindowEx(0,0,0, 'App')
