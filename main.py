#IDE / SHELL needs to be run as Admin to work!

import win32gui, pyautogui, time
import PySimpleGUI  as pg

interval = 0
gui_done = False
WAIT = 2




def build_gui():
    window = pg.Window('App', layout=[[pg.Text('Insert New Application Name: ', size=(25, 1)), pg.In(size=(25, 1), default_text='Neuer Tab - Google Chrome', key='application_name_new'), pg.Image('Example_Name.PNG'), pg.Text('(Example)')],
                             [pg.Text('Insert Old Application Name: ', size=(25, 1)), pg.In(size=(25, 1), default_text='App', key='application_name_old')],
                             [pg.Text('Insert Duration (sec): ', size=(25, 1)), pg.In(size=(25, 1), default_text='5', key='duration')],
                             [pg.Button(button_text='Start', key='btn_start')]], margins=(200, 200)).read()
    return window

window = build_gui()
while(True):
    event, value = window
    if event == 'btn_start':
        hwnd = win32gui.FindWindowEx(0, 0, 0, value['application_name_new'])
        hwnd_old = win32gui.FindWindowEx(0, 0, 0, value['application_name_old'])
        interval = int(value['duration'])
        gui_done = True
        break

#hwnd = win32gui.FindWindowEx(0,0,0, 'FINAL FANTASY XIV')
#hwnd_old = win32gui.FindWindowEx(0,0,0, 'App')

while (gui_done):
    win32gui.SetForegroundWindow(hwnd)
    time.sleep(WAIT)
    pyautogui.keyDown('a')
    time.sleep(WAIT)
    pyautogui.keyUp('a')
    win32gui.SetForegroundWindow(hwnd_old)
    time.sleep(interval)