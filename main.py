#!/usr/bin/env python3
""" Type contents of the clipboard 
to wherever the cursor is.
"""
import os, sys
import win32clipboard
from tqdm import tqdm
from time import sleep
import PySimpleGUI as sg
from pynput.keyboard import Key, Controller

# 
quiet = True

def del_prev_line():
    """ Deletes previous line and 
    moves cursor one up. """
    if not quiet:
        sys.stdout.write('\x1b[1A')
        sys.stdout.write('\x1b[2K')


def clipboard_to_keystrokes(progress_bar, textWaiting):
    """ Reads the last value in clipboard and enters them as Keyboard Keystrokes
    without using the Copy-Paste
    Args:
        progress_bar (PySimpleGUI.PySimpleGUI.ProgressBar): ProgressBar Object from PySimpleGUI
        textWaiting (PySimpleGUI.PySimpleGUI.Text): Text Object from PySimpleGUI
    """
    os.system("cls")
    keyboard = Controller()

    # get clipboard data
    win32clipboard.OpenClipboard()
    data = win32clipboard.GetClipboardData()
    win32clipboard.CloseClipboard()
    
    if not quiet: 
        print(" Text from the clipboard to be Typed\n" + '='*37)
        print(data,"\n\nCursor to desired textbox now.")

    textWaiting.update("STATUS: HOLD")
    for i in range(6):
        if not quiet: print("Waiting for", (5-i), "Seconds...")
        sleep(1)
        del_prev_line()
    textWaiting.update(("STATUS: ONGOING"))
    del_prev_line()
    # Starts typing $data
    for i in range(0, len(data)):
        keyboard.press(data[i])
        keyboard.release(data[i])
        perc = ((i+1)/len(data))*100
        #print(perc, data[i])
        progress_bar.UpdateBar(perc)
        sleep(0.035) # Typing interval
    print()



def main_app():
    """ Initializes and handles PySimpleGUI Frames and Windows
    """
    sg.SetOptions(element_padding=(0, 0))
    progressbar = [[sg.ProgressBar(100, orientation='h', size=(31, 10), key='progressbar')]]
    textWaiting = [[sg.Text('STATUS: NONE', font=('Helvetica', 10), size=(20, 1),
                    justification='center', key='textWaiting')]]
    layout = [
        [sg.Button("Click here to Type the Clipboard"), sg.Text(' | '), sg.Frame('', layout=textWaiting)],
        [sg.Frame('Progress', layout=progressbar), sg.Button('EXIT', size=(3, 1), font=('Helvetica', 8), button_color=('white', 'firebrick3'))]]
    # location=[960, 1004],
    window = sg.Window(
        title="Type Clipboard", 
        layout=layout, 
        margins=(25,10),
        no_titlebar=True, 
        keep_on_top=True,
        grab_anywhere=True,
        finalize=True
    )
    progress_bar = window['progressbar']
    textWaiting = window['textWaiting']
    # Create an event loop
    while True:
        event, values = window.read()
        if event == "Click here to Type the Clipboard":
            textWaiting.update("STATUS: HOLD")
            clipboard_to_keystrokes(progress_bar, textWaiting)
            textWaiting.update("STATUS: OK")
        elif event == "EXIT" or event == sg.WIN_CLOSED:
            break 
    window.close()


# Driver Code
if __name__ == "__main__":
    main_app()
