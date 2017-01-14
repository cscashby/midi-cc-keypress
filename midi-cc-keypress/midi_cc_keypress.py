import sys
import time
import rtmidi
import win32com.client
import ctypes

DEVICENAME = u'nanoKONTROL2'
WINDOWTITLE = 'MultiPlay'
DEBUG = False

COMMANDS = (
    ([176,41,0], "{ }"),
    ([176,42,0], "{ESC}"),
    ([176,43,0], "{UP}"),
    ([176,44,0], "{DOWN}")
)

ERR_DEVICENOTFOUND = 100
ERR_WINDOWNOTFOUND = 101

# Useful Windows Functions
EnumWindows = ctypes.windll.user32.EnumWindows
EnumWindowsProc = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.POINTER(ctypes.c_int), ctypes.POINTER(ctypes.c_int))
GetWindowText = ctypes.windll.user32.GetWindowTextW
GetWindowTextLength = ctypes.windll.user32.GetWindowTextLengthW
IsWindowVisible = ctypes.windll.user32.IsWindowVisible
GetForegroundWindow = ctypes.windll.user32.GetForegroundWindow
SetForegroundWindow = ctypes.windll.user32.SetForegroundWindow

def main():
    # Check application is running
    global shell
    shell = win32com.client.Dispatch("WScript.Shell")
    if findWindow(WINDOWTITLE) == "":
        print "Window with title {} not found, please run the application and try again".format(WINDOWTITLE)
        exit(ERR_WINDOWNOTFOUND)

    # Set up MIDI
    midiin = rtmidi.MidiIn()
    available_ports = midiin.get_ports()
    devicePort = -1    

    for idx, port in enumerate(available_ports):
        if DEBUG:
            print "{} at index {}".format(port, idx)
        if port.find(DEVICENAME) != -1:
            devicePort = idx
            break

    if devicePort == -1:
        print "{} device not found".format(DEVICENAME)
        sys.exit(ERR_DEVICENOTFOUND)

    midiin.set_callback(process_message)
    port = midiin.open_port(devicePort)
    
    while True:
        time.sleep(0.01)

    del midiin

def findWindow(title):
    windows = []
    def foreach_window(hwnd, lParam):
        if IsWindowVisible(hwnd):
            length = GetWindowTextLength(hwnd)
            buff = ctypes.create_unicode_buffer(length + 1)
            GetWindowText(hwnd, buff, length + 1)
            if buff.value.find(title) != -1:
                windows.append(hwnd)
                return False
        return True
    EnumWindows(EnumWindowsProc(foreach_window), 0)
    if len(windows) > 0:
        return windows[0]
    else:
        return ""

def getActiveWindowText():
    hwnd = GetForegroundWindow()
    length = GetWindowTextLength(hwnd)
    buff = ctypes.create_unicode_buffer(length + 1)
    GetWindowText(hwnd, buff, length + 1)
    return buff.value

def process_message(message, time_stamp):
    if DEBUG:    
        print "{} received at {}".format(message, time_stamp)
    c = findCommand(message[0])    
    if c:
        oldHWnd = GetForegroundWindow()
        SetForegroundWindow(findWindow(WINDOWTITLE))
        shell.SendKeys(repr(c), 0)
        time.sleep(0.5)
        SetForegroundWindow(oldHWnd)
        
def findCommand(cmd):
    return [value for key, value in COMMANDS if repr(key) == repr(cmd)]
 
if __name__ == '__main__':
    main()
