import sys
import time
import rtmidi
import ctypes
import win32con
import win32com.client
import win32api
import win32gui

DEVICENAME = u'nanoKONTROL2'
WINDOWTITLE = 'MultiPlay'
DEBUG = True

COMMANDS = (
    ([176,41,0], win32con.VK_SPACE),
    ([176,42,0], win32con.VK_ESCAPE),
    ([176,43,0], win32con.VK_UP),
    ([176,44,0], win32con.VK_DOWN)
)

ERR_DEVICENOTFOUND = 100
ERR_WINDOWNOTFOUND = 101

# Useful Windows Functions
EnumWindows = ctypes.windll.user32.EnumWindows
EnumWindowsProc = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.POINTER(ctypes.c_int), ctypes.POINTER(ctypes.c_int))
GetWindowText = ctypes.windll.user32.GetWindowTextW
GetWindowTextLength = ctypes.windll.user32.GetWindowTextLengthW
IsWindowVisible = ctypes.windll.user32.IsWindowVisible
FindWindow = ctypes.windll.user32.FindWindowA
SendMessage = ctypes.windll.user32.SendMessageA
GetLastError = ctypes.GetLastError

def main():
    # Check application is running
    global shell
    shell = win32com.client.Dispatch("WScript.Shell")
    if findWindow(WINDOWTITLE) == "":
        print "Window with title {} not found, please run the application and try again".format(WINDOWTITLE)
        time.sleep(5)
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
                windows.append(buff.value)
                return False
        return True
    EnumWindows(EnumWindowsProc(foreach_window), 0)
    if len(windows) > 0:
        hwnd = win32gui.FindWindow(None, windows[0])
        if hwnd == 0:
            print "Window not found (hwnd == 0)"
        return hwnd
    else:
        return ""

def process_message(message, time_stamp):
    c = findCommand(message[0])    
    if DEBUG:    
        print "{} received at {}".format(message, time_stamp)
    if c:
        hwnd = findWindow(WINDOWTITLE)
        if hwnd != "":
            if DEBUG:
                print "Window: {}".format(hwnd)
            if isinstance(c, basestring):
                m = ord(c)
            if isinstance(c, int):
                m = c
            else:
                print "Unknown key type"
            SendMessage(hwnd, win32con.WM_KEYDOWN, m, 0)
            SendMessage(hwnd, win32con.WM_CHAR, m, 0)
            SendMessage(hwnd, win32con.WM_KEYUP, m, 0)
        
def findCommand(cmd):
    for key, value in COMMANDS:
        if repr(key) == repr(cmd):
            return value
    return False
 
if __name__ == '__main__':
    main()
