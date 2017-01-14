import sys
import time
import rtmidi

DEVICENAME = u'nanoKONTROL2'
DEBUG = False

COMMANDS = (
    ([176,41,0], 'Play'),
    ([176,42,0], 'Stop')
)

ERR_DEVICENOTFOUND = 100

def main():
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

def process_message(message, time_stamp):
    if DEBUG:    
        print "{} received at {}".format(message, time_stamp)
    c = findCommand(message[0])    
    if c:
        print c

def findCommand(cmd):
    return [value for key, value in COMMANDS if repr(key) == repr(cmd)]
 
if __name__ == '__main__':
    main()
