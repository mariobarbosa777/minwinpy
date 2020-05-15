import win32gui
import win32api
import win32con
import win32ui
import time

# VK_CODE is a dictionary to simplify some Virtual Key Codes, more codes and information in 
#https://docs.microsoft.com/en-us/windows/win32/inputdev/virtual-key-codes 

VK_CODE = {'backspace':0x08,
           'tab':0x09,
           'clear':0x0C,
           'enter':0x0D,
           'shift':0x10,
           'ctrl':0x11,
           'alt':0x12,
           'pause':0x13,
           'caps_lock':0x14,
           'esc':0x1B,
           'spacebar':0x20,
           'page_up':0x21,
           'page_down':0x22,
           'end':0x23,
           'home':0x24,
           'left_arrow':0x25,
           'up_arrow':0x26,
           'right_arrow':0x27,
           'down_arrow':0x28,
           'select':0x29,
           'print':0x2A,
           'execute':0x2B,
           'print_screen':0x2C,
           'ins':0x2D,
           'del':0x2E,
           'help':0x2F,
           '0':0x30,
           '1':0x31,
           '2':0x32,
           '3':0x33,
           '4':0x34,
           '5':0x35,
           '6':0x36,
           '7':0x37,
           '8':0x38,
           '9':0x39,
           'a':0x41,
           'b':0x42,
           'c':0x43,
           'd':0x44,
           'e':0x45,
           'f':0x46,
           'g':0x47,
           'h':0x48,
           'i':0x49,
           'j':0x4A,
           'k':0x4B,
           'l':0x4C,
           'm':0x4D,
           'n':0x4E,
           'o':0x4F,
           'p':0x50,
           'q':0x51,
           'r':0x52,
           's':0x53,
           't':0x54,
           'u':0x55,
           'v':0x56,
           'w':0x57,
           'x':0x58,
           'y':0x59,
           'z':0x5A,
           'numpad_0':0x60,
           'numpad_1':0x61,
           'numpad_2':0x62,
           'numpad_3':0x63,
           'numpad_4':0x64,
           'numpad_5':0x65,
           'numpad_6':0x66,
           'numpad_7':0x67,
           'numpad_8':0x68,
           'numpad_9':0x69,
           'multiply_key':0x6A,
           'add_key':0x6B,
           'separator_key':0x6C,
           'subtract_key':0x6D,
           'decimal_key':0x6E,
           'divide_key':0x6F,
           'F1':0x70,
           'F2':0x71,
           'F3':0x72,
           'F4':0x73,
           'F5':0x74,
           'F6':0x75,
           'F7':0x76,
           'F8':0x77,
           'F9':0x78,
           'F10':0x79,
           'F11':0x7A,
           'F12':0x7B,
           'F13':0x7C,
           'F14':0x7D,
           'F15':0x7E,
           'F16':0x7F,
           'F17':0x80,
           'F18':0x81,
           'F19':0x82,
           'F20':0x83,
           'F21':0x84,
           'F22':0x85,
           'F23':0x86,
           'F24':0x87,
           'num_lock':0x90,
           'scroll_lock':0x91,
           'left_shift':0xA0,
           'right_shift ':0xA1,
           'left_control':0xA2,
           'right_control':0xA3,
           'left_menu':0xA4,
           'right_menu':0xA5,
           'browser_back':0xA6,
           'browser_forward':0xA7,
           'browser_refresh':0xA8,
           'browser_stop':0xA9,
           'browser_search':0xAA,
           'browser_favorites':0xAB,
           'browser_start_and_home':0xAC,
           'volume_mute':0xAD,
           'volume_Down':0xAE,
           'volume_up':0xAF,
           'next_track':0xB0,
           'previous_track':0xB1,
           'stop_media':0xB2,
           'play/pause_media':0xB3,
           'start_mail':0xB4,
           'select_media':0xB5,
           'start_application_1':0xB6,
           'start_application_2':0xB7,
           'attn_key':0xF6,
           'crsel_key':0xF7,
           'exsel_key':0xF8,
           'play_key':0xFA,
           'zoom_key':0xFB,
           'clear_key':0xFE,
           '+':0xBB,
           ',':0xBC,
           '-':0xBD,
           '.':0xBE,
           '/':0xBF,
           '`':0xC0,
           ';':0xBA,
           '[':0xDB,
           '\\':0xDC,
           ']':0xDD,
           "'":0xDE,}



class minwinpy:

    def __init__(self, WindowsName):
        self._originalWindonsName= WindowsName
        self._WindonsName= WindowsName
        self._hwnd = win32gui.FindWindow(None,WindowsName)
        if self._hwnd == 0:
            raise ValueError("WindowsNanme Not Found")

    def GetHWND(self):
        """
        Return HWND of the WindowsName 
        """
        return self._hwnd

    def PressKey(self, key, pressTime=0.1):
        """
            press key for a time in seconds 
            (KeyDOWN, wait a time ,KeyUP),
            key is a string from VK_CODE
            Lettes in lower case for example "a", "b" , "c"
            F keys in upper case for example "F1", "F2", "F3"
        """
        win32api.SendMessage(self._hwnd, win32con.WM_KEYDOWN,VK_CODE[key], 0)
        time.sleep(pressTime)
        win32api.SendMessage(self._hwnd, win32con.WM_KEYUP,VK_CODE[key], 0)
        return None

    def LeftClick(self, pos, pressTime=0.1):
        """
            Left click in  pos[1],pos[2] a seconds
            (LButtonDowm, wait a time ,LButtonUp),
        """
        tmp=win32api.MAKELONG(pos[1],pos[2])
        win32gui.SendMessage(self._hwnd, win32con.WM_LBUTTONDOWN,win32con.MK_LBUTTON,tmp)
        time.sleep(pressTime)
        win32gui.SendMessage(self._hwnd, win32con.WM_LBUTTONUP,win32con.MK_LBUTTON,tmp)
    
    def RightClick(self, x, y, pressTime=0.1):
        """
            WARNING NOT WORKING 
            Left click for a  
            (LButtonDowm, wait a time ,LButtonUp),
        """
        tmp=win32api.MAKELONG(x,y)
        win32gui.SendMessage(self._hwnd, win32con.WM_LBUTTONDOWN,win32con.MK_LBUTTON,tmp)
        time.sleep(pressTime)
        win32gui.SendMessage(self._hwnd, win32con.WM_LBUTTONUP,win32con.MK_LBUTTON,tmp)


    def FullScreenshot(self, filename):
        """
            Takes a Full Screenshot of a WindonsName and save this in filename.png
        """
        if not  filename.endswith(".png"):
            filename=filename+".png"


        rect = win32gui.GetWindowRect(self._hwnd)
        w = abs(rect[2] - rect[0])
        h = abs(rect[3] - rect[1])
        x, y = 0, 0
        hwndDC = win32gui.GetWindowDC(self._hwnd)
        mfcDC = win32ui.CreateDCFromHandle(hwndDC)
        saveDC = mfcDC.CreateCompatibleDC()
        saveBitMap = win32ui.CreateBitmap()
        saveBitMap.CreateCompatibleBitmap(mfcDC, w, h)
        saveDC.SelectObject(saveBitMap)
        saveDC.BitBlt((0, 0), (w, h), mfcDC, (x, y), win32con.SRCCOPY)
        saveBitMap.SaveBitmapFile(saveDC, filename)
        mfcDC.DeleteDC()
        saveDC.DeleteDC()
        win32gui.ReleaseDC(self._hwnd, hwndDC)
        win32gui.DeleteObject(saveBitMap.GetHandle())


    def Screenshot(self,filename,start,size):
        """
            Takes a partcial Screenshot of a WindonsName and save this in filename.png
            
            starts= [x,y]: Initial pixel with x and y coordinates
            size = [w, h]: width, and height of image 
        """
        if not  filename.endswith(".png"):
            filename=filename+".png"

        x = start[0]
        y = start[1]
        w = size[0]
        h = size[1]

        hwndDC = win32gui.GetWindowDC(self._hwnd)
        mfcDC = win32ui.CreateDCFromHandle(hwndDC)
        saveDC = mfcDC.CreateCompatibleDC()
        saveBitMap = win32ui.CreateBitmap()
        saveBitMap.CreateCompatibleBitmap(mfcDC, w, h)
        saveDC.SelectObject(saveBitMap)
        saveDC.BitBlt((0, 0), (w, h), mfcDC, (x, y), win32con.SRCCOPY)
        saveBitMap.SaveBitmapFile(saveDC, filename)
        mfcDC.DeleteDC()
        saveDC.DeleteDC()
        win32gui.ReleaseDC(self._hwnd, hwndDC)
        win32gui.DeleteObject(saveBitMap.GetHandle())

    
    def SetWindowsPos(self,start):
        """
            Change initial position  x,y  of Windonsname without changes its size 
            x = start[0]
            y = start[1]
            
        """

        rect = win32gui.GetWindowRect(self._hwnd)
        x = start[0]
        y = start[1]
        w = rect[2] - x
        h = rect[3] - y
        win32gui.SetWindowPos(self._hwnd, win32con.HWND_NOTOPMOST, 0, 0, w, h, 0)


    def SetWindowsName(self,newName):
        """
            Change windows name 
            
        """
        win32gui.SetWindowText(self._hwnd,newName) 
        _WindonsName=newName

    def GetWindowsName(self):
        return self._WindonsName


    def GetPixelRGBColor(self,pos):

        rect = win32gui.GetWindowRect(self._hwnd)
        w = abs(rect[2] - rect[0])
        h = abs(rect[3] - rect[1])

        hwndDC = win32gui.GetWindowDC(self._hwnd)
        mfcDC = win32ui.CreateDCFromHandle(hwndDC)
        saveDC = mfcDC.CreateCompatibleDC()
        saveBitMap = win32ui.CreateBitmap()
        saveBitMap.CreateCompatibleBitmap(mfcDC, w, h)
        saveDC.SelectObject(saveBitMap)


        ret=win32gui.GetPixel(hwndDC,pos[0],pos[1])
        r, g, b = ret & 0xff, (ret >> 8) & 0xff, (ret >> 16) & 0xff
        
        mfcDC.DeleteDC()
        saveDC.DeleteDC()
        win32gui.ReleaseDC(self._hwnd,hwndDC)
        win32gui.DeleteObject(saveBitMap.GetHandle())

        return [r, g, b]
    

        



if __name__ == "__main__" :

    TantraGame = minwinpy("Mateito")
    namePJ="Nexo Game"

    TantraGame.SetWindowsName(namePJ)
    print(TantraGame.GetWindowsName())








    while 1:
        TantraGame.PressKey("e", pressTime=0.5)
        time.sleep(1)


        

