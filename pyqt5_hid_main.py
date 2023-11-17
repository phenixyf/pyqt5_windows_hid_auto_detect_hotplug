import sys
import ctypes
from ctypes.wintypes import MSG
import ctypes.wintypes as wintypes

import hid

from PyQt5.QtWidgets import *
from PyQt5.QtCore import *

NULL = 0
INVALID_HANDLE_VALUE = -1
DEVICE_NOTIFY_WINDOW_HANDLE = 0x00000000
WM_DEVICECHANGE = 0x0219            # windows 系统设备变动事件序号
DBT_DEVTYP_DEVICEINTERFACE = 5
DBT_DEVICEREMOVECOMPLETE = 0x8004   # windows 系统设备移出信息序号
DBT_DEVICEARRIVAL = 0x8000          # windows 系统设备插入信息序号


user32 = ctypes.windll.user32
RegisterDeviceNotification = user32.RegisterDeviceNotificationW
UnregisterDeviceNotification = user32.UnregisterDeviceNotification


class GUID(ctypes.Structure):
    _pack_ = 1
    _fields_ = [("Data1", ctypes.c_ulong),
                ("Data2", ctypes.c_ushort),
                ("Data3", ctypes.c_ushort),
                ("Data4", ctypes.c_ubyte * 8)]


class DEV_BROADCAST_DEVICEINTERFACE(ctypes.Structure):
    _pack_ = 1
    _fields_ = [("dbcc_size", wintypes.DWORD),
                ("dbcc_devicetype", wintypes.DWORD),
                ("dbcc_reserved", wintypes.DWORD),
                ("dbcc_classguid", GUID),
                ("dbcc_name", ctypes.c_wchar * 260)]


class DEV_BROADCAST_HDR(ctypes.Structure):
    _fields_ = [("dbch_size", wintypes.DWORD),
                ("dbch_devicetype", wintypes.DWORD),
                ("dbch_reserved", wintypes.DWORD)]


# GUID_DEVCLASS_PORTS = GUID(0x4D36E978, 0xE325, 0x11CE,
#                            (ctypes.c_ubyte * 8)(0xBF, 0xC1, 0x08, 0x00, 0x2B, 0xE1, 0x03, 0x18))
GUID_DEVINTERFACE_USB_DEVICE = GUID(0xA5DCBF10, 0x6530, 0x11D2,
                                    (ctypes.c_ubyte * 8)(0x90, 0x1F, 0x00, 0xC0, 0x4F, 0xB9, 0x51, 0xED))


target_pid = 0xfe07  # 用你的目标PID替换这里
target_vid = 0x1a86  # 用你的目标VID替换这里


class Window(QWidget):
    hidBdg = hid.device()   # add hid device object
    hidStatus = False       # False - hid open failed
                            # True - hid open successful

    def __init__(self, parent=None):
        super(Window, self).__init__(parent)

        self.setupNotification()            # 调用注册 HID 设备热插拔事件函数
        self.initUI()

    def initUI(self):
        self.resize(QSize(600, 320))
        self.setWindowTitle("Device Notify")
        vbox = QVBoxLayout(self)
        vbox.addWidget(QLabel("Log window:", self))
        self.logEdit = QPlainTextEdit(self)
        vbox.addWidget(self.logEdit)
        self.setLayout(vbox)
        self.open_hid()         # 程序启动即尝试打开 HID

    def setupNotification(self):
        dbh = DEV_BROADCAST_DEVICEINTERFACE()
        dbh.dbcc_size = ctypes.sizeof(DEV_BROADCAST_DEVICEINTERFACE)
        dbh.dbcc_devicetype = DBT_DEVTYP_DEVICEINTERFACE
        dbh.dbcc_classguid = GUID_DEVINTERFACE_USB_DEVICE  # GUID_DEVCLASS_PORTS
        self.hNofity = RegisterDeviceNotification(int(self.winId()),
                                                  ctypes.byref(dbh),
                                                  DEVICE_NOTIFY_WINDOW_HANDLE)
        if self.hNofity == NULL:
            print(ctypes.FormatError(), int(self.winId()))
            print("RegisterDeviceNotification failed")
        else:
            print("register successfully")


    def nativeEvent(self, eventType, msg):
        message = MSG.from_address(msg.__int__())       # 获取系统传输消息参数的信息
        if message.message == WM_DEVICECHANGE:          # 本例只处理设备热插拔，所以只添加 WM_DEVICECHANGE 事件的处理代码
            self.onDeviceChanged(message.wParam, message.lParam)    # 调用自己写的设备热插拔处理函数
        return False, 0


    def onDeviceChanged(self, wParam, lParam):
        if DBT_DEVICEARRIVAL == wParam:
            dev_info = ctypes.cast(lParam, ctypes.POINTER(DEV_BROADCAST_DEVICEINTERFACE)).contents
            device_path = ctypes.c_wchar_p(dev_info.dbcc_name).value
            cycCnt = 0
            if f"VID_{target_vid:04X}&PID_{target_pid:04X}" in device_path:
                while (self.open_hid() is not True) and (cycCnt < 5):
                    self.open_hid()
                    cycCnt += 1
                    print(f'Target USB device inserted')

        elif DBT_DEVICEREMOVECOMPLETE == wParam:
            dev_info = ctypes.cast(lParam, ctypes.POINTER(DEV_BROADCAST_DEVICEINTERFACE)).contents
            device_path = ctypes.c_wchar_p(dev_info.dbcc_name).value
            if f"VID_{target_vid:04X}&PID_{target_pid:04X}" in device_path:
                self.close_hid()
                print(f'Target USB device removed')


    def open_hid(self):
        try:
            if self.hidStatus == False:
                self.hidBdg.open(0x1A86, 0xFE07)  # VendorID/ProductID
                self.hidBdg.set_nonblocking(1)  # hid device enable non-blocking modepp
                self.hidStatus = True
                self.logEdit.appendHtml("<font color=blue>Device Arrival: connected</font>")
                return self.hidStatus
            else:
                return self.hidStatus
        except:
            self.logEdit.appendHtml("<font color=red>Open HID failed:</font>")
            self.hidStatus = False
            return self.hidStatus


    def close_hid(self):
        try:
            if self.hidStatus == True:
                self.hidBdg.close()
                self.hidStatus = False
                self.logEdit.appendHtml("<font color=red>Device Removed: disconnected</font>")
                return self.hidStatus
            else:
                return self.hidStatus
        except:
            self.logEdit.appendHtml("<font color=red>Close HID failed:</font>")
            self.hidStatus = True



if __name__ == '__main__':
    app = QApplication(sys.argv)
    w = Window()
    w.show()
    sys.exit(app.exec_())
