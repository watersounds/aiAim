import win32api
import win32gui
import win32con
import win32print
import time


def getWindow():
    # 1 获取句柄
    # 1.1 通过坐标获取窗口句柄,鼠标所在窗口
    handle = win32gui.WindowFromPoint(win32api.GetCursorPos())  # (259, 185)
    # 1.2 获取最前窗口句柄
    # handle = win32gui.GetForegroundWindow()
    # 1.3 通过类名或查标题找窗口
    # handle = win32gui.FindWindow('cls_name', "title")
    # 1.4 找子窗体
    # sub_handle = win32gui.FindWindowEx(handle, None, 'Edit', None)  # 子窗口类名叫“Edit”


    # 句柄操作
    title = win32gui.GetWindowText(handle)
    cls_name = win32gui.GetClassName(handle)
    print({'类名': cls_name, '标题': title})
    # 获取窗口位置
    info = win32gui.GetWindowRect(handle)
    print(f"info:{info}")
    # 设置为最前窗口
    # win32gui.SetForegroundWindow(handle)

'''
# 2.按键-看键盘码
# 获取鼠标当前位置的坐标
cursor_pos = win32api.GetCursorPos()
# 将鼠标移动到坐标处
win32api.SetCursorPos((200, 200))
# 回车
win32api.keybd_event(13, 0, win32con.KEYEVENTF_EXTENDEDKEY, 0)
win32api.keybd_event(13, 0, win32con.KEYEVENTF_KEYUP, 0)
# 左单键击
win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP | win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
# 右键单击
win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTUP | win32con.MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0)
# 鼠标左键按下-放开
win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
# 鼠标右键按下-放开
win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
# TAB键
win32api.keybd_event(win32con.VK_TAB, 0, 0, 0)
win32api.keybd_event(win32con.VK_TAB, 0, win32con.KEYEVENTF_KEYUP, 0)
# 快捷键Alt+F
win32api.keybd_event(18, 0, 0, 0)  # Alt
win32api.keybd_event(70, 0, 0, 0)  # F
win32api.keybd_event(70, 0, win32con.KEYEVENTF_KEYUP, 0)  # 释放按键
win32api.keybd_event(18, 0, win32con.KEYEVENTF_KEYUP, 0)

# 3.Message
win = win32gui.FindWindow('Notepad', None)
tid = win32gui.FindWindowEx(win, None, 'Edit', None)
# 输入文本
win32gui.SendMessage(tid, win32con.WM_SETTEXT, None, '你好hello word!')
# 确定
win32gui.SendMessage(handle, win32con.WM_COMMAND, 1, btnhld)
# 关闭窗口
win32gui.PostMessage(win32gui.FindWindow('Notepad', None), win32con.WM_CLOSE, 0, 0)
# 回车
win32gui.PostMessage(tid, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)  # 插入一个回车符,post 没有返回值，执行完马上返回
win32gui.PostMessage(tid, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
print("%x" % win)
print("%x" % tid)
# 选项框
res = win32api.MessageBox(None, "Hello Pywin32", "pywin32", win32con.MB_YESNOCANCEL)
print(res)
win32api.MessageBox(0, "Hello PYwin32", "MessageBox", win32con.MB_OK | win32con.MB_ICONWARNING)  # 加警告标志

# 4.截屏
from PIL import ImageGrab

# 利用PIL截屏
im = ImageGrab.grab()
im.save('aa.jpg')

# 5.文件的读写
import win32file, win32api, win32con
import os


def SimpleFileDemo():
    testName = os.path.join(win32api.GetTempPath(), "opt_win_file.txt")

    if os.path.exists(testName):
        os.unlink(testName)  # os.unlink() 方法用于删除文件,如果文件是一个目录则返回一个错误。
    # 写
    handle = win32file.CreateFile(testName,
                                  win32file.GENERIC_WRITE,
                                  0,
                                  None,
                                  win32con.CREATE_NEW,
                                  0,
                                  None)
    test_data = "Hello0there".encode("ascii")
    win32file.WriteFile(handle, test_data)
    handle.Close()
    # 读
    handle = win32file.CreateFile(testName, win32file.GENERIC_READ, 0, None, win32con.OPEN_EXISTING, 0, None)
    rc, data = win32file.ReadFile(handle, 1024)
    handle.Close()  # 此处也可使用win32file.CloseHandle(handle)来关闭句柄
    if data == test_data:
        print("Successfully wrote and read a file")
    else:
        raise Exception("Got different data back???")
    os.unlink(testName)


if __name__ == '__main__':
    SimpleFileDemo()
    # print(win32api.GetTempPath())  # 获取临时文件夹路径

# 6. ShellExecute
win32api.ShellExecute(None, "open", "C:Test.txt", None, None, SW_SHOWNORMAL)  # 打开C:Test.txt 文件
win32api.ShellExecute(None, "open", "http:#www.google.com", None, None, SW_SHOWNORMAL)  # 打开网页www.google.com
win32api.ShellExecute(None, "explore", "D:C++", None, None, SW_SHOWNORMAL)  # 打开目录D:C++
win32api.ShellExecute(None, "print", "C:Test.txt", None, None, SW_HIDE)  # 打印文件C:Test.txt
win32api.ShellExecute(None, "open", "mailto:", None, None, SW_SHOWNORMAL)  # 打开邮箱
win32api.ShellExecute(None, "open", "calc.exe", None, None, SW_SHOWNORMAL)  # 调用计算器
win32api.ShellExecute(None, "open", "NOTEPAD.EXE", None, None, SW_SHOWNORMAL)  # 调用记事本
#  ShellExecute 不支持定向输出

# 打印--只能打印txt word excel 不能打印图片及pdf
from win32con import SW_HIDE, SW_SHOWNORMAL

for fn in ['2.txt', '3.txt']:
    print(fn)
    res = win32api.ShellExecute(0,  # 指定父窗口句柄

                                'print',  # 指定动作, 譬如: open、print、edit、explore、find

                                fn,  # 指定要打开的文件或程序

                                win32print.GetDefaultPrinter(),
                                # 给要打开的程序指定参数;GetDefaultPrinter　　取得默认打印机名称 <type 'str'>,GetDefaultPrinterW　　取得默认打印机名称 <type 'unicode'>

                                "./downloads/",  # 目录路径

                                SW_SHOWNORMAL)  # 打开选项,SW_HIDE = 0; {隐藏},SW_SHOWNORMAL = 1; {用最近的大小和位置显示, 激活}
    print(res)  # 返回值大于32表示执行成功,返回值小于32表示执行错误


# 打印 -pdf
def print_pdf(pdf_file_name):
    """
    静默打印pdf
    :param pdf_file_name:
    :return:
    """
    # GSPRINT_PATH = resource_path + 'GSPRINTgsprint'
    GHOSTSCRIPT_PATH = resource_path + 'GHOSTSCRIPTbingswin32c'  # gswin32c.exe
    currentprinter = config.printerName  # "printerName":"FUJI XEROX ApeosPort-VI C3370""
    currentprinter = win32print.GetDefaultPrinter()
    arg = '-dPrinted '
    '-dBATCH '
    '-dNOPAUSE '
    '-dNOSAFER '
    '-dFitPage '
    '-dNORANGEPAGESIZE '
    '-q '
    '-dNumCopies=1 '
    '-sDEVICE=mswinpr2 '
    '-sOutputFile="spool'
    + currentprinter + " " +
    pdf_file_name


log.info(arg)
win32api.ShellExecute(
    0,
    'open',
    GHOSTSCRIPT_PATH,
    arg,
    ".",
    0
)
# os.remove(pdf_file_name)
'''

if __name__ == '__main__':
    for i in range(10):
        getWindow()
        time.sleep(1)
