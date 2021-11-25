import cv2.cv2 as cv2
import numpy as np
import win32gui
import win32ui
import win32con
import win32api


def grab_screen(region=None):
    """
    截取当前屏幕
    :param region: 截取范围（x1, y1, x2, y2）
    :return:一个数据类型为uint8的ndarray
    """
    # GetDesktopWindow()返回Windows桌面窗口的句柄。桌面窗口覆盖整个屏幕。桌面窗口是其上所有图标和其他窗口的区域。
    hwin = win32gui.GetDesktopWindow()
    print(f"hwin:{hwin}")

    if region:
        x1, y1, x2, y2 = region
        width = x2 - x1 + 1
        height = y2 - y1 + 1
    else:
        width = win32api.GetSystemMetrics(win32con.SM_CXVIRTUALSCREEN)
        height = win32api.GetSystemMetrics(win32con.SM_CYVIRTUALSCREEN)
        x1 = win32api.GetSystemMetrics(win32con.SM_XVIRTUALSCREEN)
        y1 = win32api.GetSystemMetrics(win32con.SM_YVIRTUALSCREEN)
    print(f"width:{width}, height:{height}, x1:{x1}, y1:{y1}")

    hwindc = win32gui.GetWindowDC(hwin)
    srcdc = win32ui.CreateDCFromHandle(hwindc)
    memdc = srcdc.CreateCompatibleDC()
    bmp = win32ui.CreateBitmap()
    bmp.CreateCompatibleBitmap(srcdc, width, height)
    memdc.SelectObject(bmp)
    memdc.BitBlt((0, 0), (width, height), srcdc, (x1, y1), win32con.SRCCOPY)

    signedIntsArray = bmp.GetBitmapBits(True)
    img = np.fromstring(signedIntsArray, dtype='uint8')
    img.shape = (height, width, 4)

    srcdc.DeleteDC()
    memdc.DeleteDC()
    win32gui.ReleaseDC(hwin, hwindc)
    win32gui.DeleteObject(bmp.GetHandle())

    return img


if __name__ == '__main__':
    region = (0, 0, 1920, 1080)
    image_array = grab_screen()
    # 获取屏幕，(0, 0, 1280, 720)表示从屏幕坐标（0,0）即左上角，截取往右1280和往下720的画面
    # array_to_image = Image.fromarray(image_array, mode='RGB')  # 将array转成图像，才能送入yolo进行预测
    cv2.imshow('window', image_array)
    if cv2.waitKey(0) & 0xFF == ord('q'):  # 按q退出，记得输入切成英语再按q
        cv2.destroyAllWindows()

