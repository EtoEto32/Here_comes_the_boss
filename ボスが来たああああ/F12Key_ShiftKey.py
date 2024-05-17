import os
import keyboard
import pyautogui
import pygetwindow as gw
def hide_windows():
    # すべてのウィンドウを最小化
    pyautogui.hotkey('win', 'd')
    # Excelが既に開いているかどうかを確認
    try:
        excel = gw.getWindowsWithTitle('Excel')[0]  # Excelのウィンドウを探す
        excel.maximize()  # Excelのウィンドウを最大化する
    except IndexError:  # Excelのウィンドウが見つからない場合
        os.system('start excel.exe')  # 新たにExcelを開く
keyboard.add_hotkey("Esc", hide_windows)

keyboard.wait()
