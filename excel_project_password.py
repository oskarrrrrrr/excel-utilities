from win32com.client import Dispatch
import tkinter
from tkinter.filedialog import askdirectory
import win32gui


WM_SETTEXT = 12
BM_CLICK = 245

ok_hwnd = None
verbose = False


def winfun(hwnd, lparam):
    s = win32gui.GetWindowText(hwnd)
    if len(s) > 0 and verbose: print("winfun | child_hwnd: %d   txt: %s" % (hwnd, s))
    if s == 'OK':
        global ok_hwnd
        ok_hwnd = hwnd
    return 1


def enter_excel_password(password):
    vba_project_pass_hwnd = win32gui.FindWindow(None, 'VBAProject Password')

    if vba_project_pass_hwnd == 0:
        if verbose:
            print('No password window found!')
        return

    if verbose:
        print(vba_project_pass_hwnd)

    edit_hwnd = win32gui.FindWindowEx(vba_project_pass_hwnd, 0, 'Edit', '')
    if verbose:
        print(edit_hwnd)
    win32gui.SendMessage(edit_hwnd, WM_SETTEXT, False, password)

    win32gui.EnumChildWindows(vba_project_pass_hwnd, winfun, None)
    win32gui.SendMessage(ok_hwnd, BM_CLICK, 0, '')


def kill_project_properties_window():
    # luckily OK buttons inside VBA Project pop-up windows seem to have the same ID
    # so we can reuse the previously found one
    win32gui.SendMessage(ok_hwnd, BM_CLICK, 0, '')


def get_excel_file_path():
    return tkinter.filedialog.askopenfilename()


def get_user_selected_excel():
    export_file_full_path = get_excel_file_path()
    excel = Dispatch('Excel.Application')
    excel.Visible = False
    wk = excel.Workbooks.Open(export_file_full_path, False, True, None, '')
    return excel, wk


def open_excel(password):
    excel, wk = get_user_selected_excel()
    excel.Application.CommandBars.Item(26).Controls(4).Execute()
    excel.VBE.CommandBars(1).FindControl(1, 2578, "", True, True).Execute()
    enter_excel_password(password)
    return excel, wk
