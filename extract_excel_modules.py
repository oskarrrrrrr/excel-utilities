import tkinter
from tkinter.filedialog import askdirectory
import os
import datetime
from excel_project_password import *


msoAutomationSecurityLow = 1
msoAutomationSecurityByUI = 2
msoAutomationSecurityForceDisable = 3


password = "SomeRandomPassword"


class VBModType:
    vbStdModule = 1
    vbClass = 2
    vbForm = 3
    vbDocument = 100


vb_mod_type_exp = {
    VBModType.vbStdModule: 'bas',
    VBModType.vbClass: 'cls',
    VBModType.vbForm: 'frm',
    VBModType.vbDocument: 'bas',
    }

vb_mode_dir_name = {
    VBModType.vbStdModule: 'Modules',
    VBModType.vbClass: 'Classes',
    VBModType.vbForm: 'Forms',
    VBModType.vbDocument: 'Documents',
}


def get_folder_path():
    return tkinter.filedialog.askdirectory()


def get_excel_file_path():
    return tkinter.filedialog.askopenfilename()


def export_vba():
    excel, wk = open_excel(password)
    export_file_full_path = wk.FullName

    export_dir = get_folder_path()
    if export_dir == '': return

    now = datetime.datetime.now()
    export_dir = export_dir + '/' + os.path.splitext(os.path.basename(export_file_full_path))[0] + '_' + now.strftime("%Y_%m_%d_%H_%M")
    os.makedirs(export_dir)

    for VBComp in wk.VBProject.VBComponents:
        if VBComp.Type in vb_mod_type_exp:

            cur_mod_export_dir = export_dir + '/' + vb_mode_dir_name[VBComp.Type]
            if not os.path.exists(cur_mod_export_dir):
                os.makedirs(cur_mod_export_dir)

            VBComp.Export(cur_mod_export_dir + '/' + VBComp.Name + '.' + vb_mod_type_exp[VBComp.Type])

    wk.Close(False)
    if excel.Workbooks.Count == 0:
        excel.Quit()


root = tkinter.Tk()
root.withdraw()

export_vba()
