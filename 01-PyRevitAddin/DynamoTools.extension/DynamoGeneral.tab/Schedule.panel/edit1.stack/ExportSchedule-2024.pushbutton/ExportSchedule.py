import clr
clr.AddReference('ProtoGeometry')
clr.AddReference('RevitAPIUI')
clr.AddReference('RevitServices')
clr.AddReference('System.Windows.Forms')

import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager
doc = DocumentManager.Instance.CurrentDBDocument
view = doc.ActiveView
uiapp = DocumentManager.Instance.CurrentUIApplication
uidoc = uiapp.ActiveUIDocument
app = uiapp.Application
# sdkNumber = int(app.VersionNumber)

from Autodesk.DesignScript.Geometry import *
from Autodesk.Revit.UI import *
from Autodesk.Revit.UI.Selection import *
from System.Collections.Generic import *

clr.AddReferenceByName('Microsoft.Office.Interop.Excel, Version=11.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c')
from Microsoft.Office.Interop import Excel


clr.AddReferenceToFileAndPath("C:\\Program Files\\Autodesk\\Revit 2021\\AddIns\\DynamoForRevit\\IronPython.Wpf.dll") 
import wpf 
from System.Windows.Markup import XamlReader, XamlWriter
from System.Windows import Window, Application, MessageBox
import System.Windows.Media.Imaging
from System.Windows.Media.Imaging import BitmapImage
from System import Uri
from System.Windows.Controls import Button, ListView, GridView, TextBox
from System.Windows.Forms import OpenFileDialog,DialogResult
from System.Windows.Controls import CheckBox




allSchedule = IN[0]
schedule_Name = []

# Kiểm tra xem allSchedule có phải là danh sách hay không
if isinstance(allSchedule, list):
    for x in allSchedule:
        # Kiểm tra an toàn trước khi truy cập thuộc tính Name
        if hasattr(x, 'Name'):
            schedule_Name.append(x.Name)
        elif isinstance(x, str):
            schedule_Name.append(x)
        else:
            # Nếu không có thuộc tính Name và không phải string, thử chuyển đổi thành string
            schedule_Name.append(str(x))
else:
    # Nếu allSchedule không phải là danh sách, xử lý như một item đơn lẻ
    if hasattr(allSchedule, 'Name'):
        schedule_Name.append(allSchedule.Name)
    elif isinstance(allSchedule, str):
        # Kiểm tra xem có phải đường dẫn file không
        if allSchedule.endswith('.txt') or '\\' in allSchedule or '/' in allSchedule:
            # Nếu là đường dẫn file, tạo danh sách rỗng thay vì thêm đường dẫn
            schedule_Name = []
        else:
            schedule_Name.append(allSchedule)
    else:
        schedule_Name.append(str(allSchedule))
filePathLink = []
# sheetExcel = []
ScheduleExport = []
FilePath = "none"


class MyWindow(Window): 
    def __init__(self):
        self.openFileDialogCalled = False
        self.winLoad = wpf.LoadComponent(self, r"G:\My Drive\CV-STRUCTON\01-Dynamo\DynamoTools.extension\DynamoGeneral.tab\Schedule.panel\edit1.stack\ExportSchedule-2024.pushbutton\UI-ExportSchedule.xaml")

        # Tìm và tải các button từ file XAML
        self.bt_SelectPath = self.winLoad.FindName("bt_SelectPath")
        self.bt_Add = self.winLoad.FindName("bt_Add")
        self.bt_Remove = self.winLoad.FindName("bt_Remove")
        self.bt_Export = self.winLoad.FindName("bt_Export")
        self.bt_Cancel = self.winLoad.FindName("bt_Cancel")
        self.txtB_ExcelPath = self.winLoad.FindName("txtB_ExcelPath")

        self.bt_SelectPath.Click +=  self.bt_SelectPathClick
        self.bt_Add.Click += self.bt_AddClick
        self.bt_Remove.Click += self.bt_RemoveClick
        self.bt_Export.Click += self.bt_ExportClick
        self.bt_Cancel.Click += self.bt_CancelClick

        self.FilePath = ""
        
        self.lv_AllScheduleInRevit = self.winLoad.FindName("lv_AllScheduleInRevit")
        self.lv_AllScheduleInRevit.ItemsSource = schedule_Name

        self.lv_AllScheduleInExcel = self.winLoad.FindName("lv_AllScheduleInExcel")
        self.lv_AllScheduleInExcel.ItemsSource = []

        self.SeachInRevit.TextChanged += self.tb_SeachInRevit
        self.SeachInExcel.TextChanged += self.tb_SeachInExcel


        self.refreshListView()

    def refreshListView(self):
        self.lv_AllScheduleInRevit.Items.Refresh()
        self.lv_AllScheduleInExcel.Items.Refresh()

    def bt_ExportClick(self, sender, e):
        if self.FilePath != "none":
            for x in self.lv_AllScheduleInExcel.ItemsSource:
                ScheduleExport.append(x)
            self.Close()            
        else:
            MessageBox.Show("Please select File Excel First","Message")
    def bt_SelectPathClick(self, sender, e):
        if self.openFileDialogCalled == False:
            openFileDialog = OpenFileDialog()
            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx|All Files|*.*"
            result = openFileDialog.ShowDialog()
            if result == DialogResult.OK:
                selectedFilePath = openFileDialog.FileName
                self.FilePath = selectedFilePath
                self.txtB_ExcelPath.Text = self.FilePath
                out_Value["PathLink"] = self.FilePath
                sheetName_remove = []
                sheetName = []
                try:
                    excelApp = Excel.ApplicationClass()
                    workbook = excelApp.Workbooks.Open(selectedFilePath)
                    for item in workbook.Sheets:
                        sheetName.append(item.Name)
                    workbook.Close()
                    excelApp.Quit()
                    for y in sheetName:
                        if y in schedule_Name:
                            sheetName_remove.append(y)
                    for y in sheetName_remove:
                        if y in self.lv_AllScheduleInRevit.ItemsSource:
                            self.lv_AllScheduleInRevit.ItemsSource.remove(y)
                            self.lv_AllScheduleInExcel.ItemsSource.append(y)
                except Exception as e:
                    raise
            else:
                MessageBox.Show("You do not select excel File","Message")
            self.openFileDialogCalled = True
        else:
            self.openFileDialogCalled = False
        self.refreshListView()
    def tb_SeachInRevit(self, sender,e):

        seach_TextRV = sender.Text.lower()

        filtered_ItemsRV = []

        if seach_TextRV:
            for x in self.lv_AllScheduleInRevit.Items:
                if seach_TextRV in x.lower():
                    filtered_ItemsRV.append(x)
        if filtered_ItemsRV:
            self.lv_AllScheduleInRevit.ItemsSource = filtered_ItemsRV
        else:
            self.lv_AllScheduleInRevit.ItemsSource = schedule_Name
        self.refreshListView()

    def tb_SeachInExcel(self, sender,e):
        seach_TextEX = sender.Text.lower()

        filtered_ItemsEX = []
        
        if seach_TextEX:
            for x in self.lv_AllScheduleInExcel.Items:
                if seach_TextEX in x.lower():
                    filtered_ItemsEX.append(x)
        if filtered_ItemsEX:
            self.lv_AllScheduleInExcel.ItemsSource = filtered_ItemsEX
        else:
            # Sử dụng danh sách rỗng thay vì biến sheetExcel không tồn tại
            self.lv_AllScheduleInExcel.ItemsSource = []
        self.refreshListView()
    
    def bt_AddClick(self, sender, e):
        selectedItems = list(self.lv_AllScheduleInRevit.SelectedItems)
        if selectedItems:
            for x in selectedItems:
                self.lv_AllScheduleInRevit.ItemsSource.remove(x)
                self.lv_AllScheduleInExcel.ItemsSource.append(x)
        self.refreshListView()

    def bt_RemoveClick(self, sender, e):
        selectedItems = list(self.lv_AllScheduleInExcel.SelectedItems)
        if selectedItems:
            for x in selectedItems:
                self.lv_AllScheduleInExcel.ItemsSource.remove(x)
                self.lv_AllScheduleInRevit.ItemsSource.append(x)

        self.refreshListView()
    def bt_CancelClick(self, sender, e):
        self.Close()

out_Value = {"PathLink": FilePath,"ScheduleExport": ScheduleExport}
OUT = out_Value
myWindow = MyWindow()
myWindow.ShowDialog()
