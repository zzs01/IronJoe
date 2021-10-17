Attribute VB_Name = "MainProcess"
Option Explicit


Sub 项目供需平衡BOM合并()

MyForm.Show

End Sub


Function MainFunc()
'主要操作
'
'

Dim wb0 As Workbook
Dim shtNameBom As String
Dim shtNamePur As String
Dim bomFilePath As String
Dim purFilePath As String

With MyForm
    
    Set wb0 = MainWb(.ProjectNameBox.Text, .WholePathBox.Text)
    
    shtNameBom = "BOM-Original"
    shtNamePur = "PUR-Original"
    bomFilePath = .bomPathBox.Text
    purFilePath = .purPathBox.Text
    
    GeneralProcess.ExcelMerge wb0, shtNameBom, bomFilePath '无返回值，不带括号
    GeneralProcess.ExcelMerge wb0, shtNamePur, purFilePath '无返回值，不带括号
End With

Dim tm1 As Date
tm1 = Now()

BOMProcess wb0, shtNameBom
Debug.Print "BOM处理总耗时:" & Format(Now() - tm1, "hh:mm:ss")

Dim tm2 As Date
tm2 = Now()

PURProcess wb0, shtNamePur
Debug.Print "PUR处理总耗时:" & Format(Now() - tm2, "hh:mm:ss")

Application.DisplayAlerts = False
wb0.Sheets("sheet1").Delete
wb0.Save
wb0.Close
Application.DisplayAlerts = True

End Function

Function MainWb(ByVal ProjectName As String, ByVal filePath As String) As Workbook

'-----------------------
'生成BOM合成后Excel的表
'参数：
'项目名称：ProjectName
'文件路径：FilePath
'-----------------------
'屏幕不刷新
Application.ScreenUpdating = False

Set MainWb = Workbooks.Add

Dim fileFullPath As String
fileFullPath = filePath & "\" & ProjectName & "供需平衡BOM_" & Format(DateTime.Date, "yyyymmdd") & ".xlsx"


'文件另存为，自动覆盖文件
Application.DisplayAlerts = False
MainWb.SaveAs fileFullPath, ConflictResolution:=2
Application.DisplayAlerts = True

End Function

Function BOMProcess(ByVal wb As Workbook, ByVal shtBomName As String)
'-------------------------
'BOM处理拆分数据
'放量以及实际需求
'-------------------------

'增加工作表并修改名称
Dim shtAD As Worksheet 'BOM-实际需求
Dim shtLS As Worksheet 'BOM-放量

Set shtAD = wb.Sheets.Add
shtAD.Name = "BOM-实际需求"

Set shtLS = wb.Sheets.Add
shtLS.Name = "BOM-放量"

Dim shtBom As Worksheet
Set shtBom = wb.Sheets(shtBomName)

'复制表头
shtBom.Rows(1).Copy Destination:=shtAD.Rows(1)
shtBom.Rows(1).Copy Destination:=shtLS.Rows(1)

'数据写入处理
Dim arr As Variant
arr = shtBom.UsedRange

'数组的行列数目取出
Dim r1 As Long
Dim c1 As Integer
r1 = UBound(arr, 1)
c1 = UBound(arr, 2)

'设置行数，进行循环查找
Dim rNum As Long

Dim keyword As String
keyword = "放量"

'列数I对应的数字
Dim cNum As Integer
cNum = 9

'BOM实际需求与放量实时更新的行数与列数
Dim rAD As Long
Dim rLS As Long
Dim cAL As Integer

rAD = 1
rLS = 1

Dim cellLS As Range
Dim cellAD As Range

Set cellLS = shtLS.Cells
Set cellAD = shtAD.Cells

Application.ScreenUpdating = False
For rNum = 2 To r1
    '判断数据是否是放量
    If InStr(arr(rNum, cNum), keyword) <> 0 Then
        rLS = rLS + 1

        '表内列数与数组列数相同
        For cAL = 1 To c1
            'shtLS.Cells(rLS, cAL).Value = arr(rNum, cAL)
            cellLS(rLS, cAL).Value = arr(rNum, cAL)

        Next
    Else
        '实际需求，不包含放量关键词
        rAD = rAD + 1

        '表内列数与数组列数相同
        For cAL = 1 To c1
            'shtAD.Cells(rAD, cAL).Value = arr(rNum, cAL)
            cellAD(rAD, cAL).Value = arr(rNum, cAL)
        Next
    End If
'DoEvents
Next
Application.ScreenUpdating = True

Application.DisplayAlerts = False
shtBom.Delete
Application.DisplayAlerts = True
End Function

Function PURProcess(ByVal wb As Workbook, ByVal shtPurName As String)
'-------------------------
'PUR处理拆分数据
'放量以及实际需求
'-------------------------

'增加工作表并修改名称
Dim sht1 As Worksheet 'BOM-实际需求

Set sht1 = wb.Sheets.Add
sht1.Name = "PUR"

Dim shtPur As Worksheet
Set shtPur = wb.Sheets(shtPurName)

'复制表头
shtPur.Rows(1).Copy Destination:=sht1.Rows(1)

'数据写入处理
Dim arr As Variant
arr = shtPur.UsedRange

'数组的行列数目取出
Dim r1 As Long
Dim c1 As Integer
r1 = UBound(arr, 1)
c1 = UBound(arr, 2)

'设置行数，进行循环查找
Dim rNum As Long

'列数I对应的数字
Dim cNum1 As Integer
Dim cNum2 As Integer
Dim cNum3 As Integer
cNum1 = 8
cNum2 = 10
cNum3 = 28

'BOM实际需求与放量实时更新的行数与列数
Dim r2 As Long
Dim c2 As Integer

r2 = 1

Dim cell1 As Range

Set cell1 = sht1.Cells

Application.ScreenUpdating = False
For rNum = 2 To r1
    '判断数据是否是放量
    If Len(arr(rNum, cNum1)) = 0 Then
        If Len(arr(rNum, cNum2)) = 0 Then
            If Len(arr(rNum, cNum3)) = 0 Then
                r2 = r2 + 1
                
                '表内列数与数组列数相同
                For c2 = 1 To c1
                    'shtLS.Cells(rLS, cAL).Value = arr(rNum, cAL)
                    cell1(r2, c2).Value = arr(rNum, c2)
                Next
            End If
        End If
    End If
'DoEvents
Next
Application.ScreenUpdating = True

Application.DisplayAlerts = False
shtPur.Delete
Application.DisplayAlerts = True
End Function
