Attribute VB_Name = "GeneralProcess"
Option Explicit

Function PathSelectionText() As String

'-------------------------------
'路径选择
'-------------------------------

Dim FolderDialogObject As Object
Dim paths As Object

'新建一个对话框对象
Set FolderDialogObject = Application.FileDialog(msoFileDialogFolderPicker)

'配置对话框

With FolderDialogObject

    .Title = "请选择要查找的文件夹"

    .InitialFileName = "C:\"

End With
'显示对话框

FolderDialogObject.Show

'获取选择对话框选择的文件夹

Set paths = FolderDialogObject.SelectedItems

If paths.Count <> 0 Then
    PathSelectionText = paths(1)
Else
    PathSelectionText = ""
End If

End Function

Function excelFileList(ByVal startPath As String) As Variant
'''''''''''''''''''
'取出所有xls或xlsx文件,忽略大小写
'返回包含文件绝对路径的数组
'
'
'''''''''''''''''''
Dim Filename As String
Dim k As Integer
Dim arr() As Variant

'获取文件夹下文件名称
Filename = Dir(startPath & "\", vbNormal)
Do
    '使用like判断是否为excel文件
    If Filename Like "*.[Xx][Ll][Ss]*" Then
        k = k + 1
        '重新定义数组大小，保持前面数据，一维数组可以使用
        ReDim Preserve arr(1 To k)
        '满足条件的写入数组
        arr(k) = Filename
    End If
    '重新传入下一个文件名称
    Filename = Dir
'循环至文件名为空
Loop Until Filename = ""

excelFileList = arr

End Function

Function ExcelMerge(ByVal wb0 As Workbook, ByVal shtName As String, ByVal filePath As String)
'''''''''''''''''
'Excel合并
'合并文件的路径
'主要的工作簿
'合并打开的工作簿
'工作表的名称
'
'''''''''''''''''
Dim i As Integer
Dim r As Long
Dim rg As Range
Dim y  As Long
Dim path As String
Dim c As Integer
Dim ad As String
Dim wb1 As Workbook

'屏幕不刷新
Application.ScreenUpdating = False

'新建工作簿，修改工作表名称
Dim sh As Worksheet

Set sh = wb0.Sheets.Add

sh.Name = shtName

'获取合并文件列表，取出数组上下标号
Dim arr As Variant
Dim b1 As Integer
Dim b2 As Integer
arr = excelFileList(filePath)
b1 = LBound(arr)
b2 = UBound(arr)

'错误处理
On Error Resume Next

Dim x As Integer
For x = b1 To b2
    '打开原始文件位置
    Set wb1 = Workbooks.Open(filePath & "\" & arr(x))
    If x = b1 Then
        '复制单元格数据
        wb1.Sheets(1).UsedRange.Copy wb0.Sheets(shtName).Range("a1")
    Else
        
        r = wb1.Sheets(1).Range("a1").CurrentRegion.Rows.Count
        c = wb1.Sheets(1).Range("a1").CurrentRegion.Columns.Count
        ad = Cells(r, c).Address
        ad = VBA.Split(ad, "$")(1)
        Set rg = Range(ad & r)
        y = wb0.Sheets(1).Range("a1").CurrentRegion.Rows.Count
        wb1.Sheets(1).Range("a2", rg).Copy wb0.Sheets(shtName).Range("a" & (y + 1))
        
    End If
    wb1.Close
    '交接控制权
    DoEvents
Next

Application.ScreenUpdating = True

End Function
