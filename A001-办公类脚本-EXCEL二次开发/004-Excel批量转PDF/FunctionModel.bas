Attribute VB_Name = "FunctionModel"
Option Explicit

Sub Excel批量转PDF()
MyForm.Show
End Sub

Function docFileList(ByVal startPath As String) As Variant

'''''''''''''''''''
'取出所有xls或xlsx文件,忽略大小写
'返回包含文件绝对路径的数组
'
'
'''''''''''''''''''
Dim FileName As String
Dim k As Integer
Dim arr() As Variant

FileName = Dir(startPath & "\", vbNormal)
Do
    If FileName Like "*.[Xx][Ll][Ss]*" Then
        k = k + 1
        ReDim Preserve arr(1 To k)
        arr(k) = FileName
        'Debug.Print Filename
    End If
    FileName = Dir
Loop Until FileName = ""

docFileList = arr

End Function

Function DocOpen(ByVal startPath As String)

'屏幕不刷新
Application.ScreenUpdating = False

Dim arr As Variant
arr = docFileList(startPath)

On Error Resume Next

Dim b1 As Integer
Dim b2 As Integer
Dim x As Integer

b1 = LBound(arr)
b2 = UBound(arr)

Dim wb0 As Workbook

Dim nameArr As Variant
Dim nb2 As Integer
Dim suffix As String
Dim FileName As String
Dim FullName As String

For x = b1 To b2
    DoEvents
    nameArr = Split(arr(x), ".")
    nb2 = UBound(nameArr)
    suffix = "." & nameArr(nb2)
    FileName = Replace(arr(x), suffix, "")
    
    '自动覆盖文件
    Application.DisplayAlerts = False

    Set wb0 = Workbooks.Open(startPath & "\" & arr(x))
    
    '文件完整路径名
    FullName = startPath & "\" & FileName & ".pdf"
    
    wb0.ExportAsFixedFormat Type:=xlTypePDF, _
                            FileName:=FullName, _
                            Quality:=xlQualityStandard, _
                            IncludeDocProperties:=True, _
                            IgnorePrintAreas:=False, _
                            OpenAfterPublish:=False
    wb0.Close
    
    Application.DisplayAlerts = True
    
Next

End Function

