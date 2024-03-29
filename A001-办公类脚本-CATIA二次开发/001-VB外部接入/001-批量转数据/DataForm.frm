VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DataForm 
   Caption         =   "数据批处理"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "DataForm.dsx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "DataForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public dealine As Date

Private Sub cancelBTN_Click()
End
End Sub

Private Sub confirmBTN_Click()
    
    Dim t1 As Date
    t1 = Now

    Dim startPath As String
    Dim txtKey As String
    Dim filePattern As String
    Dim pFlag As Integer
    
    Dim folderinput As String
    Dim folderoutput As String
    folderinput = inPath.Text
    folderoutput = outPath.Text
    
    '=========================
    '建立文件夹
    Dim difFolder As String
    Dim fullFolder As String
    
    Dim objFSO As Object
    Dim oFile As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    difFolder = "A0-不一致文件"
    fullFolder = folderoutput & "\" & difFolder
    
    If Not objFSO.FolderExists(fullFolder) Then '判断文件夹是否存在
       objFSO.CreateFolder (fullFolder) '创建文件夹
    End If
    '==========================
    
    pFlag = 0
    If dataSuffix = "igs" Or dataSuffix = "stp" Then
        pFlag = 1
        Call solidBatch(folderinput, folderoutput, dataSuffix)
    ElseIf dataSuffix = "dwg" Then
        filePattern = "*.[Dd][Ww][Gg]*"
        Call drawingBatch(folderinput, folderoutput, dataSuffix)
    ElseIf dataSuffix = "pdf" Then
        filePattern = "*.[Pp][Dd][Ff]*"
        Call drawingBatch(folderinput, folderoutput, dataSuffix)
    End If
    

    '=========================
    '图纸文件删除
    txtKey = "图纸细节"
    If pFlag = 0 Then
        Call deleteFile(folderoutput, filePattern, txtKey)

        Call renameFile(folderoutput, filePattern)
    End If
    '=========================
    
    
    '=========================
    '文件夹判断删除
    '返回文件夹对象
    Dim objFolder As Object
    Set objFolder = objFSO.GetFolder(fullFolder)
    If objFolder.Files.Count = 1 Then
        objFolder.Delete (True)
    End If
    '=========================
    
    Dim t2 As Date
    t2 = Now
    
    MsgBox "程序运行结束，请检查，谢谢！" & vbCrLf _
        & vbCrLf & "程序开始时间：" & t1 _
        & vbCrLf & "程序结束时间：" & t2 _
        & vbCrLf & "程序耗时    ：" & VBA.DateDiff("s", t1, t2) & "s"
    End

End Sub


Private Sub igsBTN_Click()
    stpBTN.Value = False
    dwgBTN.Value = False
    pdfBTN.Value = False
End Sub


Private Sub stpBTN_Click()
    igsBTN.Value = False
    dwgBTN.Value = False
    pdfBTN.Value = False
End Sub

Private Sub dwgBTN_Click()
    igsBTN.Value = False
    stpBTN.Value = False
    pdfBTN.Value = False
End Sub

Private Sub pdfBTN_Click()
    igsBTN.Value = False
    stpBTN.Value = False
    dwgBTN.Value = False
End Sub
Private Sub inputBRO_Click()
    inPath.Value = FolderBroswer
End Sub

Private Sub outputBRO_Click()
    outPath.Value = FolderBroswer
End Sub

Function dataSuffix() As String

    Select Case OptionButton

    Case igsBTN.Value = False
        dataSuffix = "igs"
    Case stpBTN.Value = False
        dataSuffix = "stp"
    Case dwgBTN.Value = False
        dataSuffix = "dwg"
    Case pdfBTN.Value = False
        dataSuffix = "pdf"
    End Select

End Function

Function solidBatch(ByVal folderinput As String, ByVal folderoutput As String, ByVal suffix As String)
    '========================================
    '*************零件转igs&stp**************
    '========================================

    Dim CATIA As Object
    On Error Resume Next
    Set CATIA = GetObject(, "CATIA.Application")
    If Err.Number <> 0 Then
        Set CATIA = CreateObject("CATIA.Application")
        CATIA.Visible = True
    End If

    On Error GoTo 0
    
    Dim fs As Object
    Dim fp As Object
    Dim f1 As Object
    Dim fc As Object
    
    Set fs = CATIA.FileSystem          'CATIA文档系统
    Set fp = fs.GetFolder(folderinput) '获取输入路径位置
    Set fc = fp.Files                  '文件集合
    
    '后缀名指定
    Dim SuffixAdd As String
    
    SuffixAdd = "." & suffix
    
    '遍历集合内的文件
    Dim documents1 As Documents
    Dim PartDocument1 'As PartDocument
    Dim fNum As Integer
    fNum = 0
    
    Dim originOutPath As String
    For Each f1 In fc
        
        originOutPath = folderoutput
        Set documents1 = CATIA.Documents
    
        Dim INTRARE As String
        INTRARE = CATIA.FileSystem.ConcatenatePaths(folderinput, f1.Name)
        
        '支持产品和零件
        If InStr(f1.Name, ".CATPart") = 0 And InStr(f1.Name, ".CATProduct") = 0 Then
            GoTo Continue
        End If
    
        Set PartDocument1 = documents1.Open(INTRARE)
    
        '字符串分割后为数组
        Dim docname As String
        Dim pathname As String
        Dim IESIRE As String
        Dim partNum As String
 
        '文件名与零件编号对比
        partNum = PartDocument1.Product.PartNumber
        docname = Split(f1.Name, ".")(0)
        
        '写入txt,并改写文件夹路径位置
        Dim difFolder As String
        Dim fullFolder As String
        difFolder = "A0-不一致文件"
        fullFolder = folderoutput & "\" & difFolder
        
        Dim txtPath As String
        txtPath = fullFolder & "\" & "A0-编号对比结果.txt"
        
        If partNum <> docname Then
            fNum = fNum + 1
            'Open txtPath For Output As #1
            Open txtPath For Append As #1
                Print #1, "P" & fNum
                Print #1, "文件名称:" & docname
                Print #1, "零件编号:" & partNum
                Print #1, VBA.Chr(13)
            Close #1
            originOutPath = fullFolder
        End If
        
        pathname = originOutPath & "\" & docname
    
        IESIRE = pathname & SuffixAdd
        
        CATIA.DisplayFileAlerts = False
    
        PartDocument1.ExportData IESIRE, suffix
    
        CATIA.ActiveDocument.Close
        
        CATIA.DisplayFileAlerts = True
    
Continue:
        Next

End Function

Function drawingBatch(ByVal suffix As String)


Dim CATIA As Object
On Error Resume Next
Set CATIA = GetObject(, "CATIA.Application")
If Err.Number <> 0 Then
    Set CATIA = CreateObject("CATIA.Application")
    CATIA.Visible = True
End If

On Error GoTo 0

Dim folderinput As String
Dim folderoutput As String

folderinput = inPath.Text
folderoutput = outPath.Text

Dim fs As Object
Dim fp As Object
Dim f1 As Object
Dim fc As Object

Set fs = CATIA.FileSystem  'CATIA文档系统
Set fp = fs.GetFolder(folderinput) '获取输入路径位置
Set fc = fp.Files                  '文件集合

'后缀名指定
Dim SuffixAdd As String

SuffixAdd = "." & suffix

'遍历集合内的文件
For Each f1 In fc

    '输入文件名称
    Dim INTRARE As String
    INTRARE = CATIA.FileSystem.ConcatenatePaths(folderinput, f1.Name)

    If InStr(f1.Name, ".CATDrawing") = 0 Then
        GoTo Continue
    End If

    '输出文件完整路径名
    '字符串分割后为数组
    docname = Split(f1.Name, ".")(0)
    pathname = folderoutput & "\" & docname

    '创建图形文件集合
    Set drawingDocuments = CATIA.Documents

    '打开图形文件
    Dim drawingDocument1 As Document
    Set drawingDocument1 = drawingDocuments.Open(INTRARE)

    Dim drawingSheets1 As DrawingSheets
    Set drawingSheets1 = drawingDocument1.Sheets

    Dim drawingSheet1 As DrawingSheet
    Set drawingSheet1 = drawingSheets1.ActiveSheet

    Dim IESIRE As String
    IESIRE = pathname & SuffixAdd

    drawingDocument1.ExportData IESIRE, suffix
    
    CATIA.ActiveDocument.Close
Continue:
    Next

End Function

Function FolderBroswer() As String

'=================================================================================================
' TITLE: BrowseForFolderDialogBox
'
' PURPOSE: Open Window dialog box to choose directory to process.
' Return the choosen path to a variable.
'
' HOW TO USE: strResult = BrowseForFolderDialogBox()
' ==================================================

Const WINDOW_HANDLE = 0
Const NO_OPTIONS = &H10
Dim objShellApp
Dim objFolder
Dim objFldrItem
Dim objPath

Set objShellApp = CreateObject("Shell.Application")
Set objFolder = objShellApp.BrowseForFolder(WINDOW_HANDLE, strTitle, NO_OPTIONS)

On Error Resume Next
If (objFolder = "") Then
    WScript.Echo "You choose to cancel. This will stop this script."
    WScript.Quit
Else
    Set objFldrItem = objFolder.Self
    objPath = objFldrItem.Path

    FolderBroswer = objPath

    Set objShellApp = Nothing
    Set objFolder = Nothing
    Set objFldrItem = Nothing
End If

End Function

Function FileList(ByVal startPath As String, ByVal filePattern As String) As Variant
'''''''''''''''''''
'取出所有xls或xlsx文件,忽略大小写
'返回包含文件绝对路径的数组
'
'filePattern = "*.[Xx][Ll][Ss]*"
'''''''''''''''''''
Dim Filename As String
Dim k As Integer
Dim arr() As Variant


'获取文件夹下文件名称
Filename = Dir(startPath & "\", vbNormal)
Do
    '使用like判断是否为excel文件
    If Filename Like filePattern Then
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

FileList = arr

End Function


'增加数组判断


Function deleteFile(ByVal startPath As String, ByVal filePattern As String, ByVal txtKey As String)
''''''''''''''''''''
'删除满足条件的文件
'
''''''''''''''''''''


Dim b1 As Integer
Dim b2 As Integer
Dim arr As Variant


arr = FileList(startPath, filePattern)
b1 = LBound(arr)
b2 = UBound(arr)


Dim x As Integer
Dim fullName As String


Dim MyFile As Object
Set MyFile = CreateObject("Scripting.FileSystemObject")
For x = b1 To b2
    If InStr(arr(x), txtKey) > 0 Then
        fullName = startPath & "\" & arr(x)
        MyFile.deleteFile fullName
    End If
Next x
End Function

Function renameFile(ByVal startPath As String, ByVal filePattern As String)

''''''''''''''''''''
'重命名文件,删除文件固定后缀
'
''''''''''''''''''''

Dim b1 As Integer
Dim b2 As Integer
Dim arr As Variant

arr = FileList(startPath, filePattern)
b1 = LBound(arr)
b2 = UBound(arr)

Dim x As Integer
Dim fullName As String
Dim deTxt As String
Dim nameArr As Variant
Dim nb2 As Integer
Dim suffix As String

Dim MyFile As Object
Dim oFile As Object
Set MyFile = CreateObject("Scripting.FileSystemObject")
For x = b1 To b2
    nameArr = Split(arr(x), "_")
    nb2 = UBound(nameArr)
    deTxt = "_" & nameArr(nb2)
    suffix = Split(nameArr(nb2), ".")(1)
    fullName = startPath & "\" & arr(x)
    Debug.Print deTxt
    Set oFile = MyFile.GetFile(fullName)
    '重命名不包括后缀名
    oFile.Name = Replace(arr(x), deTxt, "") & "." & suffix
Next x

End Function

Function DDL() As Boolean
'''''''''''''
'程序过期时间判定
'
''''''''''''
dealine = "2021/12/30"
DDL = True

Dim timeDifference
Dim dateTime As Date
dateTime = VBA.Date
Debug.Print dateTime
timeDifference = dealine - dateTime
If dateTime < dealine Then
    If timeDifference < 4 Then
        MsgBox "程序还有 " & timeDifference & " 天过期！"
    End If
ElseIf dateTime = dealine Then
    MsgBox "程序今日过后将过期！请尽快联系管理员处理，谢谢！"
Else
    DDL = False
    MsgBox "程序已过期，请联系管理员处理，谢谢！"
End If

End Function


Private Sub UserForm_Initialize()
'Call DDL

'DDLTxt.Caption = "过期时间：" & dealine
DDLTxt.Caption = "尊贵的VIP客户，现已对您进行永久授权！"
DDLTxt.AutoSize = True
End Sub
