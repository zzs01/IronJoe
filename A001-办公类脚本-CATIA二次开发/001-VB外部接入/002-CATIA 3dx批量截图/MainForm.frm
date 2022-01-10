VERSION 5.00
Begin VB.Form MainForm 
   Caption         =   "3dx批量截图"
   ClientHeight    =   5775
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7920
   LinkTopic       =   "Form1"
   ScaleHeight     =   5775
   ScaleWidth      =   7920
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame1 
      Caption         =   "SetWindowSize"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   20
      Top             =   3720
      Width           =   3615
      Begin VB.TextBox TextBoxWidth 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         TabIndex        =   27
         Text            =   "800"
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox TextBoxHeight 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         TabIndex        =   26
         Text            =   "800"
         Top             =   360
         Width           =   735
      End
      Begin VB.HScrollBar ScrollWidth 
         Height          =   255
         LargeChange     =   10
         Left            =   960
         Max             =   1200
         Min             =   400
         TabIndex        =   23
         Top             =   840
         Value           =   800
         Width           =   1575
      End
      Begin VB.HScrollBar ScrollHeight 
         Height          =   255
         LargeChange     =   10
         Left            =   960
         Max             =   1200
         Min             =   400
         TabIndex        =   21
         Top             =   420
         Value           =   800
         Width           =   1575
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Width"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   24
         Top             =   840
         Width           =   600
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Height"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   22
         Top             =   420
         Width           =   720
      End
   End
   Begin VB.CommandButton BTNCancel 
      Caption         =   "取消"
      Height          =   495
      Left            =   6240
      TabIndex        =   9
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton BTNConfirm 
      Caption         =   "确定"
      Height          =   495
      Left            =   4440
      TabIndex        =   8
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton BTNFolderSelect2 
      Caption         =   "..."
      Height          =   495
      Left            =   7200
      TabIndex        =   7
      Top             =   1080
      Width           =   615
   End
   Begin VB.TextBox TextBox2 
      Height          =   495
      Left            =   1200
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   1080
      Width           =   5895
   End
   Begin VB.CommandButton BTNFolderSelect1 
      Caption         =   "..."
      Height          =   495
      Left            =   7200
      TabIndex        =   4
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox TextBox1 
      Height          =   495
      Left            =   1200
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   480
      Width           =   5895
   End
   Begin VB.Frame FrameRGB 
      Caption         =   "SetBackgroundColor"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   10
      Top             =   1680
      Width           =   7695
      Begin VB.TextBox TextBoxRed 
         Enabled         =   0   'False
         Height          =   375
         Left            =   4800
         TabIndex        =   28
         Text            =   "0"
         Top             =   360
         Width           =   615
      End
      Begin VB.HScrollBar ScrollBarGreen 
         Height          =   300
         LargeChange     =   10
         Left            =   1080
         Max             =   255
         TabIndex        =   19
         Top             =   870
         Width           =   3615
      End
      Begin VB.HScrollBar ScrollBarBlue 
         Height          =   300
         LargeChange     =   10
         Left            =   1080
         Max             =   255
         TabIndex        =   18
         Top             =   1350
         Width           =   3615
      End
      Begin VB.TextBox TextBoxGreen 
         Enabled         =   0   'False
         Height          =   375
         Left            =   4800
         TabIndex        =   17
         Text            =   "0"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox TextBoxBlue 
         Enabled         =   0   'False
         Height          =   375
         Left            =   4800
         TabIndex        =   16
         Text            =   "0"
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox ShowRGB 
         Height          =   1455
         Left            =   5640
         TabIndex        =   15
         Top             =   240
         Width           =   1815
      End
      Begin VB.HScrollBar ScrollBarRed 
         Height          =   300
         LargeChange     =   10
         Left            =   1080
         Max             =   255
         TabIndex        =   11
         Top             =   360
         Width           =   3615
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Red"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   270
         TabIndex        =   14
         Top             =   480
         Width           =   450
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Blue"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   195
         TabIndex        =   13
         Top             =   1440
         Width           =   600
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Green"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   750
      End
   End
   Begin VB.Label LabelCurrentEqu 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1800
      TabIndex        =   30
      Top             =   5280
      Width           =   120
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "当前设备："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   360
      TabIndex        =   29
      Top             =   5280
      Width           =   1275
   End
   Begin VB.Label DDLTxt 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3600
      TabIndex        =   25
      Top             =   120
      Width           =   120
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "图片位置:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   945
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "产品位置:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   2
      Top             =   622
      Width           =   945
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Copyright By IronJoe"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2700
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "抗锯齿开到0.1，超级采样开到最大"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3840
      TabIndex        =   0
      Top             =   3840
      Width           =   3975
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const dealine = "2022/12/31"

Function DDL()
'''''''''''''
'程序过期时间判定
'
''''''''''''
Dim timeDifference
Dim dateTime As Date

dateTime = Date

timeDifference = CInt(CDate(dealine) - dateTime)

If dateTime < dealine Then
    If timeDifference < 4 Then
        Debug.Print 0
        MsgBox "程序还有 " & timeDifference & " 天过期！"
    End If
ElseIf dateTime = dealine Then
    MsgBox "程序今日过后将过期！请尽快联系管理员处理，谢谢！"
Else
    MsgBox "程序已过期，请联系管理员处理，谢谢！"
    End
End If

End Function

Private Sub Form_Initialize()

    Call DDL
    DDLTxt.Caption = "过期时间：" & dealine
'    DDLTxt.Caption = "尊贵的VIP客户，现已对您进行永久授权！"
    DDLTxt.AutoSize = True
    ShowRGB.BackColor = RGB(ScrollBarRed.Value, ScrollBarGreen.Value, ScrollBarBlue)
    
End Sub

Private Sub BTNCancel_Click()
    End
End Sub

Private Sub BTNConfirm_Click()

    On Error Resume Next
    '设置时间
    Dim t1 As Date
    t1 = Timer
    
    '获得环境语言
    Dim flag As Boolean
    flag = MainProcess.GetLanguageFlag
    
    '设置截图尺寸大小
    Dim screenShotHeight As Long
    Dim screenShotWidth As Long
    
    screenShotHeight = ScrollHeight.Value
    screenShotWidth = ScrollWidth.Value
    
    Call MainProcess.SetScreenSize(screenShotHeight, screenShotWidth)
    
    '设置背景颜色
    Dim dblColorArray(2)
    dblColorArray(0) = ScrollBarRed.Value
    dblColorArray(1) = ScrollBarGreen.Value
    dblColorArray(2) = ScrollBarBlue.Value
    
    Dim CATIA As Object
    On Error Resume Next
    Set CATIA = GetObject(, "CATIA.Application")
    If Err.Number <> 0 Then
        Set CATIA = CreateObject("CATIA.Application")
        CATIA.Visible = True
    End If
    
    CATIA.RefreshDisplay = False
        
    On Error Resume Next
    
    Dim folderPath1 As String
    folderPath1 = MainForm.TextBox1.Text
    
    Debug.Print folderPath1
     
    If folderPath1 = "" Then Exit Sub
    
    Dim fs As Object
    Dim fp As Object
    Dim f1 'As Object
    Dim fc As Object
    
    Set fs = CATIA.FileSystem  'CATIA文档系统
    Set fp = fs.GetFolder(folderPath1) '获取输入路径位置
    Set fc = fp.Files
    
    '取出每个文件夹下的cgr文件
    Dim SingleFileList As Variant
    SingleFileList = GetSuffixFile(fc, ".3dxml")
    
    '图片保存位置
    Dim folderPath2 As String
    folderPath2 = MainForm.TextBox2.Text
    If folderPath2 = "" Then Exit Sub
    
    Dim FilePath As String
    
    
    
    For Each f1 In SingleFileList
        
        FilePath = folderPath1 & "\" & f1
        
        LabelCurrentEqu.Caption = f1
        
        Call MainProcess.ExportImage(folderPath2, FilePath, dblColorArray, flag)
    
    Next
    
    CATIA.RefreshDisplay = True
        
    MsgBox "程序运行结束，请检查，谢谢！" & vbCrLf & "程序运行时间：" & Format$(Timer - t1, "0.00") & "s"

    End

End Sub

Private Sub BTNFolderSelect1_Click()
    Dim FilePath As String
    FilePath = GetFolderPath
    
    MainForm.TextBox1.Text = FilePath
End Sub

Private Sub BTNFolderSelect2_Click()
    Dim FilePath As String
    FilePath = GetFolderPath
    
    MainForm.TextBox2.Text = FilePath
End Sub

Private Sub ScrollBarBlue_Change()

    TextBoxBlue.Text = ScrollBarBlue.Value
    ShowRGB.BackColor = RGB(ScrollBarRed.Value, ScrollBarGreen.Value, ScrollBarBlue)
    
End Sub

Private Sub ScrollBarGreen_Change()

    TextBoxGreen.Text = ScrollBarGreen.Value
    ShowRGB.BackColor = RGB(ScrollBarRed.Value, ScrollBarGreen.Value, ScrollBarBlue)
    
End Sub

Private Sub ScrollBarRed_Change()

    TextBoxRed.Text = ScrollBarRed.Value
    ShowRGB.BackColor = RGB(ScrollBarRed.Value, ScrollBarGreen.Value, ScrollBarBlue)
    
End Sub

'Private Sub TextBoxBlue_Change()
'
'    If TextBoxBlue.Value > 255 Then TextBoxBlue.Text = 255
'    If TextBoxBlue.Value < 0 Then TextBoxBlue.Text = 0
'
'    ScrollBarBlue.Value = TextBoxBlue.Text
'
'End Sub
'
'Private Sub TextBoxGreen_Change()
'
'    If TextBoxGreen.Value > 255 Then TextBoxGreen.Text = 255
'    If TextBoxGreen.Value < 0 Then TextBoxGreen.Text = 0
'
'    ScrollBarGreen.Value = TextBoxGreen.Text
'
'End Sub
'
'Private Sub TextBoxRed_Change()
'
'    If TextBoxRed.Value > 255 Then TextBoxRed.Text = 255
'    If TextBoxRed.Value < 0 Then TextBoxRed.Text = 0
'
'    ScrollBarRed.Value = TextBoxRed.Text
'
'End Sub



Function GetFolderPath() As String

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
Dim strTitle

Set objShellApp = CreateObject("Shell.Application")
Set objFolder = objShellApp.BrowseForFolder(WINDOW_HANDLE, strTitle, NO_OPTIONS)

'获得文件夹的名称
objPath = ""
'*****************************'
''''''''''''''
'当前文件夹名称
''''''''''''''
'Dim BrowseForFolderDialogBox
'BrowseForFolderDialogBox = objPath

'If objFolder Is Nothing Then
'
'    objFolder = ""
'
'End If

'*****************************'

If objFolder Is Nothing Then Exit Function

'文件夹绝对路径
Set objFldrItem = objFolder.Self
objPath = objFldrItem.Path

GetFolderPath = CStr(objPath)

Set objShellApp = Nothing
Set objFolder = Nothing
Set objFldrItem = Nothing

End Function

Function GetSuffixFile(ByVal FileVar, ByVal fileSuffix As String) As Variant  'ByVal FileVar As Files

'======================================================'
'***********文件筛选处理，文件夹后缀取出***************'
'======================================================'

Dim file As Object
Dim fileName As String
Dim GetFileList() As Variant

Dim i As Integer

'遍历集合内的文件

For Each file In FileVar
    
    fileName = file.Name
    If InStr(LCase(fileName), fileSuffix) <> 0 Then
        
        i = i + 1
        
        ReDim Preserve GetFileList(1 To i)
        
        GetFileList(i) = fileName

    End If
    
    DoEvents
Next

GetSuffixFile = GetFileList

End Function

Private Sub ScrollHeight_Change()
     TextBoxHeight.Text = ScrollHeight.Value
End Sub

Private Sub ScrollWidth_Change()
    TextBoxWidth.Text = ScrollWidth.Value
End Sub
