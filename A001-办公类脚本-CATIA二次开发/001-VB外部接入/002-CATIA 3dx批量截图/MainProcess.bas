Attribute VB_Name = "MainProcess"
Option Explicit

Public CATIA As INFITF.Application
Const suffix = ".JPEG" '保存后缀 'png 4
Const imgFormatMode = 5

'抗锯齿开到0.1，超级采样开到最大
'增加隐藏目录树的焊枪以及参考，工件 功能
'增加多个视图截图调整
'增加背景颜色调节
'增加透明度处理

Sub Main()
    MainForm.Show
End Sub

Function ExportImage(ByVal folderPath As String, ByVal FilePath As String, ByVal dblColorArray, ByVal languageFlag As Boolean)

    'folderPath 图片保存位置
    'filePath 文件打开位置
    'IsLFlag 左侧截图标志
    
    Dim CATIA As Object
    On Error Resume Next
    Set CATIA = GetObject(, "CATIA.Application")
    If Err.Number <> 0 Then
        Set CATIA = CreateObject("CATIA.Application")
        CATIA.Visible = True
    End If
    
    Dim documents1 As Documents
    Set documents1 = CATIA.Documents
    
    Dim Document1 As Document
    Set Document1 = documents1.Open(FilePath)
    
    Debug.Print FilePath
    
    Dim productDocument1 As ProductDocument
    Set productDocument1 = CATIA.ActiveDocument
    
    Dim product1 As Product
    Set product1 = productDocument1.Product
    
    '保存名称
    Dim partname As String
    partname = product1.PartNumber
    
    Debug.Print partname
    
    '完整名称
    Dim strName As String
    strName = folderPath & "\" & partname
    
    '======================================================'
    '**********************设置文档线宽********************'
    '======================================================'
    
    Dim selection1 As Selection
    Set selection1 = productDocument1.Selection
    
    selection1.Clear
    selection1.Add product1
    
    Dim Obj As VisPropertySet
    Set Obj = selection1.VisProperties
    
'    '获取初始线宽
'    Dim width As Long
'    Obj.GetRealWidth width
    
    '设置线宽
    Obj.SetRealWidth 55#, 1
    selection1.Clear
    
    '======================================================'
    '*************设置背景渐变样式以及背景颜色*************'
    '======================================================'
    Dim settingControllers1 As SettingControllers
    Set settingControllers1 = CATIA.SettingControllers
    
    Dim visualizationSettingAtt1 'As VisualizationSettingAtt
    Set visualizationSettingAtt1 = settingControllers1.Item("CATVizVisualizationSettingCtrl")
    
    '获得初始的背景颜色，用于后面恢复
    Dim DBLBackArray(2)
    visualizationSettingAtt1.GetBackgroundRGB DBLBackArray(0), DBLBackArray(1), DBLBackArray(2)
    
    '设置背景颜色
    visualizationSettingAtt1.SetBackgroundRGB dblColorArray(0), dblColorArray(1), dblColorArray(2)
    
    Dim showFlag As Boolean
    showFlag = visualizationSettingAtt1.ColorBackgroundMode
    
    'True为渐变，False为不渐变
    If showFlag Then
        visualizationSettingAtt1.ColorBackgroundMode = False
    End If
    
    visualizationSettingAtt1.SaveRepository
    
    '======================================================'
    '**********************文档居中显示********************'
    '======================================================'
    Dim specsAndGeomWindow1 As SpecsAndGeomWindow
    Set specsAndGeomWindow1 = CATIA.ActiveWindow
    
    Dim viewer3D1 'As Viewer3D
    Set viewer3D1 = specsAndGeomWindow1.ActiveViewer
    
    '隐藏结构树
    specsAndGeomWindow1.Layout = catWindowGeomOnly
    
    '======================================================'
    '**********************截图多种角度********************'
    '======================================================'
    
    '***************************************
    '隐藏指南针
    
    If languageFlag Then
        CATIA.StartCommand ("指南针")
    Else
        CATIA.StartCommand ("Compass")
    End If
    
    
    '======================================================'
    
    '截图本体
    Call SetViewPoint(viewer3D1, strName)
    
    '显示指南针
    If languageFlag Then
        CATIA.StartCommand ("指南针")
    Else
        CATIA.StartCommand ("Compass")
    End If
    
    '======================================================'
    '*************恢复背景渐变样式以及背景颜色*************'
    '======================================================'
        
    If showFlag Then
        visualizationSettingAtt1.ColorBackgroundMode = True
    End If
    
    visualizationSettingAtt1.SetBackgroundRGB DBLBackArray(0), DBLBackArray(1), DBLBackArray(2)
    
    visualizationSettingAtt1.SaveRepository
    
    '恢复结构树
    specsAndGeomWindow1.Layout = catWindowSpecsAndGeom

    '刷新背景
    viewer3D1.Update
    
    productDocument1.Close
    
End Function

Sub SetViewPoint(ByVal viewer3D1, ByVal strName As String)

    '======================================================'
    '***************截图四种样式以供选择*******************'
    '======================================================'

    
    Dim SightDir(2), UpDir(2) '定义视线方向和上方向坐标分量

'    '轴测视图
'    SightDir(0) = 1.414
'    SightDir(1) = 1.414
'    SightDir(2) = 0
'
'    UpDir(0) = 1.414
'    UpDir(1) = 0
'    UpDir(2) = 0
    
    Dim i As Integer
    
    Dim s1 As Double
    Dim s2 As Double
    Dim s3 As Double
    
    s1 = 0.5
    s2 = 0.8
    s3 = 0.3
    
    Dim u1 As Double
    Dim u2 As Double
    Dim u3 As Double
    
    u1 = 0.7
    u2 = 0.6
    u3 = 0.3
    
    For i = 1 To 4
        If i = 1 Then
            SightDir(0) = -s1
            SightDir(1) = -s2
            SightDir(2) = s3
            
            UpDir(0) = -u1
            UpDir(1) = u2
            UpDir(2) = u3
            
        ElseIf i = 2 Then
            SightDir(0) = s1
            SightDir(1) = -s2
            SightDir(2) = s3
            
            UpDir(0) = u1
            UpDir(1) = u2
            UpDir(2) = u3
            
        ElseIf i = 3 Then
            SightDir(0) = -s1
            SightDir(1) = s2
            SightDir(2) = -s3
            
            UpDir(0) = -u1
            UpDir(1) = -u2
            UpDir(2) = -u3
            
        Else
            SightDir(0) = -s1
            SightDir(1) = s2
            SightDir(2) = s3
            
            UpDir(0) = -u1
            UpDir(1) = -u2
            UpDir(2) = u3
        End If
    
        Dim viewpoint3D1 'As Viewpoint3D
        Set viewpoint3D1 = viewer3D1.Viewpoint3D
        viewpoint3D1.Zoom = 2
        
        viewpoint3D1.PutSightDirection SightDir  '从视点看向原点
        viewpoint3D1.PutUpDirection UpDir '定义上方方向分量
        
        '自适应居中
        viewer3D1.Reframe
        '放大
        'viewer3D1.ZoomIn
        
        '刷新背景
        viewer3D1.Update
          
        '''''截图程序本体，最后复原'''''
        Dim fileName As String
        
        fileName = strName & "-" & i & suffix
        viewer3D1.CaptureToFile imgFormatMode, fileName
    Next
    
End Sub

Sub SetScreenSize(ByVal screenShotHeight As Long, ByVal screenShotWidth As Long)

    Dim CATIA As Object
    On Error Resume Next
    Set CATIA = GetObject(, "CATIA.Application")
    If Err.Number <> 0 Then
        Set CATIA = CreateObject("CATIA.Application")
        CATIA.Visible = True
    End If

    '======================================================'
    '*******************设置截图窗口大小*******************'
    '======================================================'
    '新建一个Product进行设置
    
    'catia 刷新窗口显示才能捕捉到窗口大小
    
    CATIA.RefreshDisplay = True

    Dim documents1 As Documents
    Set documents1 = CATIA.Documents
'
    Dim productDocument1 As ProductDocument
    Set productDocument1 = documents1.Add("Product")
    
    Set productDocument1 = CATIA.ActiveDocument
    
    Dim product1 As Product
    Set product1 = productDocument1.Product
    
    '保存名称
    Dim partname As String
    partname = product1.PartNumber
    
    Dim specsAndGeomWindow1 As SpecsAndGeomWindow
    Set specsAndGeomWindow1 = CATIA.ActiveWindow
    
    Dim sH As Long
    Dim sW As Long
    
    sH = specsAndGeomWindow1.Height
    sW = specsAndGeomWindow1.Width
    
    CATIA.Height = CATIA.Height - (sH - screenShotHeight)
    CATIA.Width = CATIA.Width - (sW - screenShotWidth) '补偿截图
    
    specsAndGeomWindow1.Close
    productDocument1.Close
    
    '修正尺寸
    If specsAndGeomWindow1.Height <> screenShotHeight And specsAndGeomWindow1.Height <> screenShotWidth Then
        Call SetScreenSize(screenShotHeight, screenShotWidth)
    End If
    
    
End Sub

Function GetLanguageFlag() As Boolean
    
    Dim CATIA As Object
    On Error Resume Next
    Set CATIA = GetObject(, "CATIA.Application")
    If Err.Number <> 0 Then
        Set CATIA = CreateObject("CATIA.Application")
        CATIA.Visible = True
    End If
    
    '======================================================'
    '********************获取环境语言**********************'
    '======================================================'
    
    If CATIA.Windows.Count < 1 Then
        MsgBox "请保证CATIA至少有一个打开的文档窗口后重新启动此程序！"
        End
    End If
    
    Dim productDocument1 As Document
    Set productDocument1 = CATIA.ActiveDocument
   
    productDocument1.Selection.Clea
    
    On Error Resume Next
    
    Dim a
    a = Left(CATIA.StatusBar, 1)
    'a = CATIA.StatusBar
    
    If a > "~" Then   '中文
    
        GetLanguageFlag = 1
    
    ElseIf Asc(a) > 47 And Asc(a) < 123 Then    '英文
    
        GetLanguageFlag = 0
    Else
    
        MsgBox "非中英文环境"
        End
    
    End If
    On Error GoTo 0
    
End Function


