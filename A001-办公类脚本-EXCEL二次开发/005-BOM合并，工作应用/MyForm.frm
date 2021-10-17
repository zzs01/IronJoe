VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MyForm 
   Caption         =   "项目供需平衡BOM合并"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "MyForm.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "MyForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub bomBTN_Click()

'-------------------------------
'BOM文件夹路径选择
'-------------------------------

bomPathBox.Text = GeneralProcess.PathSelectionText

End Sub

Private Sub CancelBTN_Click()
End
End Sub


Private Sub ConfirmBTN_Click()

Dim tm As Date
tm = Now()
MainProcess.MainFunc

MsgBox "合并完成，请检查，谢谢！"

Debug.Print "程序总耗时:" & Format(Now() - tm, "hh:mm:ss")

End

End Sub

Private Sub ProjectPathBTN_Click()

'-------------------------------
'生成项目文件夹路径选择
'-------------------------------

WholePathBox.Text = GeneralProcess.PathSelectionText
End Sub

Private Sub purBTN_Click()

'-------------------------------
'PUR文件夹路径选择
'-------------------------------

purPathBox.Text = GeneralProcess.PathSelectionText


End Sub

