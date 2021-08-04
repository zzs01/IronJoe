VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MyForm 
   Caption         =   "Excel批量PDF"
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

Private Sub ConfirmBTN_Click()

Dim startPath As String
startPath = PathBox.Text
Call FunctionModel.DocOpen(startPath)

MsgBox "程序运行结束，请检查！"

End

End Sub

Private Sub PathChooseBTN_Click()

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
    PathBox.Text = paths(1)
Else
    PathBox.Text = ""
End If

End Sub

Private Sub CnacelBTN_Click()
End
End Sub
