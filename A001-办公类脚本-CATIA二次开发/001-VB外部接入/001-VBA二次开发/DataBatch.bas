Attribute VB_Name = "DataBatch"
Public CATIA As INFITF.Application

Sub Main()
    Dim CATIA As Object
    On Error Resume Next
    Set CATIA = GetObject(, "CATIA.Application")
    If Err.Number <> 0 Then
        Set CATIA = CreateObject("CATIA.Application")
        CATIA.Visible = True
    End If
    On Error GoTo 0
    DataForm.igsBTN.Value = True
    DataForm.Show

End Sub

