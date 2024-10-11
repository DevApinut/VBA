VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "UserForm2"
   ClientHeight    =   1632
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
Dim FilePath1 As String
Dim startX As Integer
Dim endX As Integer
Dim row_of_Sub As Integer
On Error Resume Next
FilePath1 = ActiveWorkbook.name
startX = Workbooks(FilePath1).Worksheets("main").Range("F2").Value
endX = Workbooks(FilePath1).Worksheets("main").Range("F3").Value + 1
row_of_Sub = Workbooks(FilePath1).Worksheets("main").Range("D2").Value

If CommandButton1.Enabled = True Then

Workbooks(FilePath1).Worksheets("DataBase1").Range("A" & startX & ":" & "B" & endX).EntireRow.delete
Workbooks(FilePath1).Worksheets("DataBase").Range("A" & row_of_Sub).EntireRow.delete
Workbooks(FilePath1).Worksheets("main").Range("C2").Value = ""
MsgBox "ลบสำเร็จ"
Unload Me
End If
End Sub

Private Sub CommandButton2_Click()
Unload Me
End Sub


Private Sub Label1_Click()

End Sub
