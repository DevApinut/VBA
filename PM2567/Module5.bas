Attribute VB_Name = "Module5"
Sub UserForm()
UserForm1.Show
End Sub
Sub UserFormA()
UserForm2.Show
End Sub
Sub UserFormB()
On Error Resume Next
Dim FilePath1 As String
Dim startX As Integer
Dim endX As Integer
j = 1

FilePath1 = ActiveWorkbook.name
startX = Workbooks(FilePath1).Worksheets("main").Range("F2").Value
endX = Workbooks(FilePath1).Worksheets("main").Range("F3").Value

If startX = 0 And endX = 0 Then
UserForm4.Show
Else
UserForm3.Show
End If

End Sub


