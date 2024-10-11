VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   6120
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   8844.001
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CheckBox1_Click()
Dim i As Long
On Error Resume Next

If CheckBox1.Value = True Then
    For i = 2 To TextBox5.Value
        Me.Controls("MFR" & i) = MFR1.Value
        Me.Controls("Type" & i) = Type1.Value
        Me.Controls("kV" & i) = kV1.Value
        Me.Controls("kA" & i) = kA1.Value
        Me.Controls("A" & i) = A1.Value
    Next i
   
Else
    For i = 2 To TextBox5.Value
        Me.Controls("MFR" & i) = ""
        Me.Controls("Type" & i) = ""
        Me.Controls("kV" & i) = ""
        Me.Controls("kA" & i) = ""
        Me.Controls("A" & i) = ""
    Next i
End If
End Sub

Private Sub CheckBox21_Click()
On Error Resume Next
Dim FilePath1 As String
Dim lRow As Integer

FilePath1 = ActiveWorkbook.name
lRow = Workbooks(FilePath1).Worksheets("DataBase1").Cells(Rows.count, 1).End(xlUp).row
If CheckBox21.Value = True Then
CheckBox22.Value = False

TextBox1.Value = Workbooks(FilePath1).Worksheets("main").Range("C2").Value
TextBox2.Value = Workbooks(FilePath1).Worksheets("main").Range("C5").Value
TextBox3.Value = Workbooks(FilePath1).Worksheets("main").Range("C3").Value
TextBox4.Value = Workbooks(FilePath1).Worksheets("main").Range("C4").Value
TextBox1.Enabled = False
TextBox2.Enabled = False
TextBox3.Enabled = False
TextBox4.Enabled = False

Else
TextBox1.Value = ""
TextBox2.Value = ""
TextBox3.Value = ""
TextBox4.Value = ""
TextBox1.Enabled = True
TextBox2.Enabled = True
TextBox3.Enabled = True
TextBox4.Enabled = True
CheckBox22.Value = True

End If
End Sub

Private Sub CheckBox22_Click()
On Error Resume Next
If CheckBox22.Value = True Then
CheckBox21.Value = False
Else
CheckBox21.Value = True
End If

End Sub

Private Sub CommandButton2_Click()
On Error Resume Next
Dim i As Long
If CommandButton2.Enabled = True Then
    For i = 1 To 20
        Me.Controls("code" & i) = ""
        Me.Controls("MFR" & i) = ""
        Me.Controls("Type" & i) = ""
        Me.Controls("Serial" & i) = ""
        Me.Controls("kV" & i) = ""
        Me.Controls("kA" & i) = ""
        Me.Controls("A" & i) = ""
    Next i
End If
End Sub
Private Sub CommandButton1_Click()
On Error Resume Next
Dim FilePath1 As String
Dim lRow As Long
Dim lRow1 As Long
Dim i As Integer
Dim j As Integer
j = 1

FilePath1 = ActiveWorkbook.name
lRow = Workbooks(FilePath1).Worksheets("DataBase1").Cells(Rows.count, 10).End(xlUp).row
lRow1 = Workbooks(FilePath1).Worksheets("DataBase").Cells(Rows.count, 1).End(xlUp).row
If CommandButton1.Enabled = True Then
     
  If CheckBox22.Value = True Then
     
             
            Workbooks(FilePath1).Worksheets("DataBase").Range("C" & lRow1 + 1).Value = TextBox1.Value
            Workbooks(FilePath1).Worksheets("DataBase").Range("D" & lRow1 + 1).Value = TextBox3.Value
            Workbooks(FilePath1).Worksheets("DataBase").Range("E" & lRow1 + 1).Value = TextBox2.Value
            Workbooks(FilePath1).Worksheets("DataBase").Range("A" & lRow1 + 1).Value = "start" & lRow1 + 1
            Workbooks(FilePath1).Worksheets("DataBase").Range("B" & lRow1 + 1).Value = "end" & lRow1 + 1
            Workbooks(FilePath1).Worksheets("DataBase").Range("F" & lRow1 + 1).Value = "LoadbreakX" & lRow1 + 1
            Workbooks(FilePath1).Worksheets("DataBase").Range("G" & lRow1 + 1).Value = "LoadbreakY" & lRow1 + 1
            Workbooks(FilePath1).Worksheets("DataBase").Range("H" & lRow1 + 1).Value = TextBox4.Value



        For i = lRow + 2 To lRow + 2 + TextBox5.Value
            If i = lRow + 2 Then
                Workbooks(FilePath1).Worksheets("DataBase1").Range("I" & i).Value = TextBox1.Value
                Workbooks(FilePath1).Worksheets("DataBase1").Range("J" & i).Value = "start" & lRow1 + 1
            ElseIf i = lRow + 1 + TextBox5.Value Then
        
                Workbooks(FilePath1).Worksheets("DataBase1").Range("J" & i).Value = "end" & lRow1 + 1
                     
            Else
        
                Workbooks(FilePath1).Worksheets("DataBase1").Range("J" & i).Value = ""
                
            End If
            
                Workbooks(FilePath1).Worksheets("DataBase1").Range("L" & i).Value = Me.Controls("code" & j).Value
                Workbooks(FilePath1).Worksheets("DataBase1").Range("M" & i).Value = Me.Controls("MFR" & j).Value
                Workbooks(FilePath1).Worksheets("DataBase1").Range("N" & i).Value = Me.Controls("Type" & j).Value
                Workbooks(FilePath1).Worksheets("DataBase1").Range("O" & i).Value = Me.Controls("Serial" & j).Value
                Workbooks(FilePath1).Worksheets("DataBase1").Range("P" & i).Value = Me.Controls("kV" & j).Value
                Workbooks(FilePath1).Worksheets("DataBase1").Range("Q" & i).Value = Me.Controls("kA" & j).Value
                Workbooks(FilePath1).Worksheets("DataBase1").Range("R" & i).Value = Me.Controls("A" & j).Value
                Workbooks(FilePath1).Worksheets("DataBase1").Range("S" & i).Value = Me.Controls("Feeder" & j).Value
            
                j = j + 1
        Next i
   MsgBox "Complete Insert Data"
 
 ElseIf CheckBox21.Value = True And Workbooks(FilePath1).Worksheets("main").Range("F2").Value = 0 Then
      For i = lRow + 2 To lRow + 2 + TextBox5.Value
            If i = lRow + 2 Then
                Workbooks(FilePath1).Worksheets("DataBase1").Range("I" & i).Value = Workbooks(FilePath1).Worksheets("main").Range("C" & 2).Value
                Workbooks(FilePath1).Worksheets("DataBase1").Range("J" & i).Value = "start" & Workbooks(FilePath1).Worksheets("main").Range("D" & 2).Value
            ElseIf i = lRow + 1 + TextBox5.Value Then
        
                Workbooks(FilePath1).Worksheets("DataBase1").Range("J" & i).Value = "end" & Workbooks(FilePath1).Worksheets("main").Range("D" & 2).Value
                     
            Else
        
                Workbooks(FilePath1).Worksheets("DataBase1").Range("J" & i).Value = ""
                
            End If
            
                Workbooks(FilePath1).Worksheets("DataBase1").Range("L" & i).Value = Me.Controls("code" & j).Value
                Workbooks(FilePath1).Worksheets("DataBase1").Range("M" & i).Value = Me.Controls("MFR" & j).Value
                Workbooks(FilePath1).Worksheets("DataBase1").Range("N" & i).Value = Me.Controls("Type" & j).Value
                Workbooks(FilePath1).Worksheets("DataBase1").Range("O" & i).Value = Me.Controls("Serial" & j).Value
                Workbooks(FilePath1).Worksheets("DataBase1").Range("P" & i).Value = Me.Controls("kV" & j).Value
                Workbooks(FilePath1).Worksheets("DataBase1").Range("Q" & i).Value = Me.Controls("kA" & j).Value
                Workbooks(FilePath1).Worksheets("DataBase1").Range("R" & i).Value = Me.Controls("A" & j).Value
                Workbooks(FilePath1).Worksheets("DataBase1").Range("S" & i).Value = Me.Controls("Feeder" & j).Value
            
                j = j + 1
        Next i
 MsgBox "Complete Insert Data"
 
 Else
 MsgBox "มีข้มูลเเล้วกลับไปแก้ไขข้อมูลโปรดเลือกเมนู Edit"
 
 End If
     

End If

End Sub

Private Sub UserForm_Initialize()

CheckBox21.Value = True
CheckBox22.Value = False

End Sub


