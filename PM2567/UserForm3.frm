VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3 
   Caption         =   "UserForm3"
   ClientHeight    =   6048
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   11676
   OleObjectBlob   =   "UserForm3.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CheckBox1_Click()
Dim i As Long
On Error Resume Next

If CheckBox1.Value = True Then
    For i = 2 To TextBox5.Value
        Me.Controls("code" & i) = code1.Value
        Me.Controls("MFR" & i) = MFR1.Value
        Me.Controls("Type" & i) = Type1.Value
        Me.Controls("serial" & i) = Serial1.Value
        Me.Controls("kV" & i) = kV1.Value
        Me.Controls("kA" & i) = kA1.Value
        Me.Controls("A" & i) = A1.Value
    Next i
   
'Else
    'For i = 2 To TextBox5.Value
       ' Me.Controls("MFR" & i) = ""
        'Me.Controls("Type" & i) = ""
        'Me.Controls("kV" & i) = ""
       ' Me.Controls("kA" & i) = ""
        'Me.Controls("A" & i) = ""
    'Next i
End If
End Sub

Private Sub CommandButton1_Click()
On Error Resume Next
Dim FilePath1 As String
Dim lRow As Long
Dim i As Integer
Dim j As Integer
Dim startX As Integer
Dim endX As Integer
Dim numberEq As Integer
Dim row As Integer

j = 1

FilePath1 = ActiveWorkbook.name
row = Workbooks(FilePath1).Worksheets("main").Range("D2").Value
startX = Workbooks(FilePath1).Worksheets("main").Range("F2").Value
endX = Workbooks(FilePath1).Worksheets("main").Range("F3").Value
row_of_Sub = Workbooks(FilePath1).Worksheets("main").Range("D2").Value


numberEq = endX - startX + 1

If CommandButton1.Enabled = True Then
            'Edit header
            Workbooks(FilePath1).Worksheets("DataBase").Range("C" & row).Value = TextBox1.Value
            Workbooks(FilePath1).Worksheets("DataBase").Range("D" & row).Value = TextBox3.Value
            Workbooks(FilePath1).Worksheets("DataBase").Range("E" & row).Value = TextBox4.Value
            Workbooks(FilePath1).Worksheets("DataBase").Range("A" & row).Value = "start" & row
            Workbooks(FilePath1).Worksheets("DataBase").Range("B" & row).Value = "end" & row
            Workbooks(FilePath1).Worksheets("DataBase").Range("F" & row).Value = "LoadbreakX" & row
            Workbooks(FilePath1).Worksheets("DataBase").Range("G" & row).Value = "LoadbreakY" & row
            Workbooks(FilePath1).Worksheets("DataBase").Range("H" & row).Value = TextBox2.Value
    
    
    
    If TextBox5.Value > numberEq Then
        MsgBox "เนื่องจากจำนวนอุปกรณ์มีการเปลี่ยนแปลงกรุณาลบข้อมูลเก่าแล้วทำการเพิ่มใหม่โดยกดปุ่ม แทนที่ข้อมูลเก่า"
        CommandButton2.Enabled = True
        Workbooks(FilePath1).Worksheets("main").Range("D7").Value = TextBox5.Value
    ElseIf TextBox5.Value < numberEq Then
        MsgBox "เนื่องจากจำนวนอุปกรณ์มีการเปลี่ยนแปลงกรุณาลบข้อมูลเก่าแล้วทำการเพิ่มใหม่โดยกดปุ่ม แทนที่ข้อมูลเก่า"
        CommandButton2.Enabled = True
        Workbooks(FilePath1).Worksheets("main").Range("D7").Value = TextBox5.Value
    Else
        For i = startX To endX
            Workbooks(FilePath1).Worksheets("DataBase1").Range("L" & i).Value = Me.Controls("code" & j).Value
            Workbooks(FilePath1).Worksheets("DataBase1").Range("M" & i).Value = Me.Controls("MFR" & j).Value
            Workbooks(FilePath1).Worksheets("DataBase1").Range("N" & i).Value = Me.Controls("Type" & j).Value
            Workbooks(FilePath1).Worksheets("DataBase1").Range("O" & i).Value = Me.Controls("Serial" & j).Value
            Workbooks(FilePath1).Worksheets("DataBase1").Range("P" & i).Value = Me.Controls("kV" & j).Value
            Workbooks(FilePath1).Worksheets("DataBase1").Range("Q" & i).Value = Me.Controls("kA" & j).Value
            Workbooks(FilePath1).Worksheets("DataBase1").Range("R" & i).Value = Me.Controls("A" & j).Value
            Workbooks(FilePath1).Worksheets("DataBase1").Range("S" & i).Value = Me.Controls("note" & j).Value
            Workbooks(FilePath1).Worksheets("DataBase1").Range("T" & i).Value = Me.Controls("Feeder" & j).Value
            
            j = j + 1
    Next i
   MsgBox "เพิ่มข้อมูลสำเร็จ"
    End If

End If



End Sub

Private Sub CommandButton2_Click()
Dim FilePath1 As String
Dim startX As Integer
Dim endX As Integer
Dim row_of_Sub As Integer
Dim j As Integer
On Error Resume Next

j = 1
FilePath1 = ActiveWorkbook.name
startX = Workbooks(FilePath1).Worksheets("main").Range("F2").Value
endX = Workbooks(FilePath1).Worksheets("main").Range("F3").Value + 1
row_of_Sub = Workbooks(FilePath1).Worksheets("main").Range("D2").Value

If CommandButton2.Enabled = True Then

Workbooks(FilePath1).Worksheets("DataBase1").Range("A" & startX & ":" & "B" & endX).EntireRow.delete
'Workbooks(FilePath1).Worksheets("DataBase").Range("A" & row_of_Sub).EntireRow.delete
CommandButton2.Enabled = False
lRow = Workbooks(FilePath1).Worksheets("DataBase1").Cells(Rows.count, 10).End(xlUp).row
 For i = lRow + 2 To lRow + 2 + Workbooks(FilePath1).Worksheets("main").Range("D7").Value
            If i = lRow + 2 Then
                Workbooks(FilePath1).Worksheets("DataBase1").Range("I" & i).Value = TextBox1.Value
                Workbooks(FilePath1).Worksheets("DataBase1").Range("J" & i).Value = "start" & row_of_Sub
            ElseIf i = lRow + 1 + TextBox5.Value Then
        
                Workbooks(FilePath1).Worksheets("DataBase1").Range("J" & i).Value = "end" & row_of_Sub
                     
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
                Workbooks(FilePath1).Worksheets("DataBase1").Range("S" & i).Value = Me.Controls("note" & j).Value
                Workbooks(FilePath1).Worksheets("DataBase1").Range("T" & i).Value = Me.Controls("Feeder" & j).Value
            
                j = j + 1
        Next i

MsgBox "แทนที่ข้อมูลสำเร็จ"
End If



End Sub


Private Sub CommandButton3_Click()
On Error Resume Next
Dim FilePath1 As String
Dim code As String
Dim MFR As String
Dim Typex As String
Dim Serial As String
Dim kV As String
Dim kA As String
Dim a As String
Dim note As String
Dim Feeder As String
Dim numberEq As Integer

code = Me.Controls("code" & Ch2.Value).Value
MFR = Me.Controls("MFR" & Ch2.Value).Value
Typex = Me.Controls("Type" & Ch2.Value).Value
Serial = Me.Controls("Serial" & Ch2.Value).Value
kV = Me.Controls("kV" & Ch2.Value).Value
kA = Me.Controls("kA" & Ch2.Value).Value
a = Me.Controls("A" & Ch2.Value).Value
note = Me.Controls("note" & Ch2.Value).Value
Feeder = Me.Controls("Feeder" & Ch2.Value).Value

FilePath1 = ActiveWorkbook.name
lRow = Workbooks(FilePath1).Worksheets("DataBase1").Cells(Rows.count, 10).End(xlUp).row

If CommandButton3.Enabled = True Then

Me.Controls("code" & Ch2.Value).Value = Me.Controls("code" & Ch1.Value).Value
Me.Controls("MFR" & Ch2.Value).Value = Me.Controls("MFR" & Ch1.Value).Value
Me.Controls("Type" & Ch2.Value).Value = Me.Controls("Type" & Ch1.Value).Value
Me.Controls("Serial" & Ch2.Value).Value = Me.Controls("Serial" & Ch1.Value).Value
Me.Controls("kV" & Ch2.Value).Value = Me.Controls("kV" & Ch1.Value).Value
Me.Controls("kA" & Ch2.Value).Value = Me.Controls("kA" & Ch1.Value).Value
Me.Controls("A" & Ch2.Value).Value = Me.Controls("A" & Ch1.Value).Value
Me.Controls("Feeder" & Ch2.Value).Value = Me.Controls("Feeder" & Ch1.Value).Value
'Me.Controls("note" & Ch2.Value).Value = Me.Controls("note" & Ch1.Value).Value

Me.Controls("code" & Ch1.Value).Value = code
Me.Controls("MFR" & Ch1.Value).Value = MFR
Me.Controls("Type" & Ch1.Value).Value = Typex
Me.Controls("Serial" & Ch1.Value).Value = Serial
Me.Controls("kV" & Ch1.Value).Value = kV
Me.Controls("kA" & Ch1.Value).Value = kA
Me.Controls("A" & Ch1.Value).Value = a
Me.Controls("Feeder" & Ch1.Value).Value = Feeder
'Me.Controls("note" & Ch1.Value).Value = note

Me.Controls("note" & Ch1.Value).Value = "สลับมาจากอุปกรณ์" & Ch2.Value
Me.Controls("note" & Ch2.Value).Value = "สลับมาจากอุปกรณ์" & Ch1.Value

End If


End Sub

Private Sub CommandButton4_Click()
On Error Resume Next
Dim i As Integer
Dim j As Integer
Dim k As Integer



If CommandButton4.Enabled = True Then
 TextBox5.Value = TextBox5.Value + 1
 For i = 1 To TextBox5.Value + 2
    j = TextBox5.Value - i + 1
    
    
        Me.Controls("code" & j).Value = Me.Controls("code" & j - 1).Value
        Me.Controls("MFR" & j).Value = Me.Controls("MFR" & j - 1).Value
        Me.Controls("Type" & j).Value = Me.Controls("Type" & j - 1).Value
        Me.Controls("Serial" & j).Value = Me.Controls("Serial" & j - 1).Value
        Me.Controls("kV" & j).Value = Me.Controls("kV" & j - 1).Value
        Me.Controls("kA" & j).Value = Me.Controls("kA" & j - 1).Value
        Me.Controls("A" & j).Value = Me.Controls("A" & j - 1).Value
        Me.Controls("note" & j).Value = Me.Controls("note" & j - 1).Value
        Me.Controls("Feeder" & j).Value = Me.Controls("Feeder" & j - 1).Value
              
Next i
        Me.Controls("code1").Value = ""
        Me.Controls("MFR1").Value = ""
        Me.Controls("Type1").Value = ""
        Me.Controls("Serial1").Value = ""
        Me.Controls("kV1").Value = ""
        Me.Controls("kA1").Value = ""
        Me.Controls("A1").Value = ""
        Me.Controls("note1").Value = ""
        Me.Controls("Feeder").Value = ""
       
For i = 1 To TextBox6.Value - 1
    
    
        Me.Controls("code" & i).Value = Me.Controls("code" & i + 1).Value
        Me.Controls("MFR" & i).Value = Me.Controls("MFR" & i + 1).Value
        Me.Controls("Type" & i).Value = Me.Controls("Type" & i + 1).Value
        Me.Controls("Serial" & i).Value = Me.Controls("Serial" & i + 1).Value
        Me.Controls("kV" & i).Value = Me.Controls("kV" & i + 1).Value
        Me.Controls("kA" & i).Value = Me.Controls("kA" & i + 1).Value
        Me.Controls("A" & i).Value = Me.Controls("A" & i + 1).Value
        Me.Controls("note" & i).Value = Me.Controls("note" & i + 1).Value
        Me.Controls("Feeder" & i).Value = Me.Controls("Feeder" & i + 1).Value
              
Next i

        Me.Controls("code" & TextBox6.Value).Value = ""
        Me.Controls("MFR" & TextBox6.Value).Value = ""
        Me.Controls("Type" & TextBox6.Value).Value = ""
        Me.Controls("Serial" & TextBox6.Value).Value = ""
        Me.Controls("kV" & TextBox6.Value).Value = ""
        Me.Controls("kA" & TextBox6.Value).Value = ""
        Me.Controls("A" & TextBox6.Value).Value = ""
        Me.Controls("note" & TextBox6.Value).Value = ""
        Me.Controls("Feeder" & TextBox6.Value).Value = ""
End If

End Sub

Private Sub CommandButton5_Click()

On Error Resume Next
Dim i As Integer
Dim j As Integer
Dim k As Integer



If CommandButton5.Enabled = True Then
 TextBox5.Value = TextBox5.Value - 1
 For i = TextBox6.Value To TextBox5.Value
    
    
        Me.Controls("code" & i).Value = Me.Controls("code" & i + 1).Value
        Me.Controls("MFR" & i).Value = Me.Controls("MFR" & i + 1).Value
        Me.Controls("Type" & i).Value = Me.Controls("Type" & i + 1).Value
        Me.Controls("Serial" & i).Value = Me.Controls("Serial" & i + 1).Value
        Me.Controls("kV" & i).Value = Me.Controls("kV" & i + 1).Value
        Me.Controls("kA" & i).Value = Me.Controls("kA" & i + 1).Value
        Me.Controls("A" & i).Value = Me.Controls("A" & i + 1).Value
        Me.Controls("note" & i).Value = Me.Controls("note" & i + 1).Value
        Me.Controls("Feeder" & i).Value = Me.Controls("Feeder" & i + 1).Value
              
Next i

        Me.Controls("code" & i).Value = ""
        Me.Controls("MFR" & i).Value = ""
        Me.Controls("Type" & i).Value = ""
        Me.Controls("Serial" & i).Value = ""
        Me.Controls("kV" & i).Value = ""
        Me.Controls("kA" & i).Value = ""
        Me.Controls("A" & i).Value = ""
        Me.Controls("note" & i).Value = ""
        Me.Controls("Feeder" & i).Value = ""
      
End If


End Sub

Private Sub UserForm_Initialize()
'initial to add value when open user form
On Error Resume Next
Dim FilePath1 As String
Dim lRow As Long
Dim i As Integer
Dim j As Integer
Dim startX As Integer
Dim endX As Integer
Dim numberEq As Integer

j = 1

FilePath1 = ActiveWorkbook.name
lRow = Workbooks(FilePath1).Worksheets("DataBase1").Cells(Rows.count, 10).End(xlUp).row
startX = Workbooks(FilePath1).Worksheets("main").Range("F2").Value
endX = Workbooks(FilePath1).Worksheets("main").Range("F3").Value
numberEq = endX - startX + 1


TextBox1.Value = Workbooks(FilePath1).Worksheets("main").Range("C2").Value
TextBox2.Value = Workbooks(FilePath1).Worksheets("main").Range("C5").Value
TextBox3.Value = Workbooks(FilePath1).Worksheets("main").Range("C3").Value
TextBox4.Value = Workbooks(FilePath1).Worksheets("main").Range("C4").Value
TextBox5.Value = numberEq
TextBox1.Enabled = False
TextBox2.Enabled = False
TextBox3.Enabled = False
TextBox4.Enabled = False
CommandButton2.Enabled = False



For i = startX To endX

        Me.Controls("code" & j).Value = Workbooks(FilePath1).Worksheets("DataBase1").Range("L" & i).Value
        Me.Controls("MFR" & j).Value = Workbooks(FilePath1).Worksheets("DataBase1").Range("M" & i).Value
        Me.Controls("Type" & j).Value = Workbooks(FilePath1).Worksheets("DataBase1").Range("N" & i).Value
        Me.Controls("Serial" & j).Value = Workbooks(FilePath1).Worksheets("DataBase1").Range("O" & i).Value
        Me.Controls("kV" & j).Value = Workbooks(FilePath1).Worksheets("DataBase1").Range("P" & i).Value
        Me.Controls("kA" & j).Value = Workbooks(FilePath1).Worksheets("DataBase1").Range("Q" & i).Value
        Me.Controls("A" & j).Value = Workbooks(FilePath1).Worksheets("DataBase1").Range("R" & i).Value
        Me.Controls("note" & j).Value = Workbooks(FilePath1).Worksheets("DataBase1").Range("S" & i).Value
        Me.Controls("Feeder" & j).Value = Workbooks(FilePath1).Worksheets("DataBase1").Range("T" & i).Value
            j = j + 1
           
    Next i

End Sub
