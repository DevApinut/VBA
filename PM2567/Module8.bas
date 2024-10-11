Attribute VB_Name = "Module8"
Sub redata()
'Finds the last non-blank cell in a single row or column
On Error Resume Next


Dim name As String
Dim nameTh As String
Dim i      As Integer
Dim j      As Integer
Dim P      As Integer
Dim startX As Integer
Dim endX As Integer
Dim lRow As Integer
Dim Feeder As String
Dim FilePath As String

ActiveWorkbook.RefreshAll
FilePath = ActiveWorkbook.name
name = Worksheets("main").Range("D3").Value
nameTh = Worksheets("main").Range("C2").Value
lRow = Workbooks(FilePath).Worksheets("Data_online").Cells(Rows.count, 1).End(xlUp).row
startX = Workbooks(FilePath).Worksheets("main").Range("F2").Value
endX = Workbooks(FilePath).Worksheets("main").Range("F3").Value


For i = 2 To lRow
'MsgBox 1 & 2
'MsgBox Workbooks(FilePath).Worksheets("main").Range("C2").Value & Workbooks(FilePath).Worksheets("Data_online").Range("D" & i).Value

If Workbooks(FilePath).Worksheets("main").Range("C2").Value = Workbooks(FilePath).Worksheets("Data_online").Range("D" & i).Value Then

'MsgBox "1"

For j = startX To endX
  
  If Workbooks(FilePath).Worksheets("DataBase1").Range("T" & j).Value = Workbooks(FilePath).Worksheets("Data_online").Range("E" & i).Value Then
     Feeder = Workbooks(FilePath).Worksheets("DataBase1").Range("L" & j).Value
     
     For P = 7 To 24
      If Workbooks(FilePath).Worksheets(name).Range("A" & P).Value = Feeder Then
        'Contact
        
        If Workbooks(FilePath).Worksheets("Data_online").Range("P" & i).Value <> "" Then
        Workbooks(FilePath).Worksheets(name).Range("H" & P).Value = Workbooks(FilePath).Worksheets("Data_online").Range("P" & i).Value
        End If
        
        If Workbooks(FilePath).Worksheets("Data_online").Range("Q" & i).Value <> "" Then
        Workbooks(FilePath).Worksheets(name).Range("I" & P).Value = Workbooks(FilePath).Worksheets("Data_online").Range("Q" & i).Value
        End If
        
        If Workbooks(FilePath).Worksheets("Data_online").Range("R" & i).Value <> "" Then
        Workbooks(FilePath).Worksheets(name).Range("J" & P).Value = Workbooks(FilePath).Worksheets("Data_online").Range("R" & i).Value
        End If
        
        'vaccum
        
        If Workbooks(FilePath).Worksheets("Data_online").Range("F" & i).Value <> "" Then
        Workbooks(FilePath).Worksheets(name).Range("Q" & P).Value = Workbooks(FilePath).Worksheets("Data_online").Range("F" & i).Value
        End If
        
        If Workbooks(FilePath).Worksheets("Data_online").Range("G" & i).Value <> "" Then
        Workbooks(FilePath).Worksheets(name).Range("R" & P).Value = Workbooks(FilePath).Worksheets("Data_online").Range("G" & i).Value
        End If
        
        If Workbooks(FilePath).Worksheets("Data_online").Range("H" & i).Value <> "" Then
        Workbooks(FilePath).Worksheets(name).Range("S" & P).Value = Workbooks(FilePath).Worksheets("Data_online").Range("H" & i).Value
        End If
        
        
        
        'insulation
        If Workbooks(FilePath).Worksheets("Data_online").Range("I" & i).Value <> "" Then
        Workbooks(FilePath).Worksheets(name).Range("K" & P).Value = Workbooks(FilePath).Worksheets("Data_online").Range("I" & i).Value
        End If
        
        If Workbooks(FilePath).Worksheets("Data_online").Range("J" & i).Value <> "" Then
        Workbooks(FilePath).Worksheets(name).Range("L" & P).Value = Workbooks(FilePath).Worksheets("Data_online").Range("J" & i).Value
        End If
        
        If Workbooks(FilePath).Worksheets("Data_online").Range("K" & i).Value <> "" Then
        Workbooks(FilePath).Worksheets(name).Range("M" & P).Value = Workbooks(FilePath).Worksheets("Data_online").Range("K" & i).Value
        End If
        
        If Workbooks(FilePath).Worksheets("Data_online").Range("L" & i).Value <> "" Then
        Workbooks(FilePath).Worksheets(name).Range("N" & P).Value = Workbooks(FilePath).Worksheets("Data_online").Range("L" & i).Value
        End If
        
        If Workbooks(FilePath).Worksheets("Data_online").Range("M" & i).Value <> "" Then
        Workbooks(FilePath).Worksheets(name).Range("O" & P).Value = Workbooks(FilePath).Worksheets("Data_online").Range("M" & i).Value
        End If
        
        If Workbooks(FilePath).Worksheets("Data_online").Range("N" & i).Value <> "" Then
        Workbooks(FilePath).Worksheets(name).Range("P" & P).Value = Workbooks(FilePath).Worksheets("Data_online").Range("N" & i).Value
        End If
        
        '--------------------------------------------------------Counter--------------------------------------------------------------------------------------------
        If Workbooks(FilePath).Worksheets("Data_online").Range("S" & i).Value <> "" Then
        Workbooks(FilePath).Worksheets(name).Range("Z" & P).Value = Workbooks(FilePath).Worksheets("Data_online").Range("S" & i).Value
        End If
        
      End If
    Next P
  End If
 Next j
 
 End If
 
 Next i

 
End Sub


