Attribute VB_Name = "Module2"
Sub create_batt()
'Finds the last non-blank cell in a single row or column
On Error Resume Next


Dim name As String
Dim FilePath1 As String
Dim FilePath2 As String
Dim FilePath3 As String
Dim FilePath4 As String
Dim FilePath5 As String
Dim Temp As String
Dim sg As String

Dim lRow As Long
Dim lCol As Long
Dim lCol1 As Long
Dim volt As Double
Dim zb As Double
Dim rs As Double
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim l As Integer
Dim P As Integer



    Application.ScreenUpdating = False

    j = 14
    
    FilePath1 = ActiveWorkbook.name
    name = Worksheets("main").Range("D3").Value

   
   Sheets("TemplateBatt").Copy before:=Sheets("main")
   ActiveSheet.name = name & " Batt125"
   
    Workbooks.Open Workbooks(FilePath1).Worksheets("main").Range("C7").Value
    FilePath2 = ActiveWorkbook.name
    Workbooks.Open Workbooks(FilePath1).Worksheets("main").Range("C9").Value
    FilePath4 = ActiveWorkbook.name
    
    ' delete .xlxs for use name
    If InStr(FilePath2, ".") > 0 Then
     
     FilePath3 = Left(FilePath2, InStr(FilePath2, ".") - 1)
     End If
    If InStr(FilePath4, ".") > 0 Then
     
     FilePath5 = Left(FilePath4, InStr(FilePath4, ".") - 1)
     End If
    
    lRow = Workbooks(FilePath2).Worksheets(FilePath3).Cells(Rows.count, 1).End(xlUp).row
    lCol = Workbooks(FilePath2).Worksheets(FilePath3).Cells(13, Columns.count).End(xlToLeft).Column
    lRow1 = Workbooks(FilePath4).Worksheets(FilePath5).Cells(Rows.count, 1).End(xlUp).row
    
    'MsgBox FilePath1 & FilePath2 & FilePath3
    
   For i = 14 To lRow
    
    
    If i Mod 2 = 0 And j < 44 Then
    
      volt = Workbooks(FilePath2).Worksheets(FilePath3).Range("L" & i).Value
      zb = Workbooks(FilePath2).Worksheets(FilePath3).Range("E" & i).Value
      Workbooks(FilePath1).Worksheets(name & " Batt125").Range("B" & j).Value = volt
      Workbooks(FilePath1).Worksheets(name & " Batt125").Range("D" & j).Value = zb
      j = j + 1
    
    ElseIf i Mod 2 = 0 And j > 43 Then
      volt = Workbooks(FilePath2).Worksheets(FilePath3).Range("L" & i).Value
      zb = Workbooks(FilePath2).Worksheets(FilePath3).Range("E" & i).Value
      Workbooks(FilePath1).Worksheets(name & " Batt125").Range("H" & j - 30).Value = volt
      Workbooks(FilePath1).Worksheets(name & " Batt125").Range("J" & j - 30).Value = zb
      j = j + 1
        
    ElseIf i Mod 2 = 1 And j < 45 Then
      j = j - 1
      rs = Workbooks(FilePath2).Worksheets(FilePath3).Range("E" & i).Value
      Workbooks(FilePath1).Worksheets(name & " Batt125").Range("E" & j).Value = rs
      j = j + 1
    ElseIf i Mod 2 = 1 And j > 44 Then
      j = j - 1
      rs = Workbooks(FilePath2).Worksheets(FilePath3).Range("E" & i).Value
      Workbooks(FilePath1).Worksheets(name & " Batt125").Range("K" & j - 30).Value = rs
      j = j + 1
        
    End If
    
     
    Next i
        
   
    P = 2
    l = 14
    For k = 1 To lRow1
    
    Temp = Workbooks(FilePath4).Worksheets(FilePath5).Range("S" & k).Value & "." & Workbooks(FilePath4).Worksheets(FilePath5).Range("T" & k)
    sg = Workbooks(FilePath4).Worksheets(FilePath5).Range("G" & k).Value & "." & Workbooks(FilePath4).Worksheets(FilePath5).Range("H" & k)
    
    
    If Workbooks(FilePath4).Worksheets(FilePath5).Range("E" & k).Value = name Then
      
     
      If P < 32 Then
        
        Workbooks(FilePath1).Worksheets(name & " Batt125").Range("C" & P + 12).Value = sg
        Workbooks(FilePath1).Worksheets(name & " Batt125").Range("F" & P + 12).Value = Temp
        P = P + 1
      
      ElseIf P = 32 Then
        Workbooks(FilePath1).Worksheets(name & " Batt125").Range("I" & l).Value = sg
        Workbooks(FilePath1).Worksheets(name & " Batt125").Range("L" & l).Value = Temp
        l = l + 1
      ElseIf P > 32 Then
        Workbooks(FilePath1).Worksheets(name & " Batt125").Range("I" & l).Value = sg
        Workbooks(FilePath1).Worksheets(name & " Batt125").Range("L" & l).Value = Temp
        l = l + 1
      End If
    End If
    
    Next k
      
    
    
    
    
    
    Workbooks(FilePath2).Close
    Workbooks(FilePath4).Close
    
   

  
End Sub


    



