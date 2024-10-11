Attribute VB_Name = "Module1"
Sub fetch()

' Workbooks.Open Workbooks(FilePath1).Worksheets("main").Range("C7").Value
' lRow = Workbooks(FilePath2).Worksheets(FilePath3).Cells(Rows.count, 1).End(xlUp).row
' Workbooks(FilePath2).Close
For i = 1 To 38 Step 1
    path = "C:\Users\HP\Desktop\CB22\รายงานผล switchGear สถานี " & Worksheets("substation").Range("A" & i) &" 2024.xlsx"
    Workbooks.Open path
    FilePath = ActiveWorkbook.name
    lastrow = Workbooks(FilePath).Worksheets("sheet1").cellS(Rows.Count, 1).End(xlUp).Row
    For j = 7 To lastrow Step 1
        lastrow1 = Workbooks("data_template.xlsm").Worksheets("measurement").cellS(Rows.Count, 1).End(xlUp).Row
        Workbooks("data_template.xlsm").Worksheets("measurement").Range("A" & lastrow1+1).Value = Workbooks("data_template.xlsm").Worksheets("substation").Range("A" & i)
        Workbooks("data_template.xlsm").Worksheets("measurement").Range("B" & lastrow1+1).Value = Workbooks("data_template.xlsm").Worksheets("substation").Range("B" & i)
        ' Feeder
        Workbooks("data_template.xlsm").Worksheets("measurement").Range("C" & lastrow1+1).Value = Workbooks(FilePath).Worksheets("sheet1").Range("A" & j)
        ' Contact resistace 
        Workbooks("data_template.xlsm").Worksheets("measurement").Range("E" & lastrow1+1).Value = Workbooks(FilePath).Worksheets("sheet1").Range("K" & j)
        Workbooks("data_template.xlsm").Worksheets("measurement").Range("F" & lastrow1+1).Value = Workbooks(FilePath).Worksheets("sheet1").Range("L" & j)
        Workbooks("data_template.xlsm").Worksheets("measurement").Range("G" & lastrow1+1).Value = Workbooks(FilePath).Worksheets("sheet1").Range("M" & j)
        ' Insulation
        Workbooks("data_template.xlsm").Worksheets("measurement").Range("H" & lastrow1+1).Value = Workbooks(FilePath).Worksheets("sheet1").Range("N" & j)
        Workbooks("data_template.xlsm").Worksheets("measurement").Range("I" & lastrow1+1).Value = Workbooks(FilePath).Worksheets("sheet1").Range("O" & j)
        Workbooks("data_template.xlsm").Worksheets("measurement").Range("J" & lastrow1+1).Value = Workbooks(FilePath).Worksheets("sheet1").Range("P" & j)
        Workbooks("data_template.xlsm").Worksheets("measurement").Range("K" & lastrow1+1).Value = Workbooks(FilePath).Worksheets("sheet1").Range("Q" & j)
        Workbooks("data_template.xlsm").Worksheets("measurement").Range("L" & lastrow1+1).Value = Workbooks(FilePath).Worksheets("sheet1").Range("R" & j)
        Workbooks("data_template.xlsm").Worksheets("measurement").Range("M" & lastrow1+1).Value = Workbooks(FilePath).Worksheets("sheet1").Range("S" & j)
        ' Vaccum
        Workbooks("data_template.xlsm").Worksheets("measurement").Range("N" & lastrow1+1).Value = Workbooks(FilePath).Worksheets("sheet1").Range("H" & j)
        Workbooks("data_template.xlsm").Worksheets("measurement").Range("O" & lastrow1+1).Value = Workbooks(FilePath).Worksheets("sheet1").Range("I" & j)
        Workbooks("data_template.xlsm").Worksheets("measurement").Range("P" & lastrow1+1).Value = Workbooks(FilePath).Worksheets("sheet1").Range("J" & j)
        ' Counter
        Workbooks("data_template.xlsm").Worksheets("measurement").Range("Y" & lastrow1+1).Value = Workbooks(FilePath).Worksheets("sheet1").Range("T" & j)
       
       
        On Error Resume Next
    Next j    
    Workbooks(FilePath).Close
Next i
End Sub
' ---------------------------------------------- Function2------------------------------------
Sub fetch2()
Dim counterx AS Long
 path = "C:\Users\HP\Desktop\File_Control25_Test6 - Copy.xlsm"
    Workbooks.Open path
    FilePath = ActiveWorkbook.name

For i = 1 To 38 Step 1
   lastrow = Workbooks("data_template.xlsm").Worksheets("measurement").cellS(Rows.Count, 1).End(xlUp).Row
   counterx = 0 
   For j =  7 To 24 Step 1   
        For k = 1 To lastrow Step 1        
            if (Workbooks(FilePath).Worksheets(Workbooks("data_template.xlsm").Worksheets("substation").Range("B" & i).Value).Range("A" & j).value =  Workbooks("data_template.xlsm").Worksheets("measurement").Range("C" & k).Value) AND (Workbooks(FilePath).Worksheets(Workbooks("data_template.xlsm").Worksheets("substation").Range("B" & i).Value).Range("A" & j).value <> "") AND (Workbooks("data_template.xlsm").Worksheets("measurement").Range("C" & k).Value <> " ") Then
                ' close 
                Workbooks("data_template.xlsm").Worksheets("measurement").Range("Q" & k).Value = Workbooks(FilePath).Worksheets(Workbooks("data_template.xlsm").Worksheets("substation").Range("B" & i).Value).Range("Q" & j).value
                Workbooks("data_template.xlsm").Worksheets("measurement").Range("R" & k).Value = Workbooks(FilePath).Worksheets(Workbooks("data_template.xlsm").Worksheets("substation").Range("B" & i).Value).Range("R" & j).value
                Workbooks("data_template.xlsm").Worksheets("measurement").Range("S" & k).Value = Workbooks(FilePath).Worksheets(Workbooks("data_template.xlsm").Worksheets("substation").Range("B" & i).Value).Range("S" & j).value
                '  Open
                Workbooks("data_template.xlsm").Worksheets("measurement").Range("T" & k).Value = Workbooks(FilePath).Worksheets(Workbooks("data_template.xlsm").Worksheets("substation").Range("B" & i).Value).Range("K" & j).value
                Workbooks("data_template.xlsm").Worksheets("measurement").Range("U" & k).Value = Workbooks(FilePath).Worksheets(Workbooks("data_template.xlsm").Worksheets("substation").Range("B" & i).Value).Range("L" & j).value
                Workbooks("data_template.xlsm").Worksheets("measurement").Range("V" & k).Value = Workbooks(FilePath).Worksheets(Workbooks("data_template.xlsm").Worksheets("substation").Range("B" & i).Value).Range("M" & j).value
                ' coil close Current
                Workbooks("data_template.xlsm").Worksheets("measurement").Range("W" & k).Value = Workbooks(FilePath).Worksheets(Workbooks("data_template.xlsm").Worksheets("substation").Range("B" & i).Value).Range("W" & j).value
                ' coil Trip Current
                Workbooks("data_template.xlsm").Worksheets("measurement").Range("X" & k).Value = Workbooks(FilePath).Worksheets(Workbooks("data_template.xlsm").Worksheets("substation").Range("B" & i).Value).Range("V" & j).value
                counterx = counterx + 1                
            End If  
            On Error Resume Next
        Next k 
    Next j
    If counterx = 0 Then    
            For o = 7 To 24 Step 1
                if Workbooks(FilePath).Worksheets(Worksheets("substation").Range("B" & i).value).Range("A" & o).value <> "" Then
                    lastrowXX = Workbooks("data_template.xlsm").Worksheets("measurement").cellS(Rows.Count, 1).End(xlUp).Row
                    Workbooks("data_template.xlsm").Worksheets("measurement").Range("A" & lastrowXX+1).Value = Workbooks("data_template.xlsm").Worksheets("substation").Range("A" & i).value
                    Workbooks("data_template.xlsm").Worksheets("measurement").Range("B" & lastrowXX+1).Value = Workbooks("data_template.xlsm").Worksheets("substation").Range("B" & i).value
                    ' Feeder
                    Workbooks("data_template.xlsm").Worksheets("measurement").Range("C" & lastrowXX+1).Value = Workbooks(FilePath).Worksheets(Workbooks("data_template.xlsm").Worksheets("substation").Range("B" & i).Value).Range("A" & o).value

                    Workbooks("data_template.xlsm").Worksheets("measurement").Range("Q" & lastrowXX+1).Value = Workbooks(FilePath).Worksheets(Workbooks("data_template.xlsm").Worksheets("substation").Range("B" & i).Value).Range("Q" & o).value
                    Workbooks("data_template.xlsm").Worksheets("measurement").Range("R" & lastrowXX+1).Value = Workbooks(FilePath).Worksheets(Workbooks("data_template.xlsm").Worksheets("substation").Range("B" & i).Value).Range("R" & o).value
                    Workbooks("data_template.xlsm").Worksheets("measurement").Range("S" & lastrowXX+1).Value = Workbooks(FilePath).Worksheets(Workbooks("data_template.xlsm").Worksheets("substation").Range("B" & i).Value).Range("S" & o).value
                    
                    Workbooks("data_template.xlsm").Worksheets("measurement").Range("T" & lastrowXX+1).Value = Workbooks(FilePath).Worksheets(Workbooks("data_template.xlsm").Worksheets("substation").Range("B" & i).Value).Range("K" & o).value
                    Workbooks("data_template.xlsm").Worksheets("measurement").Range("U" & lastrowXX+1).Value = Workbooks(FilePath).Worksheets(Workbooks("data_template.xlsm").Worksheets("substation").Range("B" & i).Value).Range("L" & o).value
                    Workbooks("data_template.xlsm").Worksheets("measurement").Range("V" & lastrowXX+1).Value = Workbooks(FilePath).Worksheets(Workbooks("data_template.xlsm").Worksheets("substation").Range("B" & i).Value).Range("M" & o).value
                    ' coil close Current
                    Workbooks("data_template.xlsm").Worksheets("measurement").Range("W" & lastrowXX+1).Value = Workbooks(FilePath).Worksheets(Workbooks("data_template.xlsm").Worksheets("substation").Range("B" & i).Value).Range("W" & o).value
                    ' coil Trip Current
                    Workbooks("data_template.xlsm").Worksheets("measurement").Range("X" & lastrowXX+1).Value =Workbooks(FilePath).Worksheets(Workbooks("data_template.xlsm").Worksheets("substation").Range("B" & i).Value).Range("V" & o).value
                end if
            Next o
        On Error Resume Next
    End If 
    On Error Resume Next
Next i
Workbooks(FilePath).Close
End Sub

