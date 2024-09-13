Attribute VB_Name = "Module1"

Sub Upselect()
  Dim LastRow, LastRow1, LastRow2, LastRow3 As Long
  Dim Sheet1, Sheet2, Sheet3, Sheet4 As Worksheet
  Dim d, s, c As Long
  Dim sumD, sumS, sumC As Long



  Set Sheet1 = Sheets("List")
  Set Sheet2 = Sheets("destroy")
  Set Sheet3 = Sheets("sell")
  Set Sheet4 = Sheets("com")
  LastRow = Sheet1.Cells.SpecialCells(xlCellTypeLastCell).Row

  Sheet2.Range("A1:ZZ100").ClearContents
  Sheet3.Range("A1:ZZ100").ClearContents
  Sheet4.Range("A1:ZZ100").ClearContents

  For i = 1 To LastRow
    If Sheet1.Range("A" & i).Value = "destroy" Then
      d = d + 1
      sumD = Sheet1.Range("F" & i).Value + sumD
      Sheet1.Range("A" & i).EntireRow.Interior.Color = vbRed
      Sheet2.Range("A" & d) = Sheet1.Range("B" & i).Value
      Sheet2.Range("B" & d) = Sheet1.Range("C" & i).Value
      Sheet2.Range("B" & d).NumberFormat = "[$-D07041E]#"
      Sheet2.Range("C" & d) = Sheet1.Range("D" & i).Value
      Sheet2.Range("E" & d) = Sheet1.Range("E" & i).Value
      Sheet2.Range("E" & d).NumberFormat = "[$-th-TH,D07]d mmmm yyyy;@"
      Sheet2.Range("D" & d) = Sheet1.Range("F" & i).Value
      Sheet2.Range("D" & d).NumberFormat = "[$-D07041E]#,###,##0.00-"
      Sheet2.Range("F" & d) = Sheet1.Range("H" & i).Value
      Sheet2.Range("F" & d).NumberFormat = "[$-th-TH,D07]#.00-"
      Sheet1.Range("A" & i).ClearContents
    Elseif Sheet1.Range("A" & i).Value = "sell" Then
      s = s + 1
      sumS = Sheet1.Range("F" & i).Value + sumS
      Sheet1.Range("A" & i).EntireRow.Interior.Color = vbYellow
      Sheet3.Range("A" & s) = Sheet1.Range("B" & i).Value
      Sheet3.Range("B" & s) = Sheet1.Range("C" & i).Value
      Sheet3.Range("B" & s).NumberFormat = "[$-D07041E]#"
      Sheet3.Range("C" & s) = Sheet1.Range("D" & i).Value
      Sheet3.Range("E" & s) = Sheet1.Range("E" & i).Value
      Sheet3.Range("E" & s).NumberFormat = "[$-th-TH,D07]d mmmm yyyy;@"
      Sheet3.Range("D" & s) = Sheet1.Range("F" & i).Value
      Sheet3.Range("D" & s).NumberFormat = "[$-D07041E]#,###,##0.00-"
      Sheet3.Range("F" & s) = Sheet1.Range("H" & i).Value
      Sheet3.Range("F" & s).NumberFormat = "[$-th-TH,D07]#.00-"
      Sheet1.Range("A" & i).ClearContents
    Elseif Sheet1.Range("A" & i).Value = "com" Then
      c = c + 1
      sumC = Sheet1.Range("F" & i).Value + sumC
      Sheet1.Range("A" & i).EntireRow.Interior.Color = vbGreen
      Sheet4.Range("A" & c) = Sheet1.Range("B" & i).Value
      Sheet4.Range("B" & c) = Sheet1.Range("C" & i).Value
      Sheet4.Range("B" & c).NumberFormat = "[$-D07041E]#"
      Sheet4.Range("C" & c) = Sheet1.Range("D" & i).Value
      Sheet4.Range("E" & c) = Sheet1.Range("E" & i).Value
      Sheet4.Range("E" & c).NumberFormat = "[$-th-TH,D07]d mmmm yyyy;@"
      Sheet4.Range("D" & c) = Sheet1.Range("F" & i).Value
      Sheet4.Range("D" & c).NumberFormat = "[$-D07041E]#,###,##0.00-"
      Sheet4.Range("F" & c) = Sheet1.Range("H" & i).Value
      Sheet4.Range("F" & c).NumberFormat = "[$-th-TH,D07]#.00-"
      Sheet1.Range("A" & i).ClearContents
    End If
  Next i

  Sheet2.Range("D" & d + 1).Value = sumD
  Sheet2.Range("D" & d + 1).NumberFormat = "[$-D07041E]#,###,##0.00-"
  Sheet3.Range("D" & s + 1).Value = sumS
  Sheet3.Range("D" & s + 1).NumberFormat = "[$-D07041E]#,###,##0.00-"
  Sheet4.Range("D" & c + 1).Value = sumC
  Sheet4.Range("D" & c + 1).NumberFormat = "[$-D07041E]#,###,##0.00-"




End Sub

Sub Upselect22()
  Application.DisplayAlerts = False
  Dim LastRowList, LastRowDropdown, LastRowsheet, LastRowsheet2 As Long
  Dim Sheet As Worksheet
  Set SheetList = Sheets("List")
  Set SheetDropdown = Sheets("Dropdown")

  LastRowList = SheetList.Cells.SpecialCells(xlCellTypeLastCell).Row
  LastRowDropdown = SheetList.Cells.SpecialCells(xlCellTypeLastCell).Row



  For j = 1 To LastRowList


    If DoesSheetExists(SheetDropdown.Range("A" & j).Value) Then
      Worksheets(SheetDropdown.Range("A" & j).Value).Delete
    Else

    End If

  Next j


  For i = 1 To LastRowList

    If SheetList.Range("A" & i).Value = "none" Or SheetList.Range("A" & i).Value = "" Then

    Else
      If DoesSheetExists(SheetList.Range("A" & i).Value) Or SheetList.Range("A" & i).Value = "none" Then

        LastRowsheet = Worksheets(SheetList.Range("A" & i).Value).Cells.SpecialCells(xlCellTypeLastCell).Row
        '------------------------------------------------------- Header -----------------------------------------------------------------------
        Worksheets(SheetList.Range("A" & i).Value).Range("A" & 1).Value = "ที่"
        Worksheets(SheetList.Range("A" & i).Value).Range("A" & 1).Font.Name = "TH SarabunIT๙"
        Worksheets(SheetList.Range("A" & i).Value).Range("A" & 1).Font.Size = 14
        Worksheets(SheetList.Range("A" & i).Value).Range("A" & 1).Font.Bold = True
        Worksheets(SheetList.Range("A" & i).Value).Range("A" & 1).HorizontalAlignment = xlCenter
        Worksheets(SheetList.Range("A" & i).Value).Range("A" & 1).VerticalAlignment = xlCenter

        Worksheets(SheetList.Range("A" & i).Value).Range("B" & 1).Value = "รหัสสินทรัพย์"
        Worksheets(SheetList.Range("A" & i).Value).Range("B" & 1).NumberFormat = "[$-D07041E]#"
        Worksheets(SheetList.Range("A" & i).Value).Range("B" & 1).Font.Name = "TH SarabunIT๙"
        Worksheets(SheetList.Range("A" & i).Value).Range("B" & 1).Font.Size = 14
        Worksheets(SheetList.Range("A" & i).Value).Range("B" & 1).Font.Bold = True
        Worksheets(SheetList.Range("A" & i).Value).Range("B" & 1).HorizontalAlignment = xlCenter
        Worksheets(SheetList.Range("A" & i).Value).Range("B" & 1).VerticalAlignment = xlCenter

        Worksheets(SheetList.Range("A" & i).Value).Range("C" & 1).Value = "รายการ"
        Worksheets(SheetList.Range("A" & i).Value).Range("C" & 1).Font.Name = "TH SarabunIT๙"
        Worksheets(SheetList.Range("A" & i).Value).Range("C" & 1).Font.Size = 14
        Worksheets(SheetList.Range("A" & i).Value).Range("C" & 1).Font.Bold = True
        Worksheets(SheetList.Range("A" & i).Value).Range("C" & 1).HorizontalAlignment = xlCenter
        Worksheets(SheetList.Range("A" & i).Value).Range("C" & 1).VerticalAlignment = xlCenter

        Worksheets(SheetList.Range("A" & i).Value).Range("E" & 1).Value = "วันที่ได้มา"
        Worksheets(SheetList.Range("A" & i).Value).Range("E" & 1).NumberFormat = "[$-th-TH,D07]d mmmm yyyy;@"
        Worksheets(SheetList.Range("A" & i).Value).Range("E" & 1).Font.Name = "TH SarabunIT๙"
        Worksheets(SheetList.Range("A" & i).Value).Range("E" & 1).Font.Size = 14
        Worksheets(SheetList.Range("A" & i).Value).Range("E" & 1).Font.Bold = True
        Worksheets(SheetList.Range("A" & i).Value).Range("E" & 1).HorizontalAlignment = xlCenter
        Worksheets(SheetList.Range("A" & i).Value).Range("E" & 1).VerticalAlignment = xlCenter

        Worksheets(SheetList.Range("A" & i).Value).Range("D" & 1).Value = "ราคาซื้อหรือได้มา (บาท)"
        Worksheets(SheetList.Range("A" & i).Value).Range("D" & 1).NumberFormat = "[$-D07041E]#,###,##0.00-"
        Worksheets(SheetList.Range("A" & i).Value).Range("D" & 1).Font.Name = "TH SarabunIT๙"
        Worksheets(SheetList.Range("A" & i).Value).Range("D" & 1).Font.Size = 14
        Worksheets(SheetList.Range("A" & i).Value).Range("D" & 1).Font.Bold = True
        Worksheets(SheetList.Range("A" & i).Value).Range("D" & 1).HorizontalAlignment = xlCenter
        Worksheets(SheetList.Range("A" & i).Value).Range("D" & 1).VerticalAlignment = xlCenter

        Worksheets(SheetList.Range("A" & i).Value).Range("F" & 1).Value = "มูลค่าคงเหลือ (บาท)"
        Worksheets(SheetList.Range("A" & i).Value).Range("F" & 1).NumberFormat = "[$-th-TH,D07]#.00-"
        Worksheets(SheetList.Range("A" & i).Value).Range("F" & 1).Font.Name = "TH SarabunIT๙"
        Worksheets(SheetList.Range("A" & i).Value).Range("F" & 1).Font.Size = 14
        Worksheets(SheetList.Range("A" & i).Value).Range("F" & 1).Font.Bold = True
        Worksheets(SheetList.Range("A" & i).Value).Range("F" & 1).HorizontalAlignment = xlCenter
        Worksheets(SheetList.Range("A" & i).Value).Range("F" & 1).VerticalAlignment = xlCenter

        '-----------------------------------------------------------------------------------------------------------------------------------------

        Worksheets(SheetList.Range("A" & i).Value).Range("A" & LastRowsheet + 1) = LastRowsheet
        Worksheets(SheetList.Range("A" & i).Value).Range("A" & LastRowsheet + 1).NumberFormat = "[$-th-TH,D07]#"
        Worksheets(SheetList.Range("A" & i).Value).Range("A" & LastRowsheet + 1).Font.Name = "TH SarabunIT๙"
        Worksheets(SheetList.Range("A" & i).Value).Range("A" & LastRowsheet + 1).Font.Size = 14
        Worksheets(SheetList.Range("A" & i).Value).Range("A" & LastRowsheet + 1).Font.Bold = False
        Worksheets(SheetList.Range("A" & i).Value).Range("A" & LastRowsheet + 1).HorizontalAlignment = xlCenter
        Worksheets(SheetList.Range("A" & i).Value).Range("A" & LastRowsheet + 1).VerticalAlignment = xlCenter


        Worksheets(SheetList.Range("A" & i).Value).Range("B" & LastRowsheet + 1) = SheetList.Range("C" & i).Value
        Worksheets(SheetList.Range("A" & i).Value).Range("B" & LastRowsheet + 1).NumberFormat = "[$-D07041E]#"
        Worksheets(SheetList.Range("A" & i).Value).Range("B" & LastRowsheet + 1).Font.Name = "TH SarabunIT๙"
        Worksheets(SheetList.Range("A" & i).Value).Range("B" & LastRowsheet + 1).Font.Size = 14
        Worksheets(SheetList.Range("A" & i).Value).Range("B" & LastRowsheet + 1).Font.Bold = False
        Worksheets(SheetList.Range("A" & i).Value).Range("B" & LastRowsheet + 1).HorizontalAlignment = xlCenter
        Worksheets(SheetList.Range("A" & i).Value).Range("B" & LastRowsheet + 1).VerticalAlignment = xlCenter


        Worksheets(SheetList.Range("A" & i).Value).Range("C" & LastRowsheet + 1) = SheetList.Range("D" & i).Value
        Worksheets(SheetList.Range("A" & i).Value).Range("C" & LastRowsheet + 1).Font.Name = "TH SarabunIT๙"
        Worksheets(SheetList.Range("A" & i).Value).Range("C" & LastRowsheet + 1).Font.Size = 14
        Worksheets(SheetList.Range("A" & i).Value).Range("C" & LastRowsheet + 1).Font.Bold = False
        Worksheets(SheetList.Range("A" & i).Value).Range("C" & LastRowsheet + 1).HorizontalAlignment = xlLeft
        Worksheets(SheetList.Range("A" & i).Value).Range("C" & LastRowsheet + 1).VerticalAlignment = xlCenter
        Worksheets(SheetList.Range("A" & i).Value).Range("C" & LastRowsheet + 1).WrapText = True


        Worksheets(SheetList.Range("A" & i).Value).Range("E" & LastRowsheet + 1) = SheetList.Range("E" & i).Value
        Worksheets(SheetList.Range("A" & i).Value).Range("E" & LastRowsheet + 1).NumberFormat = "[$-th-TH,D07]d mmmm yyyy;@"
        Worksheets(SheetList.Range("A" & i).Value).Range("E" & LastRowsheet + 1).Font.Name = "TH SarabunIT๙"
        Worksheets(SheetList.Range("A" & i).Value).Range("E" & LastRowsheet + 1).Font.Size = 14
        Worksheets(SheetList.Range("A" & i).Value).Range("E" & LastRowsheet + 1).Font.Bold = False
        Worksheets(SheetList.Range("A" & i).Value).Range("E" & LastRowsheet + 1).HorizontalAlignment = xlCenter
        Worksheets(SheetList.Range("A" & i).Value).Range("E" & LastRowsheet + 1).VerticalAlignment = xlCenter


        Worksheets(SheetList.Range("A" & i).Value).Range("D" & LastRowsheet + 1) = SheetList.Range("F" & i).Value
        Worksheets(SheetList.Range("A" & i).Value).Range("D" & LastRowsheet + 1).NumberFormat = "[$-D07041E]#,###,##0.00-"
        Worksheets(SheetList.Range("A" & i).Value).Range("D" & LastRowsheet + 1).Font.Name = "TH SarabunIT๙"
        Worksheets(SheetList.Range("A" & i).Value).Range("D" & LastRowsheet + 1).Font.Size = 14
        Worksheets(SheetList.Range("A" & i).Value).Range("D" & LastRowsheet + 1).Font.Bold = False
        Worksheets(SheetList.Range("A" & i).Value).Range("D" & LastRowsheet + 1).HorizontalAlignment = xlCenter
        Worksheets(SheetList.Range("A" & i).Value).Range("D" & LastRowsheet + 1).VerticalAlignment = xlCenter


        Worksheets(SheetList.Range("A" & i).Value).Range("F" & LastRowsheet + 1) = SheetList.Range("H" & i).Value
        Worksheets(SheetList.Range("A" & i).Value).Range("F" & LastRowsheet + 1).NumberFormat = "[$-th-TH,D07]#.00-"
        Worksheets(SheetList.Range("A" & i).Value).Range("F" & LastRowsheet + 1).Font.Name = "TH SarabunIT๙"
        Worksheets(SheetList.Range("A" & i).Value).Range("F" & LastRowsheet + 1).Font.Size = 14
        Worksheets(SheetList.Range("A" & i).Value).Range("F" & LastRowsheet + 1).Font.Bold = False
        Worksheets(SheetList.Range("A" & i).Value).Range("F" & LastRowsheet + 1).HorizontalAlignment = xlCenter
        Worksheets(SheetList.Range("A" & i).Value).Range("F" & LastRowsheet + 1).VerticalAlignment = xlCenter



        Worksheets(SheetList.Range("A" & i).Value).Columns("A").ColumnWidth = 2.5
        Worksheets(SheetList.Range("A" & i).Value).Columns("B").ColumnWidth = 18
        Worksheets(SheetList.Range("A" & i).Value).Columns("C").ColumnWidth = 18
        Worksheets(SheetList.Range("A" & i).Value).Columns("D").ColumnWidth = 13
        Worksheets(SheetList.Range("A" & i).Value).Columns("E").ColumnWidth = 18
        Worksheets(SheetList.Range("A" & i).Value).Columns("F").ColumnWidth = 10
      Else

        Sheets.Add(After:=Sheets(2)).Name = SheetList.Range("A" & i).Value
        LastRowsheet = Worksheets(SheetList.Range("A" & i).Value).Cells.SpecialCells(xlCellTypeLastCell).Row
        '------------------------------------------------------- Header -----------------------------------------------------------------------
        Worksheets(SheetList.Range("A" & i).Value).Range("A" & 1).Value = "ที่"
        Worksheets(SheetList.Range("A" & i).Value).Range("A" & 1).Font.Name = "TH SarabunIT๙"
        Worksheets(SheetList.Range("A" & i).Value).Range("A" & 1).Font.Size = 14
        Worksheets(SheetList.Range("A" & i).Value).Range("A" & 1).Font.Bold = True
        Worksheets(SheetList.Range("A" & i).Value).Range("A" & 1).HorizontalAlignment = xlCenter
        Worksheets(SheetList.Range("A" & i).Value).Range("A" & 1).VerticalAlignment = xlCenter

        Worksheets(SheetList.Range("A" & i).Value).Range("B" & 1).Value = "รหัสสินทรัพย์"
        Worksheets(SheetList.Range("A" & i).Value).Range("B" & 1).NumberFormat = "[$-D07041E]#"
        Worksheets(SheetList.Range("A" & i).Value).Range("B" & 1).Font.Name = "TH SarabunIT๙"
        Worksheets(SheetList.Range("A" & i).Value).Range("B" & 1).Font.Size = 14
        Worksheets(SheetList.Range("A" & i).Value).Range("B" & 1).Font.Bold = True
        Worksheets(SheetList.Range("A" & i).Value).Range("B" & 1).HorizontalAlignment = xlCenter
        Worksheets(SheetList.Range("A" & i).Value).Range("B" & 1).VerticalAlignment = xlCenter

        Worksheets(SheetList.Range("A" & i).Value).Range("C" & 1).Value = "รายการ"
        Worksheets(SheetList.Range("A" & i).Value).Range("C" & 1).Font.Name = "TH SarabunIT๙"
        Worksheets(SheetList.Range("A" & i).Value).Range("C" & 1).Font.Size = 14
        Worksheets(SheetList.Range("A" & i).Value).Range("C" & 1).Font.Bold = True
        Worksheets(SheetList.Range("A" & i).Value).Range("C" & 1).HorizontalAlignment = xlCenter
        Worksheets(SheetList.Range("A" & i).Value).Range("C" & 1).VerticalAlignment = xlCenter

        Worksheets(SheetList.Range("A" & i).Value).Range("E" & 1).Value = "วันที่ได้มา"
        Worksheets(SheetList.Range("A" & i).Value).Range("E" & 1).NumberFormat = "[$-th-TH,D07]d mmmm yyyy;@"
        Worksheets(SheetList.Range("A" & i).Value).Range("E" & 1).Font.Name = "TH SarabunIT๙"
        Worksheets(SheetList.Range("A" & i).Value).Range("E" & 1).Font.Size = 14
        Worksheets(SheetList.Range("A" & i).Value).Range("E" & 1).Font.Bold = True
        Worksheets(SheetList.Range("A" & i).Value).Range("E" & 1).HorizontalAlignment = xlCenter
        Worksheets(SheetList.Range("A" & i).Value).Range("E" & 1).VerticalAlignment = xlCenter

        Worksheets(SheetList.Range("A" & i).Value).Range("D" & 1).Value = "ราคาซื้อหรือได้มา (บาท)"
        Worksheets(SheetList.Range("A" & i).Value).Range("D" & 1).NumberFormat = "[$-D07041E]#,###,##0.00-"
        Worksheets(SheetList.Range("A" & i).Value).Range("D" & 1).Font.Name = "TH SarabunIT๙"
        Worksheets(SheetList.Range("A" & i).Value).Range("D" & 1).Font.Size = 14
        Worksheets(SheetList.Range("A" & i).Value).Range("D" & 1).Font.Bold = True
        Worksheets(SheetList.Range("A" & i).Value).Range("D" & 1).HorizontalAlignment = xlCenter
        Worksheets(SheetList.Range("A" & i).Value).Range("D" & 1).VerticalAlignment = xlCenter

        Worksheets(SheetList.Range("A" & i).Value).Range("F" & 1).Value = "มูลค่าคงเหลือ (บาท)"
        Worksheets(SheetList.Range("A" & i).Value).Range("F" & 1).NumberFormat = "[$-th-TH,D07]#.00-"
        Worksheets(SheetList.Range("A" & i).Value).Range("F" & 1).Font.Name = "TH SarabunIT๙"
        Worksheets(SheetList.Range("A" & i).Value).Range("F" & 1).Font.Size = 14
        Worksheets(SheetList.Range("A" & i).Value).Range("F" & 1).Font.Bold = True
        Worksheets(SheetList.Range("A" & i).Value).Range("F" & 1).HorizontalAlignment = xlCenter
        Worksheets(SheetList.Range("A" & i).Value).Range("F" & 1).VerticalAlignment = xlCenter


        '-----------------------------------------------------------------------------------------------------------------------------------------

        Worksheets(SheetList.Range("A" & i).Value).Range("A" & LastRowsheet + 1) = LastRowsheet
        Worksheets(SheetList.Range("A" & i).Value).Range("A" & LastRowsheet + 1).NumberFormat = "[$-th-TH,D07]#"
        Worksheets(SheetList.Range("A" & i).Value).Range("A" & LastRowsheet + 1).Font.Name = "TH SarabunIT๙"
        Worksheets(SheetList.Range("A" & i).Value).Range("A" & LastRowsheet + 1).Font.Size = 14
        Worksheets(SheetList.Range("A" & i).Value).Range("A" & LastRowsheet + 1).Font.Bold = False
        Worksheets(SheetList.Range("A" & i).Value).Range("A" & LastRowsheet + 1).HorizontalAlignment = xlCenter
        Worksheets(SheetList.Range("A" & i).Value).Range("A" & LastRowsheet + 1).VerticalAlignment = xlCenter


        Worksheets(SheetList.Range("A" & i).Value).Range("B" & LastRowsheet + 1) = SheetList.Range("C" & i).Value
        Worksheets(SheetList.Range("A" & i).Value).Range("B" & LastRowsheet + 1).NumberFormat = "[$-D07041E]#"
        Worksheets(SheetList.Range("A" & i).Value).Range("B" & LastRowsheet + 1).Font.Name = "TH SarabunIT๙"
        Worksheets(SheetList.Range("A" & i).Value).Range("B" & LastRowsheet + 1).Font.Size = 14
        Worksheets(SheetList.Range("A" & i).Value).Range("B" & LastRowsheet + 1).Font.Bold = False
        Worksheets(SheetList.Range("A" & i).Value).Range("B" & LastRowsheet + 1).HorizontalAlignment = xlCenter
        Worksheets(SheetList.Range("A" & i).Value).Range("B" & LastRowsheet + 1).VerticalAlignment = xlCenter


        Worksheets(SheetList.Range("A" & i).Value).Range("C" & LastRowsheet + 1) = SheetList.Range("D" & i).Value
        Worksheets(SheetList.Range("A" & i).Value).Range("C" & LastRowsheet + 1).Font.Name = "TH SarabunIT๙"
        Worksheets(SheetList.Range("A" & i).Value).Range("C" & LastRowsheet + 1).Font.Size = 14
        Worksheets(SheetList.Range("A" & i).Value).Range("C" & LastRowsheet + 1).Font.Bold = False
        Worksheets(SheetList.Range("A" & i).Value).Range("C" & LastRowsheet + 1).HorizontalAlignment = xlLeft
        Worksheets(SheetList.Range("A" & i).Value).Range("C" & LastRowsheet + 1).VerticalAlignment = xlCenter
        Worksheets(SheetList.Range("A" & i).Value).Range("C" & LastRowsheet + 1).WrapText = True


        Worksheets(SheetList.Range("A" & i).Value).Range("E" & LastRowsheet + 1) = SheetList.Range("E" & i).Value
        Worksheets(SheetList.Range("A" & i).Value).Range("E" & LastRowsheet + 1).NumberFormat = "[$-th-TH,D07]d mmmm yyyy;@"
        Worksheets(SheetList.Range("A" & i).Value).Range("E" & LastRowsheet + 1).Font.Name = "TH SarabunIT๙"
        Worksheets(SheetList.Range("A" & i).Value).Range("E" & LastRowsheet + 1).Font.Size = 14
        Worksheets(SheetList.Range("A" & i).Value).Range("E" & LastRowsheet + 1).Font.Bold = False
        Worksheets(SheetList.Range("A" & i).Value).Range("E" & LastRowsheet + 1).HorizontalAlignment = xlCenter
        Worksheets(SheetList.Range("A" & i).Value).Range("E" & LastRowsheet + 1).VerticalAlignment = xlCenter


        Worksheets(SheetList.Range("A" & i).Value).Range("D" & LastRowsheet + 1) = SheetList.Range("F" & i).Value
        Worksheets(SheetList.Range("A" & i).Value).Range("D" & LastRowsheet + 1).NumberFormat = "[$-D07041E]#,###,##0.00-"
        Worksheets(SheetList.Range("A" & i).Value).Range("D" & LastRowsheet + 1).Font.Name = "TH SarabunIT๙"
        Worksheets(SheetList.Range("A" & i).Value).Range("D" & LastRowsheet + 1).Font.Size = 14
        Worksheets(SheetList.Range("A" & i).Value).Range("D" & LastRowsheet + 1).Font.Bold = False
        Worksheets(SheetList.Range("A" & i).Value).Range("D" & LastRowsheet + 1).HorizontalAlignment = xlCenter
        Worksheets(SheetList.Range("A" & i).Value).Range("D" & LastRowsheet + 1).VerticalAlignment = xlCenter


        Worksheets(SheetList.Range("A" & i).Value).Range("F" & LastRowsheet + 1) = SheetList.Range("H" & i).Value
        Worksheets(SheetList.Range("A" & i).Value).Range("F" & LastRowsheet + 1).NumberFormat = "[$-th-TH,D07]#.00-"
        Worksheets(SheetList.Range("A" & i).Value).Range("F" & LastRowsheet + 1).Font.Name = "TH SarabunIT๙"
        Worksheets(SheetList.Range("A" & i).Value).Range("F" & LastRowsheet + 1).Font.Size = 14
        Worksheets(SheetList.Range("A" & i).Value).Range("F" & LastRowsheet + 1).Font.Bold = False
        Worksheets(SheetList.Range("A" & i).Value).Range("F" & LastRowsheet + 1).HorizontalAlignment = xlCenter
        Worksheets(SheetList.Range("A" & i).Value).Range("F" & LastRowsheet + 1).VerticalAlignment = xlCenter

        Worksheets(SheetList.Range("A" & i).Value).Columns("A").ColumnWidth = 2.5
        Worksheets(SheetList.Range("A" & i).Value).Columns("B").ColumnWidth = 18
        Worksheets(SheetList.Range("A" & i).Value).Columns("C").ColumnWidth = 18
        Worksheets(SheetList.Range("A" & i).Value).Columns("D").ColumnWidth = 13
        Worksheets(SheetList.Range("A" & i).Value).Columns("E").ColumnWidth = 18
        Worksheets(SheetList.Range("A" & i).Value).Columns("F").ColumnWidth = 10
      End If
    End If

  Next i



  For Each ws In ActiveWorkbook.Worksheets
    If ws.Name <> "List" And ws.Name <> "Dropdown" Then
      LastRowsheet1 = ws.Cells.SpecialCells(xlCellTypeLastCell).Row
      ws.Range("B2").EntireRow.AutoFit
      ws.Range("B" & LastRowsheet1 + 1).Value = "รวมราคาซื้อมาหรือได้มา (บาท)"


      ws.Range("D" & LastRowsheet1 + 1).Value = "=SUM(D2:D" & LastRowsheet1 & ")"
      ws.Range("D" & LastRowsheet1 + 1).NumberFormat = "[$-D07041E]#,###,##0.00-"
      ws.Range("D" & LastRowsheet1 + 2).Value = "=SUM(D2:D" & LastRowsheet1 & ")"
      ws.Range("D" & LastRowsheet1 + 2).NumberFormat = "#,##0.00_);(#,##0.00)"
      ws.Range("D" & LastRowsheet1 + 1).Font.Name = "TH SarabunIT๙"
      ws.Range("D" & LastRowsheet1 + 1).Font.Size = 14
      ws.Range("D" & LastRowsheet1 + 1).HorizontalAlignment = xlCenter


      ws.Range("B" & LastRowsheet1 + 1 & ":C" & LastRowsheet1 + 1).Merge

      ws.Range("B" & LastRowsheet1 + 1).Font.Name = "TH SarabunIT๙"
      ws.Range("B" & LastRowsheet1 + 1).Font.Size = 14
      ws.Range("B" & LastRowsheet1 + 1).HorizontalAlignment = xlCenter


      ws.Range("A1" & ":F" & LastRowsheet1 + 1).Borders.LineStyle = xlContinuous
      ws.Range("E" & LastRowsheet1 + 1 & ":F" & LastRowsheet1 + 1).Borders.LineStyle = none

    End If
    Next

End Sub
'------------ Check sheet -----------------------------

Function DoesSheetExists(sh As String) As Boolean
  Dim ws As Worksheet

  On Error Resume Next
  Set ws = ThisWorkbook.Sheets(sh)
  On Error Goto 0

    If Not ws Is Nothing Then DoesSheetExists = True
End Function



Sub Macro1()
  Dim Sht As Worksheet
  Dim WordApp As Object
  Dim WordDoc As Object
  Dim xName, xage, xzip, i As Long
  Dim LastRowsheet1 As Long
  Dim filePath As String, path_dest As String
  Set Sht = ActiveWorkbook.Sheets(1)
  Set WordApp = CreateObject("Word.Application")
  WordApp.Visible = True  

  For Each ws In ActiveWorkbook.Worksheets
    If ws.Name <> "List" And ws.Name <> "Dropdown" Then
      LastRowsheet1 = ws.Cells.SpecialCells(xlCellTypeLastCell).Row
      Set Sht = ActiveWorkbook.Sheets(1)
      Set WordApp = CreateObject("Word.Application")
      WordApp.Visible = True
      filePath = Worksheets("List").Range("L" & 3).Value
      Set WordDoc = WordApp.Documents.Open(filePath)
      path_dest = Worksheets("List").Range("L" & 5).Value & ws.Name & ".docx"
      ws.Range("A1:F" & LastRowsheet1 - 1).Copy      
      WordDoc.Paragraphs(7).Range.PasteExcelTable Linkedtoexcel:=False, wordformatting:=False, RTF:=False
      WordDoc.SaveAs path_dest
      WordDoc.Close
      WordApp.Quit
      Set WordDoc = Nothing
      Set WordApp = Nothing

    End If
    Next

End Sub


' -------------------- For Select File Path ---------------------------------
Sub SelectFile()

  Dim DialogBox As FileDialog

  Dim path As String

  Set DialogBox = Application.FileDialog(msoFileDialogFilePicker)

  DialogBox.Title = "Select file For " & FileType

  DialogBox.Filters.Clear

  DialogBox.Show

  If DialogBox.SelectedItems.Count = 1 Then

    path = DialogBox.SelectedItems(1)

  End If

  Worksheets("List").Range("L3").Value = path

End Sub


' -------------------- For Select Folder Path ---------------------------------
Sub SelectFolder()
  'PURPOSE: Have User Select a Folder Path And Store it To a variable
  'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault

  Dim FldrPicker As FileDialog
  Dim myFolder As String

  'Have User Select Folder To Save To With Dialog Box
  Set FldrPicker = Application.FileDialog(msoFileDialogFolderPicker)

  With FldrPicker
    .Title = "Select A Target Folder"
    .AllowMultiSelect = False
    If .Show <> -1 Then Exit Sub 'Check If user clicked cancel button
      myFolder = .SelectedItems(1) & "\"
    End With

    'Carry out rest of your code here....
    MsgBox "Folder Path is: " & myFolder
    Worksheets("List").Range("L5").Value = myFolder


End Sub
Sub replaceText()
  Dim Sht As Worksheet
  Dim WordApp As Object
  Dim WordDoc As Object
  Dim xName, xage, xzip, i As Long
  Dim LastRowsheet1 As Long
  Dim filePath As String, path_dest As String
  Set Sht = ActiveWorkbook.Sheets(1)
  Set WordApp = CreateObject("Word.Application")
  WordApp.Visible = True
  filePath = Worksheets("List").Range("L" & 3).Value
  Set WordDoc = WordApp.Documents.Open(filePath)
  LastRowsheet1 = Worksheets("List14").Cells.SpecialCells(xlCellTypeLastCell).Row

  Worksheets("List14").Range("A1:F5").Copy
  
  With WordDoc.Range.Find
    .Text = "{Test1}"
    .Replacement.Text = "TTTT"
    .Execute Replace:=2
    .Found.Select
  End With


End Sub


Sub replacetext2()

  Dim wdapp As Object
  Dim wddoc As Object
  Dim Path As String
  Set wdapp = CreateObject("Word.application")
  wdapp.Visible = True
  Path = "E:\VBA_VSCODE\projectAssESS\test.docx"

  Set wddoc = wdapp.Documents.Open(Path)

  Worksheets("sheet1").Range("A1:C3").Select    
  Selection.copy

  wdDoc.tables(1).Select
  wdDoc.tables(1).Range.PasteExcelTable LinkedToExcel:=False, WordFormatting:=True, RTF:=False
  wdDoc.tables(1).Rows.Last.Cells.Delete
  
  ' สำหรับคำสั่งนี้ต้องการเปิดเืพื่อจะทำให้ใช้งานได้ https://stackoverflow.com/questions/56668383/having-troubles-creating-and-formatting-word-tables-from-excel-vba 
  wdDoc.tables(2).Rows.Alignment = wdAlignRowCenter



End Sub