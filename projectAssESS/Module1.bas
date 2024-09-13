Attribute VB_Name = "Module1"

Sub replacetext()

    Dim wdapp As Object
    Dim wddoc As Object
    Dim Path As String
    Set wdapp = CreateObject("Word.application")
    wdapp.Visible = True
    Path = "E:\VBA_VSCODE\projectAssESS\test.docx"

    Set wddoc = wdapp.Documents.Open(Path)

    Worksheets("sheet1").Range("A1:C3").Select    
    Selection.copy

    ' Worksheets("sheet1").Range("A4:C6").Select
    ' Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks :=False, Transpose:=False


    ' With wddoc.Content.Find
    '     .Text = "{Test1}"
    '     .Range.PasteExcelTable LinkedToExcel:=False, WordFormatting:=True, RTF:=False     
    '     '.Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks :=False, Transpose:=False      
    '     .Execute Replace:=2
    ' End With
    
    wdDoc.tables(1).Select
    wdDoc.tables(1).Range.PasteExcelTable LinkedToExcel:=False, WordFormatting:=True, RTF:=False
    wdDoc.tables(1).Rows.Last.Cells.Delete    
    wdDoc.tables(2).Rows.Alignment = wdAlignRowCenter



End Sub