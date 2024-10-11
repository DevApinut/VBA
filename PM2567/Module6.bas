Attribute VB_Name = "Module6"
Sub Printpdf()
    Dim path As String
    Dim myArr As Variant, a As Variant
    Dim rngArr As Variant
    Dim Ws As Worksheet
    Dim formName As String
    Dim i As Integer
    Dim FilePath1 As String
    Dim saveXe As String
    On Error Resume Next
    Application.ScreenUpdating = False

    'https://stackoverflow.com/questions/47099711/saving-multiple-ranges-on-two-different-sheets-to-pdf-using-vba
    FilePath1 = ActiveWorkbook.name
    name = Worksheets("main").Range("D3").Value
    saveXe = Workbooks(FilePath1).Worksheets("main").Range("C8").Value

    formName = saveXe & name & ".pdf"

    myArr = Array(name & " Batt125", name) '<~~ Sheet name
    rngArr = Array("A1:L46", "A1:AA37") '<~~ print area address

    For i = 0 To UBound(myArr)
        Set Ws = Sheets(myArr(i))
        With Ws
            .PageSetup.PrintArea = .Range(rngArr(i)).Address
        End With
    Next i
    Sheets(myArr).Select

    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:= _
        formName, Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:= _
        True


End Sub
