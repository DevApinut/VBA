Attribute VB_Name = "Module3"
Sub Savepdf_exit()

Dim FilePath1 As String
Dim name As String
Dim saveXe As String
On Error Resume Next

FilePath1 = ActiveWorkbook.name
name = Worksheets("main").Range("D3").Value

saveXe = Workbooks(FilePath1).Worksheets("main").Range("C8").Value
'Workbooks(FilePath1).Worksheets(name & " Batt125").Activate
   
   
   Workbooks(FilePath1).Worksheets(name).Activate
   ActiveSheet.UsedRange.Select
   Workbooks(FilePath1).Worksheets(name & " Batt125").Activate
   ActiveSheet.UsedRange.Select
   
   
   
     ThisWorkbook.Sheets(Array(name, name & " Batt125")).Select
   Selection.ExportAsFixedFormat Type:=xlTypePDF, FileName:=saveXe & name & ".pdf", Quality:=xlQualityStandard, _
      IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:= _
      True
End Sub

Sub excel1()


Dim name As String
Dim FilePath1 As String
Dim saveXe As String


On Error Resume Next

Application.ScreenUpdating = False
name = Worksheets("main").Range("D3").Value
FilePath1 = ActiveWorkbook.name
 
saveXe = Workbooks(FilePath1).Worksheets("main").Range("C8").Value

  Sheets(name).Copy
  Set WB = ActiveWorkbook
  With WB
    .SaveAs saveXe & name & ".xlsx"
    '.Close False
  End With
   
   Workbooks.Open savee & name & ".xlsx"
   
   Workbooks(FilePath1).Worksheets(name & " Batt125").Copy After:=Workbooks(name).Worksheets(name)
     
  
    Workbooks(name).Worksheets(name).Shapes("rectang").delete
    Workbooks(name).Worksheets(name).Shapes("refresh").delete
    Workbooks(name).Worksheets(name).Shapes("excel").delete
    Workbooks(name).Worksheets(name).Shapes("pdf").delete
    
    
    
   Workbooks(name).Save
   'Workbooks(name).Close
   
   
End Sub
Sub delete()
Dim name As String
Dim FilePath1 As String

On Error Resume Next

Application.ScreenUpdating = False
Application.DisplayAlerts = False

FilePath1 = ActiveWorkbook.name
name = Worksheets("main").Range("D3").Value

Workbooks(FilePath1).Worksheets(name).delete
Workbooks(FilePath1).Worksheets(name & " Batt125").delete
Workbooks(FilePath1).Worksheets(name & " Batt48").delete


End Sub
