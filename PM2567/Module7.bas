Attribute VB_Name = "Module7"
Sub create_batt48()
'Finds the last non-blank cell in a single row or column
On Error Resume Next
Dim name As String
Dim FilePath1 As String
Dim j As Integer

    Application.ScreenUpdating = False

    
    FilePath1 = ActiveWorkbook.name
    name = Worksheets("main").Range("D3").Value

   
   Sheets("TemplateBatt").Copy before:=Sheets("main")
   ActiveSheet.name = name & " Batt48"
     
   

  
End Sub


    



