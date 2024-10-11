Attribute VB_Name = "Module4"
Sub Get_Data_From_File()
    Application.GetOpenFilename
End Sub

Sub SelectFolder()
On Error Resume Next

'PURPOSE: Have User Select a Folder Path and Store it to a variable
'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault
'https://www.thespreadsheetguru.com/vba/vba-code-to-select-folder-path
Dim FilePath As String
Dim FldrPicker As FileDialog
Dim myFolder As String


FilePath = ActiveWorkbook.name

'Have User Select Folder to Save to with Dialog Box
  Set FldrPicker = Application.FileDialog(msoFileDialogFolderPicker)

  With FldrPicker
    .Title = "Select A Target Folder"
    .AllowMultiSelect = False
    If .Show <> -1 Then Exit Sub 'Check if user clicked cancel button
    myFolder = .SelectedItems(1) & "\"
  End With
  
'Carry out rest of your code here....
'MsgBox "Folder Path is: " & myFolder
Workbooks(FilePath).Worksheets("main").Range("C6").Value = myFolder

End Sub


Sub SelectFolder1()
On Error Resume Next

'PURPOSE: Have User Select a Folder Path and Store it to a variable
'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault

Dim FilePath As String
Dim FldrPicker As FileDialog
Dim my__Folder As String


FilePath = ActiveWorkbook.name

'Have User Select Folder to Save to with Dialog Box
  Set FldrPicker = Application.FileDialog(msoFileDialogFolderPicker)

  With FldrPicker
    .Title = "Select A Target Folder"
    .AllowMultiSelect = False
    If .Show <> -1 Then Exit Sub 'Check if user clicked cancel button
    my__Folder = .SelectedItems(1) & "\"
  End With
  
'Carry out rest of your code here....
'MsgBox "Folder Path is: " & myFolder
Workbooks(FilePath).Worksheets("main").Range("C8").Value = my__Folder



End Sub

Sub File_Path()
On Error Resume Next
'https://www.exceldemy.com/excel-vba-browse-for-file-path/
Dim File_Picker As FileDialog
Dim my_path As String
Dim FilePath As String

FilePath = ActiveWorkbook.name

Set File_Picker = Application.FileDialog(msoFileDialogFilePicker)
File_Picker.Title = "Select a File" & FileType
File_Picker.Filters.Clear
File_Picker.Show
If File_Picker.SelectedItems.count = 1 Then
my_path = File_Picker.SelectedItems(1)
End If

'ActiveSheet.Range("C4").Value = my_path
'MsgBox my_path
Workbooks(FilePath).Worksheets("main").Range("C7").Value = my_path
'SendKeys "Ctrl{F2}"

End Sub
Sub File_Path1()
On Error Resume Next
'https://www.exceldemy.com/excel-vba-browse-for-file-path/
Dim File_Picker As FileDialog
Dim my_path As String
Dim FilePath As String

FilePath = ActiveWorkbook.name

Set File_Picker = Application.FileDialog(msoFileDialogFilePicker)
File_Picker.Title = "Select a File" & FileType
File_Picker.Filters.Clear
File_Picker.Show
If File_Picker.SelectedItems.count = 1 Then
my_path = File_Picker.SelectedItems(1)
End If

'ActiveSheet.Range("C4").Value = my_path
'MsgBox my_path
Workbooks(FilePath).Worksheets("main").Range("C9").Value = my_path
'SendKeys "Ctrl{F2}"

End Sub

