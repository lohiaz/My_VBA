Sub LoopAllExcelFilesInFolder()
'PURPOSE: To loop through all Excel files in a user specified folder and perform a set task on them

Dim wb As Workbook
Dim myPath As String
Dim myFile As String
Dim myExtension As String
Dim FldrPicker As FileDialog
Dim lastRow As Long
Dim sht1 As Worksheet
Dim NewWb As Workbook

Set NewWb = Workbooks.Add
    ChDir "C:\Users\plohia\Desktop"
    ActiveWorkbook.SaveAs Filename:="C:\Users\plohia\Desktop\ExcelCombined.xlsx", _
        FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False

'Optimize Macro Speed
  Application.ScreenUpdating = False
  Application.EnableEvents = False
  Application.Calculation = xlCalculationManual

'Retrieve Target Folder Path From User
  Set FldrPicker = Application.FileDialog(msoFileDialogFolderPicker)

    With FldrPicker
      .Title = "Select A Target Folder"
      .AllowMultiSelect = False
        If .Show <> -1 Then GoTo NextCode
        myPath = .SelectedItems(1) & "\"
    End With

'In Case of Cancel
NextCode:
  myPath = myPath
  If myPath = "" Then GoTo ResetSettings

'Target File Extension (must include wildcard "*")
  myExtension = "*.xlsx*"

'Target Path with Ending Extention
  myFile = Dir(myPath & myExtension)

'Loop through each Excel file in folder
  Do While myFile <> ""
    'Set variable equal to opened workbook
      Set wb = Workbooks.Open(Filename:=myPath & myFile)
    
    'Ensure Workbook has opened before moving on to next line of code
      DoEvents
    
    'Find Last row on Sheet 1
    lastRow = Cells(wb.Worksheets(1).Rows.Count, "A").End(xlUp).Row
    
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    NewWb.Activate
    Range("A1").Select
    Selection.End(xlDown).Select
    Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    
      
    'Save and Close Workbook
      wb.Close SaveChanges:=True
      
    'Ensure Workbook has closed before moving on to next line of code
      DoEvents

    'Get next file name
      myFile = Dir
  Loop

'Message Box when tasks are completed
  MsgBox "Task Complete!"

ResetSettings:
  'Reset Macro Optimization Settings
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

End Sub


Sub Macro3()
'
' Macro3 Macro
'

'
    Columns("A:K").Select
    ActiveSheet.Range("$A$1:$K$1115").AutoFilter Field:=1, Criteria1:=Array( _
        "Date", "Total", "="), Operator:=xlFilterValues

    Range(“A1”).Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow.Delete

    AutoFilterMode = False



    Selection.End(xlToLeft).Select
    Selection.End(xlUp).Select
    Range("A101").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Selection.AutoFilter
End Sub
