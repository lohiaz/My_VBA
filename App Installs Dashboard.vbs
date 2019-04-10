Sub AppAll()
'
' AppInstall Macro
'
'Dim mypath
Dim mypath As String
mypath = ActiveWorkbook.Path
'mypath = Environ("USERPROFILE") & "\Desktop"
'
    'Sheets.Add After:=ActiveSheet
    Application.DisplayAlerts = False
    
    '
    '-------------------------------------------- This is Lotto --------------------------------------------
    '
  
    
    Workbooks.Add
    ChDir mypath
    ActiveWorkbook.SaveAs Filename:=mypath & "\Install.xlsx", _
        FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
        

    ChDir mypath & "\1. Lotto"
    
    Workbooks.Open Filename:= _
        mypath & "\1. Lotto\Lotto Ios.csv"

    Workbooks.Open Filename:= _
        mypath & "\1. Lotto\Lotto Android.csv"
        
    Windows("Lotto Android.csv").Activate
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "Android"
    Range("I2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.FillDown
    Range("A2").Select
    
    
    Windows("Lotto Ios.csv").Activate
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "IOS"
    Range("I2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.FillDown
    Range("A2").Select
    
    
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    
    Windows("Lotto Android.csv").Activate
    Range("A2").Select
    Selection.End(xlDown).Select
    Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    Cells.Select
    Selection.Columns.AutoFit
    Range("A1").Select
    Selection.AutoFilter
    
    ActiveWorkbook.Worksheets("Lotto Android").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Lotto Android").AutoFilter.Sort.SortFields.Add Key _
        :=Range("A1:A10058"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Lotto Android").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Selection.AutoFilter
    Range("B:B,E:E,G:G,J:N").Select
    Range("J1").Activate
    Selection.Delete Shift:=xlToLeft
    
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Windows("Install").Activate
    ActiveSheet.Paste
    Selection.Columns.AutoFit

    Sheets("Sheet1").Select
    Sheets.Add After:=ActiveSheet
    Windows("Lotto Android.csv").Activate
    ActiveWindow.Close
    Windows("Lotto Ios.csv").Activate
    ActiveWindow.Close
    
    Windows("AppInstalls_File.xlsm").Activate

    
    '
    '-------------------------------------------- This is PCH App --------------------------------------------
    '
    

    'Sheets.Add After:=ActiveSheet

    ChDir mypath & "\2. Pch App"
    Workbooks.Open Filename:= _
        mypath & "\2. Pch App\Pch Ios.csv"

    Workbooks.Open Filename:= _
        mypath & "\2. Pch App\Pch Android.csv"
        
    Windows("Pch Android.csv").Activate
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "Android"
    Range("I2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.FillDown
    Range("A2").Select
    
    
    Windows("Pch Ios.csv").Activate
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "IOS"
    Range("I2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.FillDown
    Range("A2").Select
    
    
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    
    Windows("Pch Android.csv").Activate
    Range("A2").Select
    Selection.End(xlDown).Select
    Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    Cells.Select
    Selection.Columns.AutoFit
    Range("A1").Select
    Selection.AutoFilter
    
    ActiveWorkbook.Worksheets("Pch Android").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Pch Android").AutoFilter.Sort.SortFields.Add Key _
        :=Range("A1:A10058"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Pch Android").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Selection.AutoFilter
    Range("B:B,E:E,G:G,J:N").Select
    Range("J1").Activate
    Selection.Delete Shift:=xlToLeft
    
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Windows("Install").Activate
    ActiveSheet.Paste
    Selection.Columns.AutoFit

    Sheets("Sheet2").Select
    Sheets.Add After:=ActiveSheet
    Windows("Pch Android.csv").Activate
    ActiveWindow.Close
    Windows("Pch Ios.csv").Activate
    ActiveWindow.Close
    
    
    
    '
    '-------------------------------------------- This is FrontPage --------------------------------------------
    '
    
    
    Windows("AppInstalls_File.xlsm").Activate

    ChDir mypath & "\3. Frontpage"
    Workbooks.Open Filename:= _
        mypath & "\3. Frontpage\Frontpage Ios.csv"

    Workbooks.Open Filename:= _
        mypath & "\3. Frontpage\Frontpage Android.csv"
        
    Windows("Frontpage Android.csv").Activate
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "Android"
    Range("I2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.FillDown
    Range("A2").Select
    
    
    Windows("Frontpage Ios.csv").Activate
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "IOS"
    Range("I2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.FillDown
    Range("A2").Select
    
    
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    
    Windows("Frontpage Android.csv").Activate
    Range("A2").Select
    Selection.End(xlDown).Select
    Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    Cells.Select
    Selection.Columns.AutoFit
    Range("A1").Select
    Selection.AutoFilter
    
    ActiveWorkbook.Worksheets("Frontpage Android").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Frontpage Android").AutoFilter.Sort.SortFields.Add Key _
        :=Range("A1:A10058"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Frontpage Android").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Selection.AutoFilter
    Range("B:B,E:E,G:G,J:N").Select
    Range("J1").Activate
    Selection.Delete Shift:=xlToLeft
    
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Windows("Install").Activate
    ActiveSheet.Paste
    Selection.Columns.AutoFit

    Sheets("Sheet3").Select
    Sheets.Add After:=ActiveSheet
    Windows("Frontpage Android.csv").Activate
    ActiveWindow.Close
    Windows("Frontpage Ios.csv").Activate
    ActiveWindow.Close
    
    Windows("AppInstalls_File.xlsm").Activate
    
    '
    '-------------------------------------------- This is Game --------------------------------------------
    '
    
    ChDir mypath & "\4. Pch Game"
    Workbooks.Open Filename:= _
        mypath & "\4. Pch Game\Game Ios.csv"

    Workbooks.Open Filename:= _
        mypath & "\4. Pch Game\Game Android.csv"
        
    Windows("Game Android.csv").Activate
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "Android"
    Range("I2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.FillDown
    Range("A2").Select
    
    
    Windows("Game Ios.csv").Activate
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "IOS"
    Range("I2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.FillDown
    Range("A2").Select
    
    
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    
    Windows("Game Android.csv").Activate
    Range("A2").Select
    Selection.End(xlDown).Select
    Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    Cells.Select
    Selection.Columns.AutoFit
    Range("A1").Select
    Selection.AutoFilter
    
    ActiveWorkbook.Worksheets("Game Android").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Game Android").AutoFilter.Sort.SortFields.Add Key _
        :=Range("A1:A10058"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Game Android").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Selection.AutoFilter
    Range("B:B,E:E,G:G,J:N").Select
    Range("J1").Activate
    Selection.Delete Shift:=xlToLeft
    
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Windows("Install").Activate
    ActiveSheet.Paste
    Selection.Columns.AutoFit

    Sheets("Sheet4").Select
    Sheets.Add After:=ActiveSheet
    Windows("Game Android.csv").Activate
    ActiveWindow.Close
    Windows("Game Ios.csv").Activate
    ActiveWindow.Close
    
    '
    '-------------------------------------------- This is WinCredible --------------------------------------------
    '
    
    ChDir mypath & "\5. Wincredible"
    Workbooks.Open Filename:= _
        mypath & "\5. Wincredible\Wincredible Ios.csv"

    Windows("Wincredible Ios.csv").Activate
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "IOS"
    Range("I2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.FillDown
    Range("A2").Select
    
    
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    
    Cells.Select
    Selection.Columns.AutoFit
    Range("A1").Select
    Selection.AutoFilter
    
    ActiveWorkbook.Worksheets("Wincredible Ios").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Wincredible Ios").AutoFilter.Sort.SortFields.Add Key _
        :=Range("A1:A10058"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Wincredible Ios").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Selection.AutoFilter
    Range("B:B,E:E,G:G,J:N").Select
    Range("J1").Activate
    Selection.Delete Shift:=xlToLeft
    
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Windows("Install").Activate
    ActiveSheet.Paste
    Selection.Columns.AutoFit

    Sheets("Sheet5").Select
    Sheets.Add After:=ActiveSheet
    Windows("Wincredible Ios.csv").Activate
    ActiveWindow.Close
    
    Windows("AppInstalls_File.xlsm").Activate
    
    Application.DisplayAlerts = True
End Sub
