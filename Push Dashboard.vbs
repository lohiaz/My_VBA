Public Sub MTR_Files()
    Dim MyPath As String
    Dim MyFile As String
    Dim dirName As String

    Dim rng As Range 'store the range you want to delete
    Dim c 'total count of columns
    Dim i 'an index
    Dim j 'another index
    Dim headName As String 'The text on the header
    Dim Status As String 'This vars is just to get the code cleaner
    Dim Name As String
    Dim Age As String
    Dim sht As String

    Status = "Sent"
    Name = "Bounce"
    Age = "Open"
    sht = "Date"

    With Application.FileDialog(msoFileDialogFolderPicker)
        ' Optional: set folder to start in
        .InitialFileName = "C:\Users\plohia\Downloads\PushReporting\"
        .Title = "Select the folder to process"
        If .Show = True Then
            dirName = .SelectedItems(1) & "\"
        End If
    End With

    MyPath = dirName & "*.csv"
    MyFile = Dir(MyPath)
    If MyFile > "" Then MyFile = dirName & MyFile

    Do While MyFile <> ""
        If Len(MyFile) = 0 Then Exit Do

        Workbooks.Open MyFile

        With ActiveWorkbook
            For Each wks In .Worksheets
            'From A1 to the left at the end, and then store the number of the column, that is, the last column
                c = Range("A1").End(xlToRight).Column
            j = 0 'initialize the var
            For i = 1 To c 'all the numbers (heres is the columns) from 1 to c
                headName = Cells(1, i).Value
                If (headName <> Status) And (headName <> Name) And (headName <> Age) And (headName <> sht) Then
                'if the header of the column is differente of any of the options
                    j = j + 1 ' ini the counter
                    If j = 1 Then 'if is the first then
                        Set rng = Columns(i)
                    Else
                        Set rng = Union(rng, Columns(i))
                    End If
                 End If
            Next i
                Columns(1).AutoFit
                rng.Delete 'then brutally erased from leaf

                wks.Range("B15:D15").Select
                Selection.FormulaR1C1 = "=SUM(R[-13]C:R[-1]C)"
            Next
        End With

        'ActiveWorkbook.Close SaveChanges:=True
        'ActiveWorkbook.Open SaveChanges:=True

        MyFile = Dir
        If MyFile > "" Then MyFile = dirName & MyFile
    Loop
End Sub


Public Sub X_Files()
    Dim MyPath As String
    Dim MyFile As String
    Dim dirName As String

    Dim rng As Range 'store the range you want to delete
    Dim c 'total count of columns
    Dim i 'an index
    Dim j 'another index
    Dim headName As String 'The text on the header
    Dim Status As String 'This vars is just to get the code cleaner
    Dim Name As String
    Dim Age As String
    Dim sht As String

    Status = "Sent"
    Name = "Bounce"
    Age = "Open"
    sht = "Date"

    With Application.FileDialog(msoFileDialogFolderPicker)
        ' Optional: set folder to start in
        .InitialFileName = "C:\Users\plohia\Downloads\PushReporting\"
        .Title = "Select the folder to process"
        If .Show = True Then
            dirName = .SelectedItems(1) & "\"
        End If
    End With

    MyPath = dirName & "*.csv"
    MyFile = Dir(MyPath)
    If MyFile > "" Then MyFile = dirName & MyFile

    Do While MyFile <> ""
        If Len(MyFile) = 0 Then Exit Do

        Workbooks.Open MyFile

        With ActiveWorkbook
            For Each wks In .Worksheets
            'From A1 to the left at the end, and then store the number of the column, that is, the last column
                c = Range("A1").End(xlToRight).Column
            j = 0 'initialize the var
            For i = 1 To c 'all the numbers (heres is the columns) from 1 to c
                headName = Cells(1, i).Value
                If (headName <> Status) And (headName <> Name) And (headName <> Age) And (headName <> sht) Then
                'if the header of the column is differente of any of the options
                    j = j + 1 ' ini the counter
                    If j = 1 Then 'if is the first then
                        Set rng = Columns(i)
                    Else
                        Set rng = Union(rng, Columns(i))
                    End If
                 End If
            Next i
                Columns(1).AutoFit
                rng.Delete 'then brutally erased from leaf
            Next
        End With

        'ActiveWorkbook.Close SaveChanges:=True
        'ActiveWorkbook.Open SaveChanges:=True

        MyFile = Dir
        If MyFile > "" Then MyFile = dirName & MyFile
    Loop
End Sub



'below macro deletes all the files placed in the directory
Sub Delete()
'You can use this to delete all the files in the folder Test
    On Error Resume Next
    Kill "C:\Users\plohia\Downloads\PushReporting\*.*"
    MsgBox "DONE"
    On Error GoTo 0
End Sub


'below macro closes all the open .csv files
Sub x()
 For Each wbk In Workbooks
        If wbk.FileFormat = xlCSV Then
            wbk.Close True
        End If
    Next wbk
End Sub

