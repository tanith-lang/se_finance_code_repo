Function exportCSVFile(sheetName)

   ' Exports specified sheet name as a .csv file with all values double-quoted

    Dim DestFile As String
    Dim FileNum As Integer
    Dim ColumnCount As Integer
    Dim RowCount As Integer
    Dim Target As Range
    
    Dim dateStart As String
    Dim dateEnd As String

    sheetName.Activate
    ActiveSheet.ListObjects(1).Range.Select
    tabName = ActiveSheet.Name
    
    dateStart = Format(fYear & "-" & tabName, "yyyy-mm-dd")
    dateEnd = Format(CDate(WorksheetFunction.EoMonth(dateStart, 0)), "yyyy-mm-dd")
    
    ' Set source (Target) and destination (DestFile)
    Set Target = Range("A1").CurrentRegion
    DestFile = fPath & fName & " " & dateStart & "_" & dateEnd & ".csv"
    
    ' Obtain next free file handle number.
    FileNum = FreeFile()
    
    ' Turn error checking off.
    On Error Resume Next
    
    ' Attempt to open destination file for output.
    Open DestFile For Output As #FileNum
    
    If Err <> 0 Then
        MsgBox "Cannot open filename " & DestFile
        End
    End If
    
    ' Turn error checking on.
    On Error GoTo 0
    
    ' Loop for each row in selection.
    For RowCount = 1 To Target.Rows.Count

        ' Loop for each column in selection.
        For ColumnCount = 1 To Target.Columns.Count

            ' Write current cell's text to file with quotation marks.
            Print #FileNum, """" & Target.Cells(RowCount, _
            ColumnCount).Text & """";

            ' Check if cell is in last column.
            If ColumnCount = Target.Columns.Count Then
                ' If so, then write a blank line.
                Print #FileNum,
            Else
                ' Otherwise, write a comma.
                Print #FileNum, ",";
            End If
            
        ' Start next iteration of ColumnCount loop.
        Next ColumnCount
        
    ' Start next iteration of RowCount loop.
    Next RowCount

    ' Close destination file.
    Close #FileNum

End Function