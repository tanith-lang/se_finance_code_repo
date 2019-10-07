Option Explicit
    Dim FSO As Object
    Dim fName As String
    Dim fPath As String
    Dim fYear As String
    Dim tabName As String
    Dim testRange As Range


Sub refreshQueries()

    ActiveWorkbook.RefreshAll
    
End Sub


Sub exportCSV()

'    Assign module level variables
    fName = "historical__component_booking_base"
    fYear = sht_controls.Range("F12").Value

    Application.ScreenUpdating = False
    On Error GoTo errHandler

    Set FSO = CreateObject("Scripting.FileSystemObject")
    fPath = sht_controls.Range("F20").Value & "/"

    If Not FSO.FolderExists(fPath) Then
        FSO.CreateFolder (fPath)
    End If
    
    Application.DisplayAlerts = False

    Call ExportCSVFile(sht_01)
    Call ExportCSVFile(sht_02)
    Call ExportCSVFile(sht_03)
    Call ExportCSVFile(sht_04)
    Call ExportCSVFile(sht_05)
    Call ExportCSVFile(sht_06)
    Call ExportCSVFile(sht_07)
    Call ExportCSVFile(sht_08)
    Call ExportCSVFile(sht_09)
    Call ExportCSVFile(sht_10)
    Call ExportCSVFile(sht_11)
    Call ExportCSVFile(sht_12)

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    MsgBox "CSV Export Complete"

Exit Sub

errHandler:

    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Number & " " & Err.Description

End Sub

Function exportFile(sheetName)

    Dim dateStart As String
    Dim dateEnd As String

    sheetName.Activate
    tabName = ActiveSheet.Name
    dateStart = Format(fYear & "-" & tabName, "yyyy-mm-dd")
    dateEnd = Format(CDate(WorksheetFunction.EoMonth(dateStart, 0)), "yyyy-mm-dd")
    
    Range("A1").CurrentRegion.Copy
    Workbooks.Add
    Range("A1").PasteSpecial (xlPasteValuesAndNumberFormats)
    ActiveWorkbook.SaveAs Filename:=fPath & fName & " " & dateStart & "_" & dateEnd & ".csv", _
    FileFormat:=xlCSV, CreateBackup:=False, Local:=True
    
    Application.CutCopyMode = False
    ActiveWorkbook.Close

End Function


Function ExportCSVFile(sheetName)

    Dim DestFile As String
    Dim FileNum As Integer
    Dim ColumnCount As Integer
    Dim RowCount As Integer
    Dim Target As Range
    
    Dim dateStart As String
    Dim dateEnd As String
    Dim stringStart As String
    Dim stringEnd As String

    sheetName.Activate
    ActiveSheet.ListObjects(1).Range.Select
    tabName = ActiveSheet.Name
    
    dateStart = Format(fYear & "-" & tabName, "yyyy-mm-dd")
    dateEnd = Format(CDate(WorksheetFunction.EoMonth(dateStart, 0)), "yyyy-mm-dd")
    stringStart = Format(dateStart, "yyyymmdd")
    stringEnd = Format(dateEnd, "yyyymmdd")
    
    ' Set source (Target) and destination (DestFile)
    Set Target = Range("A1").CurrentRegion
    DestFile = fPath & fName & " " & stringStart & "_" & stringEnd & ".csv"
    
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
