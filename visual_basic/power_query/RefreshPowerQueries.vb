Public Sub RefreshPowerQueries()
    ' Subroutine to refresh all Power Query scripts in a workbook

    Dim i As Long, conn As WorkbookConnection
    On Error Resume Next
    
    ' Loop over all queries in workbook and refresh if no error
    For Each conn In ThisWorkbook.Connections
        i = InStr(1, conn.OLEDBConnection.Connection, "Provider=Microsoft.Mashup.OleDb.1", vbTextCompare)
        If Err.Number <> 0 Then
            Err.Clear
            Exit For
        End If
        If i > 0 Then
            conn.Refresh
            DoEvents
        End If
    Next conn

End Sub