Option Explicit

    Dim FSO As Object
    Dim fd As FileDialog
    Dim fPath As String
    Dim fDate As String
    Dim currCode As String
    Dim statusBar As Range
    Dim statusNum As Integer
    Dim exportDate As String
    Dim queryString As String
 
Public Sub getFolderPath_Import()
    
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    shtSummary.Activate
    
    If fd.Show = -1 Then
        fPath = fd.SelectedItems(1)
    End If
    
    Range("I10").Value = fPath & "\"

End Sub

Public Sub getFolderPath_Export()
    
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    shtSummary.Activate
    
    If fd.Show = -1 Then
        fPath = fd.SelectedItems(1)
    End If
    
    Range("I19").Value = fPath & "\"

End Sub

Public Sub getData()

    On Error GoTo errHandler:
    shtSummary.Activate
    
    ' Check Import Folder is present
    If Range("I10").Value = "" Then
        MsgBox "Please select an import folder"
        Exit Sub
    End If
    
    resetStatus
        
    ' Refresh queries
    
    refreshQuery ("fPaymentRiskReport(1)")
    Range("I16").Interior.Color = 3394611
    
    refreshQuery ("fEmailageAcceptedUsers")
    Range("M16").Interior.Color = 3394611
    
    refreshQuery ("fEmailageRejectedUsers")
    Range("O16").Interior.Color = 3394611
    
    refreshQuery ("fDailyFraudChecks")
    Range("Q16").Interior.Color = 3394611
    
    refreshQuery ("fZeroDepositBookings")
    Range("S16").Interior.Color = 3394611
    
    MsgBox "Import Completed"
    resetStatus
    Range("I17").Value = "Ready"

Exit Sub

errHandler:
    MsgBox "Error no: " & Err.Number & " has occurred" _
    & vbNewLine & "Description: " & Err.Description

End Sub

Public Sub exportCSV()

    On Error GoTo errHandler
    Application.ScreenUpdating = False

    ' Initialise
    shtSummary.Activate
    Set FSO = CreateObject("Scripting.FileSystemObject")
    fDate = Range("I13").Value
    fPath = Range("I19").Value
    exportDate = Format(fDate, "yyyy-mm-dd")

    ' Check if export folder exists, create new folder if not.
    If Not FSO.FolderExists(fPath) Then
        FSO.CreateFolder (fPath)
    End If

    ' Export Acquirer Batch Details worksheet .csv
    exportFile ("Booking Data Import")
        
    ' Update Summary sheet
    shtSummary.Activate
    Range("I22").Value = Now()
    Range("C22").Select
    
    ' Restore application defaults
    Application.CutCopyMode = False
    Application.ScreenUpdating = True

    MsgBox "CSV Export Complete"
    resetStatus
    Range("I17").Value = "Ready"
    
Exit Sub

errHandler:
    MsgBox "Error no: " & Err.Number & " has occurred" _
    & vbNewLine & "Description: " & Err.Description

End Sub

Function resetStatus()

    Set statusBar = Range("I16:S16")

    statusBar.Select
    With Selection.Interior
        .Pattern = xlVertical
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark2
        .TintAndShade = -0.899990844447157
        .PatternTintAndShade = 0
    End With
    
    Range("C16").Select

End Function

Function refreshQuery(query As String)
    
    queryString = "Query - " & query
    Range("I17").Value = "Loading " & queryString
    ActiveWorkbook.Connections(queryString).Refresh
    DoEvents
    
End Function

Function exportFile(exportSht As String)
    
    Sheets(exportSht).Activate
    Range("A1").CurrentRegion.Copy
    
    Workbooks.Add
    Range("A1").PasteSpecial (xlPasteValuesAndNumberFormats)
    
    Cells.Select
    Cells.Replace What:="(blank)", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
    ActiveWorkbook.SaveAs Filename:=fPath & exportDate & "_" & "DailyFraudChecks" & ".csv", _
        FileFormat:=xlCSVUTF8, _
        CreateBackup:=False, _
        Local:=True
    ActiveWorkbook.Close

End Function
