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
    Dim inputCheck As Integer
    Dim outputCheck As Integer
 
Public Sub GetImportPath()
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    shtSummary.Activate
    
    If fd.Show = -1 Then
        fPath = fd.SelectedItems(1)
    End If
    
    Range("I10").Value = fPath

End Sub

Public Sub GetMappingPath()
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    shtSummary.Activate
    
    If fd.Show = -1 Then
        fPath = fd.SelectedItems(1)
    End If
    
    Range("I13").Value = fPath

End Sub

Public Sub GetExportPath()
    
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    shtSummary.Activate
    
    If fd.Show = -1 Then
        fPath = fd.SelectedItems(1)
    End If
    
    Range("I19").Value = fPath & "\"

End Sub

Public Sub getData()

    On Error GoTo errhandler:
    shtSummary.Activate
    
    ' Check Import Folder is present
    If Range("I10").Value = "" Then
        MsgBox "Please select an import folder"
        Exit Sub
    End If
    
    resetStatus
        
    ' Refresh queries
    refreshQuery ("iGLExtract")
    Range("I16").Interior.Color = 3394611
    
    refreshQuery ("sGLExtract")
    Range("L16").Interior.Color = 3394611
    
    refreshQuery ("Journal Gross Check (Table)")
    Range("O16").Interior.Color = 3394611
    
    refreshQuery ("Invoice / Credit Check (Table)")
    Range("R16").Interior.Color = 3394611
    
    MsgBox "Import Completed"
    resetStatus
    Range("I17").Value = "Ready"

Exit Sub

errhandler:
    MsgBox "Error no: " & Err.Number & " has occurred" _
    & vbNewLine & "Description: " & Err.Description

End Sub

Public Sub exportCSV()

    On Error GoTo errhandler
    Application.ScreenUpdating = False

    ' Initialise
    shtSummary.Activate
    Set FSO = CreateObject("Scripting.FileSystemObject")
    fDate = Range("I22").Value
    fPath = Range("I19").Value
    exportDate = Format(fDate, "yyyy-mm-dd")
    inputCheck = Range("Q25").Value
    outputCheck = Range("S25").Value

    ' Check if export folder exists, create new folder if not.
    If Not FSO.FolderExists(fPath) Then
        FSO.CreateFolder (fPath)
    End If
    
    If inputCheck <> outputCheck Then
        MsgBox "Input and output row counts do not match!"
        Exit Sub
    End If

    ' Export Acquirer Batch Details worksheet .csv
    exportFile "Invoices", "CONCUR_INVOICES"
    exportFile "Credit Notes", "CONCUR_CREDIT_NOTES"
        
    ' Update Summary sheet
    shtSummary.Activate
    Range("I17").Value = "Ready"
    Range("I25").Value = Now()
    Range("C22").Select
    
    ' Restore application defaults
    Application.CutCopyMode = False
    Application.ScreenUpdating = True

    MsgBox "CSV Export Complete"
    resetStatus
    
Exit Sub

errhandler:
    MsgBox "Error no: " & Err.Number & " has occurred" _
    & vbNewLine & "Description: " & Err.Description

End Sub

Sub generateInvoices()

    Dim testSheet As Worksheet
    Dim tblGLExtract As ListObject
    Dim testCell As Range

    Set testSheet = shtGLExtract
    Set tblGLExtract = testSheet.ListObjects(1)
    Set testCell = tblGLExtract.DataBodyRange.Find(What:="Error: ", LookIn:=xlValues, SearchDirection:=xlNext, MatchCase:="True")

    If testCell Is Nothing Then
        On Error GoTo errhandler
    
        'Refresh Invoice Table
        ActiveWorkbook.Connections("Query - Invoices (Table)").Refresh
        DoEvents
        
        'Refresh GL Extract Table
        ActiveWorkbook.Connections("Query - Credit Notes (Table)").Refresh
        DoEvents
    Else
        Application.ScreenUpdating = True
        MsgBox "Errors exist in GL Extract table"
        
        Exit Sub
    End If

    Application.ScreenUpdating = "True"
    MsgBox "Invoices and credits generated"

Exit Sub

errhandler:

    Application.ScreenUpdating = "True"
    MsgBox "Invoice export error"

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

Function exportFile(exportSht As String, titleText As String)

    Dim outputPath As String
    
    Sheets(exportSht).Activate
    Range("A1").CurrentRegion.Copy
    
    Workbooks.Add
    Range("A1").PasteSpecial (xlPasteValuesAndNumberFormats)
    
    Cells.Select
    Cells.Replace What:="(blank)", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
        outputPath = fPath & exportDate & "_" & titleText & ".csv"
    
        ActiveWorkbook.SaveAs Filename:=outputPath, _
        FileFormat:=xlCSV, CreateBackup:=False, Local:=True
        ActiveWorkbook.Close

End Function