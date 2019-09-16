Option Explicit

    ' File system object
    Dim FSO As Object
    Dim fd As FileDialog
    
    ' File paths
    Dim file_path As String
    
    'Filter date
    Dim filter_date As Date
    
    ' File name components
    Public file_date As String
    Public currency_code As String
    
 
Sub get_import_path()
    
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    sht_controls.Activate
    
    If fd.Show = -1 Then
        file_path = fd.SelectedItems(1)
    End If
    
    Range("I10").Value = file_path & "\"

End Sub


Sub get_export_path()
    
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    sht_controls.Activate
    
    If fd.Show = -1 Then
        file_path = fd.SelectedItems(1)
    End If
    
    Range("I19").Value = file_path & "\"

End Sub


Sub import_data()

    sht_controls.Activate
    
    ' Check Import Folder is present
    If Range("I10").Value = "" Then
        MsgBox "Please select an import folder"
        Exit Sub
    End If
    
    ActiveWorkbook.RefreshAll
    
    MsgBox "Import Completed"
    resetStatus

    Exit Sub

End Sub


Sub exportCSV()

    Application.ScreenUpdating = False

    ' Initialise
    Set FSO = CreateObject("Scripting.FileSystemObject")
    currency_code = sht_controls.Range("I13").Value
    file_path = sht_controls.Range("I19").Value
    file_date = Format(sht_controls.Range("P13").Value, "yyyy-MM")
    filter_date = Format(sht_controls.Range("P13").Value, "MM/dd/yyyy")
    
    ' Filter NBT506 and Bank Statement tables
    Call filter_output_tables

    ' Check if export folder exists, create new folder if not.
    If Not FSO.FolderExists(file_path) Then
        FSO.CreateFolder (file_path)
    End If
    
    
    ' Export APM Transfers reconciliation
    Call export_file(sht_apm_transfers, file_date, "APM_BATCH_RECON", "WORLDPAY")
    
    ' Export Acquirer Batch reconciliation
    Call export_file(sht_aq_batch, file_date, "CARD_BATCH_RECON", "WORLDPAY")
    
    ' Export Interchange reconciliation
    Call export_file(sht_interchange, file_date, "INTERCHANGE_RECON", "WORLDPAY")
    
    ' Export NBT506 reconciliation
    Call export_file(sht_nbt506_emis, file_date, "NBT506_RECON", "WORLDPAY")
    
    ' Export RatePay reconciliation
    Call export_file(sht_ratepay, file_date, "RATEPAY_RECON", "RATEPAY")
    
    ' Export Bank reconciliation
    Call export_file(sht_bank, file_date, "BNK_RECON", "BANK_STATEMENT")
     
     
    ' Update Summary sheet
    sht_controls.Activate
    Range("I25").Value = Now()
    Range("C25").Select
    
    
    ' Restore application defaults
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    
    MsgBox "CSV Export Complete"

    Exit Sub

End Sub


Function resetStatus()

    Range("I16:T16").Select
    With Selection.Interior
        .Pattern = xlVertical
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark2
        .TintAndShade = -0.899990844447157
        .PatternTintAndShade = 0
    End With
    Range("I17").Value = ""
    Range("C16").Select

End Function


Function statusMessage(message As String)

    Range("I17").Value = message
    
End Function


Function refreshQuery(query As String)

    Dim queryString As String
    queryString = "Query - " & query

    ActiveWorkbook.Connections(queryString).Refresh
    Range("I17").Value = ActiveWorkbook.Connections(queryString).Name
    DoEvents

End Function


Function export_file(export_sheet As Worksheet, file_date As String, batch_type As String, psp As String)
    
    Dim full_path As String
    
    full_path = file_path & psp & "\" & file_date & "_" & batch_type & "_" & currency_code & ".csv"
    
    export_sheet.Activate
    Range("A1").CurrentRegion.Copy
    
    Workbooks.Add
    Range("A1").PasteSpecial (xlPasteValuesAndNumberFormats)
    ActiveWorkbook.SaveAs Filename:=full_path, _
        FileFormat:=xlCSVUTF8, _
        CreateBackup:=False, _
        Local:=True
    ActiveWorkbook.Close

End Function


Function filter_output_tables()

    ' Filter NBT506 table
    sht_nbt506_emis.ListObjects("fNBT506_EMIS").AutoFilter.ShowAllData
    sht_nbt506_emis.ListObjects("fNBT506_EMIS").Range.AutoFilter Field:=3, Operator _
    :=xlFilterValues, Criteria2:=Array(1, filter_date)
    
    ' Filter Bank Statement table
    sht_bank.ListObjects("fBankStatement").AutoFilter.ShowAllData
    sht_bank.ListObjects("fBankStatement").Range.AutoFilter Field:=1, _
    Operator:=xlFilterValues, Criteria2:=Array(1, filter_date)

End Function
