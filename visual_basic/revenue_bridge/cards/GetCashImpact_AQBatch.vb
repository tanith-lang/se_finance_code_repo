Option Explicit

Sub getCashImpactAQBatch()

    Application.ScreenUpdating = False

    'Declare NBT339 variables
    Dim NBT339_merchantId As String
    Dim NBT339_tradingDate As Date
    Dim NBT339_processingDate As Date
    Dim NBT339_amount As Double
    Dim cashImpact As Date
    
    'Declare Acquirer Batch variables
    Dim AQ_merchantId As String
    Dim AQ_tradingDate As Date
    Dim AQ_amount As Double
    
    'Declare data table arrays
    Dim testArray() As Variant
    Dim resultArray() As Variant
    
    'Declare counters
    Dim t As Long
    Dim r As Long
    Dim startRow As Integer
    
    'Initialise arrays
    testArray = shtNBT339.ListObjects(1).DataBodyRange
    resultArray = shtAcquirerBatch.ListObjects(1).DataBodyRange
    
    'Enter results start row here (Default = enter '1' to skip table header):
    startRow = 1
    
    'Outer loop over Acquirer Batch table
    For r = LBound(resultArray(), 1) To UBound(resultArray(), 1) Step 1
        AQ_merchantId = resultArray(r, 2)
        AQ_tradingDate = resultArray(r, 13)
        AQ_amount = Round(resultArray(r, 12), 2)
        
        'Inner loop over NBT339 table
        For t = LBound(testArray(), 1) To UBound(testArray(), 1) Step 1
            NBT339_merchantId = testArray(t, 1)
            NBT339_tradingDate = testArray(t, 2)
            NBT339_processingDate = testArray(t, 3)
            NBT339_amount = Round(testArray(t, 5), 2)
            cashImpact = testArray(t, 7)
            
            'Conditional logic to return Cash Impact
            If NBT339_merchantId = AQ_merchantId _
            And NBT339_tradingDate = AQ_tradingDate _
            And NBT339_amount = AQ_amount Then
                shtAcquirerBatch.Range("N" & r + startRow).Value = NBT339_processingDate
                shtAcquirerBatch.Range("O" & r + startRow).Value = cashImpact
            End If
        Next
    Next
    
    'Clear system RAM
    Erase testArray()
    Erase resultArray()
    
    Application.ScreenUpdating = True
    MsgBox "Task completed at " & Now()

End Sub
