Option Explicit

Sub getCashImpact_NBT506()

    Application.ScreenUpdating = False

    'Declare NBT506 variables
    Dim NBT506_merchantId As String
    Dim NBT506_processingDate As Date
    Dim NBT506_processingMax As Date
    Dim NBT506_amount As Double
    Dim bankCharge As Double
    
    'Declare Bank Statement variables
    Dim cashImpact As Date
    Dim bank_merchantId As String
    Dim bank_amount As Double
    
    'Declare Arrays
    Dim testArray() As Variant
    Dim resultArray() As Variant
    
    'Declare Counters
    Dim t As Long
    Dim r As Long
    Dim startRow As Integer
    
    'Initialise Arrays
    testArray = shtBank.ListObjects(1).DataBodyRange
    resultArray = shtNBT506.ListObjects(1).DataBodyRange
    
    'Enter results start row here (Default = enter '1' to skip table header):
    startRow = 1
    
    '''''''''''''''''''''''''''''''
    'Enter bank charge value here:'
    '''''''''''''''''''''''''''''''
    bankCharge = shtNBT506.Range("N8").Value
    
    'Outer loop over NBT506 table
    For r = LBound(resultArray, 1) To UBound(resultArray, 1) Step 1
        NBT506_merchantId = resultArray(r, 1)
        NBT506_processingDate = resultArray(r, 3)
        NBT506_processingMax = DateAdd("d", 7, NBT506_processingDate)
        NBT506_amount = Round(resultArray(r, 10), 2)
    
        'Inner loop over Bank Statement table
        For t = LBound(testArray, 1) To UBound(testArray, 1) Step 1
            cashImpact = testArray(t, 1)
            bank_merchantId = testArray(t, 4)
            bank_amount = Round(testArray(t, 6), 2)
            
            'Conditional logic to return cash impact date
            If NBT506_amount = 0 Then
                shtNBT506.Range("K" & r + startRow).Value = resultArray(r, 3)
            ElseIf cashImpact >= NBT506_processingDate _
                And cashImpact <= NBT506_processingMax _
                And bank_merchantId = NBT506_merchantId _
                And bank_amount <= NBT506_amount _
                And bank_amount >= NBT506_amount - bankCharge Then
                shtNBT506.Range("K" & r + startRow).Value = cashImpact
            End If
        Next
    Next
    
    'Clear system RAM
    Erase testArray()
    Erase resultArray()
    
    Application.ScreenUpdating = True
    
    'Completion Message
    MsgBox "Task completed"

End Sub
