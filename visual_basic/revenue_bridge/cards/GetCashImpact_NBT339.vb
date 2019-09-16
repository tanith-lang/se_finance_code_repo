Option Explicit

Sub getCashImpact_NBT339()

    Application.ScreenUpdating = False

    'Declare NBT506 variables
    Dim NBT506_merchantId As String
    Dim NBT506_processingDate As Date
    Dim NBT506_amount As Double
    Dim cashImpact As Date
    
    'Declare NBT339 variables
    Dim NBT339_merchantId As String
    Dim NBT339_processingDate As Date
    Dim NBT339_amount As Double
    
    'Declare Arrays
    Dim testArray() As Variant
    Dim resultArray() As Variant
    
    'Declare Counters
    Dim t As Long
    Dim r As Long
    Dim startRow As Integer
    
    'Initialise Arrays
    testArray = shtNBT506.ListObjects(1).DataBodyRange
    resultArray = shtNBT339.ListObjects(1).DataBodyRange

    'Enter results start row here (Default = enter '1' to skip table header):
    startRow = 1
    
    'Outer loop over NBT339 table
    For r = LBound(resultArray(), 1) To UBound(resultArray(), 1) Step 1
        NBT339_merchantId = resultArray(r, 1)
        NBT339_processingDate = resultArray(r, 3)
        NBT339_amount = Round(resultArray(r, 6), 2)
        
        'Inner loop over NBT506 table
        For t = LBound(testArray(), 1) To UBound(testArray(), 1) Step 1
            NBT506_merchantId = testArray(t, 1)
            NBT506_processingDate = testArray(t, 3)
            NBT506_amount = Round(testArray(t, 5), 2)
            cashImpact = testArray(t, 11)
        'Conditional logic to return Cash Impact
            If NBT339_merchantId = NBT506_merchantId _
            And NBT339_processingDate = NBT506_processingDate _
            And NBT339_amount = NBT506_amount Then
                shtNBT339.Range("G" & r + startRow).Value = cashImpact
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
