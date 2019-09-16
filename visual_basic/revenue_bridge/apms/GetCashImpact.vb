Option Explicit

Sub getCashImpact()

'    This subroutine will deduct user specified bank charges from acquirer batch
'    transfers, then match the transfers to the bank statement. Please refer to
'    the comment block below to set the bank charge value.

    'Declare APM Transfers variables
    Dim APM_merchantId As String
    Dim APM_eventType As String
    Dim APM_transferDate As Date
    Dim APM_transferMax As Date
    Dim APM_amount As Double
    
    'Declare bank statement variables
    Dim cashImpact As Date
    Dim bank_charge As Double
    Dim bank_merchantId As String
    Dim bank_amount As Double
    
    'Declare arrays
    Dim testArray() As Variant
    Dim resultArray() As Variant
    
    'Declare counters
    Dim t As Long
    Dim r As Long
    Dim startRow As Integer
    
    'Initialise arrays
    testArray() = shtBank.ListObjects(1).DataBodyRange
    resultArray() = shtAPMRec.ListObjects(1).DataBodyRange
    startRow = 1
    
    ''''''''''''''''''''''''''''
    'Set bank charge value here'
    ''''''''''''''''''''''''''''
    bank_charge = shtAPMRec.Range("W2").Value
    
    'Outer loop over APM Reconciliation table
    For r = LBound(resultArray(), 1) To UBound(resultArray(), 1) Step 1
        APM_merchantId = resultArray(r, 2)
        APM_eventType = resultArray(r, 5)
        APM_transferDate = resultArray(r, 18)
        APM_transferMax = DateAdd("d", 7, resultArray(r, 18))
        APM_amount = Round(resultArray(r, 16), 2) * -1
        
        'Check event status
        If APM_eventType = "PAYMENT_TO_MERCHANT_INITIATED" Then
        
            'Inner loop over bank statement table
            For t = LBound(testArray(), 1) To UBound(testArray(), 1) Step 1
                cashImpact = testArray(t, 1)
                bank_merchantId = Left(testArray(t, 4), 3)
                bank_amount = testArray(t, 6)
                
                'Conditional logic to return Cash Impact
                If cashImpact >= APM_transferDate _
                And cashImpact <= APM_transferMax _
                And bank_merchantId = "BTX" _
                And bank_amount <= APM_amount And bank_amount >= APM_amount - bank_charge Then
                    shtAPMRec.Range("S" & r + startRow).Value = cashImpact
                End If
            Next
        End If
    Next
    
    'Completion message
    MsgBox "Task completed at: " & Now()
    
End Sub
