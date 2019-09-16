Option Explicit

Sub fillCashImpact()

    'Declare APM Transfers variables
    Dim eventType As String
    Dim cashImpact As Date
    
    'Declare array
    Dim resultArray() As Variant
    
    'Declare counters
    Dim r As Long
    Dim startRow As Integer
    
    'Initialise array
    resultArray = shtAPMRec.ListObjects(1).DataBodyRange
    startRow = 1
    
    'Loop backwards over APM Reconciliation table
    For r = UBound(resultArray(), 1) To LBound(resultArray(), 1) Step -1
        eventType = resultArray(r, 5)
        If eventType = "PAYMENT_TO_MERCHANT_INITIATED" Then
            cashImpact = resultArray(r, 19)
        ElseIf eventType <> "PAYMENT_TO_MERCHANT_INITIATED" Then
            shtAPMRec.Range("S" & r + startRow).Value = cashImpact
        End If
    Next

End Sub
