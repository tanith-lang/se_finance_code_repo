let
    Source = Table.NestedJoin(sAcquirerDetails,{"Batch Id", "Merchant Code", "Bank Account", "Merchant ID"},sAcquirerBatches,{"Batch Id", "Merchant Code", "Bank Account", "Merchant ID"},"sAcquirerBatches",JoinKind.LeftOuter),
    #"Expanded sAcquirerBatches" = Table.ExpandTableColumn(Source, "sAcquirerBatches", {"Period Begin Date", "Period End Date", "#Transactions", "Batch Total"}, {"Period Begin Date", "Period End Date", "#Transactions", "Batch Total"}),
    #"Add [Period Begin]" = Table.AddColumn(#"Expanded sAcquirerBatches", "Period End", each DateTime.Date( [Period End Date] )),
    #"Left Join sAcquirerSettlements" = Table.NestedJoin(#"Add [Period Begin]",{"Period End", "Merchant ID"},sAcquirerSettlements,{"Period End", "Merchant ID"},"sAcquirerSettlements",JoinKind.LeftOuter),
    #"Expanded sAcquirerSettlements" = Table.ExpandTableColumn(#"Left Join sAcquirerSettlements", "sAcquirerSettlements", {"Settlement Total"}, {"Settlement Total"}),
    #"Remove [Period End]" = Table.RemoveColumns(#"Expanded sAcquirerSettlements",{"Period End"}),
    #"Reordered Columns" = Table.ReorderColumns(#"Remove [Period End]",{"Merchant ID", "Merchant Code", "Bank Account", "Batch Id", "Date", "Period Begin Date", "Period End Date", "Order Code", "Payment Method", "Status", "Currency", "#Transactions", "Amount", "Batch Total"}),
    #"Renamed Columns" = Table.RenameColumns(#"Reordered Columns",{{"Merchant ID", "company_code"}, {"Merchant Code", "merchant_code"}, {"Bank Account", "bank_account"}, {"Batch Id", "batch_id"}, {"Date", "date"}, {"Period Begin Date", "period_begin_date"}, {"Period End Date", "period_end_date"}, {"Order Code", "order_code"}, {"Payment Method", "payment_method"}, {"Status", "status"}, {"Currency", "currency"}, {"#Transactions", "no_transactions"}, {"Amount", "amount"}, {"Batch Total", "batch_total"}, {"Settlement Total", "settlement_total"}}),
    #"Sorted Rows" = Table.Sort(#"Renamed Columns",{{"company_code", Order.Ascending}, {"batch_id", Order.Ascending}, {"merchant_code", Order.Ascending}})
in
    #"Sorted Rows"