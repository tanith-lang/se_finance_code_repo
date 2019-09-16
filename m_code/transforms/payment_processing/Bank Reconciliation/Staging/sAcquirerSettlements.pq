let
    Source = Table.NestedJoin(sAcquirerDetails,{"Batch Id", "Merchant Code", "Bank Account", "Merchant ID"},sAcquirerBatches,{"Batch Id", "Merchant Code", "Bank Account", "Merchant ID"},"sAcquirerBatches",JoinKind.LeftOuter),
    #"Expanded sAcquirerBatches" = Table.ExpandTableColumn(Source, "sAcquirerBatches", {"Period Begin Date", "Period End Date", "#Transactions", "Batch Total"}, {"Period Begin Date", "Period End Date", "#Transactions", "Batch Total"}),
    #"Reordered Columns" = Table.ReorderColumns(#"Expanded sAcquirerBatches",{"Merchant ID", "Merchant Code", "Bank Account", "Batch Id", "Date", "Period Begin Date", "Period End Date", "Order Code", "Payment Method", "Status", "Currency", "#Transactions", "Amount", "Batch Total"}),
    #"Add [Period End]" = Table.AddColumn(#"Reordered Columns", "Period End", each DateTime.Date( [Period End Date] )),
    #"Grouped Rows" = Table.Group(#"Add [Period End]", {"Period End", "Merchant ID"}, {{"Settlement Total", each List.Sum([Amount]), type number}}),
    #"< Set Data Types >" = Table.TransformColumnTypes(#"Grouped Rows",{{"Settlement Total", Currency.Type}, {"Period End", type date}})
in
    #"< Set Data Types >"