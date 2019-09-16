let
    Source = iRatePayTransactions,
    #"Left Join sRatePaySettlements" = Table.NestedJoin(Source,{"SHOP_ID", "PAYMENTDATE"},sRatePaySettlements,{"SHOP_ID", "PAYMENTDATE"},"iRatePaySettlements",JoinKind.LeftOuter),
    #"Expanded iRatePaySettlements" = Table.ExpandTableColumn(#"Left Join sRatePaySettlements", "iRatePaySettlements", {"SETTLEMENT_AMOUNT"}, {"settlement_amount"}),
    #"Renamed Columns" = Table.RenameColumns(#"Expanded iRatePaySettlements",{{"SHOP_ID", "company_code"}, {"PAYMENTDATE", "payment_date"}, {"PAYMENTMONTH", "payment_month"}, {"SHOPNAME", "merchant_code"}, {"AMOUNT", "amount"}, {"DESCRIPTOR", "descriptor"}, {"SHOPINVOICE_ID", "shop_invoice_id"}, {"SHOPSORDER_ID", "shop_order_id"}, {"INVOICENUMBER", "invoice_number"}, {"DESCRIPTION", "description"}, {"FEETYPE", "fee_type"}, {"ORDERDATE", "order_date"}, {"SENTDATE", "sent_date"}, {"TRANSACTION_ID", "transaction_id"}}),
    #"Reordered Columns" = Table.ReorderColumns(#"Renamed Columns",{"company_code", "merchant_code", "settlement_currency", "payment_date", "payment_month", "amount", "descriptor", "shop_invoice_id", "shop_order_id", "invoice_number", "description", "fee_type", "order_date", "sent_date", "transaction_id", "settlement_amount"})
in
    #"Reordered Columns"