let
    import_currency = Excel.CurrentWorkbook(){[Name="import_currency"]}[Content]{0}[Column1],

    Source = iTransfersByPeriod,
    #"Filter [Settlement Currency]" = Table.SelectRows(Source, each ([Settlement Currency] = import_currency)),
    #"Replace ""PAYMENT_TO_""" = Table.ReplaceValue(#"Filter [Settlement Currency]","PAYMENT_TO_","Z_PAYMENT_TO_",Replacer.ReplaceText,{"Event Type"}),
    #"Replace "". .""" = Table.ReplaceValue(#"Replace ""PAYMENT_TO_""",". .","",Replacer.ReplaceText,{"Journal Description"}),
    #"Replace "".  .""" = Table.ReplaceValue(#"Replace "". .""",".  .","",Replacer.ReplaceText,{"Journal Description"}),
    #"Replace null" = Table.ReplaceValue(#"Replace "".  .""","",null,Replacer.ReplaceValue,{"Service Level", "Order Code", "Reference", "Payment Method", "Payment Currency", "Settlement Currency", "Journal Description"}),
    #"Uppercased Text" = Table.TransformColumns(#"Replace null",{{"Journal Description", Text.Upper}}),
    #"Add [Transfer Date]" = Table.AddColumn(#"Uppercased Text", "Transfer Date", each try (Date.FromText(Text.BetweenDelimiters([Journal Description], " ", " ", 0, 0))) otherwise null, type date),
    #"Left Join dMerchantAccounts" = Table.NestedJoin(#"Add [Transfer Date]",{"Merchant"},dMerchantAccounts_PSP,{"merchant_code"},"dMerchantAccounts",JoinKind.LeftOuter),
    #"Expanded dMerchantAccounts" = Table.ExpandTableColumn(#"Left Join dMerchantAccounts", "dMerchantAccounts", {"company_code"}, {"company_code"}),
    #"Sorted Rows" = Table.Sort(#"Expanded dMerchantAccounts",{{"company_code", Order.Ascending}, {"Batch ID", Order.Ascending}, {"Event Type", Order.Ascending}}),
    #"Filled Up" = Table.FillUp(#"Sorted Rows",{"Transfer Date"}),
    #"Filled Down" = Table.FillDown(#"Filled Up",{"Payment Currency"}),
    #"Replace ""Z""" = Table.ReplaceValue(#"Filled Down","Z_","",Replacer.ReplaceText,{"Event Type"}),
    #"Renamed Columns" = Table.RenameColumns(#"Replace ""Z""", {{"Merchant", "merchant_code"}, {"Batch ID", "batch_id"}, {"Service Level", "service_level"}, {"Event Type", "event_type"}, {"Date", "date"}, {"Order Code", "order_code"}, {"Reference", "reference"}, {"Payment Method", "payment_method"}, {"Payment Currency", "payment_currency"}, {"Payment Amount", "payment_amount"}, {"Settlement Currency", "settlement_currency"}, {"Commission Amount", "commission_amount"}, {"Guarantee Amount", "guarantee_amount"}, {"Net Amount", "net_amount"}, {"Journal Description", "journal_description"}, {"Transfer Date", "transfer_date"}}),
    #"Reordered Columns" = Table.ReorderColumns(#"Renamed Columns",{"company_code", "merchant_code", "batch_id", "service_level", "event_type", "date", "order_code", "reference", "payment_method", "payment_currency", "payment_amount", "settlement_currency", "commission_amount", "guarantee_amount", "net_amount", "journal_description", "transfer_date"})
in
    #"Reordered Columns"