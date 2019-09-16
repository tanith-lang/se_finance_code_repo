// iNBT339_EMIS
let
    folder_path = Excel.CurrentWorkbook(){[Name="folder_path"]}[Content]{0}[Column1],
    import_currency = Excel.CurrentWorkbook(){[Name="import_currency"]}[Content]{0}[Column1],
    import_month = Excel.CurrentWorkbook(){[Name="import_month"]}[Content]{0}[Column1],

    Source = Folder.Files(folder_path),
    #"Filter [Folder Path]" = Table.SelectRows(Source, each Text.Contains([Folder Path], "WORLDPAY") and Text.Contains([Folder Path], "NBT339")),
    #"Filter [Extension]" = Table.SelectRows(#"Filter [Folder Path]", each ([Extension] = ".CSV")),
    #"Filter [Name]" = Table.SelectRows(#"Filter [Extension]", each Text.Contains([Name], "SECX339")),
    #"Invoke Custom Function" = Table.AddColumn(#"Filter [Name]", "getNBT339_EMIS", each getNBT339_EMIS([Folder Path], [Name])),
    #"Removed Other Columns" = Table.SelectColumns(#"Invoke Custom Function",{"getNBT339_EMIS"}),
    #"Expanded getNBT339" = Table.ExpandTableColumn(#"Removed Other Columns", "getNBT339_EMIS", {"record_type", "trading_date", "company_Id", "company_name", "merchant_Id", "merchant_name", "settlement_currency", "reporting_currency", "transaction_currency", "exchange_rate", "reporting_rate", "card_type", "no_of_sales", "value_of_sales", "no_of_refunds", "value_of_refunds", "total_transactions", "net_value", "transaction_currency_code", "no_of_transactions", "settlement_value", "settlement_currency_code", "no_of_settlements", "reporting_value", "reporting_currency_code", "processed_date"}, {"record_type", "trading_date", "company_Id", "company_name", "merchant_Id", "merchant_name", "settlement_currency", "reporting_currency", "transaction_currency", "exchange_rate", "reporting_rate", "card_type", "no_of_sales", "value_of_sales", "no_of_refunds", "value_of_refunds", "total_transactions", "net_value", "transaction_currency_code", "no_of_transactions", "settlement_value", "settlement_currency_code", "no_of_settlements", "reporting_value", "reporting_currency_code", "processed_date"})
in
    #"Expanded getNBT339"