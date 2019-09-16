let
    import_currency = Excel.CurrentWorkbook(){[Name="import_currency"]}[Content]{0}[Column1],

    Source = iNBT339_EMIS,
    #"Filter [record_type]" = Table.SelectRows(Source, each ([record_type] = "01")),
    #"Parse [trading_date]" = Table.TransformColumns(#"Filter [record_type]",{{"trading_date", each Date.From(DateTimeZone.From(_)), type date}}),
    #"Rename [company_Id]" = Table.RenameColumns(#"Parse [trading_date]",{{"company_Id", "company_Id_temp"}}),
    #"Add [company_Id]" = Table.AddColumn(#"Rename [company_Id]", "company_Id", each if Text.Length([company_Id_temp]) <> 6 then "0" & [company_Id_temp] else [company_Id_temp]),
    #"Rename [process_date]" = Table.RenameColumns(#"Add [company_Id]",{{"processed_date", "processed_date_temp"}}),
    #"Add [process_date]" = Table.AddColumn(#"Rename [process_date]", "processed_date", each Text.Start([processed_date_temp], 4) & "-"
       & Text.Range([processed_date_temp], 4, 2) & "-"
       & Text.Range([processed_date_temp], 6, 2)),
    #"< Set Data Types >" = Table.TransformColumnTypes(#"Add [process_date]",{{"company_Id", type text}, {"processed_date", type date}, {"merchant_Id", type text}, {"merchant_name", type text}, {"exchange_rate", type number}, {"reporting_rate", type number}, {"card_type", type text}, {"no_of_sales", Int64.Type}, {"value_of_sales", Currency.Type}, {"no_of_refunds", Int64.Type}, {"value_of_refunds", Currency.Type}, {"total_transactions", Int64.Type}, {"net_value", Currency.Type}, {"transaction_currency_code", type text}, {"no_of_transactions", Int64.Type}, {"settlement_value", Currency.Type}, {"settlement_currency_code", type text}}),
    #"Divide [total_sales]" = Table.TransformColumns(#"< Set Data Types >", {{"value_of_sales", each _ / 100, Currency.Type}}),
    #"Divide [value_refunds]" = Table.TransformColumns(#"Divide [total_sales]", {{"value_of_refunds", each _ / 100, Currency.Type}}),
    #"Divide [net_value]" = Table.TransformColumns(#"Divide [value_refunds]", {{"net_value", each _ / 100, Currency.Type}}),
    #"Divide [settlement_value]" = Table.TransformColumns(#"Divide [net_value]", {{"settlement_value", each _ / 100, Currency.Type}}),
    #"Filter [settlement_currency_code]" = Table.SelectRows(#"Divide [settlement_value]", each [settlement_currency_code] = import_currency),
    #"Left Join Trading Totals" = Table.NestedJoin(#"Filter [settlement_currency_code]",{"merchant_Id", "trading_date", "processed_date"},sNBT339_EMIS_Trading_Totals,{"merchant_Id", "trading_date", "processed_date"},"sNBT339_EMIS_Totals",JoinKind.LeftOuter),
    #"Expanded sNBT339_EMIS_Totals" = Table.ExpandTableColumn(#"Left Join Trading Totals", "sNBT339_EMIS_Totals", {"trading_date_total_lcy", "trading_date_total"}, {"trading_date_total_lcy", "trading_date_total"}),
    #"Left Join Processed Totals" = Table.NestedJoin(#"Expanded sNBT339_EMIS_Totals",{"company_Id", "processed_date"},sNBT339_EMIS_Processed_Totals,{"company_Id", "processed_date"},"sNBT339_EMIS_Processed_Totals",JoinKind.LeftOuter),
    #"Expanded Processed Totals" = Table.ExpandTableColumn(#"Left Join Processed Totals", "sNBT339_EMIS_Processed_Totals", {"processed_date_total_lcy", "processed_date_total", "processed_exchange_rate"}, {"processed_date_total_lcy", "processed_date_total", "processed_exchange_rate"}),
    #"Left Join dMerchantAccounts" = Table.NestedJoin(#"Expanded Processed Totals",{"merchant_Id"},dMerchantAccounts_AQ_Batch,{"out_id"},"dMerchantAccounts_AQ_Batch",JoinKind.LeftOuter),
    #"Expanded dMerchantAccounts_AQ_Batch" = Table.ExpandTableColumn(#"Left Join dMerchantAccounts", "dMerchantAccounts_AQ_Batch", {"merchant_code"}, {"merchant_code"}),
    #"Removed Columns" = Table.RemoveColumns(#"Expanded dMerchantAccounts_AQ_Batch",{"record_type", "company_name", "merchant_name", "company_Id_temp", "processed_date_temp", "reporting_rate", "settlement_currency", "reporting_currency", "transaction_currency", "no_of_settlements", "reporting_value", "reporting_currency_code", "no_of_transactions", "merchant_Id"}),
    #"Reordered Columns" = Table.ReorderColumns(#"Removed Columns",{"company_Id", "trading_date", "processed_date", "merchant_code", "exchange_rate", "card_type", "no_of_sales", "value_of_sales", "no_of_refunds", "value_of_refunds", "total_transactions", "net_value", "transaction_currency_code", "settlement_value", "settlement_currency_code", "trading_date_total_lcy", "processed_date_total_lcy", "trading_date_total", "processed_date_total", "processed_exchange_rate"}),
    #"Renamed Columns" = Table.RenameColumns(#"Reordered Columns",{{"transaction_currency_code", "trading_currency"}, {"settlement_currency_code", "settlement_currency"}, {"company_Id", "company_code"}, {"total_transactions", "transaction_count"}, {"net_value", "transaction_value_lcy"}}),
    #"Sorted Rows" = Table.Sort(#"Renamed Columns",{{"company_code", Order.Ascending}, {"processed_date", Order.Ascending}, {"trading_date", Order.Ascending}, {"merchant_code", Order.Ascending}})
in
    #"Sorted Rows"