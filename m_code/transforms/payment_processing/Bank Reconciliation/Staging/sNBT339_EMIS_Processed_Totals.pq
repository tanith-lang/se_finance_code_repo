// sNBT339_EMIS_Processed_Totals
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
    #"Removed Columns" = Table.RemoveColumns(#"Add [process_date]",{"record_type", "company_name", "company_Id_temp", "processed_date_temp"}),
    #"< Set Data Types >" = Table.TransformColumnTypes(#"Removed Columns",{{"company_Id", type text}, {"processed_date", type date}, {"merchant_Id", type text}, {"merchant_name", type text}, {"settlement_currency", type text}, {"reporting_currency", type text}, {"transaction_currency", type text}, {"exchange_rate", type number}, {"reporting_rate", type number}, {"card_type", type text}, {"no_of_sales", Int64.Type}, {"value_of_sales", Currency.Type}, {"no_of_refunds", Int64.Type}, {"value_of_refunds", Currency.Type}, {"total_transactions", Int64.Type}, {"net_value", Currency.Type}, {"transaction_currency_code", type text}, {"no_of_transactions", Int64.Type}, {"settlement_value", Currency.Type}, {"settlement_currency_code", type text}, {"no_of_settlements", Int64.Type}, {"reporting_value", Currency.Type}, {"reporting_currency_code", type text}}),
    #"Divide [total_sales]" = Table.TransformColumns(#"< Set Data Types >", {{"value_of_sales", each _ / 100, Currency.Type}}),
    #"Divide [value_refunds]" = Table.TransformColumns(#"Divide [total_sales]", {{"value_of_refunds", each _ / 100, Currency.Type}}),
    #"Divide [net_value]" = Table.TransformColumns(#"Divide [value_refunds]", {{"net_value", each _ / 100, Currency.Type}}),
    #"Divide [settlement_value]" = Table.TransformColumns(#"Divide [net_value]", {{"settlement_value", each _ / 100, Currency.Type}}),
    #"Divide [reporting_value]" = Table.TransformColumns(#"Divide [settlement_value]", {{"reporting_value", each _ / 100, Currency.Type}}),
    #"Filter [settlement_currency_code]" = Table.SelectRows(#"Divide [reporting_value]", each [settlement_currency_code] = import_currency),
    // Rows grouped by [company_id] and [processed_date] to generate daily batch totals.
    #"Group Rows" = Table.Group(#"Filter [settlement_currency_code]", {"company_Id", "processed_date"}, {{"processed_date_total_lcy", each List.Sum([net_value]), Currency.Type}, {"processed_date_total", each List.Sum([settlement_value]), Currency.Type}}),
    #"Add [processed_exchange_rate]" = Table.AddColumn(#"Group Rows", "processed_exchange_rate", each [processed_date_total_lcy] / [processed_date_total], type number),
    #"Sorted Rows" = Table.Sort(#"Add [processed_exchange_rate]",{{"company_Id", Order.Ascending}, {"processed_date", Order.Ascending}})
in
    #"Sorted Rows"