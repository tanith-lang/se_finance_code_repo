let
    import_currency = Excel.CurrentWorkbook(){[Name="prmCurrency"]}[Content]{0}[Column1],

    Source = iTransSummary,
    #"Renamed Columns" = Table.RenameColumns(Source,{{"Administration Code", "administration_code"}, {"Merchant Code", "merchant_code"}, {"OrderCode", "order_code"}, {"Date", "date"}, {"PaymentMethod", "payment_method"}, {"Status", "status"}, {"CurrencyCode", "currency"}, {"Amount", "amount"}, {"Commission", "commission"}, {"Batch ID", "batch_id"}, {"Refusal Reason", "refusal_reason"}}),
    #"Removed Columns" = Table.RemoveColumns(#"Renamed Columns",{"refusal_reason"}),
    #"Add [cash_impact]" = Table.AddColumn(#"Removed Columns", "cash_impact", each null, type date),
    #"Filter [status]" = Table.SelectRows(#"Add [cash_impact]", each ([status] = "AUTHORISED" or [status] = "CANCELLED"))
in
    #"Filter [status]"