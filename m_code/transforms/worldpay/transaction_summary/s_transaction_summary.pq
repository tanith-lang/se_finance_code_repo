let
    Source = i_worldpay_transactions,
    #"Replace Text Whitespace" = Table.ReplaceValue(Source,"",null,Replacer.ReplaceValue,{"admin_code", "merchant_code", "order_code", "payment_method", "status", "currency", "refusal_reason"}),
    #"Replace Number Nulls" = Table.ReplaceValue(#"Replace Text Whitespace",null,0,Replacer.ReplaceValue,{"amount", "commission"}),
    #"Filter [status]" = Table.SelectRows(#"Replace Number Nulls", each ([status] = "AUTHORISED" or [status] = "CANCELLED")),
    #"Add [date]" = Table.AddColumn(#"Filter [status]", "date", each DateTime.Date([datetime]), type date),
    #"Add [moto_check]" = Table.AddColumn(#"Add [date]", "moto_check", each if Text.StartsWith([order_code], "C1") then true else false, type logical),
    #"Add [psp]" = Table.AddColumn(#"Add [moto_check]", "psp", each "WorldPay", type text),
    #"Left Join d_trading_currencies" = Table.NestedJoin(#"Add [psp]",{"currency"},d_trading_currencies,{"trading_currency"},"d_trading_currencies",JoinKind.LeftOuter),
    #"Expanded d_trading_currencies" = Table.ExpandTableColumn(#"Left Join d_trading_currencies", "d_trading_currencies", {"trading_currency_Id"}, {"trading_currency_Id"}),
    #"Removed Columns" = Table.RemoveColumns(#"Expanded d_trading_currencies",{"admin_code", "commission", "batch_id", "refusal_reason"}),
    #"Reordered Columns" = Table.ReorderColumns(#"Removed Columns",{"merchant_code", "order_code", "date", "datetime", "payment_method", "status", "currency", "trading_currency_Id", "moto_check", "amount"})
in
    #"Reordered Columns"