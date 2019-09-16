let
    Source = i_ratepay_transactions,
    #"Renamed Columns" = Table.RenameColumns(Source,{{"shop_name", "merchant_code"}, {"shop_order_id", "order_code"}, {"description", "status"}, {"order_date", "date"}}),
    #"Removed Columns" = Table.RemoveColumns(#"Renamed Columns",{"shop_id", "payment_date", "payment_month", "descriptor", "shop_invoice_id", "invoice_number", "fee_type", "sent_date", "transaction_id"}),
    #"Filter [status]" = Table.SelectRows(#"Removed Columns", each ([status] = "Factoring-Kaufpreis FmV")),
    #"Replace ""Factoring-Kaufpreis FmV""" = Table.ReplaceValue(#"Filter [status]","Factoring-Kaufpreis FmV","AUTHORISED",Replacer.ReplaceText,{"status"}),
    #"Uppercase [merchant_code]" = Table.TransformColumns(#"Replace ""Factoring-Kaufpreis FmV""",{{"merchant_code", Text.Upper, type text}}),
    #"Multiplied Column" = Table.TransformColumns(#"Uppercase [merchant_code]", {{"amount", each _ * -1, type number}}),
    #"Add [datetime]" = Table.AddColumn(#"Multiplied Column", "datetime", each DateTime.From([date]), type datetime),
    #"Add [payment_method]" = Table.AddColumn(#"Add [datetime]", "payment_method", each "RATEPAY", type text),
    #"Add [moto_check]" = Table.AddColumn(#"Add [payment_method]", "moto_check", each false, type logical),
    #"Add [psp]" = Table.AddColumn(#"Add [moto_check]", "psp", each "RatePay", type text),
    #"Left Join d_merchants" = Table.NestedJoin(#"Add [psp]",{"merchant_code"},d_merchants,{"merchant_code"},"d_merchants",JoinKind.LeftOuter),
    #"Expanded d_merchants" = Table.ExpandTableColumn(#"Left Join d_merchants", "d_merchants", {"trading_currency"}, {"currency"}),
    #"Left Join d_trading_currencies" = Table.NestedJoin(#"Expanded d_merchants",{"currency"},d_trading_currencies,{"trading_currency"},"d_trading_currencies",JoinKind.LeftOuter),
    #"Expanded d_trading_currencies" = Table.ExpandTableColumn(#"Left Join d_trading_currencies", "d_trading_currencies", {"trading_currency_Id"}, {"trading_currency_Id"}),
    #"Reordered Columns" = Table.ReorderColumns(#"Expanded d_trading_currencies",{"merchant_code", "order_code", "date", "datetime", "payment_method", "status", "currency", "trading_currency_Id", "moto_check", "amount"})
in
    #"Reordered Columns"