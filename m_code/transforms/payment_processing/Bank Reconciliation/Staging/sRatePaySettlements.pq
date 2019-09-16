let
    Source = iRatePayTransactions,
    #"Multiply [AMOUNT]" = Table.TransformColumns(Source, {{"AMOUNT", each _ * -1, Currency.Type}}),
    #"Grouped Rows" = Table.Group(#"Multiply [AMOUNT]", {"SHOP_ID", "PAYMENTDATE"}, {{"SETTLEMENT_AMOUNT", each List.Sum([AMOUNT]), Currency.Type}})
in
    #"Grouped Rows"