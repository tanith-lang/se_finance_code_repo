let
    Source = iAcquirerBatchDetails,
    #"Renamed Columns" = Table.RenameColumns(Source,{{"Event Type", "Status"}, {"Account", "Bank Account"}}),
    #"Remove [BackOffice Code]" = Table.RemoveColumns(#"Renamed Columns",{"BackOffice Code"}),
    #"Filter [Payment Method]" = Table.SelectRows(#"Remove [BackOffice Code]", each ([Payment Method] <> "AMEX-SSL" and [Payment Method] <> "PAYPAL-EXPRESS")),
    #"Left Join sMerchants" = Table.NestedJoin(#"Filter [Payment Method]",{"Bank Account"},dMerchantAccounts_AQ_Batch,{"bank_account"},"sMerchants",JoinKind.LeftOuter),
    #"Expanded sMerchants" = Table.ExpandTableColumn(#"Left Join sMerchants", "sMerchants", {"company_id", "merchant_code"}, {"Merchant ID", "Merchant Code"})
in
    #"Expanded sMerchants"