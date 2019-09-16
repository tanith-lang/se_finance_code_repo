let
    Source = iAcquirerBatchSummaries,
    #"Renamed Columns" = Table.RenameColumns(Source,{{"Merchant", "Merchant Code"}, {"Amount", "Batch Total"}}),
    #"Left Join sMerchants" = Table.NestedJoin(#"Renamed Columns",{"Merchant Code"},dMerchantAccounts_AQ_Batch,{"merchant_code"},"sMerchants",JoinKind.LeftOuter),
    #"Expanded sMerchants" = Table.ExpandTableColumn(#"Left Join sMerchants", "sMerchants", {"company_id", "bank_account"}, {"Merchant ID", "Bank Account"})
in
    #"Expanded sMerchants"