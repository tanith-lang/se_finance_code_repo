let
    folder_path = Excel.CurrentWorkbook(){[Name="folder_path"]}[Content]{0}[Column1],
    import_currency = Excel.CurrentWorkbook(){[Name="import_currency"]}[Content]{0}[Column1],
    import_month = Excel.CurrentWorkbook(){[Name="import_month"]}[Content]{0}[Column1],

    Source = Folder.Files(folder_path),
    #"Filter [Folder Path]" = Table.SelectRows(Source, each Text.Contains([Folder Path], "WORLDPAY") and Text.Contains([Folder Path], "ACQRR_BTCH_PRDS") and Text.Contains([Folder Path], "SMMRY_BY_BTCH")),
    #"Filter [Extension]" = Table.SelectRows(#"Filter [Folder Path]", each [Extension] = ".csv"),
    #"Filter [Name]" = Table.SelectRows(#"Filter [Extension]", each Text.Contains([Name], "ACQRR_SMMRY") and Text.Contains([Name], "2018") and Text.Contains([Name], "-" & import_month)),
    #"Removed Other Columns" = Table.SelectColumns(#"Filter [Name]",{"Name", "Folder Path"}),
    #"Add <getAcquirerBatches>" = Table.AddColumn(#"Removed Other Columns", "getAcquirerBatches", each getAcquirerBatches( [Folder Path], [Name] )),
    #"Expand <getAcquirerBatches>" = Table.ExpandTableColumn(#"Add <getAcquirerBatches>", "getAcquirerBatches", {"Merchant", "Batch Id", "Period Begin Date", "Period End Date", "#Transactions", "Currency", "Amount"}, {"Merchant", "Batch Id", "Period Begin Date", "Period End Date", "#Transactions", "Currency", "Amount"}),
    #"Remove Source Columns" = Table.RemoveColumns(#"Expand <getAcquirerBatches>",{"Name", "Folder Path"}),
    #"< Set Data Types >" = Table.TransformColumnTypes(#"Remove Source Columns",{{"Merchant", type text}, {"Batch Id", type text}, {"Currency", type text}, {"Period Begin Date", type datetime}, {"Period End Date", type datetime}, {"#Transactions", Int64.Type}, {"Amount", Currency.Type}}),
    #"Filter [Currency]" = Table.SelectRows(#"< Set Data Types >", each [Currency] = import_currency)
in
    #"Filter [Currency]"