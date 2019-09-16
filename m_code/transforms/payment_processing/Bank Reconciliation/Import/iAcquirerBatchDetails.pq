let
    import_currency = Excel.CurrentWorkbook(){[Name="import_currency"]}[Content]{0}[Column1],
    import_path = Excel.CurrentWorkbook(){[Name="import_path"]}[Content]{0}[Column1],
    import_month = Excel.CurrentWorkbook(){[Name="import_month"]}[Content]{0}[Column1],

    Source = Folder.Files(import_path),
    #"Filter [Folder Path]" = Table.SelectRows(Source, each Text.Contains([Folder Path], "WORLDPAY") and Text.Contains([Folder Path], "ACQRR_BTCH_PRDS") and Text.Contains([Folder Path], "DTLS")),
    #"Filter [Extension]" = Table.SelectRows(#"Filter [Folder Path]", each [Extension] = ".csv"),
    #"Filter [Name]" = Table.SelectRows(#"Filter [Extension]", each Text.Contains([Name], "ACQRR_DTLS") and Text.Contains([Name], "2018") and Text.Contains([Name], "-" & import_month)),
    #"Removed Other Columns" = Table.SelectColumns(#"Filter [Name]",{"Name", "Folder Path"}),
    #"Invoke <getAcquirerDetails>" = Table.AddColumn(#"Removed Other Columns", "getAcquirerDetails", each getAcquirerDetails( [Folder Path], [Name] )),
    #"Expand <getAcquirerDetails>" = Table.ExpandTableColumn(#"Invoke <getAcquirerDetails>", "getAcquirerDetails", {"Batch Id", "Account", "Event Type", "Order Code", "BackOffice Code", "Payment Method", "Date", "Currency", "Amount"}, {"Batch Id", "Account", "Event Type", "Order Code", "BackOffice Code", "Payment Method", "Date", "Currency", "Amount"}),
    #"Removed Columns" = Table.RemoveColumns(#"Expand <getAcquirerDetails>",{"Name", "Folder Path"}),
    #"< Set Data Types >" = Table.TransformColumnTypes(#"Removed Columns",{{"Batch Id", type text}, {"Account", type text}, {"Event Type", type text}, {"Order Code", type text}, {"BackOffice Code", type text}, {"Payment Method", type text}, {"Currency", type text}, {"Date", type date}, {"Amount", Currency.Type}}),
    #"Filter [Currency]" = Table.SelectRows(#"< Set Data Types >", each [Currency] = import_currency)
in
    #"Filter [Currency]"