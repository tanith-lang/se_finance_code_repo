let
    import_path = Excel.CurrentWorkbook(){[Name="import_path"]}[Content]{0}[Column1],
    import_month = Excel.CurrentWorkbook(){[Name="import_month"]}[Content]{0}[Column1],

    Source = Folder.Files(import_path),
    #"Filter [Folder Path]" = Table.SelectRows(Source, each Text.Contains([Folder Path], "WORLDPAY") and Text.Contains([Folder Path], "TRNSFRS_BY_PRD")),
    #"Filter [Extension]" = Table.SelectRows(#"Filter [Folder Path]", each [Extension] = ".csv"),
    #"Filter [Name]" = Table.SelectRows(#"Filter [Extension]", each Text.Contains([Name], "Transfers-by-Period") and Text.Contains([Name], "2018")),
    #"Add [Import Month]" = Table.AddColumn(#"Filter [Name]", "Import Month", each Text.BetweenDelimiters([Name], "-", "_", 0, 0), type text),
    #"Filter [Import Month]" = Table.SelectRows(#"Add [Import Month]", each [Import Month] = import_month),
    #"Invoke Custom Function" = Table.AddColumn(#"Filter [Import Month]", "read_csv", each read_csv([Folder Path], [Name], ",", 15)),
    #"Remove Source Columns" = Table.RemoveColumns(#"Invoke Custom Function",{"Content", "Name", "Extension", "Date accessed", "Date modified", "Date created", "Attributes", "Folder Path", "Import Month"}),
    #"Expand read_csv" = Table.ExpandTableColumn(#"Remove Source Columns", "read_csv", {"Merchant", "Batch ID", "Service Level", "Event Type", "Date", "Order Code", "Reference", "Payment Method", "Payment Currency", "Payment Amount", "Settlement Currency", "Commission Amount", "Guarantee Amount", "Net Amount", "Journal Description"}, {"Merchant", "Batch ID", "Service Level", "Event Type", "Date", "Order Code", "Reference", "Payment Method", "Payment Currency", "Payment Amount", "Settlement Currency", "Commission Amount", "Guarantee Amount", "Net Amount", "Journal Description"}),
    #"< Set Data Types >" = Table.TransformColumnTypes(#"Expand read_csv",{{"Merchant", type text}, {"Batch ID", type text}, {"Service Level", type text}, {"Event Type", type text}, {"Order Code", type text}, {"Reference", type text}, {"Payment Method", type text}, {"Payment Currency", type text}, {"Date", type date}, {"Payment Amount", Currency.Type}, {"Settlement Currency", type text}, {"Commission Amount", Currency.Type}, {"Guarantee Amount", Currency.Type}, {"Net Amount", Currency.Type}, {"Journal Description", type text}})
in
    #"< Set Data Types >"